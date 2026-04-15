VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form P_ExpPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Planificación Minutas"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2775
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      Begin VB.Frame Frame1 
         Caption         =   "Regimen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   1320
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   1
            Left            =   2205
            TabIndex        =   12
            Top             =   360
            Width           =   795
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   3120
            Picture         =   "P_EIPlan.frx":0000
            Top             =   195
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Servicios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Index           =   2
         Left            =   4200
         TabIndex        =   8
         Top             =   1320
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   3
            Left            =   2205
            TabIndex        =   10
            Top             =   360
            Width           =   795
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   2
            Left            =   330
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   3120
            Picture         =   "P_EIPlan.frx":030A
            Top             =   195
            Width           =   480
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   915
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
         Text            =   "07/2007"
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   2
         Top             =   210
         Width           =   1335
         _Version        =   196608
         _ExtentX        =   2364
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1260
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   540
         Width           =   5895
         _Version        =   196608
         _ExtentX        =   10398
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
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   2
         AutoAdvance     =   0   'False
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
      Begin MSComctlLib.ProgressBar PB 
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   2325
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Exp."
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
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Casino"
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
         Index           =   9
         Left            =   120
         TabIndex        =   4
         Top             =   285
         Width           =   585
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3060
         TabIndex        =   3
         Top             =   210
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2550
         Picture         =   "P_EIPlan.frx":0614
         Top             =   120
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3105
         TabIndex        =   6
         Top             =   255
         Width           =   4095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3030
      Left            =   8400
      TabIndex        =   7
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   5345
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   15
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
      _Version        =   393216
      _ExtentX        =   3201
      _ExtentY        =   661
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "P_EIPlan.frx":091E
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   375
      Index           =   1
      Left            =   0
      TabIndex        =   16
      Top             =   2280
      Visible         =   0   'False
      Width           =   1815
      _Version        =   393216
      _ExtentX        =   3201
      _ExtentY        =   661
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
      MaxCols         =   3
      MaxRows         =   1
      SpreadDesigner  =   "P_EIPlan.frx":0BA5
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "P_ExpPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 3510
Me.Width = 9030
Msgtitulo = "Exportar Planificación Minuta"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1), True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpText(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
MoverGrilla
End Sub

Sub MoverGrilla()
'------- Mover regimen
vaSpread1(0).MaxRows = 0
RS.Open "SELECT * FROM  a_regimen ORDER BY reg_codigo", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
      vaSpread1(0).Row = vaSpread1(0).MaxRows
      vaSpread1(0).Col = 1: vaSpread1(0).text = "1"
      vaSpread1(0).Col = 2: vaSpread1(0).text = RS!reg_codigo
      vaSpread1(0).Col = 3: vaSpread1(0).text = Trim(RS!reg_nombre)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing

'------- Mover servicio
vaSpread1(1).MaxRows = 0
RS.Open "SELECT * FROM  a_servicio ORDER BY ser_codigo", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
      vaSpread1(1).Row = vaSpread1(1).MaxRows
      vaSpread1(1).Col = 1: vaSpread1(1).text = "1"
      vaSpread1(1).Col = 2: vaSpread1(1).text = RS!ser_codigo
      vaSpread1(1).Col = 3: vaSpread1(1).text = Trim(RS!ser_nombre)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_Change(Index As Integer)
RS.Open "SELECT * FROM b_clientes WHERE cli_codigo='" & fpText(0).text & "' AND cli_tipo=0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos (*.mdb)|*.mdb"
CD.DefaultExt = "*.mdb"
CD.ShowSave
If CD.Filename = "" Then fpText1.text = "" Else fpText1.text = CD.Filename 'Dir(CD.Filename)
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Casino"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText(0).text = Trim(vg_codigo)
    fpayuda(0).Caption = vg_nombre
Case 1
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText(0).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Regimen", Me.vaSpread1, fpText(0).text, "", 0, 0, "0", "FacCli", 0, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
Case 2
   vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If fpText(0).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, fpText(0).text, "", 0, 0, "0", "FacCli", 1, "'1'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Dim inddex As Integer
inddex = IIf(Index = 0, 0, 1)
Select Case Index
Case 0, 2
    For i = 1 To vaSpread1(inddex).MaxRows
        vaSpread1(inddex).Row = i
        vaSpread1(inddex).Col = 1: vaSpread1(inddex).text = "1"
    Next i
    If Index = 0 Then Image1(1).Enabled = False Else Image1(2).Enabled = False
Case 1
   Image1(1).Enabled = True
Case 3
   Image1(2).Enabled = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codreg As String, codser As String
Select Case Button.Index
Case 1
    '------- Validar si existe archivo dbgt
    If Dir(CD.Filename) = BaseDeDato Then MsgBox "Base de dato no puede ser la misma del sistema, cambie de nombre", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar ruta
    If Trim(fpText1.text) = "" Then fg_descarga: MsgBox "Carpeta no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar cencos
    If Trim(fpayuda(0).Caption) = "" Then fg_descarga: MsgBox "Casino debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    codreg = "": codser = ""
    '------- Validar regimen
    For i = 1 To vaSpread1(0).MaxRows
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then vaSpread1(0).Col = 2: codreg = codreg & "" & vaSpread1(0).text & ","
    Next i
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar servicio
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then vaSpread1(1).Col = 2: codser = codser & "" & vaSpread1(1).text & ","
    Next i
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar planificaciòn de minutas
    RS.Open "SELECT DISTINCT b.min_cencos " & _
            "FROM b_minuta b, b_minutadet c WHERE b.min_codigo=c.mid_codigo AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND  b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "Planificaciòn no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga ""
    Frame1(0).Enabled = False: Toolbar1.Enabled = False
    PB.Visible = True: PB.Min = 0: PB.Value = 0: PB.Max = 24
    Dim cDBO As String
    '------- Crear directorio para generar planificación
'    If Dir(CD.Filename, vbVolume) = "" Then MkDir CD.Filename
    '------- Generar base padre
    cDBO = dir_trabajo & BaseDeDato
    If Dir(CD.Filename) <> "" Then Kill CD.Filename 'borrar base datos si existe
    '------- generar archivo mdb
    Set db1 = DBEngine(0).CreateDatabase(CD.Filename, dbLangGeneral)
    '------- tabla relacionada a productos
    db1.Execute "CREATE TABLE a_tipopro (tip_codigo int, tip_nombre char(35), tip_previo int)"
    db1.Execute "CREATE TABLE a_unidad (uni_codigo int, uni_nombre char(10), uni_nomcor char(5))"
    db1.Execute "CREATE TABLE a_embalaje (emb_codigo int, emb_nombre char(20), emb_nomcor char(5))"
    db1.Execute "CREATE TABLE a_ctacontable (cta_codigo char(10), cta_nombre char(40))"
    db1.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255))"
    db1.Execute "CREATE TABLE a_impuesto (imp_codigo int, imp_nombre char(15), imp_pctimp double, imp_inccos int, imp_codsap char(20), imp_indmod char(1))"
    db1.Execute "CREATE TABLE a_unidadmed (unm_codigo int, unm_nombre char(10), unm_nomcor char(5))"
    db1.Execute "CREATE TABLE a_nutriente (nut_codigo int, nut_nombre char(30), nut_nomuni char(5), nut_indpri int, nut_secnro int)"
    db1.Execute "CREATE TABLE b_productos (pro_codigo char(20), pro_codbar char(20), pro_codcom char(20), pro_codtip int, pro_nombre char(50), pro_coduni int, pro_facing double, pro_facsto double, pro_codemb int, pro_uniemb double, pro_upreco double, pro_fecuco datetime, pro_propon double, pro_ctacon char(10), pro_fecven int, pro_ctrsto int)"
    db1.Execute "CREATE TABLE b_productosimp (ipr_codpro char(20), ipr_codimp int)"
    db1.Execute "CREATE TABLE b_productosing (pri_codpro char(20), pri_coding char(20))"
    db1.Execute "CREATE TABLE b_ingrediente (ing_codigo char(20), ing_nombre char(50), ing_nomfan char(50), ing_unimed int, ing_pctapr double, ing_pctcoc double, ing_pctnut double, ing_facnut double, ing_indpav int, ing_indgrv int, ing_precos double, ing_feccos int, ing_codcom char(20), ing_codped char(20))"
    db1.Execute "CREATE TABLE b_productonut (pnu_codpro char(20), pnu_codapo int, pnu_canapo double)"
    db1.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50))"
    '------- tabla relacionada a recetas
    db1.Execute "CREATE TABLE a_recetacatdie (car_codigo int, car_nombre char(50), car_previo int)"
    db1.Execute "CREATE TABLE a_recetatippla (tip_codigo int, tip_nombre char(50), tip_previo int)"
    db1.Execute "CREATE TABLE b_receta (rec_codigo int, rec_catdie int, rec_tippla int, rec_nombre char(80), rec_nomfan char(80), rec_metpre longtext, rec_conche longtext, rec_sugere longtext, rec_basrac int, rec_tiprec char(1), rec_fecvig int, rec_gruvul longtext)"
    db1.Execute "CREATE TABLE b_recetadet (red_codigo int, red_nroite int, red_codpro char(20), red_canpro double, red_cospro double, red_pctapr double, red_pctcoc double, red_pctnut double, red_tiprec int)"
    '------- tabla relacionada a planificación
    db1.Execute "CREATE TABLE b_costopatron (cpa_cencos char(10), cpa_codreg int, cpa_codser int, cpa_anomes int, cpa_descripcion char(10), cpa_valor double)"
    db1.Execute "CREATE TABLE b_gramofamproducto (gfp_cencos char(10), gfp_codreg int, gfp_catdie int, gfp_tiprec int, gfp_fampro int, gfp_graini double, gfp_grafin double)"
    db1.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
    db1.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(50), ser_orden int)"
    db1.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double)"
    db1.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int, Constraint b_minuta_pk Primary Key (min_codigo))"
    db1.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac int, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer int, mid_rec5eta char(1), Constraint b_minutadet_pk Primary Key (mid_codigo, mid_numlin))"
    PB.Value = PB.Value + 1
    '------- generar familia productos
    db1.Execute "INSERT INTO a_tipopro SELECT tip_codigo, tip_nombre, tip_previo FROM a_tipopro IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar unidad medida productos
    db1.Execute "INSERT INTO a_unidad SELECT DISTINCT uni_codigo, uni_nombre, uni_nomcor FROM a_unidad IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar embalaje productos
    db1.Execute "INSERT INTO a_embalaje SELECT DISTINCT emb_codigo, emb_nombre, emb_nomcor FROM a_embalaje IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar cuentas contables productos
    db1.Execute "INSERT INTO a_ctacontable SELECT DISTINCT cta_codigo, cta_nombre FROM a_ctacontable IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar parametros cuentas contables
    db1.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor FROM a_param IN '" & cDBO & "' WHERE par_codigo IN ('ctagastos','ctagastos2','ctainsumo','ctalimdes')"
    PB.Value = PB.Value + 1
    '------- generar impuesto productos
    db1.Execute "INSERT INTO a_impuesto SELECT imp_codigo, imp_nombre, imp_pctimp, imp_inccos, imp_codsap, imp_indmod FROM a_impuesto IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar unidad medida ingrediente
    db1.Execute "INSERT INTO a_unidadmed SELECT DISTINCT unm_codigo, unm_nombre, unm_nomcor FROM a_unidadmed IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar nutriente aporte
    db1.Execute "INSERT INTO a_nutriente SELECT DISTINCT nut_codigo, nut_nombre, nut_nomuni, nut_indpri, nut_secnro FROM a_nutriente IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar proveedores
    DoEvents
    db1.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro FROM b_proveedor IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- Generar productos
    DoEvents
    db1.Execute "INSERT INTO b_productos SELECT DISTINCT pro_codigo, pro_codbar, pro_codcom, pro_codtip, pro_nombre, pro_coduni, pro_facing, pro_facsto, pro_codemb, " & _
                "pro_uniemb, pro_upreco, pro_fecuco, pro_propon, pro_ctacon, pro_fecven, pro_ctrsto FROM b_productos IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar productos impuestos
    DoEvents
    db1.Execute "INSERT INTO b_productosimp SELECT DISTINCT ipr_codpro, ipr_codimp FROM b_productosimp IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar productos ingredientes & ingredientes
    DoEvents
    db1.Execute "INSERT INTO b_productosing SELECT DISTINCT pri_codpro, pri_coding FROM b_productosing IN '" & cDBO & "'"
    DoEvents: PB.Value = PB.Value + 1
    db1.Execute "INSERT INTO b_ingrediente SELECT DISTINCT ing_codigo , ing_nombre, ing_nomfan, ing_unimed, ing_pctapr, ing_pctcoc, ing_pctnut, ing_facnut, ing_indpav, " & _
                "ing_indgrv, ing_precos, ing_feccos, ing_codcom, ing_codped FROM b_ingrediente IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar nutriente del ingrediente
    db1.Execute "INSERT INTO b_productonut SELECT DISTINCT pnu_codpro, pnu_codapo, pnu_canapo FROM b_productonut IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar categoria dietetica
    db1.Execute "INSERT INTO a_recetacatdie SELECT DISTINCT car_codigo, car_nombre, car_previo FROM a_recetacatdie IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- generar tipo plato
    db1.Execute "INSERT INTO a_recetatippla SELECT DISTINCT tip_codigo, tip_nombre, tip_previo FROM a_recetatippla IN '" & cDBO & "'"
    PB.Value = PB.Value + 1
    '------- Generar encabezado receta
    db1.Execute "INSERT INTO b_receta SELECT DISTINCT a.rec_codigo, a.rec_catdie, a.rec_tippla, a.rec_nombre, a.rec_nomfan, a.rec_metpre, a.rec_conche, a.rec_sugere, a.rec_basrac, a.rec_tiprec " & _
                "FROM b_receta a, b_minuta b, b_minutadet c IN '" & cDBO & "' WHERE b.min_codigo=c.mid_codigo AND c.mid_codrec=a.rec_codigo AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND  b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'"
    DoEvents: PB.Value = PB.Value + 1
    '------- generar detalle recetas
    db1.Execute "INSERT INTO b_recetadet SELECT a.red_codigo, a.red_nroite, a.red_codpro, a.red_canpro, a.red_cospro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, red_tiprec " & _
                "FROM b_recetadet a, b_receta b, b_minuta c, b_minutadet d IN '" & cDBO & "' WHERE c.min_codigo=d.mid_codigo AND d.mid_codrec=b.rec_codigo AND b.rec_codigo=a.red_codigo " & _
                "AND  c.min_cencos='" & fpText(0).text & "' AND c.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND  c.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(c.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND d.mid_tipmin='1'"
    DoEvents: PB.Value = PB.Value + 1
    '------- generar regimen
    db1.Execute "INSERT INTO a_regimen SELECT DISTINCT a.reg_codigo, a.reg_nombre FROM a_regimen a, b_minuta b, b_minutadet c IN '" & cDBO & "' WHERE b.min_codigo=c.mid_codigo AND a.reg_codigo=b.min_codreg AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'"
    PB.Value = PB.Value + 1
    '------- generar servicio
    db1.Execute "INSERT INTO a_servicio SELECT DISTINCT a.ser_codigo, a.ser_nombre, a.ser_orden FROM a_servicio a, b_minuta b, b_minutadet c IN '" & cDBO & "' WHERE b.min_codigo=c.mid_codigo AND a.ser_codigo=b.min_codser AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'"
    PB.Value = PB.Value + 1
    '------- generar estructura servicio
    db1.Execute "INSERT INTO a_estservicio SELECT DISTINCT a.ess_codser, a.ess_codigo, a.ess_nombre, a.ess_orden, a.ess_codsec, a.ess_racmin FROM a_estservicio a, a_servicio b, b_minuta c, b_minutadet d IN '" & cDBO & "' WHERE c.min_codigo=d.mid_codigo AND b.ser_codigo=c.min_codser AND b.ser_codigo=a.ess_codser AND c.min_cencos='" & fpText(0).text & "' AND c.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND c.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(c.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND d.mid_tipmin='1'"
    PB.Value = PB.Value + 1
    '------- generar encabezado planificación minutas
    db1.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN '" & cDBO & "' WHERE a.min_codigo=b.mid_codigo AND a.min_cencos='" & fpText(0).text & "' AND a.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND a.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(a.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND b.mid_tipmin='1'"
    PB.Value = PB.Value + 1
    '------- generar detalle planificación minutas
    db1.Execute "INSERT INTO b_minutadet SELECT DISTINCT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec FROM b_minutadet a, b_minuta b IN '" & cDBO & "' WHERE b.min_codigo=a.mid_codigo AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND a.mid_tipmin='1'"
    PB.Value = PB.Value + 1
    '------- costo patron
'    db1.Execute "INSERT INTO b_costopatron (cpa_cencos, cpa_codreg, cpa_codser, cpa_anomes, cpa_descripcion, cpa_valor) SELECT pcp_cencos, pcp_codreg, pcp_codser, pcp_anomes, pcp_descripcion, pcp_valor FROM b_paramcostopatron IN " & DBO & " WHERE pcp_anomes=" & Fecha & ""
    'PB.Value = PB.Value + 1
    '------- gramo familia producto
'    db1.Execute "INSERT INTO b_gramofamproducto (gfp_cencos, gfp_codreg, gfp_catdie, gfp_tiprec, gfp_fampro, gfp_graini, gfp_grafin) SELECT gfp_subseg, gfp_codreg, gfp_catdie, gfp_tiprec, gfp_fampro, gfp_graini, gfp_grafin FROM b_gramofamproducto IN " & DBO & ""
    'PB.Value = PB.Value + 1
    db1.Close
    If Dir(Mid(CD.Filename, 1, Len(CD.Filename) - Len(Dir(CD.Filename))) & Mid(Dir(CD.Filename), 1, (Len(Dir(CD.Filename)) - 4)) & ".zip") <> "" Then Kill Mid(Dir(CD.Filename), 1, (Len(Dir(CD.Filename)) - 4)) & ".zip" 'borrar base datos si existe
    AZ1.CreateZip Mid(Dir(CD.Filename), 1, (Len(Dir(CD.Filename)) - 4)) & ".zip", "": AZ1.AddFile CD.Filename, "", True, "": AZ1.Close
    If Dir(CD.Filename) <> "" Then Kill CD.Filename 'borrar base datos si existe
    fg_descarga
    MsgBox "Proceso de Exportación Finalizado", vbInformation + vbOKOnly, Msgtitulo
    PB.Visible = False: Frame1(0).Enabled = True: Toolbar1.Enabled = True
Case 3
    Me.Hide
    Unload Me
End Select
End Sub
