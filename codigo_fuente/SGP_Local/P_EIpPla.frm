VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form P_EIpPla 
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
      TabIndex        =   5
      Top             =   120
      Width           =   8175
      Begin ACTIVEZIPLib.ActiveZip AZ1 
         Left            =   7320
         Top             =   2760
         _Version        =   65536
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
      End
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
         TabIndex        =   14
         Top             =   1320
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   0
            Left            =   330
            TabIndex        =   3
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   1
            Left            =   2205
            TabIndex        =   4
            Top             =   360
            Width           =   795
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   3120
            Picture         =   "P_EIpPla.frx":0000
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
         TabIndex        =   11
         Top             =   1320
         Width           =   3855
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   225
            Index           =   3
            Left            =   2205
            TabIndex        =   13
            Top             =   360
            Width           =   795
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   225
            Index           =   2
            Left            =   330
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   3120
            Picture         =   "P_EIpPla.frx":030A
            Top             =   195
            Width           =   480
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1620
         TabIndex        =   2
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
         Text            =   "11/2017"
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
         Left            =   1620
         TabIndex        =   0
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
         Left            =   1620
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   555
         Width           =   5940
         _Version        =   196608
         _ExtentX        =   10477
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
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   2100
         Visible         =   0   'False
         Width           =   600
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
         Left            =   480
         TabIndex        =   18
         Top             =   630
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
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   480
         TabIndex        =   7
         Top             =   285
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3420
         TabIndex        =   6
         Top             =   210
         Width           =   4095
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2910
         Picture         =   "P_EIpPla.frx":0614
         Top             =   120
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3465
         TabIndex        =   9
         Top             =   255
         Width           =   4095
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3030
      Left            =   8400
      TabIndex        =   10
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
      SpreadDesigner  =   "P_EIpPla.frx":091E
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
      SpreadDesigner  =   "P_EIpPla.frx":0C23
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "P_EIpPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim opcion As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 3510
Me.Width = 9030
'Msgtitulo = "Exportar Planificación Minuta"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1), True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
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
RS.Open "SELECT * FROM  a_servicio ORDER BY ser_orden, ser_codigo", vg_db, adOpenStatic
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
RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(0).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)
Cd.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
Cd.Filter = IIf(opcion = "E", "Todos los archivos (*.mdb)|*.mdb", "Todos los archivos (*.zip)|*.zip")
Cd.DefaultExt = IIf(opcion = "E", "*.mdb", "*.zip")
If opcion = "E" Then Cd.ShowSave Else Cd.ShowOpen
If Cd.Filename = "" Then fpText1.text = "" Else fpText1.text = Cd.Filename 'Dir(CD.Filename)
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
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
Dim codreg As String, codser As String, oError As Boolean, sql1 As String
Select Case Button.Index
Case 1
    '------- Validar si existe archivo dbgt
    If Dir(Cd.Filename) = BaseDeDato Then MsgBox "Base de dato no puede ser la misma del sistema, cambie de nombre", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar ruta
    If Trim(fpText1.text) = "" Then fg_descarga: MsgBox "Carpeta no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    '------- Validar cencos
    If Trim(fpayuda(0).Caption) = "" Then fg_descarga: MsgBox "Contrato debe ser informado", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If opcion = "E" Then
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
       sql1 = IIf(vg_tipbase = "1", " val(mid(b.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) ")
       RS.Open "SELECT DISTINCT b.min_cencos " & _
               "FROM b_minuta b, b_minutadet c WHERE b.min_codigo = c.mid_codigo AND b.min_cencos = '" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
               "AND  b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND " & sql1 & " = " & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin = '1'", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "Planificaciòn no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
       RS.Close: Set RS = Nothing
    End If
    If opcion = "E" Then
       oError = IIf(ExportarDatos(Trim(Dir(Cd.Filename)), codreg, codser), False, True)
    Else
       If vg_tipbase = "1" Then
          oError = IIf(ImportarDatosAccess(Trim(Dir(Cd.Filename))), False, True)
       Else
          oError = IIf(ImportarDatosSql(Trim(Dir(Cd.Filename))), False, True)
       End If
    End If
    If oError Then
       MsgBox IIf(opcion = "E", "Proceso de Exportar Falló", "Proceso de Importar Falló"), vbInformation + vbOKOnly, Msgtitulo
    Else
       If Not oError And Dir(Cd.Filename) <> "" And opcion = "I" Then
           Name Trim(fpText1.text) As Mid(Trim(fpText1.text), 1, Len(Trim(fpText1.text)) - 3) & "dwl"
           fpText1.text = ""
       End If
       MsgBox IIf(opcion = "E", "Proceso de Exportar Finalizado", "Proceso de Importar Finalizado"), vbInformation + vbOKOnly, Msgtitulo
    End If
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Function ImportarDatosSql(ByVal cdbz As String) As Long
Dim fso As New FileSystemObject, cdbi As String, indice As Long, cDBO As String, DBO As String, spid  As Long
On Error GoTo Man_Error
ImportarDatosSql = False
If Not fso.FileExists(Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & cdbz) Then MsgBox "No se encuentra el archivo para importar datos...", vbExclamation + vbOKOnly, Msgtitulo: Exit Function
cdbi = Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
AZ1.OpenZip Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & cdbz
AZ1.ExtractFile AZ1.Filename(0), Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb", ""
AZ1.Close
Set dbI = New ADODB.Connection
dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbI.ConnectionTimeout = 3600
dbI.CommandTimeout = 3600
dbI.Open
RS.Open "SELECT * FROM a_procesa WHERE pro_codigo = '1'", dbI, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Archivo a procesar no existe, proceso cancelado...", vbInformation + vbOKOnly, Msgtitulo: ImportarDatosSql = True: Exit Function
RS.Close: Set RS = Nothing
'RS.Open "SELECT DISTINCT substring(convert(varchar(8),min_fecmin),1,6) AS fecha FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND min_fecmin IN (SELECT min_fecmin FROM b_minuta IN '" & cDBI & "') AND min_indblo = 1", vg_db, adOpenStatic
RS.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_fecmin IN (SELECT DISTINCT min_fecmin FROM b_minuta IN " & DBO & " WHERE (min_indblo = 1 or min_indblo IS NULL) AND min_cencos = '" & MuestraCasino(1) & "')", dbI, adOpenStatic
If Not RS.EOF Then
'   dbI.Close: Set dbI = Nothing: RS.Close: Set RS = Nothing: fso.DeleteFile cDBI: ImportarDatosSql = False
   dbI.Close: Set dbI = Nothing: fso.DeleteFile cdbi: ImportarDatosSql = False
'   Name cdbz As Mid(cdbz, 1, Len(cdbz) - 3) & "dwl"
   fpText1.text = ""
   MsgBox "Planificación minuta esta bloqueada, proceso cancelado...", vbInformation + vbOKOnly, Msgtitulo: ImportarDatosSql = False: Exit Function
End If
RS.Close: Set RS = Nothing
'------- Actualizar centro costo
dbI.Execute "UPDATE b_minuta SET min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "'"
PB.Min = 0: PB.Value = 0: PB.max = 30
Label1(2).Visible = True: PB.Visible = True
'-------> Borrar tabla paso receta -
vg_db.Execute "DELETE paso_receta WHERE rec_spid = @@spid AND rec_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_recetadet WHERE red_spid = @@spid AND red_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_productosing WHERE pri_spid = @@spid AND pri_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_regimen WHERE reg_spid = @@spid AND reg_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_productospmpdia WHERE ppd_spid = @@spid AND ppd_usuario = '" & vg_NUsr & "'"
'-------> Buscar spid
Set RS = vg_db.Execute("SELECT @@spid spid")
If Not RS.EOF Then spid = RS!spid
RS.Close: Set RS = Nothing


'------- Tipos de Producto
Label1(2).Caption = "Importando Tipos de Producto": DoEvents
RS1.Open "SELECT * FROM a_tipopro", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipopro WHERE tip_codigo=" & RS1!tip_codigo
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
'------- Parametro Despacho
Label1(2).Caption = "Importando Parametro de Despacho": DoEvents
RS1.Open "SELECT b.tip_codigo, b.tip_nombre FROM a_tipopro a INNER JOIN a_tipopro AS b ON a.tip_codigo=b.tip_previo WHERE a.tip_previo = 0", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      RS2.Open "SELECT DISTINCT pad_codigo FROM b_paramdesp WHERE pad_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND pad_codigo = " & RS1!tip_codigo & "", vg_db, adOpenStatic
      If RS2.EOF Then vg_db.Execute "INSERT INTO b_paramdesp VALUES (" & RS1!tip_codigo & ", 'S', '" & LimpiaDato(Trim(fpText(0).text)) & "', '')"
      RS2.Close: Set RS2 = Nothing
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
 
'------- Unidades de medida
Label1(2).Caption = "Importando Unidades de Medida": DoEvents
RS1.Open "SELECT * FROM a_unidadmed", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidadmed WHERE unm_codigo=" & RS1!unm_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Unidades de stock
Label1(2).Caption = "Importando Unidades de Stock"
DoEvents
RS1.Open "SELECT * FROM a_unidad", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidad WHERE uni_codigo=" & RS1!uni_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Unidades de embalaje
Label1(2).Caption = "Importando unidades de embalaje"
DoEvents
RS1.Open "SELECT * FROM a_embalaje", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_embalaje WHERE emb_codigo=" & RS1!emb_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Cuentas Contables
Label1(2).Caption = "Importando Cuentas Contables"
DoEvents
RS1.Open "SELECT * FROM a_ctacontable", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_ctacontable WHERE cta_codigo='" & RS1!cta_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Parametro
dbI.Execute "UPDATE a_param SET par_cencos='" & MuestraCasino(1) & "'"
Label1(2).Caption = "Importando Parametros"
DoEvents
RS1.Open "SELECT * FROM a_param", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_param WHERE par_cencos='" & RS1!par_cencos & "' AND par_codigo='" & RS1!par_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Impuestos
Label1(2).Caption = "Importando Impuestos"
DoEvents
RS1.Open "SELECT * FROM a_impuesto", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_impuesto WHERE imp_codigo=" & RS1!imp_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Nutrientes
Label1(2).Caption = "Importando Nutrientes"
DoEvents
RS1.Open "SELECT * FROM a_nutriente", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_nutriente WHERE nut_codigo=" & RS1!nut_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Articulos de Stock
Label1(2).Caption = "Importando Artículos de Stock": DoEvents
RS1.Open "SELECT * FROM b_productos", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productos WHERE pro_codigo='" & RS1!pro_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Impuestos Articulos de Stock
Label1(2).Caption = "Importando Impuestos Relacionados": DoEvents
RS1.Open "SELECT * FROM b_productosimp", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosimp WHERE ipr_codpro='" & RS1!ipr_codpro & "' AND ipr_codimp=" & RS1!ipr_codimp
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Ingredientes
Label1(2).Caption = "Importando Ingredientes": DoEvents
RS1.Open "SELECT * FROM b_ingrediente", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_ingrediente WHERE ing_codigo='" & RS1!ing_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Ingredientes Articulos de Stock
Label1(2).Caption = "Importando Ingredientes Relacionados": DoEvents
RS1.Open "SELECT * FROM b_productosing", dbI, adOpenStatic
Do While Not RS1.EOF
   vg_db.Execute "INSERT INTO paso_productosing VALUES (" & spid & ", '" & vg_NUsr & "', '" & RS1!pri_codpro & "')"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
vg_db.Execute "DELETE FROM b_productosing WHERE pri_codpro IN (SELECT pri_codpro FROM paso_productosing WHERE pri_spid = " & spid & " AND pri_usr = '" & vg_NUsr & "')"
RS1.Open "SELECT * FROM b_productosing", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosing WHERE pri_codpro='" & RS1!pri_codpro & "' AND pri_coding='" & RS1!pri_coding & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Aportes Nutricionales Ingrediente
Label1(2).Caption = "Importando Aportes Nutricionales Ingrediente": DoEvents
RS1.Open "SELECT * FROM b_productonut", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productonut WHERE pnu_codpro='" & RS1!pnu_codpro & "' AND pnu_codapo=" & RS1!pnu_codapo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Proveedores
Label1(2).Caption = "Importando Proveedores": DoEvents
'    vg_db.Execute "delete from b_proveedor where prv_codigo not in (select toc_rutpro from b_totcompras)"
RS1.Open "SELECT * FROM b_proveedor", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_proveedor WHERE prv_codigo='" & RS1!prv_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'*'------- Actualizar ingrediente que tengan precio negativo
'* vg_db.Execute "UPDATE b_contlistprepro INNER JOIN (b_contlistpreing INNER JOIN b_productos ON (b_contlistpreing.cpi_codped=b_productos.pro_codigo) AND (b_contlistpreing.cpi_codcom = b_productos.pro_codigo)) ON b_contlistprepro.cpp_codpro = b_productos.pro_codigo SET b_contlistpreing.cpi_precos=iif(b_contlistprepro.cpp_propon<0 or b_productos.pro_facing<=0,0,b_contlistprepro.cpp_propon/b_productos.pro_facing) " & _
'*              "WHERE b_contlistpreing.cpi_precos<0 AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "' AND b_contlistprepro.cpp_cencos='" & MuestraCasino(1) & "' AND b_productos.pro_ctacon='410001'"

'------- Mover zero al stock si es negativo
vg_db.Execute "UPDATE b_bodegas set bod_canmer=0 WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<0"

'------- Categoría de Receta
Label1(2).Caption = "Importando Categoría de Receta": DoEvents
RS1.Open "SELECT * FROM a_recetacatdie", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetacatdie WHERE car_codigo=" & RS1!car_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Tipo de Plato
Label1(2).Caption = "Importando Tipo de Plato": DoEvents
RS1.Open "SELECT * FROM a_recetatippla", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetatippla WHERE tip_codigo=" & RS1!tip_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Recetas
Label1(2).Caption = "Importando Recetas": DoEvents
RS1.Open "SELECT * FROM b_receta", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_receta WHERE rec_codigo=" & RS1!rec_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Ingredientes de Recetas
'------- Agregar campo
dbI.Execute "UPDATE b_recetadet SET red_cencos = '0' WHERE red_tiprec = 0"
dbI.Execute "UPDATE b_recetadet SET red_cencos = '" & MuestraCasino(1) & "' WHERE red_tiprec <> 0"

Label1(2).Caption = "Importando Ingredientes Recetas": DoEvents
RS1.Open "SELECT DISTINCT red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos FROM b_recetadet", dbI, adOpenStatic
Do While Not RS1.EOF
   vg_db.Execute "INSERT INTO paso_recetadet VALUES (" & spid & ", '" & vg_NUsr & "', " & RS1!red_codigo & ", " & RS1!red_nroite & ", '" & RS1!red_codpro & "', " & RS1!red_canpro & ", " & RS1!red_cospro & ", " & RS1!red_pctapr & ", " & RS1!red_pctcoc & ", " & RS1!red_pctnut & ", " & RS1!red_tiprec & ", '" & RS1!red_cencos & "')"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
vg_db.Execute "DELETE FROM b_recetadet WHERE red_codigo IN (SELECT red_codigo FROM paso_recetadet WHERE ((red_tiprec <> 0 AND red_cencos = '" & MuestraCasino(1) & "') OR (red_tiprec = 0 AND red_cencos = '0'))) AND red_tiprec IN (SELECT red_tiprec FROM paso_recetadet WHERE red_cencos='" & MuestraCasino(1) & "' OR red_cencos='0')"
RS1.Open "SELECT * FROM b_recetadet", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_recetadet WHERE red_codigo=" & RS1!red_codigo & " AND red_nroite=" & RS1!red_nroite & " AND red_tiprec=" & RS1!red_tiprec & " AND red_cencos='" & RS1!red_cencos & "'"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Regimen
Label1(2).Caption = "Importando Regimen": DoEvents
RS1.Open "SELECT * FROM a_regimen", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_regimen WHERE reg_codigo=" & RS1!reg_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Servicio
Label1(2).Caption = "Importando Servicio": DoEvents
RS1.Open "SELECT * FROM a_servicio", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_servicio WHERE ser_codigo=" & RS1!ser_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Sector
Label1(2).Caption = "Importando Sector": DoEvents
RS1.Open "SELECT * FROM a_sector", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_sector WHERE sec_codigo=" & RS1!sec_codigo & ""
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Estructura Servicio
Label1(2).Caption = "Importando Estructura Servicio": DoEvents
dbI.Execute "UPDATE a_estservicio SET ess_cencos='" & MuestraCasino(1) & "'"
dbI.Execute "UPDATE a_estservicio SET ess_racmin=0 WHERE ess_racmin is null"
RS1.Open "SELECT * FROM a_estservicio", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_estservicio WHERE ess_codser=" & RS1!ess_codser & " AND ess_codigo=" & RS1!ess_codigo & " AND ess_cencos='" & RS1!ess_cencos & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Validar si existe planificación minutas
vg_db.BeginTrans
indice = 0
Label1(2).Caption = "Validar Planificación Minutas": DoEvents
RS1.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha, min_codreg, reg_nombre, min_codser, ser_nombre FROM b_minuta a, a_regimen b, a_servicio c WHERE a.min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND a.min_codreg = b.reg_codigo AND a.min_codser = c.ser_codigo", dbI, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      RS2.Open "SELECT DISTINCT convert(int,substring(convert(varchar(8),min_fecmin),1,6)) AS fecha FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & RS1!Fecha & " AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & "", vg_db, adOpenStatic
      If Not RS2.EOF Then
         If MsgBox("Existe planificación minuta, desea borrar la información existente... " & VgLinea & VgLinea & "Regimen  : " & RS1!min_codreg & " " & Trim(RS1!reg_nombre) & VgLinea & "Servicio   :  " & RS1!min_codser & " " & Trim(RS1!ser_nombre), vbQuestion + vbYesNo, Msgtitulo) = vbYes Then
            '------- Borrar planificación contrato
            vg_db.Execute "DELETE b_minutadet FROM b_minuta, b_minutadet WHERE b_minuta.min_codigo = b_minutadet.mid_codigo AND b_minuta.min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND convert(int,substring(convert(varchar(8),b_minuta.min_fecmin),1,6)) = " & RS1!Fecha & " AND b_minuta.min_codreg = " & RS1!min_codreg & " AND b_minuta.min_codser = " & RS1!min_codser & ""
            vg_db.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND convert(int,substring(convert(varchar(8),min_fecmin),1,6)) = " & RS1!Fecha & " AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & ""
         Else
            '------- Borrar planificación de la base carga
            dbI.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo WHERE b_minuta.min_cencos = '" & MuestraCasino(1) & "'AND VAL(MID(b_minuta.min_fecmin,1,6)) = " & RS1!Fecha & " AND b_minuta.min_codreg = " & RS1!min_codreg & " AND b_minuta.min_codser = " & RS1!min_codser & ""
            dbI.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND VAL(MID(min_fecmin,1,6)) = " & RS1!Fecha & " AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & ""
         End If
      End If
      RS2.Close: Set RS2 = Nothing
      '------- Traer ultimo correlativo
      If indice = 0 Then
         RS2.Open "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC", vg_db, adOpenStatic
         If Not RS2.EOF Then RS2.MoveFirst: indice = RS2!min_codigo + 1 Else indice = 1
         RS2.Close: Set RS2 = Nothing
      End If
      '------- actualizar correlativo planificación base externa
      RS2.Open "SELECT DISTINCT min_codigo, min_codreg FROM b_minuta WHERE min_cencos='" & MuestraCasino(1) & "' AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & "", dbI, adOpenStatic
      If Not RS2.EOF Then
         Do While Not RS2.EOF
            dbI.Execute "UPDATE b_minutadet SET mid_codigo=" & indice & " WHERE mid_codigo=" & RS2!min_codigo & ""
            dbI.Execute "UPDATE b_minuta SET min_codigo = " & indice & " WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codigo = " & RS2!min_codigo & ""
            RS2.MoveNext: indice = indice + 1
         Loop
      End If
      RS2.Close: Set RS2 = Nothing
      '------- actualizar nro. raciones totales
      RS2.Open "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos = '" & MuestraCasino(1) & "' AND sra_codser = " & RS1!min_codser & " GROUP BY sra_serdia ORDER BY sra_serdia", vg_db, adOpenStatic
      If Not RS2.EOF Then
         Do While Not RS2.EOF
            dbI.Execute "UPDATE b_minuta SET min_racteo = " & RS2!raciones & " WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & " AND IIF(datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)) = 1,7,datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)-1))=" & RS2!sra_serdia & ""
            RS2.MoveNext
         Loop
      End If
      RS2.Close: Set RS2 = Nothing
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
vg_db.CommitTrans

'------- Encabezado Planificación
Label1(2).Caption = "Importando Planificación Encabezado": DoEvents
RS1.Open "SELECT * FROM b_minuta", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minuta WHERE min_codigo = " & RS1!min_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Detalle Planificación
Label1(2).Caption = "Importando Planificación Detalle": DoEvents
RS1.Open "SELECT * FROM b_minutadet", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutadet WHERE mid_codigo=" & RS1!mid_codigo & " AND mid_tipmin='" & RS1!mid_tipmin & "' AND mid_numlin=" & RS1!mid_numlin & ""
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Costo Patron
Label1(2).Caption = "Importando Costo Patron": DoEvents
dbI.Execute "UPDATE b_costopatron SET cpa_cencos='" & MuestraCasino(1) & "'"
RS1.Open "SELECT * FROM b_costopatron", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_costopatron WHERE cpa_cencos='" & RS1!cpa_cencos & "' AND cpa_codreg=" & RS1!cpa_codreg & " AND cpa_codser=" & RS1!cpa_codser & " AND cpa_anomes=" & RS1!cpa_anomes & " AND cpa_descripcion='" & RS1!cpa_descripcion & "'"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Gramos Familia Producto
Label1(2).Caption = "Importando Gramos Familia Producto": DoEvents
RS1.Open "SELECT * FROM b_gramofamproducto", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_gramofamproducto WHERE gfp_cencos='" & RS1!gfp_cencos & "' AND gfp_codreg=" & RS1!gfp_codreg & " AND gfp_catdie=" & RS1!gfp_catdie & " AND gfp_tiprec=" & RS1!gfp_tiprec & " AND gfp_fampro=" & RS1!gfp_fampro & ""
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Tipo de Servicio
Label1(1).Caption = "Importando tipo de servicio"
DoEvents
RS1.Open "SELECT * FROM a_tiposervicio", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tiposervicio WHERE tis_codigo=" & RS1!tis_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Segmento
Label1(1).Caption = "Importando segmento"
DoEvents
RS1.Open "SELECT * FROM a_segmento", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_segmento WHERE seg_codigo=" & RS1!seg_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

vg_db.BeginTrans
'------- Actualizar tabla lista producto y lista ingrediente
'* vg_db.Execute "INSERT INTO b_contlistprepro (cpp_cencos, cpp_codpro, cpp_upreco, cpp_fecuco, cpp_propon) SELECT '" & MuestraCasino(1) & "', pro_codigo, 0, null, 0 FROM b_productos WHERE pro_codigo NOT IN (SELECT DISTINCT cpp_codpro FROM b_contlistprepro WHERE cpp_cencos='" & MuestraCasino(1) & "')"
vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & MuestraCasino(1) & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente WHERE ing_codigo NOT IN (SELECT DISTINCT cpi_coding FROM b_contlistpreing WHERE cpi_cencos='" & MuestraCasino(1) & "')"
vg_db.CommitTrans

'-------> Borrar tabla paso receta -
vg_db.Execute "DELETE paso_receta WHERE rec_spid = " & spid & " AND rec_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_recetadet WHERE red_spid = " & spid & " AND red_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_productosing WHERE pri_spid = " & spid & " AND pri_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_regimen WHERE reg_spid = " & spid & " AND reg_usr = '" & vg_NUsr & "'"
vg_db.Execute "DELETE paso_productospmpdia WHERE ppd_spid = " & spid & " AND ppd_usuario = '" & vg_NUsr & "'"

dbI.Close: Set dbI = Nothing
'vg_db.BeginTrans
fso.DeleteFile cdbi
'vg_db.Execute "insert into log_actualizacion values ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', cdate('" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm") & "'))"
ImportarDatosSql = True
'vg_db.CommitTrans
'------- Rutina validar producto vigente
ValidarProductoVigente
Label1(2).Visible = False: PB.Visible = False

Exit Function
Man_Error:
If Err = -2147217865 Or Err = 3265 Then
   dbI.Close: Set dbI = Nothing
'    vg_db.BeginTrans
'    fso.DeleteFile cDBI
'    vg_db.Execute "insert into log_actualizacion values ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', cdate('" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm") & "'))"
   ImportarDatosSql = True
'    vg_db.CommitTrans
   Exit Function
ElseIf Err = -2147467259 Then
   MsgBox "Archivo no corresponde a un MDB, proceso cancelado...", vbInformation + vbOKOnly, Msgtitulo
   ImportarDatosSql = False
   Exit Function
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Function
vg_db.RollbackTrans
If Err.Number = -2147467259 Then
    MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Function

Private Function ExportarDatos(ByVal cdbz As String, codreg As String, codser As String) As Long
fg_carga ""
ExportarDatos = False
Frame1(0).Enabled = False: Toolbar1.Enabled = False
PB.Visible = True: PB.Min = 0: PB.Value = 0: PB.max = 30
Dim cDBO As String
'-------> Crear directorio para generar planificación
'    If Dir(CD.Filename, vbVolume) = "" Then MkDir CD.Filename
'-------> Generar base padre
cDBO = dir_trabajo & BaseDeDato
If Dir(Cd.Filename) <> "" Then Kill Cd.Filename 'borrar base datos si existe
'-------> generar archivo mdb
Set db1 = DBEngine(0).CreateDatabase(Cd.Filename, dbLangGeneral)
'-------> tabla relacionada a productos
db1.Execute "CREATE TABLE a_procesa (pro_codigo char(1))"
db1.Execute "CREATE TABLE a_tipopro (tip_codigo int, tip_nombre char(35), tip_previo int)"
db1.Execute "CREATE TABLE a_unidad (uni_codigo int, uni_nombre char(10), uni_nomcor char(5))"
db1.Execute "CREATE TABLE a_embalaje (emb_codigo int, emb_nombre char(20), emb_nomcor char(5))"
db1.Execute "CREATE TABLE a_ctacontable (cta_codigo char(10), cta_nombre char(40))"
db1.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
db1.Execute "CREATE TABLE a_impuesto (imp_codigo int, imp_nombre char(15), imp_pctimp double, imp_inccos int, imp_codsap char(20), imp_indmod char(1))"
db1.Execute "CREATE TABLE a_unidadmed (unm_codigo int, unm_nombre char(10), unm_nomcor char(5))"
db1.Execute "CREATE TABLE a_nutriente (nut_codigo int, nut_nombre char(30), nut_nomuni char(5), nut_indpri int, nut_secnro int)"
db1.Execute "CREATE TABLE b_productos (pro_codigo char(20), pro_codbar char(20), pro_codcom char(20), pro_codtip int, pro_nombre char(50), pro_coduni int, pro_facing double, pro_facsto double, pro_codemb int, pro_uniemb double, pro_upreco double, pro_fecuco datetime, pro_propon double, pro_ctacon char(10), pro_fecven int, pro_ctrsto int)"
db1.Execute "CREATE TABLE b_productosimp (ipr_codpro char(20), ipr_codimp int)"
db1.Execute "CREATE TABLE b_productosing (pri_codpro char(20), pri_coding char(20))"
db1.Execute "CREATE TABLE b_ingrediente (ing_codigo char(20), ing_nombre char(50), ing_nomfan char(50), ing_unimed int, ing_pctapr double, ing_pctcoc double, ing_pctnut double, ing_facnut double, ing_indpav int, ing_indgrv int, ing_precos double, ing_feccos int, ing_codcom char(20), ing_codped char(20))"
db1.Execute "CREATE TABLE b_productonut (pnu_codpro char(20), pnu_codapo int, pnu_canapo double)"
db1.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo date, prv_origen char(1))"
db1.Execute "CREATE TABLE a_tiposervicio (tis_codigo int, tis_nombre char(50))"
db1.Execute "CREATE TABLE a_segmento (seg_codigo int, seg_nombre char(50))"
'-------> tabla relacionada a recetas
db1.Execute "CREATE TABLE a_recetacatdie (car_codigo int, car_nombre char(50), car_previo int)"
db1.Execute "CREATE TABLE a_recetatippla (tip_codigo int, tip_nombre char(50), tip_previo int)"
db1.Execute "CREATE TABLE b_receta (rec_codigo int, rec_catdie int, rec_tippla int, rec_nombre char(80), rec_nomfan char(80), rec_metpre longtext, rec_conche longtext, rec_sugere longtext, rec_basrac int, rec_tiprec int, rec_fecvig int, rec_gruvul longtext)"
db1.Execute "CREATE TABLE b_recetadet (red_codigo int, red_nroite int, red_codpro char(20), red_canpro double, red_cospro double, red_pctapr double, red_pctcoc double, red_pctnut double, red_tiprec int, red_cencos char(10))"
'-------> tabla relacionada a planificación
db1.Execute "CREATE TABLE a_sector (sec_codigo int, sec_nombre char(50), sec_orden int)"
db1.Execute "CREATE TABLE b_costopatron (cpa_cencos char(10), cpa_codreg int, cpa_codser int, cpa_anomes int, cpa_descripcion char(10), cpa_valor double)"
db1.Execute "CREATE TABLE b_gramofamproducto (gfp_cencos char(10), gfp_codreg int, gfp_catdie int, gfp_tiprec int, gfp_fampro int, gfp_graini double, gfp_grafin double)"
db1.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
db1.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(50), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob date, ser_horent date, ser_horpda date)"
db1.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double, ess_cencos char(10))"
db1.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int, Constraint b_minuta_pk Primary Key (min_codigo))"
db1.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac int, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer int, mid_rec5eta char(1), mid_cosdes double, Constraint b_minutadet_pk Primary Key (mid_codigo, mid_numlin))"
'-------> generar archivo origen
db1.Execute "INSERT INTO a_procesa VALUES('1')"
PB.Value = PB.Value + 1
'-------> generar familia productos
db1.Execute "INSERT INTO a_tipopro SELECT tip_codigo, tip_nombre, tip_previo FROM a_tipopro IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar unidad medida productos
db1.Execute "INSERT INTO a_unidad SELECT DISTINCT uni_codigo, uni_nombre, uni_nomcor FROM a_unidad IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar embalaje productos
db1.Execute "INSERT INTO a_embalaje SELECT DISTINCT emb_codigo, emb_nombre, emb_nomcor FROM a_embalaje IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar cuentas contables productos
db1.Execute "INSERT INTO a_ctacontable SELECT DISTINCT cta_codigo, cta_nombre FROM a_ctacontable IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar parametros cuentas contables
db1.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN '" & cDBO & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo IN ('ctagastos','ctagastos2','ctainsumo','ctalimdes')"
db1.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN '" & cDBO & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='5etapas'"
db1.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN '" & cDBO & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='addreceta'"
db1.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN '" & cDBO & "' WHERE par_cencos='" & MuestraCasino(1) & "' AND par_codigo='opgruvul'"
PB.Value = PB.Value + 1
'-------> generar impuesto productos
db1.Execute "INSERT INTO a_impuesto SELECT imp_codigo, imp_nombre, imp_pctimp, imp_inccos, imp_codsap, imp_indmod FROM a_impuesto IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar unidad medida ingrediente
db1.Execute "INSERT INTO a_unidadmed SELECT DISTINCT unm_codigo, unm_nombre, unm_nomcor FROM a_unidadmed IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar nutriente aporte
db1.Execute "INSERT INTO a_nutriente SELECT DISTINCT nut_codigo, nut_nombre, nut_nomuni, nut_indpri, nut_secnro FROM a_nutriente IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar proveedores
DoEvents
db1.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen FROM b_proveedor IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> Generar productos
DoEvents
db1.Execute "INSERT INTO b_productos SELECT DISTINCT pro_codigo, pro_codbar, pro_codcom, pro_codtip, pro_nombre, pro_coduni, pro_facing, pro_facsto, pro_codemb, " & _
            "pro_uniemb, pro_upreco, pro_fecuco, pro_propon, pro_ctacon, pro_fecven, pro_ctrsto FROM b_productos IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar productos impuestos
DoEvents
db1.Execute "INSERT INTO b_productosimp SELECT DISTINCT ipr_codpro, ipr_codimp FROM b_productosimp IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar productos ingredientes & ingredientes
DoEvents
db1.Execute "INSERT INTO b_productosing SELECT DISTINCT pri_codpro, pri_coding FROM b_productosing IN '" & cDBO & "'"
DoEvents: PB.Value = PB.Value + 1
db1.Execute "INSERT INTO b_ingrediente SELECT DISTINCT ing_codigo , ing_nombre, ing_nomfan, ing_unimed, ing_pctapr, ing_pctcoc, ing_pctnut, ing_facnut, ing_indpav, " & _
            "ing_indgrv, ing_precos, ing_feccos, ing_codcom, ing_codped FROM b_ingrediente IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar nutriente del ingrediente
db1.Execute "INSERT INTO b_productonut SELECT DISTINCT pnu_codpro, pnu_codapo, pnu_canapo FROM b_productonut IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar categoria dietetica
db1.Execute "INSERT INTO a_recetacatdie SELECT DISTINCT car_codigo, car_nombre, car_previo FROM a_recetacatdie IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar tipo plato
db1.Execute "INSERT INTO a_recetatippla SELECT DISTINCT tip_codigo, tip_nombre, tip_previo FROM a_recetatippla IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> Generar encabezado receta
db1.Execute "INSERT INTO b_receta SELECT DISTINCT a.rec_codigo, a.rec_catdie, a.rec_tippla, a.rec_nombre, a.rec_nomfan, a.rec_metpre, a.rec_conche, a.rec_sugere, a.rec_basrac, a.rec_tiprec, a.rec_fecvig, a.rec_gruvul " & _
            "FROM b_receta a, b_minuta b, b_minutadet c IN '" & cDBO & "' WHERE b.min_codigo=c.mid_codigo AND c.mid_codrec=a.rec_codigo AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND  b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'"
DoEvents: PB.Value = PB.Value + 1
'-------> generar detalle recetas
db1.Execute "INSERT INTO b_recetadet SELECT a.red_codigo, a.red_nroite, a.red_codpro, a.red_canpro, a.red_cospro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, red_tiprec, red_cencos " & _
            "FROM b_recetadet a, b_receta b, b_minuta c, b_minutadet d IN '" & cDBO & "' WHERE c.min_codigo=d.mid_codigo AND d.mid_codrec=b.rec_codigo AND (d.mid_tiprec=a.red_tiprec OR a.red_tiprec=0) AND ((a.red_tiprec<>0 AND a.red_cencos='" & MuestraCasino(1) & "') OR (a.red_tiprec=0 AND a.red_cencos='0')) AND b.rec_codigo=a.red_codigo " & _
            "AND  c.min_cencos='" & fpText(0).text & "' AND c.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND  c.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(c.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND d.mid_tipmin='1'"

'            "FROM b_recetadet a, b_receta b, b_minuta c, b_minutadet d IN '" & cDBO & "' WHERE c.min_codigo=d.mid_codigo AND d.mid_codrec=b.rec_codigo AND (d.mid_tiprec=a.red_tiprec OR a.red_tiprec=0) AND ((a.red_tiprec<>0 AND a.red_cencos='" & MuestraCasino(1) & "') OR (a.red_tiprec=0 AND a.red_cencos='0')) AND b.rec_codigo=a.red_codigo " & _
DoEvents: PB.Value = PB.Value + 1
'-------> generar regimen
db1.Execute "INSERT INTO a_regimen SELECT DISTINCT a.reg_codigo, a.reg_nombre FROM a_regimen a, b_minuta b, b_minutadet c IN '" & cDBO & "' WHERE b.min_codigo=c.mid_codigo AND a.reg_codigo=b.min_codreg AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'"
PB.Value = PB.Value + 1
'-------> generar servicio
db1.Execute "INSERT INTO a_servicio SELECT DISTINCT a.ser_codigo, a.ser_nombre, a.ser_orden, a.ser_codsap, a.ser_facturable, a.ser_activo, a.ser_horcob, a.ser_horent, a.ser_horpda FROM a_servicio a, b_minuta b, b_minutadet c IN '" & cDBO & "' WHERE b.min_codigo=c.mid_codigo AND a.ser_codigo=b.min_codser AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND c.mid_tipmin='1'"
PB.Value = PB.Value + 1
'-------> generar sector
db1.Execute "INSERT INTO a_sector SELECT DISTINCT e.sec_codigo, e.sec_nombre, e.sec_orden FROM a_estservicio a, a_servicio b, b_minuta c, b_minutadet d, a_sector e IN '" & cDBO & "' WHERE c.min_codigo=d.mid_codigo AND b.ser_codigo=c.min_codser AND b.ser_codigo=a.ess_codser AND a.ess_codsec=e.sec_codigo AND c.min_cencos=a.ess_cencos AND c.min_cencos='" & fpText(0).text & "' AND c.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND c.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(c.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND d.mid_tipmin='1'"
PB.Value = PB.Value + 1
'-------> generar estructura servicio
db1.Execute "INSERT INTO a_estservicio SELECT DISTINCT a.ess_codser, a.ess_codigo, a.ess_nombre, a.ess_orden, a.ess_codsec, a.ess_racmin, a.ess_cencos FROM a_estservicio a, a_servicio b, b_minuta c, b_minutadet d IN '" & cDBO & "' WHERE c.min_codigo=d.mid_codigo AND b.ser_codigo=c.min_codser AND b.ser_codigo=a.ess_codser AND c.min_cencos=a.ess_cencos AND c.min_cencos='" & fpText(0).text & "' AND c.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND c.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(c.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND d.mid_tipmin='1'"
PB.Value = PB.Value + 1
'-------> generar encabezado planificación minutas
db1.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, 0 AS min_indblo, 0 AS min_racteo, 0 AS min_racrea FROM b_minuta a, b_minutadet b IN '" & cDBO & "' WHERE a.min_codigo=b.mid_codigo AND a.min_cencos='" & fpText(0).text & "' AND a.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND a.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(a.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND b.mid_tipmin='1'"
PB.Value = PB.Value + 1
'-------> generar detalle planificación minutas
db1.Execute "INSERT INTO b_minutadet SELECT DISTINCT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, 0 AS mid_fecval, a.mid_tiprec, a.mid_cosdes FROM b_minutadet a, b_minuta b IN '" & cDBO & "' WHERE b.min_codigo=a.mid_codigo AND b.min_cencos='" & fpText(0).text & "' AND b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND val(mid(b.min_fecmin,1,6))=" & Format(fpDateTime1.text, "yyyymm") & " AND a.mid_tipmin='1'"
PB.Value = PB.Value + 1
'-------> costo patron
db1.Execute "INSERT INTO b_costopatron SELECT cpa_cencos, cpa_codreg, cpa_codser, cpa_anomes, cpa_descripcion, cpa_valor FROM b_costopatron IN '" & cDBO & "' WHERE cpa_cencos='" & fpText(0).text & "' AND cpa_anomes=" & Format(fpDateTime1.text, "yyyymm") & " AND cpa_descripcion='PISO'"
PB.Value = PB.Value + 1
'-------> gramo familia producto
db1.Execute "INSERT INTO b_gramofamproducto SELECT gfp_cencos, gfp_codreg, gfp_catdie, gfp_tiprec, gfp_fampro, gfp_graini, gfp_grafin FROM b_gramofamproducto IN '" & cDBO & "' WHERE gfp_cencos='" & fpText(0).text & "'"
PB.Value = PB.Value + 1
'-------> generar tipo servicio
db1.Execute "INSERT INTO a_tiposervicio SELECT DISTINCT tis_codigo, tis_nombre FROM a_tiposervicio IN '" & cDBO & "'"
PB.Value = PB.Value + 1
'-------> generar segmento
db1.Execute "INSERT INTO a_segmento SELECT DISTINCT seg_codigo, seg_nombre FROM a_segmento IN '" & cDBO & "'"
PB.Value = PB.Value + 1
db1.Close
If Dir(Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & Mid(Dir(Cd.Filename), 1, (Len(Dir(Cd.Filename)) - 4)) & ".zip") <> "" Then Kill Mid(Dir(Cd.Filename), 1, (Len(Dir(Cd.Filename)) - 4)) & ".zip" 'borrar base datos si existe
AZ1.CreateZip Mid(Dir(Cd.Filename), 1, (Len(Dir(Cd.Filename)) - 4)) & ".zip", "": AZ1.AddFile Cd.Filename, "", True, "": AZ1.Close
If Dir(Cd.Filename) <> "" Then Kill Cd.Filename 'borrar base datos si existe
ExportarDatos = True
fg_descarga
PB.Visible = False: Frame1(0).Enabled = True: Toolbar1.Enabled = True
End Function

Private Function ImportarDatosAccess(ByVal cdbz As String) As Long
Dim fso As New FileSystemObject, cdbi As String, indice As Long, cDBO As String, DBO As String, spid As Long
On Error GoTo Man_Error
ImportarDatosAccess = False
If Not fso.FileExists(Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & cdbz) Then MsgBox "No se encuentra el archivo para importar datos...", vbExclamation + vbOKOnly, Msgtitulo: Exit Function
cdbi = Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
AZ1.OpenZip Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & cdbz
AZ1.ExtractFile AZ1.Filename(0), Mid(Cd.Filename, 1, Len(Cd.Filename) - Len(Dir(Cd.Filename))) & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb", ""
AZ1.Close
Set dbI = New ADODB.Connection
dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbI.ConnectionTimeout = 3600
dbI.CommandTimeout = 3600
dbI.Open
RS.Open "SELECT * FROM a_procesa WHERE pro_codigo = '1'", dbI, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Archivo a procesar no existe, proceso cancelado...", vbInformation + vbOKOnly, Msgtitulo: ImportarDatosAccess = True: Exit Function
RS.Close: Set RS = Nothing
If vg_tipbase = "1" Then
   RS.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND min_fecmin IN (SELECT min_fecmin FROM b_minuta IN '" & cdbi & "') AND min_indblo = 1", vg_db, adOpenStatic
Else
   RS.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND min_fecmin IN (SELECT min_fecmin FROM b_minuta IN '" & DBO & "') AND min_indblo = 1", vg_db, adOpenStatic
End If
If Not RS.EOF Then
   dbI.Close: Set dbI = Nothing: RS.Close: Set RS = Nothing: fso.DeleteFile cdbi: ImportarDatosAccess = False
'   Name cdbz As Mid(cdbz, 1, Len(cdbz) - 3) & "dwl"
   fpText1.text = ""
   MsgBox "Planificación minuta esta bloqueada, proceso cancelado...", vbInformation + vbOKOnly, Msgtitulo: ImportarDatosAccess = False: Exit Function
End If
RS.Close: Set RS = Nothing
'------- Actualizar centro costo
dbI.Execute "UPDATE b_minuta SET min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "'"
PB.Min = 0: PB.Value = 0: PB.max = 30
Label1(2).Visible = True: PB.Visible = True
'------- Tipos de Producto
Label1(2).Caption = "Importando Tipos de Producto": DoEvents
RS1.Open "SELECT * FROM a_tipopro", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipopro WHERE tip_codigo = " & RS1!tip_codigo
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
'------- Parametro Despacho
Label1(2).Caption = "Importando Parametro de Despacho": DoEvents
RS1.Open "SELECT b.tip_codigo, b.tip_nombre FROM a_tipopro a INNER JOIN a_tipopro AS b ON a.tip_codigo = b.tip_previo WHERE a.tip_previo = 0", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      RS2.Open "SELECT DISTINCT pad_codigo FROM b_paramdesp WHERE pad_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND pad_codigo = " & RS1!tip_codigo & "", vg_db, adOpenStatic
      If RS2.EOF Then vg_db.Execute "INSERT INTO b_paramdesp VALUES (" & RS1!tip_codigo & ", 'S', '" & LimpiaDato(Trim(fpText(0).text)) & "', '')"
      RS2.Close: Set RS2 = Nothing
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
 
'------- Unidades de medida
Label1(2).Caption = "Importando Unidades de Medida": DoEvents
RS1.Open "SELECT * FROM a_unidadmed", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidadmed WHERE unm_codigo=" & RS1!unm_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Unidades de stock
Label1(2).Caption = "Importando Unidades de Stock"
DoEvents
RS1.Open "SELECT * FROM a_unidad", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidad WHERE uni_codigo=" & RS1!uni_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Unidades de embalaje
Label1(2).Caption = "Importando unidades de embalaje"
DoEvents
RS1.Open "SELECT * FROM a_embalaje", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_embalaje WHERE emb_codigo=" & RS1!emb_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Cuentas Contables
Label1(2).Caption = "Importando Cuentas Contables"
DoEvents
RS1.Open "SELECT * FROM a_ctacontable", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_ctacontable WHERE cta_codigo='" & RS1!cta_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Parametro
dbI.Execute "UPDATE a_param SET par_cencos='" & MuestraCasino(1) & "'"
Label1(2).Caption = "Importando Parametros"
DoEvents
RS1.Open "SELECT * FROM a_param", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_param WHERE par_cencos='" & RS1!par_cencos & "' AND par_codigo='" & RS1!par_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Impuestos
Label1(2).Caption = "Importando Impuestos"
DoEvents
RS1.Open "SELECT * FROM a_impuesto", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_impuesto WHERE imp_codigo=" & RS1!imp_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Nutrientes
Label1(2).Caption = "Importando Nutrientes"
DoEvents
RS1.Open "SELECT * FROM a_nutriente", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_nutriente WHERE nut_codigo=" & RS1!nut_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Articulos de Stock
Label1(2).Caption = "Importando Artículos de Stock": DoEvents
RS1.Open "SELECT * FROM b_productos", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productos WHERE pro_codigo='" & RS1!pro_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Impuestos Articulos de Stock
Label1(2).Caption = "Importando Impuestos Relacionados": DoEvents
RS1.Open "SELECT * FROM b_productosimp", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosimp WHERE ipr_codpro='" & RS1!ipr_codpro & "' AND ipr_codimp=" & RS1!ipr_codimp
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Ingredientes
Label1(2).Caption = "Importando Ingredientes": DoEvents
RS1.Open "SELECT * FROM b_ingrediente", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_ingrediente WHERE ing_codigo='" & RS1!ing_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Ingredientes Articulos de Stock
Label1(2).Caption = "Importando Ingredientes Relacionados": DoEvents
'    vg_db.Execute "Delete  b_productosing from b_productosing where pri_codpro in (select pri_codpro from b_productosing in 'C:\Desarrollo\Casino Skmalaga\Actualizar\mp.mdb')"
vg_db.Execute "DELETE FROM b_productosing WHERE pri_codpro IN (SELECT pri_codpro FROM b_productosing IN '" & cdbi & "')"
RS1.Open "SELECT * FROM b_productosing", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosing WHERE pri_codpro='" & RS1!pri_codpro & "' AND pri_coding='" & RS1!pri_coding & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Aportes Nutricionales Ingrediente
Label1(2).Caption = "Importando Aportes Nutricionales Ingrediente": DoEvents
RS1.Open "SELECT * FROM b_productonut", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productonut WHERE pnu_codpro='" & RS1!pnu_codpro & "' AND pnu_codapo=" & RS1!pnu_codapo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Proveedores
Label1(2).Caption = "Importando Proveedores": DoEvents
'    vg_db.Execute "delete from b_proveedor where prv_codigo not in (select toc_rutpro from b_totcompras)"
RS1.Open "SELECT * FROM b_proveedor", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_proveedor WHERE prv_codigo='" & RS1!prv_codigo & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'*'------- Actualizar ingrediente que tengan precio negativo
'* vg_db.Execute "UPDATE b_contlistprepro INNER JOIN (b_contlistpreing INNER JOIN b_productos ON (b_contlistpreing.cpi_codped=b_productos.pro_codigo) AND (b_contlistpreing.cpi_codcom = b_productos.pro_codigo)) ON b_contlistprepro.cpp_codpro = b_productos.pro_codigo SET b_contlistpreing.cpi_precos=iif(b_contlistprepro.cpp_propon<0 or b_productos.pro_facing<=0,0,b_contlistprepro.cpp_propon/b_productos.pro_facing) " & _
'*              "WHERE b_contlistpreing.cpi_precos<0 AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "' AND b_contlistprepro.cpp_cencos='" & MuestraCasino(1) & "' AND b_productos.pro_ctacon='410001'"

'------- Mover zero al stock si es negativo
vg_db.Execute "UPDATE b_bodegas set bod_canmer=0 WHERE bod_codbod=" & vg_codbod & " AND bod_canmer<0"

'------- Categoría de Receta
Label1(2).Caption = "Importando Categoría de Receta": DoEvents
RS1.Open "SELECT * FROM a_recetacatdie", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetacatdie WHERE car_codigo=" & RS1!car_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Tipo de Plato
Label1(2).Caption = "Importando Tipo de Plato": DoEvents
RS1.Open "SELECT * FROM a_recetatippla", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetatippla WHERE tip_codigo=" & RS1!tip_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Recetas
Label1(2).Caption = "Importando Recetas": DoEvents
RS1.Open "SELECT * FROM b_receta", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_receta WHERE rec_codigo=" & RS1!rec_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Ingredientes de Recetas
'------- Agregar campo
dbI.Execute "UPDATE b_recetadet SET red_cencos='0' WHERE red_tiprec=0"
dbI.Execute "UPDATE b_recetadet SET red_cencos='" & MuestraCasino(1) & "' WHERE red_tiprec<>0"

Label1(2).Caption = "Importando Ingredientes Recetas": DoEvents
vg_db.Execute "DELETE FROM b_recetadet WHERE red_codigo IN (SELECT red_codigo FROM b_recetadet IN '" & cdbi & "' WHERE ((red_tiprec<>0 AND red_cencos='" & MuestraCasino(1) & "') OR (red_tiprec=0 AND red_cencos='0'))) AND red_tiprec IN (SELECT red_tiprec FROM b_recetadet IN '" & cdbi & "' WHERE red_cencos='" & MuestraCasino(1) & "' OR red_cencos='0')"
RS1.Open "SELECT * FROM b_recetadet", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_recetadet WHERE red_codigo=" & RS1!red_codigo & " AND red_nroite=" & RS1!red_nroite & " AND red_tiprec=" & RS1!red_tiprec & " AND red_cencos='" & RS1!red_cencos & "'"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Regimen
Label1(2).Caption = "Importando Regimen": DoEvents
RS1.Open "SELECT * FROM a_regimen", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_regimen WHERE reg_codigo=" & RS1!reg_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Servicio
Label1(2).Caption = "Importando Servicio": DoEvents
RS1.Open "SELECT * FROM a_servicio", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_servicio WHERE ser_codigo=" & RS1!ser_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Sector
Label1(2).Caption = "Importando Sector": DoEvents
RS1.Open "SELECT * FROM a_sector", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_sector WHERE sec_codigo=" & RS1!sec_codigo & ""
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Importando Estructura Servicio
Label1(2).Caption = "Importando Estructura Servicio": DoEvents
dbI.Execute "UPDATE a_estservicio SET ess_cencos='" & MuestraCasino(1) & "'"
dbI.Execute "UPDATE a_estservicio SET ess_racmin=0 WHERE ess_racmin is null"
RS1.Open "SELECT * FROM a_estservicio", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_estservicio WHERE ess_codser=" & RS1!ess_codser & " AND ess_codigo=" & RS1!ess_codigo & " AND ess_cencos='" & RS1!ess_cencos & "'"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Validar si existe planificación minutas
vg_db.BeginTrans
indice = 0
Label1(2).Caption = "Validar Planificación Minutas": DoEvents
RS1.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha, min_codreg, reg_nombre, min_codser, ser_nombre FROM b_minuta a, a_regimen b, a_servicio c WHERE a.min_cencos = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND a.min_codreg = b.reg_codigo AND a.min_codser = c.ser_codigo", dbI, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      RS2.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND VAL(MID(min_fecmin,1,6))=" & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & "", vg_db, adOpenStatic
      If Not RS2.EOF Then
         If MsgBox("Existe planificación minuta, desea borrar la información existente... " & VgLinea & VgLinea & "Regimen  : " & RS1!min_codreg & " " & Trim(RS1!reg_nombre) & VgLinea & "Servicio   :  " & RS1!min_codser & " " & Trim(RS1!ser_nombre), vbQuestion + vbYesNo, Msgtitulo) = vbYes Then
            '------- Borrar planificación contrato
            vg_db.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo WHERE b_minuta.min_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND VAL(MID(b_minuta.min_fecmin,1,6))=" & RS1!Fecha & " AND b_minuta.min_codreg=" & RS1!min_codreg & " AND b_minuta.min_codser=" & RS1!min_codser & ""
            vg_db.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos='" & LimpiaDato(Trim(fpText(0).text)) & "' AND VAL(MID(min_fecmin,1,6))=" & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & ""
         Else
            '------- Borrar planificación de la base carga
            dbI.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo WHERE b_minuta.min_cencos='" & MuestraCasino(1) & "'AND VAL(MID(b_minuta.min_fecmin,1,6))=" & RS1!Fecha & " AND b_minuta.min_codreg=" & RS1!min_codreg & " AND b_minuta.min_codser=" & RS1!min_codser & ""
            dbI.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos='" & MuestraCasino(1) & "' AND VAL(MID(min_fecmin,1,6))=" & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & ""
         End If
      End If
      RS2.Close: Set RS2 = Nothing
      '------- Traer ultimo correlativo
      If indice = 0 Then
         RS2.Open "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC", vg_db, adOpenStatic
         If Not RS2.EOF Then RS2.MoveFirst: indice = RS2!min_codigo + 1 Else indice = 1
         RS2.Close: Set RS2 = Nothing
      End If
      '------- actualizar correlativo planificación base externa
      RS2.Open "SELECT DISTINCT min_codigo, min_codreg FROM b_minuta WHERE min_cencos='" & MuestraCasino(1) & "' AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & "", dbI, adOpenStatic
      If Not RS2.EOF Then
         Do While Not RS2.EOF
            dbI.Execute "UPDATE b_minutadet SET mid_codigo=" & indice & " WHERE mid_codigo=" & RS2!min_codigo & ""
'            dbI.Execute "UPDATE b_minutadet SET mid_codigo=" & indice & ", mid_tiprec=" & RS2!min_codreg & " WHERE mid_codigo=" & RS2!min_codigo & ""
            dbI.Execute "UPDATE b_minuta SET min_codigo=" & indice & " WHERE min_cencos='" & MuestraCasino(1) & "' AND min_codigo=" & RS2!min_codigo & ""
            RS2.MoveNext: indice = indice + 1
         Loop
      End If
      RS2.Close: Set RS2 = Nothing
      '------- actualizar nro. raciones totales
      RS2.Open "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos='" & MuestraCasino(1) & "' AND sra_codser=" & RS1!min_codser & " GROUP BY sra_serdia ORDER BY sra_serdia", vg_db, adOpenStatic
      If Not RS2.EOF Then
         Do While Not RS2.EOF
            dbI.Execute "UPDATE b_minuta SET min_racteo=" & RS2!raciones & " WHERE min_cencos='" & MuestraCasino(1) & "' AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & " AND IIF(datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4))=1,7,datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)-1))=" & RS2!sra_serdia & ""
            RS2.MoveNext
         Loop
      End If
      RS2.Close: Set RS2 = Nothing
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
vg_db.CommitTrans

'------- Encabezado Planificación
Label1(2).Caption = "Importando Planificación Encabezado": DoEvents
RS1.Open "SELECT * FROM b_minuta", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minuta WHERE min_codigo=" & RS1!min_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Detalle Planificación
Label1(2).Caption = "Importando Planificación Detalle": DoEvents
RS1.Open "SELECT * FROM b_minutadet", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutadet WHERE mid_codigo=" & RS1!mid_codigo & " AND mid_tipmin='" & RS1!mid_tipmin & "' AND mid_numlin=" & RS1!mid_numlin & ""
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Costo Patron
Label1(2).Caption = "Importando Costo Patron": DoEvents
dbI.Execute "UPDATE b_costopatron SET cpa_cencos='" & MuestraCasino(1) & "'"
RS1.Open "SELECT * FROM b_costopatron", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_costopatron WHERE cpa_cencos='" & RS1!cpa_cencos & "' AND cpa_codreg=" & RS1!cpa_codreg & " AND cpa_codser=" & RS1!cpa_codser & " AND cpa_anomes=" & RS1!cpa_anomes & " AND cpa_descripcion='" & RS1!cpa_descripcion & "'"
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Gramos Familia Producto
Label1(2).Caption = "Importando Gramos Familia Producto": DoEvents
RS1.Open "SELECT * FROM b_gramofamproducto", dbI, adOpenStatic
Do While Not RS1.EOF
   ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_gramofamproducto WHERE gfp_cencos='" & RS1!gfp_cencos & "' AND gfp_codreg=" & RS1!gfp_codreg & " AND gfp_catdie=" & RS1!gfp_catdie & " AND gfp_tiprec=" & RS1!gfp_tiprec & " AND gfp_fampro=" & RS1!gfp_fampro & ""
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Tipo de Servicio
Label1(1).Caption = "Importando tipo de servicio"
DoEvents
RS1.Open "SELECT * FROM a_tiposervicio", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tiposervicio WHERE tis_codigo=" & RS1!tis_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

'------- Segmento
Label1(1).Caption = "Importando segmento"
DoEvents
RS1.Open "SELECT * FROM a_segmento", dbI, adOpenStatic
Do While Not RS1.EOF
    ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_segmento WHERE seg_codigo=" & RS1!seg_codigo
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

vg_db.BeginTrans
'------- Actualizar tabla lista producto y lista ingrediente
'* vg_db.Execute "INSERT INTO b_contlistprepro (cpp_cencos, cpp_codpro, cpp_upreco, cpp_fecuco, cpp_propon) SELECT '" & MuestraCasino(1) & "', pro_codigo, 0, null, 0 FROM b_productos WHERE pro_codigo NOT IN (SELECT DISTINCT cpp_codpro FROM b_contlistprepro WHERE cpp_cencos='" & MuestraCasino(1) & "')"
vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & MuestraCasino(1) & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente WHERE ing_codigo NOT IN (SELECT DISTINCT cpi_coding FROM b_contlistpreing WHERE cpi_cencos='" & MuestraCasino(1) & "')"
vg_db.CommitTrans

dbI.Close: Set dbI = Nothing
'vg_db.BeginTrans
fso.DeleteFile cdbi
'vg_db.Execute "insert into log_actualizacion values ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', cdate('" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm") & "'))"
ImportarDatosAccess = True
'vg_db.CommitTrans
'------- Rutina validar producto vigente
ValidarProductoVigente
Label1(2).Visible = False: PB.Visible = False

Exit Function
Man_Error:
If Err = -2147217865 Or Err = 3265 Then
   dbI.Close: Set dbI = Nothing
'    vg_db.BeginTrans
'    fso.DeleteFile cDBI
'    vg_db.Execute "insert into log_actualizacion values ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', cdate('" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm") & "'))"
   ImportarDatosAccess = True
'    vg_db.CommitTrans
   Exit Function
ElseIf Err = -2147467259 Then
   MsgBox "Archivo no corresponde a un MDB, proceso cancelado...", vbInformation + vbOKOnly, Msgtitulo
   ImportarDatosAccess = False
   Exit Function
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Function
vg_db.RollbackTrans
If Err.Number = -2147467259 Then
    MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Function

Private Function ActRegistro(RSO As Fields, DBO As ADODB.Connection, DBD As ADODB.Connection, cSql As String)
Dim RS2 As New ADODB.Recordset, i As Long, bAdd As Boolean
On Error GoTo ManError
DoEvents
bAdd = False
RS2.Open cSql, DBD, adOpenDynamic, adLockOptimistic
vg_db.BeginTrans
If RS2.EOF Then
    bAdd = True
    RS2.AddNew
End If
For i = 0 To RSO.count - 1
    If bAdd Or (RS2.Fields(i).Name <> "pro_upreco" And RS2.Fields(i).Name <> "pro_fecuco" And RS2.Fields(i).Name <> "pro_propon" And RS2.Fields(i).Name <> "ing_precos" And RS2.Fields(i).Name <> "ing_feccos" And RS2.Fields(i).Name <> "rec_tiprec" And RS2.Fields(i).Name <> "ing_codcom" And RS2.Fields(i).Name <> "ing_codped" And RS2.Fields(i).Name <> "ess_codsec") Then
        Select Case RS2.Fields(i).Type
        Case adChar, adVarChar, adVarWChar
            If TipoDato(RS2.Fields(i).Value, "") <> Trim(TipoDato(RSO.Item(i).Value, "")) Then RS2.Fields(i).Value = IIf(Trim((RSO.Item(i).Value)) = "", " ", Trim(RSO.Item(i).Value))
        Case Else
            If TipoDato(RS2.Fields(i).Value, 0) <> TipoDato(RSO.Item(i).Value, 0) Or RS2.Fields(i).Name = "red_tiprec" Or RS2.Fields(i).Name = "mid_tiprec" Or RS2.Fields(i).Name = "mid_racteo" Or RS2.Fields(i).Name = "mid_racrea" Then
               RS2.Fields(i).Value = RSO.Item(i).Value
            ElseIf RS2.Fields(i).Name = "mid_cosrec" Then
               RS2.Fields(i).Value = fg_CalCtoRecInv(RSO.Item(i - 3).Value, RSO.Item(i + 2).Value, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")))
            ElseIf RS2.Fields(i).Name = "mid_cosdes" Then
               RS2.Fields(i).Value = fg_CalCtoRecInv(RSO.Item(i - 8).Value, RSO.Item(i - 3).Value, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))) 'fg_CalCtoRecInv(RSO.Item(i - 3).Value, RSO.Item(i + 2).Value, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")))
            End If
        End Select
    ElseIf RS2.Fields(i).Name = "rec_tiprec" And RSO.Item(i).Value > 0 Then
        RS2.Fields(i).Value = RSO.Item(i).Value
    ElseIf RS2.Fields(i).Name = "ess_codsec" And RSO.Item(i).Value > 0 Then
        RS2.Fields(i).Value = RSO.Item(i).Value
    End If
Next
RS2.Update
RS2.Close: Set RS2 = Nothing
vg_db.CommitTrans
Exit Function
ManError:
If Err.Number = -2147217887 Then Resume Next
MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
End Function

Sub Inicio(mtit As String, op As String)
Msgtitulo = mtit
Me.Caption = mtit
opcion = op
If op = "I" Then
   Frame1(1).Visible = False
   Frame1(2).Visible = False
   Label1(0).Visible = False
   fpDateTime1.Visible = False
   Label1(1).Caption = "Ruta Imp."
End If
End Sub
