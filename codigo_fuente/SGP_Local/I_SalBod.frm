VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_SalBod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida y Devolución de Producción"
   ClientHeight    =   3435
   ClientLeft      =   3465
   ClientTop       =   3870
   ClientWidth     =   8430
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
   ScaleHeight     =   3435
   ScaleWidth      =   8430
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
      Height          =   2970
      Left            =   0
      TabIndex        =   6
      Top             =   375
      Width           =   8430
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   30
         Index           =   0
         Left            =   5520
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   255
         _Version        =   393216
         _ExtentX        =   450
         _ExtentY        =   53
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
         MaxRows         =   0
         SpreadDesigner  =   "I_SalBod.frx":0000
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Servicio"
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   0
         Left            =   4440
         TabIndex        =   17
         Top             =   1560
         Width           =   3735
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   255
            Index           =   1
            Left            =   2280
            TabIndex        =   19
            Top             =   300
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   3000
            Picture         =   "I_SalBod.frx":0274
            Top             =   165
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Regimen"
         ForeColor       =   &H80000008&
         Height          =   735
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   3735
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   16
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Lista"
            Height          =   255
            Index           =   3
            Left            =   2280
            TabIndex        =   15
            Top             =   300
            Width           =   735
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   3
            Left            =   3000
            Picture         =   "I_SalBod.frx":057E
            Top             =   165
            Width           =   480
         End
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H000000FF&
         Caption         =   "Salto Página"
         Height          =   255
         Left            =   6720
         TabIndex        =   13
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         Left            =   1980
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   675
         Width           =   5175
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1980
         TabIndex        =   0
         Top             =   330
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2010
         TabIndex        =   3
         Top             =   1070
         Width           =   1515
         _Version        =   196608
         _ExtentX        =   2672
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
         ButtonStyle     =   3
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
         DateCalcMethod  =   3
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
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483637
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
         Left            =   5550
         TabIndex        =   4
         Top             =   1065
         Width           =   1635
         _Version        =   196608
         _ExtentX        =   2884
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
         ButtonStyle     =   3
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
         DateCalcMethod  =   3
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
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   2
         Left            =   1995
         TabIndex        =   12
         Top             =   735
         Width           =   5205
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Informe"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   11
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   9
         Top             =   420
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3345
         Picture         =   "I_SalBod.frx":0888
         Top             =   240
         Width           =   480
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
         Top             =   345
         Width           =   4335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   8
         Top             =   1125
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha Termino"
         Height          =   285
         Index           =   4
         Left            =   4110
         TabIndex        =   7
         Top             =   1125
         Width           =   1395
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3840
         TabIndex        =   10
         Top             =   390
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8430
      _ExtentX        =   14870
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   30
      Index           =   1
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Visible         =   0   'False
      Width           =   255
      _Version        =   393216
      _ExtentX        =   450
      _ExtentY        =   53
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
      MaxRows         =   0
      SpreadDesigner  =   "I_SalBod.frx":0B92
   End
End
Attribute VB_Name = "I_SalBod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String
Dim est       As Boolean
Dim codreg    As String

Private Sub Combo1_Click(Index As Integer)

Toolbar1.Buttons(3).Enabled = False

Select Case Index

Case 2
    
    Select Case Combo1(2).ListIndex
    
    Case 0, 1, 2, 3
        
        Check1.Visible = True
         
        Toolbar1.Buttons(3).Enabled = IIf(Combo1(2).ListIndex = 2, True, False)

    Case 4, 5, 6
        
        Check1.Visible = False
    
    End Select

End Select

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Form_Load()

Me.Width = 8520
Me.Height = 3915

est = True
MsgTitulo = "Salida de Bodega a Producción"
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
fg_centra Me
Me.HelpContextID = vg_OpcM

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Excel": BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_ExploradorWindows", , tbrDefault, "A_ExploradorWindows")
BtnX.Visible = True
BtnX.ToolTipText = "Ver Carpeta"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

'----------------------------Valida permisos para impresión
Toolbar1.Buttons.Item(1).Visible = IIf(Val(Mid(ValidarUsuario(Me), 4, 1)) = 1, True, False)
'-------------------------Asigna tipo de informe-------------
With Combo1(2)
    
    .Clear
    .AddItem "Formato de Requisición Resumido"
    .AddItem "Formato de Requisición x Sector"
    .AddItem "Formato de Requisición x Estructura Servicio Detallado"
    .AddItem "Formato de Requisición x Estructura Servicio Resumido"
    .AddItem "Resumen de Salida a Bodega"
    .AddItem "Devolución de Salida a Bodega"
    .AddItem "Salida Menos Devoluciones a Bodega"
    .ListIndex = 0

End With
'-------------------------Fin Asigna tipo de informe---------
'-------------------------Asigna fecha actual del sistema para informe-------------
fpDateTime1(0).text = Date
fpDateTime1(1).text = Date
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
est = False
codreg = ""
MoverDatoGrilla

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
If IsDate(fpDateTime1(0).text) Then
    
    If CDate(fpDateTime1(0).text) > CDate(fpDateTime1(1).text) Then fpDateTime1(1).text = fpDateTime1(0).text: Exit Sub

End If

Select Case Index

Case 0
    
    If fpDateTime1(0).text = "" Then
        
        fpDateTime1(1).Enabled = False
        fpDateTime1(1).text = ""
        Exit Sub
    
    Else
        
        fpDateTime1(1).Enabled = True
    
    End If

End Select

'codreg = ""
'MoverDatoGrilla
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub fpText1_Change(Index As Integer)

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText1(1).text)), ""), vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
fpayuda(1).Caption = Trim(RS!cli_nombre)
RS.Close
Set RS = Nothing

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Image1_Click(Index As Integer)

vg_codigo = 0

Select Case Index

Case 1
    
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    If Combo1(0).Enabled = True Then Combo1(0).SetFocus

Case 2
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    If fpText1(1).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Servicio", Me.vaSpread1, fpText1(1).text, codreg, Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Combo1(2).ListIndex = 0 Or Combo1(2).ListIndex = 1 Or Combo1(2).ListIndex = 2 Or Combo1(2).ListIndex = 3, "0", "0"), "", 1, "'2'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub

Case 3
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    If fpText1(1).text = "" Then Exit Sub
    B_MTaEst.LlenaDatos "Regimen", Me.vaSpread1, fpText1(1).text, "", Val(Format(fpDateTime1(0).text, "yyyymmdd")), Val(Format(fpDateTime1(1).text, "yyyymmdd")), IIf(Combo1(2).ListIndex = 0 Or Combo1(2).ListIndex = 1 Or Combo1(2).ListIndex = 2 Or Combo1(2).ListIndex = 3, "0", "0"), "", 0, "'2'"
    B_MTaEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub

End Select

End Sub

Private Sub Option1_Click(Index As Integer)

Select Case Index

Case 0
    
    Image1(2).Enabled = False
    For i = 1 To vaSpread1(1).MaxRows
        
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1: vaSpread1(1).text = "1"
    
    Next i

Case 1
    
    Image1(2).Enabled = True

Case 2
    
    Image1(3).Enabled = False
    For i = 1 To vaSpread1(0).MaxRows
        
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1: vaSpread1(0).text = "1"
    
    Next i

Case 3
    
    Image1(3).Enabled = True

End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim codser             As String
Dim RS                 As New ADODB.Recordset
Dim MyBufferServicio   As String
Dim MyBufferRegimen    As String
Dim NombreArchivoExcel As String

On Error GoTo Error_Salir

Select Case Button.Index

Case 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText1(1).text)), " "), vg_db, adOpenStatic
    If RS.EOF Then
    
       RS.Close
       Set RS = Nothing
       fpText1(1).text = ""
       fpayuda(1).Caption = ""
       MsgBox "No existe contrato", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    RS.Close
    Set RS = Nothing
    
    If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 4, 2) <> Mid(fpDateTime1(1).text, 4, 2) Then MsgBox "Mes origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Mid(fpDateTime1(0).text, 7, 4) <> Mid(fpDateTime1(1).text, 7, 4) Then MsgBox "Año origen mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    Dim nroser As Long
    codser = ""
    nroser = 0
    codreg = ""
    codser = ""
    
    Let MyBufferRegimen = ""
    Let MyBufferRegimen = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferRegimen = MyBufferRegimen & "<Regimen>"
    
    For i = 1 To vaSpread1(0).MaxRows
        
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then
           
           vaSpread1(0).Col = 2
           codreg = codreg & "" & vaSpread1(0).text & ","
    
           MyBufferRegimen = MyBufferRegimen & " <Reg"
           MyBufferRegimen = MyBufferRegimen & " Reg = " & Chr(34) & vaSpread1(0).text & Chr(34)
           Let MyBufferRegimen = MyBufferRegimen & "/>"
        
        End If
        
    Next i
    Let MyBufferRegimen = MyBufferRegimen & "</Regimen>"
    
    Let MyBufferServicio = ""
    Let MyBufferServicio = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferServicio = MyBufferServicio & "<Servicio>"
    
    For i = 1 To vaSpread1(1).MaxRows
        
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then
        
           vaSpread1(1).Col = 2
           codser = codser & "" & vaSpread1(1).text & ","
           nroser = nroser + 1
    
           MyBufferServicio = MyBufferServicio & " <Ser"
           MyBufferServicio = MyBufferServicio & " Ser = " & Chr(34) & vaSpread1(1).text & Chr(34)
           Let MyBufferServicio = MyBufferServicio & "/>"
    
        End If
        
    Next i
    
    Let MyBufferServicio = MyBufferServicio & "</Servicio>"
    
    If Trim(codreg) = "" Then fg_descarga: MsgBox "Regimen debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(codser) = "" Then fg_descarga: MsgBox "Servicio debe ser informado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    
    Select Case Combo1(2).ListIndex
    
        Case 0, 1, 2, 3
        
            'Validar si existe información
               
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
        
            Set RS = vg_db.Execute("sgp_Sel_ValidarServicioComensalesCeroSalProduccion '" & MyBufferServicio & "', '" & MyBufferRegimen & "', '" & Trim(fpText1(1).text) & "', " & Format(Trim(fpDateTime1(0).text), "yyyymmdd") & ", " & Format(Trim(fpDateTime1(1).text), "yyyymmdd") & "")
               
            If Not RS.EOF Then
                        
                  fg_descarga
                  
                  MsgBox "Falta ingreso comensales totales, se generará archivo Excel con información faltante... ", vbInformation + vbOKOnly, MsgTitulo
                               
                  'Exportar Excel
               
                  '-------> Crear directorio FormatoRequisicion
                  If Dir(dir_trabajo_Inf & "\" & "ExcelSGP", vbDirectory) = "" Then
               
                     MkDir dir_trabajo_Inf & "\" & "ExcelSGP"
                  
                  End If
                  '-------> Fin crear directorio Excel Versión
        
                  NombreArchivoExcel = "SalidaBodegaSerComensalesValCero_" & Trim(fpText1(1).text) & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "HHMMSS")
                  Generar_ArchivoExcel "sgp_Sel_ValidarServicioComensalesCeroSalProduccion '" & MyBufferServicio & "', '" & MyBufferRegimen & "', '" & Trim(fpText1(1).text) & "', " & Format(Trim(fpDateTime1(0).text), "yyyymmdd") & ", " & Format(Trim(fpDateTime1(1).text), "yyyymmdd") & "", dir_trabajo_Inf & "ExcelSGP\", NombreArchivoExcel
                  
                  fg_descarga
                   
'                  MsgBox "Proceso generación exitosamente, archivo excel", vbExclamation + vbOKOnly, MsgTitulo
            
            End If
            RS.Close
            Set RS = Nothing
    
    End Select
    
    Select Case Combo1(2).ListIndex
        
        '---------Formato Requisición------
        Case 0
        
            I_SalBodega Trim(fpText1(1).text), MyBufferRegimen, MyBufferServicio, Trim(fpDateTime1(0).text), Trim(fpDateTime1(1).text), IIf(Check1.Value = 1, True, False)
        '------- Formato de requesición x sector
    
        Case 1
        
            I_SalBodegaSector Trim(fpText1(1).text), MyBufferRegimen, MyBufferServicio, Trim(fpDateTime1(0).text), Trim(fpDateTime1(1).text), 1, IIf(Check1.Value = 1, True, False)
        '---------Formato Requisición x Estructura Servicio Detallado------
    
        Case 2
        
            I_SalBodegaDet Trim(fpText1(1).text), MyBufferRegimen, MyBufferServicio, Trim(fpDateTime1(0).text), Trim(fpDateTime1(1).text), 0, IIf(Check1.Value = 1, True, False)
    '        I_SalBodegaxEst cencos, codreg, codser, fecini, fecter, 0
    
        Case 3

        '---------Formato Requisición x Estructura Servicio Resumido------
            I_SalBodegaxEst Trim(fpText1(1).text), MyBufferRegimen, MyBufferServicio, Trim(fpDateTime1(0).text), Trim(fpDateTime1(1).text), 1, IIf(Check1.Value = 1, True, False)
        '---------Resto de Informes ------
    
        Case 4, 5, 6
        
            I_SalidasDevolBod Trim(fpText1(1).text), codreg, codser, Trim(fpDateTime1(0).text), Trim(fpDateTime1(1).text)
    
    End Select

Case 3 ' ValidarProductosVigente & Grabar & Exportar
      
       If fpDateTime1(0).Value > fpDateTime1(1).Value Then MsgBox "Fecha origen Mayor destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       
       fg_carga ""
       
       Let MyBufferServicio = ""
       Let MyBufferServicio = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
       Let MyBufferServicio = MyBufferServicio & "<Servicio>"
   
       For i = 1 To vaSpread1(1).MaxRows
       
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
       
           If vaSpread1(1).text = "1" Then
          
              vaSpread1(1).Col = 2
              MyBufferServicio = MyBufferServicio & " <Ser"
              MyBufferServicio = MyBufferServicio & " Ser = " & Chr(34) & vaSpread1(1).text & Chr(34)
              Let MyBufferServicio = MyBufferServicio & "/>"
       
           End If
   
       Next i
       Let MyBufferServicio = MyBufferServicio & "</Servicio>"

       Let MyBufferRegimen = ""
       Let MyBufferRegimen = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
       Let MyBufferRegimen = MyBufferRegimen & "<Regimen>"
   
       For i = 1 To vaSpread1(0).MaxRows
       
           vaSpread1(0).Row = i
           vaSpread1(0).Col = 1
       
           If vaSpread1(0).text = "1" Then
          
              vaSpread1(0).Col = 2
              MyBufferRegimen = MyBufferRegimen & " <Reg"
              MyBufferRegimen = MyBufferRegimen & " Reg = " & Chr(34) & vaSpread1(0).text & Chr(34)
              Let MyBufferRegimen = MyBufferRegimen & "/>"
       
           End If
   
       Next i
       Let MyBufferRegimen = MyBufferRegimen & "</Regimen>"

       'Validar si existe información
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

       Set RS = vg_db.Execute("sgp_Sel_formatorequesicionestdetallado '" & MyBufferServicio & "', '" & MyBufferRegimen & "', '" & Trim(fpText1(1).text) & "', " & vg_codbod & " , " & Format(Trim(fpDateTime1(0).text), "yyyymmdd") & ", " & Format(Trim(fpDateTime1(1).text), "yyyymmdd") & "")
       
       If RS.EOF Then
                
          fg_descarga
          
          MsgBox "No existe información, con los parametros indicados.. ", vbInformation + vbOKOnly, MsgTitulo
             
          RS.Close
          Set RS = Nothing
          Exit Sub
          
       End If
       RS.Close
       Set RS = Nothing
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
    
       Set RS = vg_db.Execute("sgp_Upd_ValidarProductoVigente '" & Trim(fpText1(1).text) & "', " & vg_codbod & "")

       If Not RS.EOF Then
       
          If RS(0) > 0 Then
                       
             fg_descarga
             
             MsgBox RS(0) & " " & RS(1) & " Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
             
             RS.Close
             Set RS = Nothing
             
             Exit Sub
   
          End If

       
       End If
       RS.Close
       Set RS = Nothing
       
       'Grabar Solicitud
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("sgp_DelIns_formatorequesicionestdetallado '" & MyBufferServicio & "', '" & MyBufferRegimen & "', '" & Trim(fpText1(1).text) & "', " & vg_codbod & " , " & Format(Trim(fpDateTime1(0).text), "yyyymmdd") & ", " & Format(Trim(fpDateTime1(1).text), "yyyymmdd") & ", '" & vg_NUsr & "'")

       If Not RS.EOF Then
       
          If RS(0) > 0 Then
                       
             fg_descarga
             
             MsgBox RS(0) & " " & RS(1) & " Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
             
             RS.Close
             Set RS = Nothing
             
             Exit Sub
   
          End If
       
       End If
       RS.Close
       Set RS = Nothing
       
       'Exportar Excel
       
       '-------> Crear directorio FormatoRequisicion
       If Dir(dir_trabajo_Inf & "\" & "FormatoRequisicion", vbDirectory) = "" Then
       
          MkDir dir_trabajo_Inf & "\" & "FormatoRequisicion"
          
       End If
       '-------> Fin crear directorio Excel Versión

       NombreArchivoExcel = "FormatoRequisicion_" & Trim(fpText1(1).text) & "_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "HHMMSS")
       Generar_ArchivoExcel "sgp_Sel_formatorequesicionestdetallado '" & MyBufferServicio & "', '" & MyBufferRegimen & "', '" & Trim(fpText1(1).text) & "', " & vg_codbod & " , " & Format(Trim(fpDateTime1(0).text), "yyyymmdd") & ", " & Format(Trim(fpDateTime1(1).text), "yyyymmdd") & "", dir_trabajo_Inf & "FormatoRequisicion\", NombreArchivoExcel
          
       fg_descarga
           
       MsgBox "Proceso generación exitosamente", vbExclamation + vbOKOnly, MsgTitulo

Case 5

       ExplorarCarpeta dir_trabajo_Inf & "FormatoRequisicion"

Case 7
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Error_Salir:
    
    
    If RS.State = 1 Then RS.Close
    
    fg_descarga
    
    MsgBox "Error:" & Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Resume Next

End Sub

Sub MoverDatoGrilla()

Dim RS     As New ADODB.Recordset

fg_carga ""
With vaSpread1(0)

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open RutinaLectura.Regimen(1, 0, ""), vg_db, adOpenStatic
    
    .MaxRows = 0
    
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS!reg_codigo
          .Col = 3: .text = Trim(RS!reg_nombre)
          
          RS.MoveNext
       
       Loop
    End If
    RS.Close
    Set RS = Nothing

End With

With vaSpread1(1)
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Servicio(1, 0, ""), vg_db, adOpenStatic
    .MaxRows = 0
    
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          .Col = 1: .text = "1"
          .Col = 2: .text = RS!ser_codigo
          .Col = 3: .text = Trim(RS!ser_nombre)
          
          RS.MoveNext
       
       Loop
    
    End If
    RS.Close
    Set RS = Nothing

End With

fg_descarga
End Sub

