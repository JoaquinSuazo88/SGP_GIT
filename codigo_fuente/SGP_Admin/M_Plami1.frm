VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Plami1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3180
   ClientLeft      =   5835
   ClientTop       =   4935
   ClientWidth     =   8355
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
   ScaleHeight     =   3180
   ScaleWidth      =   8355
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3045
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   60
      Width           =   7515
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   0
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1880
         Width           =   3480
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Index           =   1
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1485
         Width           =   1800
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   765
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   2
         Top             =   1125
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
         NoSpecialKeys   =   2
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Top             =   2295
         Width           =   1155
         _Version        =   196608
         _ExtentX        =   2037
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
         Text            =   "07/2023"
         DateCalcMethod  =   2
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   420
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
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   1960
         TabIndex        =   24
         Top             =   1980
         Width           =   3495
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1960
         TabIndex        =   23
         Top             =   1580
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Zona"
         Height          =   225
         Index           =   0
         Left            =   195
         TabIndex        =   22
         Top             =   1985
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo  Planificación"
         Height          =   225
         Index           =   22
         Left            =   165
         TabIndex        =   21
         Top             =   1555
         Width           =   1815
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1920
         TabIndex        =   19
         Top             =   2640
         Width           =   4935
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   6840
         Picture         =   "M_Plami1.frx":0000
         Top             =   2580
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Lista Precio"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   18
         Top             =   2720
         Width           =   1020
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3285
         TabIndex        =   14
         Top             =   420
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3285
         TabIndex        =   13
         Top             =   765
         Width           =   3975
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3285
         TabIndex        =   12
         Top             =   1125
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   11
         Top             =   2340
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
         Height          =   195
         Index           =   3
         Left            =   195
         TabIndex        =   10
         Top             =   1170
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   9
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Segmento"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   480
         Width           =   1245
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2820
         Picture         =   "M_Plami1.frx":030A
         Top             =   345
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2820
         Picture         =   "M_Plami1.frx":0614
         Top             =   675
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2820
         Picture         =   "M_Plami1.frx":091E
         Top             =   1035
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3315
         TabIndex        =   15
         Top             =   465
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3315
         TabIndex        =   16
         Top             =   810
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3315
         TabIndex        =   17
         Top             =   1170
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   1950
         TabIndex        =   20
         Top             =   2700
         Width           =   4935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3180
      Left            =   7725
      TabIndex        =   6
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   5609
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Plami1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim NomFor As String
Dim Est As Boolean
Dim OpUsuario As String

Private Sub Combo2_Click(Index As Integer)
If Est Then Exit Sub
Select Case Index
Case 0
    vg_Zona = Val(fg_codigocbo(Combo2, 0, 10, ""))
Case 1
    vg_IndpprSelec = fg_codigocbo(Combo2, 1, 1, "") 'fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
End Select
MoverListaPrecio
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
fg_carga ""
Est = True
Me.HelpContextID = vg_OpcM
Me.Height = 3660
Me.Width = 8445
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar " ': btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Planificación"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1.text = Format(Date, "mm/yyyy")
vg_Zona = 1
vg_codlpr = 0
vg_IndpprSelec = 0
fg_descarga

OpUsuario = vg_Indppr
If IsNull(OpUsuario) Or Trim(OpUsuario) = "" Then
    MsgBox "Contactese con el Administrador del Sistema...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
Else
    Select Case OpUsuario
    Case "1"
        Combo2(1).Clear
        Combo2(1).AddItem "Real" & Space(150) & "(1)"
        Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
        vg_IndpprSelec = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
    Case "2"
        Combo2(1).Clear
        Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
        Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
        vg_IndpprSelec = fg_buscacbo(Combo2, 1, 1, fg_pone_cero(Str(OpUsuario), 1))
    Case "3"
        Combo2(1).Clear
        Combo2(1).AddItem "Real" & Space(150) & "(1)"
        Combo2(1).AddItem "Propuesta" & Space(150) & "(2)"
        Combo2(1).ListIndex = 0
        vg_IndpprSelec = 1
    End Select
End If

'------> Llenado primer combo Sub-Segmento
Set RS = vg_db.Execute("SELECT * FROM a_zona where zon_activo = '1'")

If Not RS.EOF Then
   
   Combo2(0).Clear
   
   Do While Not RS.EOF
      
      Combo2(0).AddItem Trim(RS!Zon_nombre) & Space(150) & "(" & fg_pone_cero(Str(RS!zon_codigo), 10) & ")"
      RS.MoveNext
   
   Loop
   RS.Close
   Set RS = Nothing
   Combo2(0).ListIndex = 0

End If

Est = False

End Sub

Private Sub fpDateTime1_Change()
If IsDate(fpDateTime1.text) = False Then Exit Sub
MoverListaPrecio
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(0).Caption = "": Exit Sub
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & " and sub_indppr='" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing:: fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    MoverListaPrecio
Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(1).Caption = "": Exit Sub
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & " and reg_indppr='" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverListaPrecio
Case 2
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(2).Caption = "": Exit Sub
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & " and ser_indppr='" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "":  Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverListaPrecio
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 0 Then Image1_Click 0
    If Index = 1 Then Image1_Click 1
    If Index = 2 Then Image1_Click 2
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
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    fpLongInteger1(1).SetFocus
Case 1
'    vg_opayuda = 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre
    fpDateTime1.SetFocus
Case 3
    vg_left = Image1(0).Left + 1920
    B_TabEst.LlenaDatos "b_listaprecio", "lpr_", "Lista de Precio", "LisPre"
    B_TabEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    vg_codlpr = Val(vg_codigo)
    Set RS = vg_db.Execute("sgpadm_s_listaprecio 4, " & Val(vg_codigo) & ", 0, '" & vg_NUsr & "'")
    If Not RS.EOF Then fpayuda(3).Caption = Trim(vg_codigo) & " - " & Trim(vg_nombre) & " - " & Mid(RS!dlp_anomes, 5, 2) & "/" & Mid(RS!dlp_anomes, 1, 4)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim tipreg As String, tipser As String, tipsus As String
vg_IndpprSelec = 0
Select Case Button.Index
Case 2 '-------> Acceso planificación minuta
'    If Trim(fpayuda(3).Caption) = "" Then MsgBox "Debe seleccionar una lista de precio", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub

If vg_Indppr = 1 Or vg_Indppr = 2 Then
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & " And sub_indppr = " & vg_Indppr & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(0).Caption = RS!sub_nombre: tipsus = RS!sub_indppr: RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & " And reg_indppr = " & vg_Indppr & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(1).Caption = RS!reg_nombre: tipreg = RS!reg_indppr: RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & "  And ser_indppr = " & vg_Indppr & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(2).Caption = RS!ser_nombre: tipser = RS!ser_indppr: RS.Close: Set RS = Nothing
    RS.Open "SELECT a_estservicio.* FROM a_estservicio With(NoLock) WHERE ess_codser=" & Val(fpLongInteger1(2).Value) & " ORDER BY ess_orden", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe estructura de servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
Else
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & " ", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(0).Caption = RS!sub_nombre: tipsus = RS!sub_indppr: RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & " ", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(1).Caption = RS!reg_nombre: tipreg = RS!reg_indppr: RS.Close: Set RS = Nothing
    RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & " ", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Servicio", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fpayuda(2).Caption = RS!ser_nombre: tipser = RS!ser_indppr: RS.Close: Set RS = Nothing
'    RS.Open "SELECT a_estservicio.* With(NoLock) FROM a_estservicio WHERE ess_codser=" & Val(fpLongInteger1(2).Value) & " ORDER BY ess_orden", vg_db, adOpenStatic
'    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe estructura de servicio", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    RS.Close: Set RS = Nothing
End If
    
    
If fg_codigocbo(Combo2, 1, 1, "") <> tipsus Or fg_codigocbo(Combo2, 1, 1, "") <> tipreg Or fg_codigocbo(Combo2, 1, 1, "") <> tipser Then MsgBox "Tipo planificación no coincide con los códigos Sub-Segmento, Regimen o Servicio ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
'If M_Plami2.ValidaMinuta(M_Plami1.fpLongInteger1(0).text, M_Plami1.fpLongInteger1(1).text, M_Plami1.fpLongInteger1(2).text, ExraeCodCombo(M_Plami1.Combo2(1).text), ExraeCodCombo(M_Plami1.Combo2(0).text), ExtraeFecha(M_Plami1.fpDateTime1)) = False Then Exit Sub
'*****************---->Validar minuta en uso <---------------------------
'------ Esta funcion crea una tabla temporal concatenando los parametros ingresaods
'------ para la minuta, de esta manera permanece una tabla temporal identificando
'------ que alguien se encuentra conectado a esa minuta, si alguien
'------ mas quiere acceder, se dara un aviso que esta en uso
'------ esta tabla temporal se destruye cuando se cierra este formulario (evento Unload)
'------ y tambien si el usuario cierra la sesion SQL Server la destruye automaticamente.
'----------------------------------------------------------------------
    
    Dim RSTempCheck As New ADODB.Recordset
    Dim RSTem As New ADODB.Recordset
    Dim RSinsert As New ADODB.Recordset
    Dim NameTemp As String
    NameTemp = fpLongInteger1(0).text & fpLongInteger1(1).text & fpLongInteger1(2).text & ExraeCodCombo(M_Plami1.Combo2(1).text) & ExraeCodCombo(M_Plami1.Combo2(0).text) & ExtraeFecha(M_Plami1.fpDateTime1)

    Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaMinuta_" & NameTemp & "'")
    
    If RSTempCheck.EOF And RSTempCheck.BOF Then
        Set RSTem = vg_db.Execute("CREATE TABLE ##ValidaMinuta_" & NameTemp & " (usu_codigo VarChar(20))")
        Set RSinsert = vg_db.Execute("INSERT INTO ##ValidaMinuta_" & NameTemp & " (usu_codigo) values ('" & vg_NUsr & "')")
'        Set RS = Nothing
'        Set RSTem = Nothing
    Else
        Set RS = vg_db.Execute("SELECT usu_codigo from ##ValidaMinuta_" & NameTemp & " ")
        If Not (RS.EOF = True And RS.BOF = True) Then
            RS.MoveFirst
            MsgBox "La minuta con los parametros ingresados, actualmente esta siendo usada por el usuario: '" & UCase(RS!usu_codigo) & "', podra ingresar cuando el usuario termine de trabajar en ella"
            RS.Close: Set RS = Nothing
            Exit Sub
        End If
        RS.Close: Set RS = Nothing
    End If

'RSTempCheck.Close
'Set RSTempCheck = Nothing
       
    vg_codsubseg = Val(fpLongInteger1(0).Value):   vg_nomsubseg = fpayuda(0)
    vg_codregimen = Val(fpLongInteger1(1).Value):  vg_nomreg = fpayuda(1)
    vg_codservicio = Val(fpLongInteger1(2).Value): vg_nomser = fpayuda(2)
    vg_fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
    vg_IndpprSelec = Val(fg_codigocbo(Combo2, 1, 1, ""))
    Dim TipMin As String
    
    If NomFor = "MINTEO" Then TipMin = "1" Else TipMin = "2"
    
    RS.Open "SELECT DISTINCT b_minuta.min_subseg " & _
            "FROM  b_minuta With(NoLock), b_minutadet With(NoLock) " & _
            "WHERE b_minuta.min_codigo=b_minutadet.mid_codigo " & _
            "AND   b_minuta.min_subseg=" & Val(fpLongInteger1(0).Value) & " " & _
            "AND   b_minuta.min_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
            "AND   b_minuta.min_codser=" & Val(fpLongInteger1(2).Value) & " " & _
            "AND   substring(convert(char(8),b_minuta.min_fecmin),1,6)=" & Val(vg_fecha) & " " & _
            "AND   min_Indppr='" & Val(fg_codigocbo(Combo2, 1, 1, "")) & "' " & _
            "AND   b_minutadet.mid_tipmin='" & TipMin & "'", vg_db, adOpenStatic
    If RS.EOF Then
      'Nueva Minuta
      RS.Close: Set RS = Nothing
      RS.Open "SELECT DISTINCT b_minuta.min_subseg " & _
            "FROM  b_minuta With(NoLock), b_minutadet With(NoLock) " & _
            "WHERE b_minuta.min_codigo=b_minutadet.mid_codigo " & _
            "AND   b_minuta.min_subseg=" & Val(fpLongInteger1(0).Value) & " " & _
            "AND   b_minuta.min_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
            "AND   b_minuta.min_codser=" & Val(fpLongInteger1(2).Value) & " " & _
            "AND   substring(convert(char(8),b_minuta.min_fecmin),1,6)=" & Val(vg_fecha) & " " & _
            "AND   b_minutadet.mid_tipmin='" & TipMin & "'", vg_db, adOpenStatic
    End If
            
    If Not RS.EOF Then
       RS.Close: Set RS = Nothing
    ElseIf RS.EOF Then
       RS.Close: Set RS = Nothing
       If (Mid(ValidarUsuario(Me), 1, 1)) = "0" Then MsgBox "No esta autorizado crear planificación, conctatece con el administrador de sistema ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       'MsgBox "Sub-segmento no tiene registros de tipo: " + Combo2(1).text, vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    Toolbar1.Enabled = False
    If NomFor = "MINTEO" Then Unload M_Plami2: M_Plami2.Show 1
    DropTebleTmp (fpLongInteger1(0).text & fpLongInteger1(1).text & fpLongInteger1(2).text & ExraeCodCombo(Combo2(1).text) & ExraeCodCombo(Combo2(0).text) & ExtraeFecha(fpDateTime1))
    
    Toolbar1.Enabled = True
'    If NomFor = "MINREA" Then
'        RS.Open "SELECT count(b_minutadet.mid_codigo) AS nreg From b_minutadet, b_minuta WHERE " & _
'                "b_minuta.min_codigo=b_minutadet.mid_codigo And val(mid(b_minuta.min_fecmin,1,6))=" & Val(vg_fecha) & " And b_minutadet.mid_tipmin='1'", vg_db, adOpenStatic
'        If RS!NReg = 0 Then RS.Close: Set RS = Nothing: MsgBox "Debe realizar la planificación teórica de este mes...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'        RS.Close: Set RS = Nothing
'        RS.Open "SELECT count(b_minutadet.mid_codigo) AS nreg From b_minutadet, b_minuta WHERE " & _
'                "b_minuta.min_codigo=b_minutadet.mid_codigo And val(mid(b_minuta.min_fecmin,1,6))=" & Val(vg_fecha) & " And b_minutadet.mid_tipmin='2'", vg_db, adOpenStatic
'        If RS!NReg = 0 Then RS.Close: Set RS = Nothing: MsgBox "Debe realizar el pedido para la planificación teórica de este mes...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'        RS.Close: Set RS = Nothing
'        Unload M_MinRea: M_MinRea.Show 1
'    End If
Case 4 '------- Historico planificación
    vg_IndpprSelec = Val(fg_codigocbo(Combo2, 1, 1, ""))
    RS.Open "SELECT * FROM a_subsegmento With(NoLock) WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpLongInteger1(0).Value = "": fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(1).Caption = "": fpLongInteger1(2).Value = "": fpayuda(2).Caption = "": MsgBox "No existe subsegmento", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    vg_codigo = ""
    If NomFor = "MINTEO" Then B_HistPm.LlenarHistPlan "Histórico Planificación Teórica", Val(fpLongInteger1(0).Value), 1, 1
'    If NomFor = "MINREA" Then B_HistPm.LlenarHistPlan "Histórico Planificación Real", fpText.Text, 2, 1
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = vg_codregimen
    If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(1).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen With(NoLock) WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    
    fpLongInteger1(2).Value = vg_codservicio
    RS.Open "SELECT * FROM a_servicio With(NoLock) WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "":  Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    fpDateTime1.text = vg_fecha
    Me.Refresh
Case 6 '------- Salir
'    Unload M_Plami2
    Unload B_Receta
    Me.Hide
    Unload Me
End Select


End Sub

Sub Partidas(tfor As String, NFor As String)
    Me.Caption = tfor
    MsgTitulo = tfor
    NomFor = NFor
End Sub

Sub MoverListaPrecio()
Dim RS As New ADODB.Recordset
vg_codlpr = 0
vg_Zona = 0
fpayuda(3).Caption = ""
Val (fg_codigocbo(Combo2, 1, 1, ""))
vg_IndpprSelec = Val(fg_codigocbo(Combo2, 1, 1, ""))
vg_Zona = Val(fg_codigocbo(Combo2, 0, 10, ""))

Set RS = vg_db.Execute("sgpadm_s_planifminuta 10, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & "," & vg_Zona & ", " & Format(fpDateTime1.text, "yyyymm") & ", 0, 0," & vg_IndpprSelec & "")
If Not RS.EOF Then fpayuda(3).Caption = RS!lpr_codigo & " - " & Trim(RS!lpr_nombre): vg_codlpr = RS!lpr_codigo: vg_Zona = RS!zon_codigo: vg_Zona = IIf(vg_Zona = "", 0, vg_Zona)
RS.Close: Set RS = Nothing
End Sub


Sub DropTebleTmp(NameTable As String)
'*****************----> Destruye Tabla temporal<---------------------------
'---- Destruye tabla temporal, de manera que desbloquee el acceso a la minuta

    Dim RSTempCheck As New ADODB.Recordset

    Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaMinuta_" & NameTable & "'")
        If Not (RSTempCheck.EOF = True And RSTempCheck.BOF = True) Then
            Set RSTem = vg_db.Execute("Drop Table ##ValidaMinuta_" & NameTable & " ")
        End If

    RSTempCheck.Close
    Set RSTempCheck = Nothing
End Sub

