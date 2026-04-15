VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_PCPatr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametro Costo Patrón"
   ClientHeight    =   5055
   ClientLeft      =   1395
   ClientTop       =   2550
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5055
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   120
      TabIndex        =   17
      Top             =   2520
      Width           =   11480
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   11205
         _Version        =   393216
         _ExtentX        =   19764
         _ExtentY        =   3625
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   14
         MaxRows         =   10
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_PCPatr.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   1905
      TabIndex        =   6
      Top             =   435
      Width           =   7965
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1280
         Width           =   1335
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1800
         TabIndex        =   4
         Top             =   1650
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
         Text            =   "2023"
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   0
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
         MinValue        =   "0"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   585
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
         MinValue        =   "0"
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
         Left            =   1800
         TabIndex        =   2
         Top             =   930
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
         MinValue        =   "0"
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Costo"
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
         Left            =   480
         TabIndex        =   19
         Top             =   1380
         Width           =   930
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2625
         Picture         =   "M_PCPatr.frx":0A93
         Top             =   140
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3075
         TabIndex        =   15
         Top             =   225
         Width           =   4335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Segmento"
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
         Left            =   480
         TabIndex        =   14
         Top             =   315
         Width           =   1245
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
         Left            =   480
         TabIndex        =   13
         Top             =   1725
         Width           =   540
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3090
         TabIndex        =   11
         Top             =   585
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2640
         Picture         =   "M_PCPatr.frx":0D9D
         Top             =   500
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3090
         TabIndex        =   9
         Top             =   945
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2640
         Picture         =   "M_PCPatr.frx":10A7
         Top             =   860
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   660
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         TabIndex        =   7
         Top             =   1040
         Width           =   705
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3120
         TabIndex        =   16
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3135
         TabIndex        =   12
         Top             =   630
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3135
         TabIndex        =   10
         Top             =   990
         Width           =   4335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   330
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_PCPatr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim RS As New ADODB.Recordset
Dim Est As Boolean
Dim MsgTitulo As String
Public lc_Aux As String

Private Sub Combo1_Click(Index As Integer)
MoverDatosGrilla
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

Me.HelpContextID = vg_OpcM
Me.Height = 5580
Me.Width = 11790
MsgTitulo = "Parametro Costo Patrón"
fg_centra Me
fpDateTime1.text = Format(Date, "yyyy")
modo = "": Est = True

Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo

Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = False

Combo1(0).Clear
If lc_Aux = "CosPis" Then
    MsgTitulo = "Parametro Costo Patrón Piso"
    Me.Caption = "Parametro Costo Patrón Piso"
    Combo1(0).AddItem "PISO" & Space(150) & "(0)"
Else
    MsgTitulo = "Parametro Costo Patrón Techo"
    Me.Caption = "Parametro Costo Patrón Techo"
    Combo1(0).AddItem "TECHO" & Space(150) & "(1)"
End If
Combo1(0).ListIndex = 0
vaSpread1.MaxRows = 0
Est = False
vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.BackColor = Shape1(1).FillColor
vaSpread1.Lock = False
vaSpread1.Row = -1
vaSpread1.Col = 1: vaSpread1.BackColor = Shape1(2).FillColor: vaSpread1.Col = 2: vaSpread1.BackColor = Shape1(2).FillColor
vaSpread1.Col = 1: vaSpread1.text = "": vaSpread1.ColHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", True, False)
vaSpread1.Col = 2: vaSpread1.text = "": vaSpread1.ColHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", True, False)
End Sub

Sub MoverDatosGrilla()
If Est Then Exit Sub
Dim i As Long, anomes As Long, cencos As String
Dim RS As New ADODB.Recordset
On Error GoTo Man_Error
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
RS.Open "sp_s_paramcostopatron 1, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1.Value, "yyyy")) & ", '" & Trim(Mid(Combo1(0).text, 1, 150)) & "'", vg_db, adOpenStatic
cencos = ""
If Not RS.EOF Then
   Do While Not RS.EOF
      If RS!pcp_cencos <> cencos Then
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         vaSpread1.RowHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0" And vaSpread1.MaxRows > 1, True, False)
         cencos = RS!pcp_cencos
      End If
      vaSpread1.Col = 1: vaSpread1.text = Trim(RS!pcp_cencos): vaSpread1.ColHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", True, False)
      vaSpread1.Col = 2: vaSpread1.text = Trim(RS!cli_nombre): vaSpread1.ColHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", True, False)
      If RS!pcp_anomes > 0 Then vaSpread1.Col = (Val(Mid(RS!pcp_anomes, 5, 2)) + 2): vaSpread1.text = IIf(RS!pcp_valor > 0, RS!pcp_valor, "")
      RS.MoveNext
   Loop
   vaSpread1.Col = -1: vaSpread1.Row = -1
   vaSpread1.Lock = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", False, True)
   vaSpread1.SetActiveCell 3, 1
'   vaSpread1.SetFocus
   
End If
RS.Close: Set RS = Nothing
vaSpread1.Visible = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub fpDateTime1_Change()
MoverDatosGrilla
End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 0
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
       Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & " AND sub_indppr = '" & vg_Indppr & "'")
    Else
       Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
    End If
    If RS.EOF Then
       RS.Close
       Set RS = Nothing
       fpayuda(0).Caption = ""
       fpLongInteger1(1).Value = ""
       fpayuda(1).Caption = ""
       fpLongInteger1(2).Value = ""
       fpayuda(2).Caption = ""
       Exit Sub
    End If
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
    fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
    MoverDatosGrilla
Case 1
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
       Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    Else
       Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrilla
Case 2
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
       Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(2).Value) & " AND ser_indppr = '" & vg_Indppr & "'")
    Else
       Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Val(fpLongInteger1(2).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "":  Exit Sub
    fpayuda(2).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    MoverDatosGrilla
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
    Combo1(0).SetFocus
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Select Case Button.Index
Case 3
    If Val(fpLongInteger1(0).Value) = 0 Or Val(fpLongInteger1(1).Value) = 0 Or Val(fpLongInteger1(2).Value) = 0 Or Trim(fpDateTime1.text) = "" Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = False
Case 5
    If Val(fpLongInteger1(0).Value) = 0 Or Val(fpLongInteger1(1).Value) = 0 Or Val(fpLongInteger1(2).Value) = 0 Or Trim(fpDateTime1.text) = "" Then Exit Sub
    If Not Est < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.Execute "DELETE b_paramcostopatron " & _
                  "WHERE  pcp_codreg = " & fpLongInteger1(1).Value & " " & _
                  "AND    pcp_codser = " & Val(fpLongInteger1(2).Value) & " " & _
                  "AND    substring(convert(char(6),pcp_anomes),1,4) = " & Format(fpDateTime1.Value, "yyyy") & " " & _
                  "AND    pcp_descripcion = '" & IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", "PISO", "TECHO") & "'"
    vaSpread1.MaxRows = 0
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
    Dim cencos As String, descripcion As String, valor As Double, IndDia As Long
    fg_carga ""
    For i = 1 To vaSpread1.MaxRows
        DoEvents
        vaSpread1.Row = i
        vaSpread1.Col = 1: cencos = vaSpread1.text
        indmes = 1
         For j = 3 To vaSpread1.maxcols
            vaSpread1.Col = j
            valor = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
            DoEvents
            If vaSpread1.ForeColor = &HFF0000 Then
               vg_db.Execute "DELETE b_paramcostopatron WHERE pcp_cencos='" & Trim(cencos) & "' " & _
                             "AND    pcp_codreg=" & Val(fpLongInteger1(1).Value) & " " & _
                             "AND    pcp_codser=" & Val(fpLongInteger1(2).Value) & " " & _
                             "AND    substring(convert(char(6),pcp_anomes),1,6)=" & Val(Format(fpDateTime1.Value, "yyyy") & fg_pone_cero(indmes, 2)) & " " & _
                             "AND    pcp_descripcion='" & IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", "PISO", "TECHO") & "'"
               vg_db.Execute "sgpadm_iu_paramcostopatron 'A', '" & cencos & "', " & Val(fpLongInteger1(1).Value) & ", " & Val(fpLongInteger1(2).Value) & ", " & Val(Format(fpDateTime1.Value, "yyyy")) & fg_pone_cero(indmes, 2) & ", '" & IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", "PISO", "TECHO") & "', " & valor & ""
            End If
            indmes = indmes + 1
        Next j
    Next i
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = False
    Frame1.Enabled = True
    fg_descarga
Case 15
    RS.Open "SELECT DISTINCT a.pcp_cencos FROM b_paramcostopatron a, b_clientes b " & _
            "WHERE a.pcp_cencos = b.cli_codigo " & _
            "AND   b.cli_subseg = " & Val(fpLongInteger1(0).Value) & " " & _
            "AND   a.pcp_codreg = " & Val(fpLongInteger1(1).Value) & " " & _
            "AND   a.pcp_codser = " & Val(fpLongInteger1(2).Value) & " " & _
            "AND   substring(convert(char(6),a.pcp_anomes),1,4) = " & Format(fpDateTime1.Value, "yyyy") & " " & _
            "AND   a.pcp_descripcion = '" & IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", "PISO", "TECHO") & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    I_ParamCostoPatron Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), Val(fpLongInteger1(2).Value), Val(Format(fpDateTime1.Value, "yyyy")), IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", 0, 1)
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
If Err = -2147467259 Or 2147217900 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
If modo = "" Then modo = "M"
If ChangeMade = True And modo = "M" Then
   If fg_codigocbo(Combo1, 0, 1, "") = "0" Then
      vaSpread1.Row = Row
      vaSpread1.Col = Col
      vaSpread1.ForeColor = &HFF0000
      valor = vaSpread1.text
      For i = 2 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = Col
          vaSpread1.ForeColor = &HFF0000
          vaSpread1.text = valor
      Next i
   Else
          vaSpread1.Row = Row
          vaSpread1.Col = Col
          vaSpread1.ForeColor = &HFF0000
'          vaSpread1.text = valor
   End If
   Gl_Ac_Botones Me, 1, 0, modo: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Frame1.Enabled = False
End If
End Sub


