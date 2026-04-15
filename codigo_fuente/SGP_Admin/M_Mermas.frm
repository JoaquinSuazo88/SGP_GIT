VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Mermas 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mermas"
   ClientHeight    =   7230
   ClientLeft      =   1800
   ClientTop       =   2010
   ClientWidth     =   9705
   FillColor       =   &H00004040&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7230
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   3795
      Left            =   225
      TabIndex        =   23
      Top             =   2550
      Width           =   9240
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3405
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8970
         _Version        =   393216
         _ExtentX        =   15822
         _ExtentY        =   6006
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   8
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Mermas.frx":0000
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   1710
      TabIndex        =   18
      Top             =   6375
      Width           =   6585
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   330
         TabIndex        =   19
         Top             =   240
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   688
         ButtonWidth     =   2963
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1635
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Mermas.frx":06B1
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Mermas.frx":09CB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad sobrepasa Stock actual"
         Height          =   450
         Index           =   1
         Left            =   4560
         TabIndex        =   20
         Top             =   225
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   4170
         Top             =   345
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   450
         Top             =   315
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Productos igresados por el Usuario"
         Height          =   450
         Index           =   0
         Left            =   840
         TabIndex        =   22
         Top             =   195
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   735
         Top             =   300
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Productos traidos de la Minuta Real"
         Height          =   450
         Index           =   2
         Left            =   1095
         TabIndex        =   21
         Top             =   180
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2145
      Left            =   585
      TabIndex        =   6
      Top             =   375
      Width           =   8580
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   2055
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1245
         Width           =   3195
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   2055
         TabIndex        =   0
         Top             =   555
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2055
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1650
         Width           =   3195
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   2055
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   195
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         ControlType     =   2
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
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
         Left            =   2055
         TabIndex        =   1
         Top             =   900
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
         AllowNull       =   -1  'True
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
         Text            =   ""
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
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   3615
         TabIndex        =   17
         Top             =   255
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Bodega"
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
         Index           =   8
         Left            =   465
         TabIndex        =   14
         Top             =   1290
         Width           =   1560
      End
      Begin VB.Label Label3 
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
         Height          =   225
         Index           =   6
         Left            =   465
         TabIndex        =   13
         Top             =   585
         Width           =   1560
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3420
         Picture         =   "M_Mermas.frx":0CE5
         Top             =   450
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "Folio"
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
         Index           =   5
         Left            =   465
         TabIndex        =   12
         Top             =   240
         Width           =   1560
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha Emisión"
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
         Index           =   7
         Left            =   465
         TabIndex        =   11
         Top             =   915
         Width           =   1560
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Merma"
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
         Index           =   10
         Left            =   465
         TabIndex        =   10
         Top             =   1695
         Width           =   1560
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3870
         TabIndex        =   8
         Top             =   555
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   3
         Left            =   2100
         TabIndex        =   7
         Top             =   1710
         Width           =   3195
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3915
         TabIndex        =   9
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2115
         TabIndex        =   15
         Top             =   1290
         Width           =   3180
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Mermas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim Est As Boolean
'Dim btnX As Button

Private Sub Combo1_Click(Index As Integer)
Dim feprod As Long, codser As Long, i As Long
If Est Then Exit Sub
Select Case Index
Case 0
    If Combo1(0).ListIndex = -1 Or Combo1(0).Text = "" Then Exit Sub
    vaSpread1.MaxRows = 0
    Gl_Ac_Botones Me, 4, 5, ""
    Toolbar2.Enabled = True
    If vaSpread1.Enabled = True Then vaSpread1.SetFocus
    Toolbar2_ButtonClick Toolbar2.Buttons.Item(1)
    If vaSpread1.Enabled = True Then vaSpread1.SetFocus
Case 1
    If vaSpread1.MaxRows = 0 Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.col = 8
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
        'REvisa color
        Dim canrea As Double, canbod As Double
        vaSpread1.col = 4: canrea = Format(Val(vaSpread1.Text), fg_Pict(9, vg_DCa))
        vaSpread1.col = 8: canbod = Format(Val(vaSpread1.Text), fg_Pict(9, vg_DCa))
        If canbod - canrea < 0 Then
            vaSpread1.col = -1: vaSpread1.BackColor = Shape1(1).FillColor
            vaSpread1.col = 7: vaSpread1.Text = "S"  'Bloqueado
        Else
            vaSpread1.col = -1: vaSpread1.BackColor = Shape1(2).FillColor
            vaSpread1.col = 7: vaSpread1.Text = "N"   'No Bloqueado
        End If
    Next i
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7635
Me.Width = 9930
fg_centra Me
Est = False
Me.HelpContextID = vg_OpcM
MsgTitulo = "Salida Producción"
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 4
vaSpread1.Row = -1
vaSpread1.col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
RS1.Open "select bod_nombre, bod_codigo from a_bodega order by bod_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Combo1(1).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
RS1.Open "select aju_nombre, aju_codigo from a_tipoajuste where aju_tipaju=0 order by aju_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        Combo1(0).AddItem RS1!aju_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!aju_codigo), 10) & ")"
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
Limpia
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame3.Left = (Me.Width \ 2) - (Frame3.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame3.Left = 255
    Frame1.Left = 585
    Frame2.Left = 1710
End If
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
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
        Gl_Ac_Botones Me, 4, 2, ""
        fpText1(1).Enabled = False
        RS1.MoveNext
    Loop
Else
    RS1.Close: Set RS1 = Nothing
    MsgBox "Casino no existe...", vbExclamation + vbOKOnly, MsgTitulo
    Limpia
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    Exit Sub
End If
RS1.Close: Set RS1 = Nothing
fpLongInteger1(0).Text = MuestraFolio(Trim(fpText1(1).Text))
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
    If Trim(vg_codigo) <> fpText1(Index) Then Limpia
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus 1
    If fpDateTime1(Index - 1).Enabled = True Then fpDateTime1(Index - 1).SetFocus
    Gl_Ac_Botones Me, 4, 2, ""
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, TipDoc As String, numdoc As Long, CodBod   As Long, codser As Long, i As Long, canact As Long, diablq As Date
Dim numlin As Long, CodMer As String, canmer As Double, predoc As Double, ptotal As Double, descri As String, total As Double
On Error GoTo Man_Error
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
Select Case Button.Index
Case 1 'Nuevo
    Limpia
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
Case 3 'Graba
    If Trim(fpText1(1).Text) = "" Or Trim(fpLongInteger1(0).Text) = "" Or Trim(Combo1(0).Text) = "" _
    Or Trim(Combo1(1).Text) = "" Or Trim(fpDateTime1(0).Text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 7
        If Left(vaSpread1.Text, 1) = "S" Then MsgBox "Existe una cantidad que exede el Stock...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    total = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.col = 6: ptotal = LimpiaDato(vaSpread1.Text)
        total = total + ptotal
    Next i
    If total = 0 Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_db.BeginTrans
    rutcli = Trim(LimpiaDato(fpText1(1).Text))
    TipDoc = "ME"
    fpLongInteger1(0).Text = MuestraFolio(Trim(fpText1(1).Text))
    numdoc = fpLongInteger1(0).Text
    CodBod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    codser = Val(fg_codigocbo(Combo1, 0, 10, ""))
    'Encabezado
    vg_db.Execute "insert into b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                  "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & CodBod & ", CDate('" & _
                  Format(fpDateTime1(0).Text, "dd/mm/yyyy") & "'), 0, " & codser & ", 0, '', '', 0)"
    'Detalle
    total = 0
    numlin = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.col = 2: descri = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.col = 4: canmer = LimpiaDato(vaSpread1.Text)
        vaSpread1.col = 5: predoc = LimpiaDato(vaSpread1.Text)
        vaSpread1.col = 6: ptotal = LimpiaDato(vaSpread1.Text)
        If canmer > 0 Then
            total = total + ptotal
            
            vg_db.Execute "insert into b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding) " & _
                          "values ('" & rutcli & "', '" & TipDoc & "', " & numdoc & ", " & numlin & ", '" & CodMer & "', " & "0" & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '')"
            'Control de Stock
            canact = 0
            RS1.Open "select bod_canmer from b_bodegas where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod, vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    canact = RS1!bod_canmer - canmer
                    RS1.MoveNext
                Loop
                vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                              "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & CodBod
            End If
            RS1.Close: Set RS1 = Nothing: numlin = numlin + 1
            
        End If
    Next i
    'Total
    vg_db.Execute "update b_totventas set tov_totdoc=" & total & " where tov_rutcli='" & Trim(LimpiaDato(fpText1(1).Text)) & "' " & _
                  "and tov_tipdoc='ME' and tov_numdoc=" & fpLongInteger1(0).Value
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    Toolbar2.Enabled = False
    vaSpread1.col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    'Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.col = 8
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    I_Mermas Me
Case 5 'Anular
    'If Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If (Format(Now, "dd/mm/yyyy") > diablq And Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Now, "mm/yyyy")) Or Format(CDate(fpDateTime1(0).Text), "mm/yyyy") < Format(Month(Now) - 1 & "/" & Year(Now), "mm/yyyy") Then MsgBox "Mes Bloqueado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    'Encabezado
    vg_db.Execute "update b_totventas set tov_estdoc='A' where tov_rutcli='" & Trim(LimpiaDato(fpText1(1).Text)) & "' " & _
                  "and tov_tipdoc='ME' and tov_numdoc=" & fpLongInteger1(0).Value
    Label1.Caption = "ANULADA"
    'Detalle
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: numlin = i
        vaSpread1.col = 1: CodMer = Trim(LimpiaDato(vaSpread1.Text))
        vaSpread1.col = 4: canmer = LimpiaDato(vaSpread1.Text)
        'Control de Stock
        canact = 0
        RS1.Open "select bod_canmer from b_bodegas where bod_codpro='" & CodMer & "' and bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")), vg_db, adOpenStatic
        If Not RS1.EOF Then
            Do While Not RS1.EOF
                canact = RS1!bod_canmer + canmer
                RS1.MoveNext
            Loop
            vg_db.Execute "update b_bodegas set bod_canmer=" & canact & " " & _
                          "where bod_codpro='" & Trim(LimpiaDato(CodMer)) & "' and bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, ""))
        End If
        RS1.Close: Set RS1 = Nothing
    Next i
    'Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_db, adOpenStatic
        vaSpread1.col = 8
        If Not RS1.EOF Then vaSpread1.Text = RS1!bod_canmer Else vaSpread1.Text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 8 'Busqueda
    If Trim(fpText1(1).Text) = "" Then MsgBox "Debe seleccionar casino...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_codigo = Trim(fpText1(1).Text)
    vg_nombre = "ME"
    B_SalBod.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    Toolbar2.Enabled = False
    Frame1.Enabled = False
    vaSpread1.col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    RS2.Open "select tov.tov_numdoc, tov.tov_codbod, tov.tov_fecemi, tov.tov_codser, tov.tov_estdoc " & _
             "from b_totventas tov, b_clientes cli " & _
             "where tov.tov_rutcli='" & LimpiaDato(Trim(fpText1(1).Text)) & "' " & _
             "and tov.tov_tipdoc='ME' " & _
             "and tov.tov_numdoc=" & Val(vg_codigo) & " " & _
             "and tov.tov_rutcli=cli.cli_codigo", vg_db, adOpenStatic
    If Not RS2.EOF Then
        Do While Not RS2.EOF
            Est = True
            fpLongInteger1(0).Text = RS2!tov_numdoc
            Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
            fpDateTime1(0).Text = RS2!tov_fecemi
            Combo1(0).ListIndex = fg_buscacbo(Combo1, 0, 10, fg_pone_cero(Str(RS2!tov_codser), 10))
            Label1.Caption = IIf(RS2!tov_estdoc = "", "", "ANULADA")
            Est = False
            RS2.MoveNext
        Loop
    End If
    RS2.Close: Set RS2 = Nothing
    RS1.Open "select dev.dev_codmer, dev.dev_canmer, dev.dev_predoc, " & _
             "dev.dev_ptotal, dev.dev_descri, uni.uni_nombre " & _
             "from b_detventas dev, b_productos pro ,a_unidad uni " & _
             "where dev.dev_rutcli='" & LimpiaDato(Trim(fpText1(1).Text)) & "'" & _
             "and dev.dev_tipdoc='ME' " & _
             "and dev.dev_numdoc=" & Val(vg_codigo) & " " & _
             "and dev.dev_codmer=pro.pro_codigo " & _
             "and pro.pro_coduni=uni.uni_codigo order by dev.dev_numlin", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = 1
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = i
            vaSpread1.col = 1: vaSpread1.Text = RS1!dev_codmer
            vaSpread1.col = 2: vaSpread1.Text = RS1!dev_descri
            vaSpread1.col = 3: vaSpread1.Text = RS1!uni_nombre
            vaSpread1.col = 4: vaSpread1.Text = RS1!dev_canmer
            vaSpread1.col = 5: vaSpread1.Text = RS1!dev_predoc
            vaSpread1.col = 6: vaSpread1.Text = RS1!dev_ptotal
            'Trae Stock
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(RS1!dev_codmer) & "'", vg_db, adOpenStatic
            vaSpread1.col = 9
            If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext: i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    vg_codigo = ""
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 9 'Imprimir
    I_Mermas Me
Case 12 'Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = 3034 Then Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Function MuestraFolio(Casino As String) As String
MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function
RS1.Open "select tov_numdoc from b_totventas where tov_tipdoc='ME' order by tov_numdoc desc", vg_db, adOpenStatic
If Not RS1.EOF Then
    RS1.MoveFirst
    MuestraFolio = RS1!tov_numdoc + 1
Else
    MuestraFolio = 1
End If
RS1.Close: Set RS1 = Nothing
End Function

Sub Limpia()
Label1.Caption = ""
Frame1.Enabled = True
fpDateTime1(0).Text = Format(Date, "dd/mm/yyyy")
Combo1(0).ListIndex = -1
Combo1(1).ListIndex = IIf(Combo1(1).ListCount = 1, 0, -1)
vaSpread1.MaxRows = 0
Toolbar2.Enabled = False
vaSpread1.col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
vaSpread1.col = 4: vaSpread1.Row = -1
vaSpread1.Lock = False
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).Text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
fpLongInteger1(0).Text = MuestraFolio(Trim(fpText1(1).Text))
Gl_Ac_Botones Me, 4, 2, ""
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long
Select Case Button.Index
Case 1
    If Trim(Combo1(1).Text) = "" Then MsgBox "Debe seleccionar bodega...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_nombre = "": vg_codigo = ""
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gen"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.Text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    RS1.Open "SELECT pro.pro_codigo, pro.pro_propon, pro.pro_nombre, uni.uni_nombre " & _
             "FROM b_productos AS pro, a_unidad AS uni " & _
             "WHERE pro.pro_coduni=uni.uni_codigo and pro.pro_codigo='" & vg_codigo & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = vaSpread1.MaxRows + 1
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.col = 1: vaSpread1.Text = RS1!pro_codigo
            vaSpread1.col = 2: vaSpread1.Text = RS1!pro_nombre
            vaSpread1.col = 3: vaSpread1.Text = RS1!uni_nombre
            vaSpread1.col = 4: vaSpread1.Text = 0
            vaSpread1.col = 5: vaSpread1.Text = RS1!pro_propon
            vaSpread1.col = 6: vaSpread1.Text = 0
            vaSpread1.col = 7: vaSpread1.Text = "N" 'No bloquedo
            'Trae Stock
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "and bod.bod_codpro=pro.pro_codigo and pro.pro_codigo='" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
            vaSpread1.col = 8
            If Not RS2.EOF Then vaSpread1.Text = RS2!bod_canmer Else vaSpread1.Text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext
            i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    If vaSpread1.MaxRows = 1 Then Gl_Ac_Botones Me, 4, 2, ""
    vaSpread1.col = 4: vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.col = 1
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    If vaSpread1.MaxRows = 0 Then Gl_Ac_Botones Me, 4, 5, ""
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal col As Long, ByVal Row As Long)
Dim canrea As Double, propon As Double, CodMer As String
vaSpread1.Row = Row
vaSpread1.col = 1: CodMer = vaSpread1.Text
vaSpread1.col = 4: canrea = Format(vaSpread1.Text, fg_Pict(9, 2))
vaSpread1.col = 5: propon = Format(vaSpread1.Text, fg_Pict(9, 2))
vaSpread1.col = 6: vaSpread1.Text = Format(canrea * propon, fg_Pict(9, 2))
vaSpread1.col = 8: canbod = Format(vaSpread1.Text, fg_Pict(9, 2))
If canbod - canrea >= 0 Then
    vaSpread1.col = -1: vaSpread1.BackColor = Shape1(2).FillColor
    vaSpread1.col = 7: vaSpread1.Text = "N" 'No Bloqueado
    Exit Sub
End If
vaSpread1.col = -1: vaSpread1.BackColor = Shape1(1).FillColor
vaSpread1.col = 7: vaSpread1.Text = "S"  'Bloqueado
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread1.Row = Row: vaSpread1.col = 8: Stock = vaSpread1.Text
vaSpread1.Row = Row: vaSpread1.col = 2: Nombre = vaSpread1.Text
TipText = "Bodega   : " & Trim(Left(Combo1(1).Text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))
End Sub
