VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
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
         Left            =   90
         TabIndex        =   19
         Top             =   240
         Width           =   3960
         _ExtentX        =   6985
         _ExtentY        =   688
         ButtonWidth     =   3307
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
         Enabled         =   0   'False
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
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
         AutoBeep        =   -1  'True
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
         Caption         =   "Nº Documento"
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
Dim est As Boolean
'Dim btnX As Button

Private Sub Combo1_Click(Index As Integer)
Dim feprod As Long, codser As Long, i As Long
If est Then Exit Sub
With vaSpread1
    Select Case Index
    Case 0
        If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then Exit Sub
        .MaxRows = 0
        Gl_Ac_Botones Me, 4, 6, ""
        Toolbar2.Enabled = True
        If .Enabled = True Then .SetFocus
        Toolbar2_ButtonClick Toolbar2.Buttons.Item(1)
        If .Enabled = True Then .SetFocus
    Case 1
        If .MaxRows = 0 Then Exit Sub
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        For i = 1 To .MaxRows
            .Row = i: .Col = 1
            RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE  bod.bod_codpro = pro.pro_codigo " & _
                     "and    bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "and    pro.pro_codigo = '" & Trim(LimpiaDato(.text)) & "' " & _
                     "and    pro.pro_ctrsto = 1", vg_db, adOpenStatic
            .Col = 8
            If Not RS1.EOF Then .text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else .text = 0
            RS1.Close: Set RS1 = Nothing
            'REvisa color
            Dim canrea As Double, canbod As Double
            .Col = 4: canrea = Format(Val(.text), fg_Pict(9, vg_DCa))
            .Col = 8: canbod = Format(Val(.text), fg_Pict(9, vg_DCa))
            If canbod - canrea < 0 Then
                .Col = -1: .BackColor = Shape1(1).FillColor
                .Col = 7: .text = "S"  'Bloqueado
            Else
                .Col = -1: .BackColor = Shape1(2).FillColor
                .Col = 7: .text = "N"   'No Bloqueado
            End If
        Next i
    End Select
End With
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7635
Me.Width = 9930
fg_centra Me
est = False
Me.HelpContextID = vg_OpcM
EspFecha fpDateTime1(0)
MsgTitulo = "Merma"
Dim X As Boolean
Gl_Mo_Botones Me, 4
With vaSpread1
    .TextTip = 2
    .TextTipDelay = 0
    X = .SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
    .Row = -1
    .Col = 4: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DCa
    .Col = 5: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DPr
    .Col = 6: .TypeNumberSeparator = vg_CSep: .TypeNumberDecimal = vg_CDec: .TypeNumberDecPlaces = vg_DPr
End With
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"
'-------> Cargar Combo Tipo Mermas
CargarDatoCombo Combo1, 0, "a_tipoajuste", "aju_", "TipAju", "NM"
Limpia 2
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
If fpText1(1).text = "" Then Exit Sub

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText1(1).text)), ""), vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        fpayuda(Index).Caption = RS1!cli_nombre
        Gl_Ac_Botones Me, 4, 2, ""
        fpText1(1).Enabled = False
        RS1.MoveNext
    Loop
Else
    RS1.Close: Set RS1 = Nothing
    MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
    Limpia 2
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    Exit Sub
End If
RS1.Close: Set RS1 = Nothing
'fpLongInteger1(0).text = MuestraFolio(Trim(fpText1(1).text))
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "ME")
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
    If Trim(vg_codigo) <> fpText1(Index) Then Limpia 2
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus 1
    If fpDateTime1(Index - 1).Enabled = True Then fpDateTime1(Index - 1).SetFocus
    Gl_Ac_Botones Me, 4, 2, ""
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, tipdoc As String, NumDoc As Long, codbod   As Long, codser As Long, i As Long, canact As Double, diablq As Date
Dim numlin As Long, fecpro As Date, codmer As String, canmer As Double, predoc As Double, ptotal As Double, descri As String, total As Double
On Error GoTo Man_Error
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre
Select Case Button.Index
Case 1, 6 '-------> Nuevo
'    Limpia
    If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Limpia IIf(Button.Index = 1, 6, 2)
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
Case 8 '-------> Graba
    If Trim(fpText1(1).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(Combo1(0).text) = "" _
    Or Trim(Combo1(1).text) = "" Or Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
       
       MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
    
       MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    'Validar inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
        
       MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
        
    'Validar ingreso documento inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
        
       MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then
    
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
    
       MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 7
        If Left(vaSpread1.text, 1) = "S" Then MsgBox "Existe una cantidad que exede el Stock...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    total = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 6: ptotal = LimpiaDato(vaSpread1.text)
        total = total + ptotal
    Next i
    If total = 0 Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
paso:
    rutcli = Trim(LimpiaDato(fpText1(1).text))
    tipdoc = "ME"
'    fpLongInteger1(0).text = MuestraFolio(Trim(fpText1(1).text))
'    numdoc = fpLongInteger1(0).text
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    codser = Val(fg_codigocbo(Combo1, 0, 10, ""))
    NumDoc = TraerCorrelativo(codbod, "ME")
'    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_parametros SET par_correlativo=" & NumDoc & " WHERE par_codbod=" & codbod & " AND par_tipdoc='ME'"
'    vg_db.CommitTrans
    fpLongInteger1(0).text = NumDoc
    DoEvents
    
    vg_db.BeginTrans
    '-------> Encabezado
    If vg_tipbase = "1" Then
       vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                     "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", CDate('" & _
                     Format(fpDateTime1(0).text, "dd/mm/yyyy") & "'), 0, " & codser & ", 0, '', '', 0)"
    Else
       vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                     "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & _
                     Format(fpDateTime1(0).text, "yyyymmdd") & "', 0, " & codser & ", 0, '', '', 0)"
    End If
    '-------> Detalle
    total = 0
    numlin = 1
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 4: canmer = LimpiaDato(vaSpread1.text)
        vaSpread1.Col = 5: predoc = LimpiaDato(vaSpread1.text)
        vaSpread1.Col = 6: ptotal = LimpiaDato(vaSpread1.text)
        If canmer > 0 Then
            total = total + ptotal
            
            vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding) " & _
                          "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', " & "0" & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', 0, " & predoc & ", '')"
            '-------< Control de Stock
            canact = 0
            RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro = '" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod = " & codbod, vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                   If (RS1!bod_canmer - canmer) < 0 Then
                       RS1.Close: Set RS1 = Nothing
                       vg_db.RollbackTrans
                       For j = 1 To vaSpread1.MaxRows
                           vaSpread1.Row = j: vaSpread1.Col = 1
                           RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                                    "WHERE bod.bod_codpro = pro.pro_codigo " & _
                                    "and   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                                    "and   pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "' " & _
                                    "and   pro.pro_ctrsto = 1", vg_db, adOpenStatic
                           vaSpread1.Col = 8
                           If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
                           RS1.Close: Set RS1 = Nothing
                       Next j
                       MsgBox "Existen productos con diferencia en la bodega, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
                       Exit Sub
                    End If
                    canact = RS1!bod_canmer - canmer
                    RS1.MoveNext
                Loop
                vg_db.Execute "UPDATE b_bodegas SET bod_canmer=" & canact & " " & _
                              "WHERE bod_codpro='" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod=" & codbod
            End If
            RS1.Close: Set RS1 = Nothing: numlin = numlin + 1
        End If
    Next i
    '-------> Total
    vg_db.Execute "UPDATE b_totventas SET tov_totdoc = " & total & " WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                  "AND tov_tipdoc = 'ME' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    Toolbar2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codpro = pro.pro_codigo " & _
                 "and   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and   pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "' " & _
                 "and   pro.pro_ctrsto = 1", vg_db, adOpenStatic
        vaSpread1.Col = 8
        If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    I_Mermas Me
Case 3 '-------> Anular
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Periodo cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then MsgBox "No puede anular documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Anula documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    '-------> Encabezado
    vg_db.Execute "UPDATE b_totventas SET tov_estdoc = 'A' WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(1).text)) & "' " & _
                  "AND tov_tipdoc = 'ME' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    Label1.Caption = "ANULADA"
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    '-------> Detalle
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: numlin = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 4: canmer = LimpiaDato(vaSpread1.text)
        '-------> Control de Stock
        canact = 0
        RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro = '" & codmer & "' and bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")), vg_db, adOpenStatic
        If Not RS1.EOF Then
            Do While Not RS1.EOF
                canact = RS1!bod_canmer + canmer
                RS1.MoveNext
            Loop
            vg_db.Execute "UPDATE b_bodegas SET bod_canmer = " & canact & " " & _
                          "WHERE bod_codpro = '" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, ""))
        End If
        RS1.Close: Set RS1 = Nothing
    Next i
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codpro = pro.pro_codigo " & _
                 "and   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "and   pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "' " & _
                 "and   pro.pro_ctrsto = 1", vg_db, adOpenStatic
        vaSpread1.Col = 8
        If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 11 '-------> Busqueda
    If Trim(fpText1(1).text) = "" Then MsgBox "Debe seleccionar contrato...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    vg_codigo = Trim(fpText1(1).text)
    vg_nombre = "ME"
    B_SalBod.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    Toolbar2.Enabled = False
    Frame1.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS2.Open "SELECT tov.tov_numdoc, tov.tov_codbod, tov.tov_fecemi, tov.tov_codser, tov.tov_estdoc " & _
             "FROM b_totventas tov, b_clientes cli " & _
             "WHERE tov.tov_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "' " & _
             "AND   tov.tov_tipdoc = 'ME' " & _
             "AND   tov.tov_numdoc = " & Val(vg_codigo) & " AND tov.tov_codbod = " & vg_codbod & " " & _
             "AND   tov.tov_rutcli = cli.cli_codigo", vg_db, adOpenStatic
    If Not RS2.EOF Then
        Do While Not RS2.EOF
            est = True
            fpLongInteger1(0).text = RS2!tov_numdoc
            Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
            fpDateTime1(0).text = RS2!tov_fecemi
            Combo1(0).ListIndex = fg_buscacbo(Combo1, 0, 10, fg_pone_cero(Str(RS2!tov_codser), 10))
            Label1.Caption = IIf(RS2!tov_estdoc = "A", "ANULADA", "")
            est = False
            RS2.MoveNext
        Loop
    End If
    RS2.Close: Set RS2 = Nothing
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT dev.dev_codmer, dev.dev_canmer, dev.dev_predoc, " & _
             "dev.dev_ptotal, dev.dev_descri, uni.uni_nombre " & _
             "FROM b_detventas dev, b_productos pro ,a_unidad uni " & _
             "WHERE dev.dev_rutcli = '" & LimpiaDato(Trim(fpText1(1).text)) & "'" & _
             "AND   dev.dev_tipdoc = 'ME' " & _
             "AND   dev.dev_numdoc = " & Val(vg_codigo) & " " & _
             "AND   dev.dev_codmer = pro.pro_codigo " & _
             "AND   pro.pro_coduni = uni.uni_codigo ORDER BY dev.dev_numlin", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = 1
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = i
            vaSpread1.Col = 1: vaSpread1.text = RS1!dev_codmer
            vaSpread1.Col = 2: vaSpread1.text = RS1!dev_descri
            vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nombre
            vaSpread1.Col = 4: vaSpread1.text = RS1!dev_canmer
            vaSpread1.Col = 5: vaSpread1.text = RS1!dev_predoc
            vaSpread1.Col = 6: vaSpread1.text = RS1!dev_ptotal

            If RS2.State = 1 Then RS2.Close
            RS2.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            'Trae Stock
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE bod.bod_codpro = pro.pro_codigo " & _
                     "AND   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "AND   pro.pro_codigo = '" & Trim(RS1!dev_codmer) & "' " & _
                     "AND   pro.pro_ctrsto = 1", vg_db, adOpenStatic
            vaSpread1.Col = 9
            If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext: i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    vg_codigo = ""
    Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
Case 12 '-------> Imprimir
    I_Mermas Me
Case 15 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: GoTo paso
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Function MuestraFolio(Casino As String) As String
MuestraFolio = ""
If Trim(Casino) = "" Then Exit Function

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT tov_numdoc FROM b_totventas WHERE tov_tipdoc='ME' AND tov_codbod=" & vg_codbod & " ORDER BY tov_numdoc DESC", vg_db, adOpenStatic
If Not RS1.EOF Then
   RS1.MoveFirst
   MuestraFolio = RS1!tov_numdoc + 1
Else
   MuestraFolio = 1
End If
RS1.Close: Set RS1 = Nothing
End Function

Sub Limpia(op As Integer)
Label1.Caption = ""
Frame1.Enabled = True
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Combo1(0).ListIndex = -1
Combo1(1).ListIndex = IIf(Combo1(1).listcount = 1, 0, -1)
Toolbar2.Enabled = False
With vaSpread1
    .MaxRows = 0
    .Col = -1: .Row = -1
    .Lock = True
    .Col = 4: .Row = -1
    .Lock = False
End With
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
'fpLongInteger1(0).text = MuestraFolio(Trim(fpText1(1).text))
fpLongInteger1(0).text = TraerCorrelativo(vg_codbod, "ME")
Gl_Ac_Botones Me, 4, op, ""
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long
With vaSpread1
    Select Case Button.Index
    Case 1
        If Trim(Combo1(1).text) = "" Then MsgBox "Debe seleccionar bodega...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        vg_nombre = "": vg_codigo = "": vg_bodega = 0: vg_bodega = Val(fg_codigocbo(Combo1, 1, 10, ""))
        vg_left = fpayuda(1).Left + 1920
        B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pbo"
        B_TabEst.Show 1
        If vg_codigo = "" Then Exit Sub
        For i = 1 To .MaxRows
            .Col = 1: .Row = i
            If Trim(.text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        Next i
        .Row = .ActiveRow
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        RS1.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nombre " & _
                 "FROM b_productos a, a_unidad b " & _
                 "WHERE a.pro_coduni = b.uni_codigo " & _
                 "AND   a.pro_codigo = '" & vg_codigo & "' AND a.pro_ctrsto = 1", vg_db, adOpenStatic
        If Not RS1.EOF Then
            i = .MaxRows + 1
            Do While Not RS1.EOF
                .MaxRows = i
                .Row = .MaxRows
                .Col = 1: .text = RS1!pro_codigo
                .Col = 2: .text = RS1!pro_nombre
                .Col = 3: .text = RS1!uni_nombre
                .Col = 4: .text = 0
                
                If RS2.State = 1 Then RS2.Close
                RS2.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient

                '-------> Traer pmp
                RS2.Open "SELECT TOP 1 ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                         "FROM b_productospmpdia " & _
                         "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
                         "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & " " & _
                         "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
                         "HAVING (ppd_propon) > 0 ORDER BY Max(ppd_fecdia) DESC", vg_db, adOpenStatic
                .Col = 5
                If Not RS2.EOF Then .text = RS2!ppd_propon Else .text = 0
                RS2.Close: Set RS2 = Nothing
                .Col = 6: .text = 0
                .Col = 7: .text = "N" 'No bloquedo
                
                If RS2.State = 1 Then RS2.Close
                RS2.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient

                '-------> Trae Stock
                RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                         "WHERE bod.bod_codpro = pro.pro_codigo " & _
                         "AND   bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                         "AND   pro.pro_codigo = '" & Trim(RS1!pro_codigo) & "' " & _
                         "AND   pro.pro_ctrsto = 1", vg_db, adOpenStatic
                .Col = 8
                If Not RS2.EOF Then .text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else .text = 0
                RS2.Close: Set RS2 = Nothing
                RS1.MoveNext
                i = i + 1
            Loop
        End If
        RS1.Close: Set RS1 = Nothing
        If .MaxRows = 1 Then Gl_Ac_Botones Me, 4, 6, ""
        .Col = 4: .Row = .MaxRows
        .SetActiveCell 4, .MaxRows
    Case 2
        If .MaxRows = 0 Then Exit Sub
        .Row = .ActiveRow
        .Col = 1
        If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        .DeleteRows .Row, 1
        .MaxRows = .MaxRows - 1
        If .MaxRows = 0 Then Gl_Ac_Botones Me, 4, 6, ""
    End Select
End With
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim canrea As Double, propon As Double, codmer As String
With vaSpread1
    .Row = Row
    .Col = 1: codmer = .text
    .Col = 4: canrea = Format(.text, fg_Pict(9, vg_DCa))
    .Col = 5: propon = Format(.text, fg_Pict(9, vg_DPr))
    .Col = 6: .text = Format(canrea * propon, fg_Pict(9, vg_DPr))
    .Col = 8: canbod = Format(.text, fg_Pict(9, vg_DCa))
    If canbod - canrea >= 0 Then
        .Col = -1: .BackColor = Shape1(2).FillColor
        .Col = 7: .text = "N" 'No Bloqueado
        Exit Sub
    End If
    .Col = -1: .BackColor = Shape1(1).FillColor
    .Col = 7: .text = "S"  'Bloqueado
End With
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String
TipWidth = 4000
ShowTip = True
MultiLine = 2
With vaSpread1
    .Row = Row: .Col = 8: Stock = Format(.text, fg_Pict(9, vg_DCa))
    .Row = Row: .Col = 2: Nombre = .text
    TipText = "Bodega   : " & Trim(Left(Combo1(1).text, 50)) & vbCrLf & _
              "Producto : " & Trim(Nombre) & vbCrLf & _
              "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))
End With
End Sub
