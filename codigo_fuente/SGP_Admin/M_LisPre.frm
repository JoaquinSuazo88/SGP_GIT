VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_LisPre 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Precio"
   ClientHeight    =   9615
   ClientLeft      =   3900
   ClientTop       =   1200
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   10140
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7335
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   2175
      Width           =   9855
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   1560
         TabIndex        =   14
         Top             =   6840
         Width           =   4605
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   15
            Top             =   135
            Width           =   4500
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   720
         TabIndex        =   12
         Top             =   6840
         Width           =   790
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   13
            Top             =   135
            Width           =   690
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6525
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   9375
         _Version        =   393216
         _ExtentX        =   16536
         _ExtentY        =   11509
         _StockProps     =   64
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
         MaxCols         =   9
         MaxRows         =   1000000
         SpreadDesigner  =   "M_LisPre.frx":0000
         VirtualMode     =   -1  'True
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   6960
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   6120
         Top             =   7080
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1680
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   9855
      Begin VB.CheckBox Check1 
         Caption         =   "Activo"
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
         Left            =   7560
         TabIndex        =   16
         Top             =   315
         Width           =   975
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   6660
         _Version        =   196608
         _ExtentX        =   11747
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
         BackColor       =   -2147483628
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
         MaxLength       =   80
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
         Left            =   1920
         TabIndex        =   2
         Top             =   960
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
         Text            =   "08/2025"
         DateCalcMethod  =   0
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   240
         Width           =   2070
         _Version        =   196608
         _ExtentX        =   3651
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
         BackColor       =   -2147483628
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
         NoSpecialKeys   =   0
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3600
         TabIndex        =   11
         Top             =   960
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BorderStyle     =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   6
         Left            =   6720
         TabIndex        =   20
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label Label1 
         Caption         =   "Cencos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   5160
         TabIndex        =   19
         Top             =   1390
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   18
         Top             =   1395
         Width           =   45
      End
      Begin VB.Label Label1 
         Caption         =   "Central Compras "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   17
         Top             =   1390
         Width           =   1455
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   4080
         Picture         =   "M_LisPre.frx":1B54
         Top             =   150
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Correlativo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Lista Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   1035
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   675
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10140
      _ExtentX        =   17886
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_LisPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim MsgTitulo  As String, modo As String
Dim Est As Boolean

Private Sub Check1_Click()
If Toolbar1.Buttons(12).Visible = True Or Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 10, 0, modo
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 10095
Me.Width = 10230
Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Lista de Precio"
vaSpread1.MaxRows = 0
modo = ""
Est = True
fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
fpDateTime1(0).text = Format(Date, "mm/yyyy")
Label1(4).Caption = ""
Label1(6).Caption = ""
Gl_Mo_Botones Me, 10
Gl_Ac_Botones Me, 10, 1, modo
Est = False
End Sub

Sub MoverDetalleListaPrecio(op As Long, codigo As Long)
Dim RS As New ADODB.Recordset
On Error GoTo Man_Error
fg_carga ""
Image1(0).Enabled = False
Label1(4).Caption = ""
Label1(6).Caption = ""
Check1.Value = IIf(op = 2, 1, 0)
vaSpread1.Visible = False
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.MaxRows = 0
Est = True
Set RS = vg_db.Execute("sgpadm_s_listaprecio 9, " & codigo & ", " & Val(Format(fpDateTime1(0).Value, "yyyymm")) & ", '" & vg_NUsr & "'")
If Not RS.EOF Then fpText1(Index).text = Trim(RS!lpr_nombre)
RS.Close: Set RS = Nothing
Set RS = vg_db.Execute("sgpadm_s_listaprecio " & op & ", " & codigo & ", " & Val(Format(fpDateTime1(0).Value, "yyyymm")) & ", '" & vg_NUsr & "'")
If Not RS.EOF Then
   Check1.Value = IIf(IsNull(RS!lpr_activo) Or RS!lpr_activo = "0", 0, 1)
'   '-------> Abrir base sac
'   AbrirBaseSac
   '-------> central de compras
   Label1(4).Caption = ""
'   RS1.Open "SELECT TABCEN_DSCEN FROM TABCEN WHERE TABCEN_CDCEN = '" & RS!dlp_codcec & "'", vg_dbsac, adOpenStatic
'   If Not RS1.EOF Then Label1(4).Caption = Trim(RS1!TABCEN_DSCEN)
'   RS1.Close: Set RS1 = Nothing
   '-------> centro de costo
   Label1(6).Caption = ""
'   RS1.Open "SELECT CADFIL_NMFIL FROM CADFIL WHERE CADFIL_CDFIL = '" & RS!dlp_codcco & "'", vg_dbsac, adOpenStatic
'   If Not RS1.EOF Then Label1(6).Caption = Trim(RS1!CADFIL_NMFIL)
'   RS1.Close: Set RS1 = Nothing
   Do While Not RS.EOF
      DoEvents
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
       
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!pro_codigo), "", Trim(RS!pro_codigo))
                
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!pro_nombre), "", Trim(RS!pro_nombre))
                
      vaSpread1.Col = 3
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!uni_nomcor), "", Trim(RS!uni_nomcor))
                
      vaSpread1.Col = 4
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.ForeColor = &HFF0000
      vaSpread1.text = IIf(IsNull(RS!dlp_precio), Format(RS!dlp_precio, fg_Pict(9, vg_DPr)), Format(RS!dlp_precio, fg_Pict(9, vg_DCa)))
         
      vaSpread1.Col = 5
'      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(op = 2, "1", "0")
       
      vaSpread1.Col = 6
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!dlp_codcec), "", Trim(RS!dlp_codcec))
       
      vaSpread1.Col = 7
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(IsNull(RS!dlp_codcco), "", Trim(RS!dlp_codcco))
       
      vaSpread1.Col = 8
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = IIf(IsNull(RS!dlp_dtsac) Or Trim(RS!dlp_dtsac) = "", "", Mid(RS!dlp_dtsac, 5, 2) & "/" & Mid(RS!dlp_dtsac, 1, 4))
       
      vaSpread1.Col = 9
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = IIf(IsNull(RS!dlp_nrosem) Or RS!dlp_nrosem = 0, "", RS!dlp_nrosem)
       
      RS.MoveNext
   Loop
   Gl_Ac_Botones Me, 10, IIf(op = 1, 1, 0), modo
Else
   RS.Close: Set RS = Nothing
   Set RS = vg_db.Execute("sgpadm_s_listaprecio 9, " & codigo & ", " & Val(Format(fpDateTime1(0).Value, "yyyymm")) & ", '" & vg_NUsr & "'")
   If Not RS.EOF Then
      vaSpread1.MaxRows = 0
      Gl_Ac_Botones Me, 10, 4, modo
   Else
      Gl_Ac_Botones Me, 10, 2, modo
   End If
End If
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
Est = False
Image1(0).Enabled = True
fg_descarga
Exit Sub
Man_Error:
Image1(0).Enabled = True
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
If Est Or modo = "A" Then Exit Sub
Dim RS As New ADODB.Recordset
Dim anomes As Long
anomes = Format(fpDateTime1(0).text, "yyyymm")
Set RS = vg_db.Execute("sgpadm_s_listaprecio 7, " & Val(fpLongInteger1(0).Value) & ", " & anomes & ", '" & vg_NUsr & "'")
If Not RS.EOF Then
   RS.Close: Set RS = Nothing
   MoverDetalleListaPrecio 1, Val(fpLongInteger1(0).Value)
Else
   RS.Close: Set RS = Nothing
   vaSpread1.MaxRows = 0
   Gl_Ac_Botones Me, 10, 4, modo
End If
End Sub

Private Sub fpText1_Change(Index As Integer)
If Toolbar1.Buttons(12).Visible = True Or Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 10, 0, modo
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = Image1(0).Left + 1920
    B_TabEst.LlenaDatos "b_listaprecio", "lpr_", "Lista de Precio", "LisPre"
    B_TabEst.Show 1
    Me.Refresh
    If Val(vg_codigo) = 0 Then Exit Sub
    Est = True
    Text1(1).text = "": Text1(2).text = ""
    fpLongInteger1(Index) = Val(vg_codigo)
    fpText1(Index).text = vg_nombre
    fpDateTime1(0).text = vg_ames
    MoverDetalleListaPrecio 1, Val(vg_codigo)
    Est = False
    modo = "M"
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2
    vaSpread1.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim i As Long, codigo As Long, Nombre As String, codpro As String, precio As Double, anomes As Long
Dim codcen As String, codcco As String, fecsac As String, nrosem As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '-------> Agregar
    Text1(1).text = "": Text1(2).text = ""
    fpLongInteger1(0).text = ""
    fpText1(0).text = ""
    Check1.Value = 1
    Image1(0).Enabled = False
    fpDateTime1(0).Enabled = True
    vaSpread1.MaxRows = 0
    modo = "A"
    MoverDetalleListaPrecio 2, 0
    fpText1(0).SetFocus
Case 3 '-------> Modificar
    If vaSpread1.MaxRows < 1 Then MsgBox "Debe seleccionar una lista precio...", vbCritical, MsgTitulo: Exit Sub
    modo = "M"
    Image1(0).Enabled = False
    Gl_Ac_Botones Me, 10, 0, modo
Case 5 '-------> Borrar
    Text1(1).text = "": Text1(2).text = ""
    codigo = Val(fpLongInteger1(0).Value)
    anomes = Format(fpDateTime1(0).text, "yyyymm")
    If vaSpread1.MaxRows < 1 Or codigo = 0 Or anomes = 0 Then Exit Sub
    If MsgBox("Eliminar Dato", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.Execute "DELETE FROM b_detlistaprecio WHERE dlp_codigo = " & codigo & " AND dlp_anomes = " & anomes & ""
    Set RS = vg_db.Execute("sgpadm_s_listaprecio 5, " & codigo & ", 0, '" & vg_NUsr & "'")
    If RS.EOF Then
       Est = True
       RS.Close: Set RS = Nothing
       vg_db.Execute "DELETE FROM b_listaprecio WHERE lpr_codigo = " & codigo & ""
       fpLongInteger1(0).text = ""
       fpText1(0).text = ""
       Est = False
    Else
       RS.Close: Set RS = Nothing
       Gl_Ac_Botones Me, 10, 4, modo
    End If
    vaSpread1.MaxRows = 0
    fpText1(0).SetFocus
Case 7 '-------> Actualizar
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Text1(1).text = "": Text1(2).text = ""
    MoverDetalleListaPrecio 1, Val(fpLongInteger1(0).Value)
Case 10 '-------> Cancelar
    fg_carga ""
    Text1(1).text = "": Text1(2).text = ""
    If modo = "A" Then
       Check1.Value = 0
       vaSpread1.MaxRows = 0
    Else
       MoverDetalleListaPrecio 1, Val(fpLongInteger1(0).Value)
    End If
    modo = "": Gl_Ac_Botones Me, 10, 1, modo
    If vaSpread1.MaxRows = 0 And Val(fpLongInteger1(0).Value) > 0 Then Gl_Ac_Botones Me, 10, 4, modo
    Image1(0).Enabled = True
    fg_descarga
Case 12 '-------> Grabar datos
    If LimpiaDato(Trim(fpText1(0).text)) = "" Or LimpiaDato(Trim(fpDateTime1(0).text)) = "" Then MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
    Text1(1).text = "": Text1(2).text = ""
    Frame3.Visible = False: Frame4.Visible = False
    Nombre = LimpiaDato(Trim(fpText1(0).text))
    anomes = Format(fpDateTime1(0).text, "yyyymm")
    fg_carga ""
    Bar1(0).Visible = True: Bar1(0).Value = 0
    codigo = Val(fpLongInteger1(0).Value)
    If codigo = 0 Then modo = "A"
    If modo = "A" Then
       Set RS = vg_db.Execute("sgpadm_s_listaprecio 3, 0, 0, '" & vg_NUsr & "'")
       If Not RS.EOF Then RS.MoveFirst: codigo = RS!lpr_codigo + 1 Else codigo = 1
       RS.Close: Set RS = Nothing
       vg_db.Execute "INSERT INTO b_listaprecio (lpr_codigo, lpr_nombre, lpr_codcec, lpr_codcco, lpr_activo) " & _
                     "VALUES ('" & codigo & "', '" & Nombre & "', null, null, '" & IIf(Check1.Value = 0, "0", "1") & "')"
       
       vg_db.Execute "INSERT INTO b_detlistaprecio (dlp_codigo, dlp_anomes, dlp_codpro, dlp_precio, dlp_usuario, dlp_codcec, dlp_codcco, dlp_dtsac, dlp_nrosem) SELECT " & codigo & ", " & anomes & ", pro_codigo, 0, '" & vg_NUsr & "', null, null, null, null " & _
                     "FROM b_productos WHERE (pro_fecven>" & Format(Date, "yyyymmdd") & " OR pro_fecven <= 0)"
       
       For i = 1 To vaSpread1.MaxRows
           DoEvents
           Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
           vaSpread1.Row = i
           
           vaSpread1.Col = 1: codpro = vaSpread1.text
           vaSpread1.Col = 4: precio = vaSpread1.text
           If precio > 0 Then
              vg_db.Execute "UPDATE b_detlistaprecio SET dlp_precio = " & precio & " WHERE dlp_codigo = " & codigo & " AND dlp_anomes = " & anomes & " AND dlp_codpro = '" & codpro & "'"
           End If
       Next i
       fpLongInteger1(0).Value = codigo
    Else
       codigo = Val(fpLongInteger1(0).Value)
       vg_db.Execute "UPDATE  b_listaprecio SET lpr_nombre = '" & Nombre & "', lpr_activo = '" & IIf(Check1.Value = 0, "0", "1") & "' WHERE lpr_codigo = " & codigo & ""
       For i = 1 To vaSpread1.MaxRows
           DoEvents
           Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
           vaSpread1.Row = i
           vaSpread1.Col = 5
           If vaSpread1.text = "1" Then
              vaSpread1.Col = 1: codpro = vaSpread1.text
              vaSpread1.Col = 4: precio = vaSpread1.text
              codcen = "": vaSpread1.Col = 6: codcen = Trim(vaSpread1.text)
              codcco = "": vaSpread1.Col = 7: codcco = Trim(vaSpread1.text)
              fecsac = "": vaSpread1.Col = 8: fecsac = Trim(vaSpread1.text)
              nrosem = 0: vaSpread1.Col = 9: nrosem = Val(vaSpread1.text)
              vg_db.Execute "DELETE FROM b_detlistaprecio WHERE dlp_codigo = " & codigo & " AND dlp_anomes = " & anomes & " AND dlp_codpro = '" & codpro & "'"
              vg_db.Execute "INSERT INTO b_detlistaprecio (dlp_codigo , dlp_anomes, dlp_codpro, dlp_precio, dlp_usuario, dlp_codcec, dlp_codcco, dlp_dtsac, dlp_nrosem) " & _
                            "VALUES ('" & codigo & "', " & anomes & ", '" & codpro & "', " & precio & ", '" & vg_NUsr & "', '" & codcen & "', '" & codcco & "', '" & fecsac & "', " & nrosem & ")"
              vaSpread1.Col = 5: vaSpread1.text = "0"
           End If
       Next i
    End If
    Bar1(0).Visible = False: Bar1(0).Value = 0
    modo = "": Gl_Ac_Botones Me, 10, 1, modo
    Image1(0).Enabled = True
    fg_descarga
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, MsgTitulo
    Frame3.Visible = True: Frame4.Visible = True
Case 15 '-------> Copiar lista precio
    M_CLisPr.LlenaDatos Val(fpLongInteger1(0))
    M_CLisPr.Show 1
Case 19 '-------> Imprimir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    I_ListaPrecio Val(fpLongInteger1(0).Value), Format(fpDateTime1(0).text, "yyyymm")
Case 22 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
Case "Desde SAC"
Case "Desde Excel"
    M_ImLprE.Show 1
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 '-------> Agregar nuevo mes
    MoverDetalleListaPrecio 2, Val(fpLongInteger1(0).Value) '0
    Gl_Ac_Botones Me, 10, 5, modo
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Or Est Then Exit Sub
vaSpread1.Row = Row
vaSpread1.Col = 5: vaSpread1.text = "1"
vaSpread1.Col = 8: vaSpread1.text = ""
vaSpread1.Col = 9: vaSpread1.text = ""
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 10, 0, modo
End Sub
