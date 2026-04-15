VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form M_CamIng 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar y Cambiar Ingrediente"
   ClientHeight    =   5655
   ClientLeft      =   1605
   ClientTop       =   1575
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5655
   ScaleWidth      =   10440
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1035
      Index           =   1
      Left            =   960
      TabIndex        =   2
      Top             =   0
      Width           =   7935
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   3
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   4
         Top             =   555
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   1
         Left            =   3600
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   555
         Width           =   4245
         _Version        =   196608
         _ExtentX        =   7488
         _ExtentY        =   556
         Enabled         =   0   'False
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   4245
         _Version        =   196608
         _ExtentX        =   7488
         _ExtentY        =   556
         Enabled         =   0   'False
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
         BackColor       =   -2147483638
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
         AlignTextH      =   0
         AlignTextV      =   2
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin VB.Label Label1 
         Caption         =   "Ingrediente Origen"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   240
         TabIndex        =   8
         Top             =   340
         Width           =   1900
      End
      Begin VB.Label Label1 
         Caption         =   "Ingrediente Reemplazar"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   655
         Width           =   1900
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3150
         Picture         =   "M_CamIng.frx":0000
         Top             =   150
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3150
         Picture         =   "M_CamIng.frx":030A
         Top             =   450
         Width           =   480
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   1080
      Width           =   9900
      _Version        =   393216
      _ExtentX        =   17463
      _ExtentY        =   8070
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModePermanent=   -1  'True
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
      MaxCols         =   11
      MaxRows         =   20
      ProcessTab      =   -1  'True
      RestrictRows    =   -1  'True
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_CamIng.frx":0614
      UserResize      =   2
      VisibleCols     =   5
      VisibleRows     =   20
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_CamIng.frx":0DB7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_CamIng.frx":10D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_CamIng.frx":13EB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5655
      Left            =   9900
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   9975
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Borrar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_CamIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset
Dim codreceta As Long, numlinea As Long
Dim i As Integer, indsel As Integer
Dim grbruto As Double, porcaprovechamiento As Double, porccoccion As Double, porcnutricional As Double
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

Me.Height = 6030
Me.Width = 10530
fg_centra Me
fg_carga (ss)

vaSpread1.MaxRows = 0: indsel = 0: numlinea = 0

fg_descarga

End Sub
Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub
Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 0 Then image1_Click 0
    If Index = 1 Then image1_Click 1
End Select
End Sub
Private Sub fpLongInteger1_LostFocus(Index As Integer)
Select Case Index
  Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then Exit Sub
    vaSpread1.MaxRows = 0
    Set ConSql = vg_db.Execute("select PB00081.Ing_Desc " & _
                 "From  PB00080, PB00081 " & _
                 "Where PB00080.Ing_No = PB00081.Ing_No " & _
                 "and   PB00080.Ing_No=" & Val(fpLongInteger1(0).Value) & " " & _
                 "and   PB00080.Del_Ind=0 " & _
                 "and   PB00081.Del_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_ingrediente 14, " & Val(fpLongInteger1(0).Value) & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(0).Text = Trim(ConSql!Ing_Desc)
    Else
       fpayuda(0).Text = "": fpLongInteger1(0).Value = "": vg_codigo = 0
       MsgBox "Ingrediente No Existe", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
    End If
    ConSql.Close: Set ConSql = Nothing
    If Val(fpLongInteger1(0).Value) > 0 Then 'And Val(fpLongInteger1(1).Value) > 0 Then
       MoverDatosGrilla
    End If
  Case 1
    If Val(fpLongInteger1(1).Value) < 1 Then Exit Sub
    Set ConSql = vg_db.Execute("select PB00081.Ing_Desc " & _
                 "From  PB00080, PB00081 " & _
                 "Where PB00080.Ing_No = PB00081.Ing_No " & _
                 "and   PB00080.Ing_No=" & Val(fpLongInteger1(1).Value) & " " & _
                 "and   PB00080.Del_Ind=0 " & _
                 "and   PB00081.Del_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_ingrediente 14, " & Val(fpLongInteger1(1).Value) & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(1).Text = Trim(ConSql!Ing_Desc)
    Else
       fpayuda(1).Text = "": fpLongInteger1(1).Value = "": vg_codigo = 0
       MsgBox "Ingrediente No Existe", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
    End If
    ConSql.Close: Set ConSql = Nothing
    If Val(fpLongInteger1(0).Value) > 0 And vaSpread1.MaxRows < 1 Then
       MoverDatosGrilla
    End If
End Select
End Sub
Sub MoverDatosGrilla()
fg_carga "ss"
vaSpread1.MaxRows = 0
Set ConSql = vg_db.Execute("select distinct PB00078.Rcpe_Desc, PB00077.Rcpe_No, " & _
             "PB00079.Rcpe_Item_Ref_No, PB00079.Rcpe_Item_No, PB00079.Rcpe_Item_Qty, " & _
             "PB00079.Rcpe_Porc_Limpieza, " & _
             "PB00079.Rcpe_Porc_Coccion, PB00079.Diet_Item_Yld_Val " & _
             "From PB00077, PB00078, PB00079, PB00083, PB00357 " & _
             "Where PB00077.Rcpe_No = PB00078.Rcpe_No " & _
             "and   PB00077.Rcpe_No = PB00079.Rcpe_No " & _
             "and   PB00077.Rcpe_No = PB00083.Rcpe_No " & _
             "and   PB00077.Rcpe_No = PB00357.Rcpe_No " & _
             "and ((PB00083.Unit_Dfnd_No=" & Val(M_Receta.fpLongInteger1(1).Value) & ") " & _
             "and  (PB00357.Diet_Cat_No=" & vg_codregimen & " or " & vg_codregimen & "=0) " & _
             "and  (PB00079.Rcpe_Item_Ref_No=" & Val(fpLongInteger1(0).Value) & ") " & _
             "and  (PB00077.Rcpe_Cat_1_No=" & vg_auxcategoria1 & " or " & vg_auxcategoria1 & "=0) " & _
             "and  (PB00077.Rcpe_Cat_2_No=" & vg_auxcategoria2 & " or " & vg_auxcategoria2 & "=0) " & _
             "and  (PB00077.Rcpe_Cat_3_No=" & vg_auxcategoria3 & " or " & vg_auxcategoria3 & "=0) " & _
             "and  (PB00077.Rcpe_Cat_4_No=" & vg_auxcategoria4 & " or " & vg_auxcategoria4 & "=0) " & _
             "and  (PB00077.Del_Ind = 0) " & _
             "and  (PB00078.Del_Ind = 0) " & _
             "and  (PB00079.Del_Ind = 0) " & _
             "and  (PB00079.Rcpe_Item_Type_No<>2 " & _
             "And   PB00079.Rcpe_Item_Type_No<>4) " & _
             "and  (PB00083.Del_Ind = 0)) " & _
             "order by PB00078.Rcpe_Desc", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_receta 14, " & Val(fpLongInteger1(0).Value) & ", " & Val(M_Receta.fpLongInteger1(1).Value) & ", " & vg_codregimen & ", " & vg_auxcategoria1 & ", " & vg_auxcategoria2 & " , " & vg_auxcategoria3 & ", " & vg_auxcategoria4 & ", ''", , adCmdStoredProc)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
         
      vaSpread1.Col = 1
      vaSpread1.CellType = 10
      vaSpread1.TypeCheckText = " "
      vaSpread1.TypeCheckCenter = True
      vaSpread1.Value = "" ' checked
      
      vaSpread1.Col = 2
      vaSpread1.Value = "(" & ConSql!Rcpe_No & ") " & Trim(ConSql!Rcpe_Desc)
              
      vaSpread1.Col = 3
      vaSpread1.Value = ConSql!Rcpe_Item_Qty
      vaSpread1.ForeColor = &HFF0000
      
      vaSpread1.Col = 4
      vaSpread1.Value = ConSql!Rcpe_No
              
      vaSpread1.Col = 5
      vaSpread1.Value = ConSql!Rcpe_Item_Ref_No
      
      vaSpread1.Col = 6
      If IsNumeric(ConSql!Rcpe_Porc_Limpieza) Then
         vaSpread1.Value = ConSql!Rcpe_Porc_Limpieza
      Else
         vaSpread1.Value = 0
      End If
      vaSpread1.ForeColor = &HFF0000
      
      vaSpread1.Col = 7
      If IsNumeric(ConSql!Rcpe_Porc_Coccion) Then
         vaSpread1.Value = ConSql!Rcpe_Porc_Coccion
      Else
         vaSpread1.Value = 0
      End If
      vaSpread1.ForeColor = &HFF0000
      
      vaSpread1.Col = 8
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = Format(((((ConSql!Rcpe_Item_Qty * ConSql!Rcpe_Porc_Limpieza) / 100) * ConSql!Rcpe_Porc_Coccion) / 100), fg_Pict(6, 2))
'      vaSpread1.ForeColor = &HFF0000
            
      vaSpread1.Col = 9
      If IsNumeric(ConSql!Diet_Item_Yld_Val) Then
         vaSpread1.Value = ConSql!Diet_Item_Yld_Val
      Else
         vaSpread1.Value = 0
      End If
      vaSpread1.ForeColor = &HFF0000
      
      vaSpread1.Col = 10
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = Format(((ConSql!Diet_Item_Yld_Val / 100) * ConSql!Rcpe_Item_Qty), fg_Pict(6, 2))
      
      vaSpread1.Col = 11
      vaSpread1.Value = ConSql!Rcpe_Item_No
      
      ConSql.MoveNext
   Loop
   ConSql.Close: Set ConSql = Nothing: fg_descarga
Else
    ConSql.Close: Set ConSql = Nothing
    fg_descarga
    MsgBox "No Existe Ingrediente Origen en Recetario", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
End If
End Sub
Private Sub image1_Click(Index As Integer)
On Error GoTo Man_Error

Select Case Index
  Case 0
    vg_codigo = 0
    vg_opcioning = 1: vg_opcioningrediente = 1
    vg_left = fpayuda(0).Left + 1770
    B_CamIng.Show 1
    If vg_codigo = 0 Then Exit Sub
    vaSpread1.MaxRows = 0
    fpLongInteger1(0).Value = vg_codigo
    Set ConSql = vg_db.Execute("select PB00081.Ing_Desc " & _
                 "From PB00080, PB00081 " & _
                 "Where PB00080.Ing_No = PB00081.Ing_No " & _
                 "and   PB00080.Ing_No=" & vg_codigo & " " & _
                 "and   PB00080.Del_Ind=0 " & _
                 "and   PB00081.Del_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_ingrediente 14, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(0).Text = Trim(ConSql!Ing_Desc)
    Else
       fpayuda(0).Text = "": fpLongInteger1(0).Value = "": vg_codigo = 0
       MsgBox "Ingrediente No Existe", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
    End If
    ConSql.Close: Set ConSql = Nothing
    If Val(fpLongInteger1(0).Value) > 0 Then 'And Val(fpLongInteger1(1).Value) > 0 Then
       MoverDatosGrilla
    End If
  Case 1
    vg_codigo = 0
    vg_opcioning = 1: vg_opcioningrediente = 1
    vg_left = fpayuda(1).Left + 1770
    B_CamIng.Show 1
    If vg_codigo = 0 Then Exit Sub
'    vaSpread1.MaxRows = 0
    fpLongInteger1(1).Value = vg_codigo
    Set ConSql = vg_db.Execute("select PB00081.Ing_Desc " & _
                 "From PB00080, PB00081 " & _
                 "Where PB00080.Ing_No = PB00081.Ing_No " & _
                 "and   PB00080.Ing_No=" & vg_codigo & " " & _
                 "and   PB00080.Del_Ind=0 " & _
                 "and   PB00081.Del_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_ingrediente 14, " & vg_codigo & ", ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       fpayuda(1).Text = Trim(ConSql!Ing_Desc)
    Else
       fpayuda(1).Text = "": fpLongInteger1(1).Value = "": vg_codigo = 0
       MsgBox "Ingrediente No Existe", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente"
    End If
    ConSql.Close: Set ConSql = Nothing
    If Val(fpLongInteger1(0).Value) > 0 And vaSpread1.MaxRows < 1 Then
       MoverDatosGrilla
    End If
End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
'vg_Area.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index
  Case 1, 3
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Set ConSql = vg_db.Execute("select PB00081.Ing_Desc " & _
                 "From PB00080, PB00081 " & _
                 "Where PB00080.Ing_No = PB00081.Ing_No " & _
                 "and   PB00080.Ing_No=" & Val(fpLongInteger1(0).Value) & " " & _
                 "and   PB00080.Del_Ind=0 " & _
                 "and   PB00081.Del_Ind=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_ingrediente 14, " & Val(fpLongInteger1(0).Value) & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: vaSpread1.MaxRows = 0: MsgBox "No Existe Ingredientes", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente": Exit Sub
    ConSql.Close: Set ConSql = Nothing
    If Val(fpLongInteger1(1).Value) > 0 Then
       Set ConSql = vg_db.Execute("select PB00081.Ing_Desc " & _
                    "From PB00080, PB00081 " & _
                    "Where PB00080.Ing_No = PB00081.Ing_No " & _
                    "and   PB00080.Ing_No=" & Val(fpLongInteger1(1).Value) & " " & _
                    "and   PB00080.Del_Ind=0 " & _
                    "and   PB00081.Del_Ind=0", , adCmdText)
'       Set ConSql = vg_db.Execute("sod_s_ingrediente 14, " & Val(fpLongInteger1(1).Value) & ", ''", , adCmdStoredProc)
       If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: vaSpread1.MaxRows = 0: MsgBox "No Existe Ingredientes a Reemplazar", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente": Exit Sub
          ConSql.Close: Set ConSql = Nothing
    End If
    Set ConSql = vg_db.Execute("select distinct PB00078.Rcpe_Desc, PB00077.Rcpe_No, " & _
                 "PB00079.Rcpe_Item_Ref_No, PB00079.Rcpe_Item_Qty, PB00079.Rcpe_Porc_Limpieza, " & _
                 "PB00079.Rcpe_Porc_Coccion, PB00079.Diet_Item_Yld_Val " & _
                 "From PB00077, PB00078, PB00079, PB00083, PB00357 " & _
                 "Where PB00077.Rcpe_No = PB00078.Rcpe_No " & _
                 "and   PB00077.Rcpe_No = PB00079.Rcpe_No " & _
                 "and   PB00077.Rcpe_No = PB00083.Rcpe_No " & _
                 "and   PB00077.Rcpe_No = PB00357.Rcpe_No " & _
                 "and ((PB00083.Unit_Dfnd_No=" & Val(M_Receta.fpLongInteger1(1).Value) & ") " & _
                 "and  (PB00357.Diet_Cat_No=" & vg_codregimen & " or " & vg_codregimen & "=0) " & _
                 "and  (PB00079.Rcpe_Item_Ref_No=" & Val(fpLongInteger1(0).Value) & ") " & _
                 "and  (PB00077.Rcpe_Cat_1_No=" & vg_auxcategoria1 & " or " & vg_auxcategoria1 & "=0) " & _
                 "and  (PB00077.Rcpe_Cat_2_No=" & vg_auxcategoria2 & " or " & vg_auxcategoria2 & "=0) " & _
                 "and  (PB00077.Rcpe_Cat_3_No=" & vg_auxcategoria3 & " or " & vg_auxcategoria3 & "=0) " & _
                 "and  (PB00077.Rcpe_Cat_4_No=" & vg_auxcategoria4 & " or " & vg_auxcategoria4 & "=0) " & _
                 "and  (PB00077.Del_Ind = 0) " & _
                 "and  (PB00078.Del_Ind = 0) " & _
                 "and  (PB00079.Del_Ind = 0) " & _
                 "and  (PB00079.Rcpe_Item_Type_No<>2 " & _
                 "And   PB00079.Rcpe_Item_Type_No<>4) " & _
                 "and  (PB00083.Del_Ind = 0)) " & _
                 "order by PB00078.Rcpe_Desc", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_receta 14, " & Val(fpLongInteger1(0).Value) & ", " & Val(M_Receta.fpLongInteger1(1).Value) & ", " & vg_codregimen & ", " & vg_auxcategoria1 & ", " & vg_auxcategoria2 & " , " & vg_auxcategoria3 & ", " & vg_auxcategoria4 & ", ''", , adCmdStoredProc)
    If ConSql.EOF Then ConSql.Close: Set ConSql = Nothing: vaSpread1.MaxRows = 0: MsgBox "No Existe Ingredientes Origen en Recetario", vbExclamation + vbOKOnly, "Buscar y Cambiar Ingrediente": Exit Sub
    ConSql.Close: Set ConSql = Nothing
    indsel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.Text = "1" Then indsel = 1: Exit For
    Next i
    If Button.Index = 1 Then
       If indsel = 0 Then MsgBox "Seleccione Uno o Más Recetas a Cambiar", vbCritical + vbOKOnly, "Cambio Ingrediente": Exit Sub
       If Val(fpLongInteger1(1).Value) > 0 Then
          msg = " Esta Seguro Reemplazar " & "(" & Trim(fpayuda(0).Text) & ")" & " Por " & "(" & Trim(fpayuda(1).Text) & ")" & " En Las Recetas Seleccionadas ?"
       Else
          msg = " Esta Seguro Remplazar Datos en " & "(" & Trim(fpayuda(0).Text) & ")" & " "
       End If
    ElseIf Button.Index = 3 Then
       If indsel = 0 Then MsgBox "Seleccione Uno o Más Recetas a Eliminar", vbCritical + vbOKOnly, "Eliminar Ingrediente": Exit Sub
       msg = " Esta Seguro Eliminar " & "(" & Trim(fpayuda(0).Text) & ")" & " En Las Recetas Seleccionadas ?"
    End If
    Style = vbYesNo + vbQuestion + vbDefaultButton2
    Help = "DEMO.HLP"
    Ctxt = 1000
    ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
    If ws_respuesta = vbNo Then Exit Sub
    fg_carga (ss)
    vg_db.BeginTrans
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 1
          If vaSpread1.Text = "1" Then
             vaSpread1.Col = 3: grbruto = Val(vaSpread1.Value)
             vaSpread1.Col = 4: codreceta = Val(vaSpread1.Value)
             vaSpread1.Col = 6: porcaprovechamiento = Val(vaSpread1.Value)
             vaSpread1.Col = 7: porccoccion = Val(vaSpread1.Value)
             vaSpread1.Col = 9: porcnutricional = Val(vaSpread1.Value)
             vaSpread1.Col = 11: numlinea = Val(vaSpread1.Value)
             If Button.Index = 1 Then
                If Val(fpLongInteger1(1).Value) > 0 Then
                   codingrediente = Val(fpLongInteger1(1).Value)
'                   vg_db.Execute "sod_p_cambioingreceta " & codreceta & ", " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & grbruto & ", " & porcaprovechamiento & ", " & porccoccion & ", " & porcnutricional & "", vg_ModoOpen
                Else
                   codingrediente = Val(fpLongInteger1(0).Value)
'                   vg_db.Execute "sod_p_cambioingreceta " & codreceta & ", " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(0).Value) & ", " & grbruto & ", " & porcaprovechamiento & ", " & porccoccion & ", " & porcnutricional & "", vg_ModoOpen
                End If
             
                Set ConSql = vg_db.Execute("select * " & _
                             "from PB00080 " & _
                             "where Ing_No=" & codingrediente & " " & _
                             "and Del_Ind=0", , adCmdText)
                If Not ConSql.EOF Then
                   ConSql.Close: Set ConSql = Nothing
                   Set ConSql = vg_db.Execute("select * " & _
                                "from PB00077 " & _
                                "where Rcpe_No=" & codreceta & " " & _
                                "and Del_Ind=0", , adCmdText)
                   If Not ConSql.EOF Then
                      ConSql.Close: Set ConSql = Nothing
                      Set ConSql = vg_db.Execute("select PB00383.Diet_Item_Yld_Val, " & _
                                   "PB00383.Diet_Item_No " & _
                                   "From  PB00383 " & _
                                   "where PB00383.Ing_No=" & codingrediente & " " & _
                                   "and   PB00383.Diet_Item_Ind=1", , adCmdText)
                      If Not ConSql.EOF Then
                         Diet_Item_No = ConSql!Diet_Item_No
                         Diet_Item_Yld_Val = ConSql!Diet_Item_Yld_Val
                      End If
                      ConSql.Close: Set ConSql = Nothing
                   Else
                      ConSql.Close: Set ConSql = Nothing
                      Diet_Item_No = 0: Diet_Item_Yld_Val = 0
                   End If
                Else
                   ConSql.Close: Set ConSql = Nothing
                End If
                valor1 = 0: valor2 = 0: valor3 = 0
                valor1 = ((grbruto * porcaprovechamiento) / 100)
                valor2 = ((porccoccion) / 100)
                valor3 = valor1 * valor2
                vg_db.Execute "Update PB00079 " & _
                              "set Diet_Item_Yld_Val=" & porcnutricional & ", " & _
                              "Diet_Item_No=" & Diet_Item_No & ", " & _
                              "Rcpe_Item_Ref_No=" & codingrediente & ", " & _
                              "Rcpe_Item_Qty=" & grbruto & ", " & _
                              "Rcpe_Porc_Coccion=" & porccoccion & ", " & _
                              "Rcpe_Porc_Limpieza=" & porcaprovechamiento & ", " & _
                              "Rcpe_Cant_Coccion=((((" & grbruto & "*" & porcaprovechamiento & ")/100)*" & porccoccion & ")/100)  " & _
                              "where PB00079.Rcpe_No=" & codreceta & " " & _
                              "and   PB00079.Rcpe_Item_Ref_No=" & Val(fpLongInteger1(0).Value) & " " & _
                              "and   PB00079.Rcpe_Item_No=" & numlinea & " " & _
                              "and   PB00079.Rcpe_Item_Type_No<>2 " & _
                              "and   PB00079.Rcpe_Item_Type_No<>4 " & _
                              "and   PB00079.Del_Ind=0"
             ElseIf Button.Index = 3 Then
                vg_db.Execute "delete PB00079 " & _
                              "where Rcpe_No=" & codreceta & " " & _
                              "and   Rcpe_Item_Ref_No=" & Val(fpLongInteger1(0).Value) & " " & _
                              "and   Rcpe_Item_No=" & numlinea & ""
             End If
          End If
      Next i
      fg_descarga
      If Button.Index = 1 Then
         MsgBox "Cambiar Ingredientes Finalizo Sin Problema", vbInformation + vbOKOnly, "Cambiar Ingredientes"
      ElseIf Button.Index = 3 Then
         MsgBox "Eliminación de Ingredientes Finalizo Sin Problema", vbInformation + vbOKOnly, "Eliminar Ingredientes"
      End If
      indsel = 0
      vaSpread1.MaxRows = 0
    vg_db.CommitTrans
  Case 5
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then
   If indsel = 0 Then
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = 10
          vaSpread1.TypeCheckText = ""
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "1" ' checked
      Next i
      indsel = 1
   Else
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = 10
          vaSpread1.TypeCheckText = " "
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "" ' checked
      Next i
      indsel = 0
   End If
End If
End Sub
Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal newrow As Long, Cancel As Boolean)
Dim paporte As Double, cbruto As Double, paporv As Double, pcoccion As Double, cantservida As Double

If vaSpread1.MaxRows < 1 Then Exit Sub
' *** Calcular Gramaje Neto *** '
          
paporte = 0: cbruto = 0: paporv = 0: pcoccion = 0: cantservida = 0

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 3
cbruto = vaSpread1.Value

vaSpread1.Col = 9
paporte = vaSpread1.Value

vaSpread1.Col = 10
vaSpread1.CellType = 5
vaSpread1.TypeHAlign = 1
vaSpread1.Value = Format(((paporte / 100) * cbruto), fg_Pict(6, 2))

' *** Calcular % Limpieza & Cocción *** '
          
vaSpread1.Col = 6
paporv = vaSpread1.Value
cantservida = CCur((paporv / 100) * cbruto)
          
vaSpread1.Col = 7
pcoccion = vaSpread1.Value
cantservida = CCur((pcoccion / 100) * cantservida)
          
vaSpread1.Col = 8
vaSpread1.CellType = 5
vaSpread1.TypeHAlign = 1
vaSpread1.Value = Format(cantservida, fg_Pict(6, 2))

End Sub
