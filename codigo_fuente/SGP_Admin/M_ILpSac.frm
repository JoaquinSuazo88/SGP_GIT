VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ILpSac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Lista de Precio Desde SAC"
   ClientHeight    =   8970
   ClientLeft      =   4305
   ClientTop       =   1170
   ClientWidth     =   7650
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8970
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "SGPADM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   13
      Top             =   3960
      Width           =   6855
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   960
         TabIndex        =   16
         Top             =   3960
         Width           =   5355
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   17
            Top             =   135
            Width           =   5250
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   3
         Top             =   285
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
         Text            =   "01/2020"
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
      Begin MSComctlLib.Toolbar Toolbar3 
         Height          =   360
         Left            =   4200
         TabIndex        =   14
         Top             =   240
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3135
         Index           =   1
         Left            =   285
         TabIndex        =   4
         Top             =   720
         Width           =   6255
         _Version        =   393216
         _ExtentX        =   11033
         _ExtentY        =   5530
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         SpreadDesigner  =   "M_ILpSac.frx":0000
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
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
         Index           =   3
         Left            =   720
         TabIndex        =   15
         Top             =   330
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "SAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   420
      TabIndex        =   5
      Top             =   0
      Width           =   6255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   225
         Width           =   4095
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   1
         Top             =   645
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
         Text            =   "01/2020"
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2265
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   1935
         _Version        =   393216
         _ExtentX        =   3413
         _ExtentY        =   3995
         _StockProps     =   64
         ButtonDrawMode  =   1
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
         MaxCols         =   2
         MaxRows         =   2
         SpreadDesigner  =   "M_ILpSac.frx":1972
         VirtualMode     =   -1  'True
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   12
         Top             =   345
         Width           =   4020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Central Compras"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Periodo"
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
         Left            =   240
         TabIndex        =   6
         Top             =   690
         Width           =   660
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8970
      Left            =   7125
      TabIndex        =   7
      Top             =   0
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   15822
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_ILpSac.frx":1CA0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   8640
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
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
      TabIndex        =   10
      Top             =   8040
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H80000003&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   5040
      Top             =   8640
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   5520
      Top             =   8640
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_ILpSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim MsgTitulo  As String
Private BtnX    As Variant

Private Sub Combo1_Click(Index As Integer)
vaSpread1(0).MaxRows = 0
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
On Error GoTo Man_Error
Me.Height = 9480
Me.Width = 7770
fg_centra Me
MsgTitulo = "Importar Lista de Precio Desde SAC"
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'Toolbar1.Buttons(1).Enabled = False
'-------> Formatear fecha
fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
fpDateTime1(0).text = Format(Date, "mm/yyyy")
fpDateTime1(1).DateTimeFormat = UserDefined
fpDateTime1(1).UserDefinedFormat = "mm/yyyy"
fpDateTime1(1).text = Format(Date, "mm/yyyy")
vaSpread1(0).Row = -1
vaSpread1(0).Col = -1
vaSpread1(0).BackColor = Shape1(0).FillColor
vaSpread1(0).MaxRows = 0

vaSpread1(1).Row = -1
vaSpread1(1).Col = -1
vaSpread1(1).BackColor = Shape1(0).FillColor
vaSpread1(1).MaxRows = 0

'-------> Llenar vector central de compras
Combo1(0).Clear
Set RS = vg_db.Execute("sgpadm_s_centralcompras_sac 1, '', ''")
Do While Not RS.EOF
   Combo1(0).AddItem IIf(IsNull(RS(0)), "", RS(0) & " - " & RS(1)) & Space(150) & "(" & fg_pone_espacio((RS(0)), 9) & ")"
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
Select Case Index
Case 0
    vaSpread1(0).MaxRows = 0
Case 1
    vaSpread1(1).MaxRows = 0
End Select
End Sub

Private Sub Text1_Change(Index As Integer)
Dim i As Long
Dim indactivo As Variant
Select Case Index
Case 1, 2
    vaSpread1(1).Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = Index
           indactivo = UCase(Trim(vaSpread1(1).Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1(1).Col = 1
           If indactivo = -1 And Trim(vaSpread1(1).text) <> "" Then
              If vaSpread1(1).RowHidden = True Then vaSpread1(1).RowHidden = False
           Else
              If vaSpread1(1).RowHidden = False Then vaSpread1(1).RowHidden = True
           End If
        Next i
        vaSpread1(1).SetActiveCell Index, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1(1).ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1(1).ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1(1).SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1(1).SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1(1).Sort -1, -1, vaSpread1(1).maxcols, vaSpread1(1).MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           If vaSpread1(1).RowHidden = True Then vaSpread1(1).RowHidden = False
       Next
       vaSpread1(1).SetActiveCell Index, vaSpread1(1).SearchCol(Index, 0, vaSpread1(1).MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(1).SetActiveCell Index, 1
    End If
    vaSpread1(1).Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim Est As Boolean
Dim i As Long, j As Long, nrosem As Long, nrolist As Long
Select Case Button.Index
Case 1 '-------> Procesar Información
    If vaSpread1(0).MaxRows < 1 Then MsgBox "Debe seleccionar el periodo SAC...", vbCritical, MsgTitulo: Exit Sub
    If vaSpread1(1).MaxRows < 1 Then MsgBox "Debe seleccionar el periodo lista precio SGPADM...", vbCritical, MsgTitulo: Exit Sub
    '-------> validar numero semana sac
    Est = False
    For i = 1 To vaSpread1(0).MaxRows
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then Est = True: Exit For
    Next i
    If Not Est Then MsgBox "Debe seleccionar de la lista Nº. Semana", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> validar lista de precios
    Est = False
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then Est = True: Exit For
    Next i
    If Not Est Then MsgBox "Debe seleccionar una lista de precio sgpadm", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> Actualizar datos
    Toolbar1.Enabled = False
    Frame2.Enabled = False
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = -1
    vaSpread1(0).Lock = True
    
    Frame1.Enabled = False
    vaSpread1(1).Row = -1
    vaSpread1(1).Col = -1
    vaSpread1(1).Lock = True
    
    fg_carga ""
    '-------> Recorrer lista sac
    For i = 1 To vaSpread1(0).MaxRows
        DoEvents
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then
           vaSpread1(0).SetActiveCell 2, vaSpread1(0).Row
           vaSpread1(0).Col = 2
           nrosem = vaSpread1(0).text
            '-------> Recorrer lista sgpadm
            For j = 1 To vaSpread1(1).MaxRows
                vaSpread1(1).Row = j
                vaSpread1(1).Col = 1
                If vaSpread1(1).text = "1" Then
                   vaSpread1(1).SetActiveCell 2, vaSpread1(1).Row
                   vaSpread1(1).Col = 3
                   nrolist = vaSpread1(1).text
                   vg_db.Execute ("sgpadm_p_actualizarlistaprecio_sac '" & vg_pais & "', '" & Trim(Mid(Combo1(0).text, 1, 150)) & "', '" & Format(fpDateTime1(0).Value, "yyyymm") & "', '" & nrosem & "', " & nrolist & ", " & Format(fpDateTime1(1).Value, "yyyymm") & ", '" & vg_NUsr & "'")
                End If
            Next j
        End If
    Next i
    fg_descarga
    MsgBox "Actualización finalizado sin problema", vbInformation + vbOKOnly, MsgTitulo
    Frame2.Enabled = True
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = -1
    vaSpread1(0).Lock = False
    
    Frame1.Enabled = True
    vaSpread1(1).Row = -1
    vaSpread1(1).Col = -1
    vaSpread1(1).Lock = False
    
    Toolbar1.Enabled = True
Case 3 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
    Frame2.Enabled = True
    vaSpread1(0).Row = -1
    vaSpread1(0).Col = -1
    vaSpread1(0).Lock = False
    
    Frame1.Enabled = True
    vaSpread1(1).Row = -1
    vaSpread1(1).Col = -1
    vaSpread1(1).Lock = False
    
    Toolbar1.Enabled = True

fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Select Case Button.Index
Case 1 '-------> Traer semana de sac
    fg_carga ""
    Set RS = vg_db.Execute("sgpadm_s_numerosemana_sac 1, '" & Trim(fg_codigocbo(Combo1, 0, 9, "")) & "', '" & Format(fpDateTime1(0).Value, "yyyymm") & "'")
    vaSpread1(0).MaxRows = 0
    Do While Not RS.EOF
       vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
       vaSpread1(0).Row = vaSpread1(0).MaxRows
       vaSpread1(0).Col = 1
       vaSpread1(0).text = "0"
       vaSpread1(0).Col = 2
       vaSpread1(0).text = Str(IIf(IsNull(RS!STACOT_NRSEM), "", RS!STACOT_NRSEM))
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    fg_descarga
End Select
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Select Case Button.Index
Case 1 '-------> Traer semana de sac
    fg_carga ""
    Text1(2).text = ""
    Set RS = vg_db.Execute("sgpadm_s_listaprecio 6, 0, " & Format(fpDateTime1(1).text, "yyyymm") & ", '" & vg_NUsr & "'")
    vaSpread1(1).Visible = 0
    vaSpread1(1).MaxRows = 0
    If Not RS.EOF Then
       Do While Not RS.EOF
          vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
          vaSpread1(1).Row = vaSpread1(1).MaxRows
          vaSpread1(1).Col = 1: vaSpread1(1).text = "0"
          vaSpread1(1).Col = 2: vaSpread1(1).text = RS!lpr_codigo & " - " & Trim(RS!lpr_nombre)
          vaSpread1(1).Col = 3: vaSpread1(1).text = RS!lpr_codigo
          vaSpread1(1).Col = 4: vaSpread1(1).text = RS!dlp_anomes
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    vaSpread1(1).Visible = True
    fg_descarga
End Select
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub
