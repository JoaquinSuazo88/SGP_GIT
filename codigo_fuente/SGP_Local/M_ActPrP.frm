VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_ActPrP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Actualizar Lista de Precio Planificación"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1850
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   12255
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         ItemData        =   "M_ActPrP.frx":0000
         Left            =   1560
         List            =   "M_ActPrP.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   960
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "M_ActPrP.frx":0004
         Left            =   1560
         List            =   "M_ActPrP.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   600
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "M_ActPrP.frx":0008
         Left            =   1560
         List            =   "M_ActPrP.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   240
         Width           =   4095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Indicadores Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6000
         TabIndex        =   11
         Top             =   120
         Width           =   6135
         Begin VB.Label Label1 
            Caption         =   "Producto con precio cero ó bien el ingrediente no tiene asignado producto."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   4
            Left            =   3840
            TabIndex        =   14
            Top             =   180
            Width           =   2175
         End
         Begin VB.Label Label1 
            Caption         =   "Proceso OK"
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
            Left            =   2235
            TabIndex        =   13
            Top             =   285
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   2
            Left            =   3360
            Picture         =   "M_ActPrP.frx":000C
            Stretch         =   -1  'True
            Top             =   240
            Width           =   360
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   1
            Left            =   1800
            Picture         =   "M_ActPrP.frx":6296
            Stretch         =   -1  'True
            Top             =   240
            Width           =   360
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   0
            Left            =   120
            Picture         =   "M_ActPrP.frx":C520
            Stretch         =   -1  'True
            Top             =   240
            Width           =   360
         End
         Begin VB.Label Label1 
            Caption         =   "Sin procesar"
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
            Left            =   555
            TabIndex        =   12
            Top             =   300
            Width           =   1095
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Top             =   1400
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
         Text            =   "04/2011"
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
         Left            =   3360
         TabIndex        =   5
         Top             =   1320
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
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
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   20
         Top             =   1050
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Régimen"
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
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   690
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   16
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo Planificación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Planificación"
      TabPicture(0)   =   "M_ActPrP.frx":127AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Bar1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "vaSpread2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Asociar Ingredientes"
      TabPicture(1)   =   "M_ActPrP.frx":127C6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(0)"
      Tab(1).Control(1)=   "vaSpread3"
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "Frame5"
      Tab(1).ControlCount=   5
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   -68980
         TabIndex        =   25
         Top             =   5560
         Width           =   4605
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   4500
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   -74760
         TabIndex        =   23
         Top             =   5560
         Width           =   790
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   24
            Top             =   135
            Width           =   690
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   -73880
         TabIndex        =   21
         Top             =   5560
         Width           =   4605
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   22
            Top             =   135
            Width           =   4500
         End
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   4335
         Left            =   -74835
         TabIndex        =   9
         Top             =   1200
         Width           =   11895
         _Version        =   393216
         _ExtentX        =   20981
         _ExtentY        =   7646
         _StockProps     =   64
         ButtonDrawMode  =   1
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
         MaxCols         =   7
         SpreadDesigner  =   "M_ActPrP.frx":127E2
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   135
         Left            =   7080
         TabIndex        =   8
         Top             =   5760
         Visible         =   0   'False
         Width           =   1455
         _Version        =   393216
         _ExtentX        =   2566
         _ExtentY        =   238
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
         SpreadDesigner  =   "M_ActPrP.frx":1429F
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5175
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   11940
         _Version        =   393216
         _ExtentX        =   21061
         _ExtentY        =   9128
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
         MaxCols         =   11
         SpreadDesigner  =   "M_ActPrP.frx":169C7
         VirtualScrollBuffer=   -1  'True
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   5680
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Planificación"
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
         Left            =   -74835
         TabIndex        =   10
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   3240
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
            Picture         =   "M_ActPrP.frx":19ED1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_ActPrP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String, est As Boolean
Dim Msgtitulo  As String
Dim codSubseg, codreg, codser As Integer

Private Sub Combo1_Click(Index As Integer)
Select Case Index
Case 0
    If Combo1(0).text <> "Todos" Then
       codSubseg = Val(fg_codigocbo(Combo1, 0, 10, ""))
       If codSubseg = 0 Then
          codSubseg = Val(fg_codigocbo(Combo1, 0, 10, ""))
        End If
    End If
Case 1
    codreg = Val(fg_codigocbo(Combo1, 1, 10, ""))
Case 2
    codser = Val(fg_codigocbo(Combo1, 2, 10, ""))
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
Me.HelpContextID = vg_OpcM
Me.Height = 9045
Me.Width = 12600
Msgtitulo = "Actualizar Lista de Precio Planificación"
fg_centra Me
modo = ""
est = True
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 3, modo
SSTab1.TabVisible(1) = False
vaSpread1.MaxRows = 0
est = False
codSubseg = 0: codreg = 0: codser = 0
'Llenado primer combo Sub-Segmento
Set RS = vg_db.Execute("sgpadm_s_subsegmento 11, 0, '', '" & vg_Indppr & "'")
Combo1(0).Clear
Combo1(0).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 10) & ")"
Do While Not RS.EOF
    Combo1(0).AddItem IIf(IsNull(RS!sub_codigo), "", RS!sub_codigo & " - " & Trim(RS!sub_nombre)) & Space(150) & "(" & fg_pone_cero(Str(RS!sub_codigo), 10) & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(0).ListIndex = 0

'Llenado primer combo Regimen
Combo1(1).Clear
Set RS = vg_db.Execute("sgpadm_s_regimen 7, " & vg_Indppr & ",''")
Combo1(1).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 10) & ")"
Do While Not RS.EOF
    Combo1(1).AddItem IIf(IsNull(RS!reg_nombre), "", RS!reg_codigo & " - " & Trim(RS!reg_nombre)) & Space(150) & "(" & fg_pone_cero(Str(RS!reg_codigo), 10) & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(1).ListIndex = 0

'Llenado primer combo Servicio
Combo1(2).Clear
Set RS = vg_db.Execute("sgpadm_s_servicio 8, '' , " & vg_Indppr & ", ''")
Combo1(2).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 10) & ")"
Do While Not RS.EOF
    Combo1(2).AddItem IIf(IsNull(RS!ser_nombre), "", RS!ser_codigo & " - " & Trim(RS!ser_nombre)) & Space(150) & "(" & fg_pone_cero(Str(RS!ser_codigo), 10) & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(2).ListIndex = 0
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2, 3
    vaSpread3.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           vaSpread3.Col = Index: nom = UCase(Trim(vaSpread3.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread3.text) <> "" Then
              If vaSpread3.RowHidden = True Then vaSpread3.RowHidden = False
           Else
              If vaSpread3.RowHidden = False Then vaSpread3.RowHidden = True
           End If
        Next i
        vaSpread3.SetActiveCell Index, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread3.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread3.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread3.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread3.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread3.Sort -1, -1, vaSpread1.MaxCols, vaSpread3.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           If vaSpread3.RowHidden = True Then vaSpread3.RowHidden = False
       Next
       vaSpread3.SetActiveCell Index, vaSpread3.SearchCol(Index, 0, vaSpread3.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread3.SetActiveCell Index, 1
    End If
    vaSpread3.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim i As Long, codsub As Long, codreg As Long, codser As Long, codlpr As Long, anomes As Long, esting As Boolean
Dim codpro As String, coding As String, precio As Double
Dim noming As String
On Error GoTo Man_Error
Select Case Button.Index
Case 3 '-------> Modificar
    If vaSpread1.MaxRows < 1 Then MsgBox "Debe seleccionar una lista precio...", vbCritical, Msgtitulo: Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 10, 0, modo
Case 10 '-------> Cancelar
    fg_carga ""
    Frame1(1).Enabled = True
    est = True
    If SSTab1.TabEnabled(0) = True Then
       If vaSpread1.MaxRows > 0 Then
          vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = "0"
          vaSpread2.Row = 1: vaSpread2.Col = 1
          vaSpread1.Row = -1: vaSpread1.Col = 10: vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
       End If
    Else
       SSTab1.TabEnabled(0) = True
       vaSpread1_Click vaSpread1.ActiveCol, vaSpread1.ActiveRow
    End If
    est = False
    Gl_Ac_Botones Me, 1, 3, modo
    fg_descarga
Case 12 '-------> Grabar datos
    fg_carga ""
    est = True
    anomes = Format(fpDateTime1(0).text, "yyyymm")
    Bar1(0).Visible = True: Bar1(0).Value = 0
    esting = True
    If SSTab1.TabEnabled(0) = True Then
       For i = 1 To vaSpread1.MaxRows
           DoEvents
           vaSpread1.Row = i
           vaSpread1.Col = 1
           Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
           If vaSpread1.text = "1" Then
              vaSpread1.SetActiveCell 3, vaSpread1.Row
              vaSpread1.Col = 2: codsub = vaSpread1.text
              vaSpread1.Col = 4: codreg = vaSpread1.text
              vaSpread1.Col = 6: codser = vaSpread1.text
              vaSpread1.Col = 8: codlpr = vaSpread1.text
              '-------> Rutina validar relación ingrediente & producto ó bien ingrediente en valor cero
              Set RS = vg_db.Execute("sgpadm_s_validaringplanif " & codsub & ", " & codreg & ", " & codser & ", " & codlpr & ", " & anomes & ", 0,'" & vg_NUsr & "'")
              If Not RS.EOF Then
                coding = ""
                noming = ""
                 vg_db.Execute "UPDATE b_minuta SET min_estact = '2' WHERE min_subseg = " & codsub & " AND min_codreg = " & codreg & " AND min_codser = " & codser & " AND substring(convert(char(8),min_fecmin),1,6) = " & anomes & ""
                 vaSpread2.Row = 1: vaSpread2.Col = 3: vaSpread1.Col = 10: vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
                 vaSpread1.Col = 11: vaSpread1.text = 1
                 esting = False
                 Do While Not RS.EOF
                    Set RS1 = vg_db.Execute("SELECT ing_codigo, ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & RS!ing_codigo & "'")
                    If Not RS1.EOF Then
                       If Trim(noming) = "" Then
                          noming = VgLinea & "Esto(s) ingrediente(s) no tiene productos asociados" & VgLinea
                          noming = noming & VgLinea & RS!ing_codigo & " - " & Trim(RS1!ing_nombre)
                       Else
                          noming = noming & VgLinea & RS!ing_codigo & " - " & Trim(RS1!ing_nombre)
                       End If
                    Else
                       If Trim(noming) = "" Then
                          noming = VgLinea & "Esto(s) Ingrediente(s) no tiene productos asociados" & VgLinea
                          noming = noming & VgLinea & RS!ing_codigo & " - " & "Ingrediente fue eliminado" & VgLinea
                       Else
                          noming = noming & VgLinea & RS!ing_codigo & " - " & "Ingrediente fue eliminado" & VgLinea
                       End If
                    End If
                    RS1.Close: Set RS1 = Nothing
                    
                    RS.MoveNext
                 Loop
                 RS.Close: Set RS = Nothing
              Else
                 '-------> Rutina actualizar precio planificación
                 RS.Close: Set RS = Nothing
                 vg_db.Execute "sgpadm_p_actuaplanif " & codsub & ", " & codreg & ", " & codser & ", " & codlpr & ", " & anomes & ""
                 '-------> Actualizar encabezado minuta
                 vg_db.Execute "UPDATE b_minuta SET min_codlpr = " & codlpr & ", min_estact = '1' WHERE min_subseg = " & codsub & " AND min_codreg = " & codreg & " AND min_codser = " & codser & " AND substring(convert(char(8),min_fecmin),1,6) = " & anomes & ""
                 vaSpread2.Row = 1: vaSpread2.Col = 2: vaSpread1.Col = 10: vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
                 vaSpread1.Col = 11: vaSpread1.text = 0
              End If
            End If
       Next i
    Else
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 2: codsub = vaSpread1.text
       vaSpread1.Col = 4: codreg = vaSpread1.text
       vaSpread1.Col = 6: codser = vaSpread1.text
       vaSpread1.Col = 8: codlpr = vaSpread1.text
       For i = 1 To vaSpread3.MaxRows
           Bar1(0).Value = Val((i / vaSpread3.MaxRows) * 100)
           vaSpread3.Row = i
           vaSpread3.Col = 7
           If vaSpread3.text = "1" Then
              vaSpread3.SetActiveCell 1, vaSpread3.Row
              vaSpread3.Col = 1: coding = vaSpread3.text
              vaSpread3.Col = 3
              If vaSpread3.TypeComboBoxCurSel <> -1 Then
                 vaSpread3.Col = 4: codpro = vaSpread3.text
                 vaSpread3.Col = 6: precio = vaSpread3.text
                 '-------> Actualizar asociación producto ingrediente
                 Set RS = vg_db.Execute("select * from b_asociaproductosing WHERE api_codsse = " & codsub & " AND api_codreg = " & codreg & " AND api_codser = " & codser & " AND api_coding = '" & coding & "' AND api_anomes = " & anomes & "")
                 If RS.EOF Then
                    vg_db.Execute "INSERT INTO b_asociaproductosing (api_codsse, api_codreg, api_codser, api_codpro, api_coding, api_anomes) values (" & codsub & ", " & codreg & ", " & codser & ", '" & codpro & "', '" & coding & "',  " & anomes & " )"
                 Else
                    vg_db.Execute "UPDATE b_asociaproductosing SET api_codpro = '" & codpro & "' WHERE api_codsse = " & codsub & " AND api_codreg = " & codreg & " AND api_codser = " & codser & " AND api_coding = '" & coding & "' AND api_anomes = " & anomes & ""
                 End If
'                 vg_db.Execute "UPDATE b_productosing SET pri_propre = 0 WHERE pri_coding = '" & coding & "'"
'                 vg_db.Execute "UPDATE b_productosing SET pri_propre = 1 WHERE pri_coding = '" & coding & "' AND pri_codpro = '" & codpro & "'"
                 '-------> Actualizar lista precio
                 vg_db.Execute "UPDATE b_detlistaprecio SET dlp_precio = " & precio & " WHERE dlp_codigo = " & codlpr & " AND dlp_anomes = " & anomes & " AND dlp_codpro = '" & codpro & "'"
              End If
           End If
       Next i
       SSTab1.TabEnabled(0) = True
    End If
    Bar1(0).Visible = False: Bar1(0).Value = 0
    Gl_Ac_Botones Me, 1, 3, modo
    vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = "0"
    Frame1(1).Enabled = True
    est = False
    fg_descarga
    MsgBox "Generación grabado finalizado " & IIf(esting, "sin problema", "con problema") & VgLinea & noming, vbInformation + vbOKOnly, Msgtitulo
Case 18 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim subseg As Integer: Dim Subreg As Integer: Dim servicio As Integer
Dim auxsub As Long, auxreg As Long, auxser As Long, auxlpr As Long
subseg = fg_codigocbo(Combo1, 0, 10, "") ' Val(fg_codigocbo(Combo1(0).ListIndex, 1, 1, ""))
'Val(fg_codigocbo(Combo2, 1, 1, ""))
Select Case Button.Index
Case 1 '------> Traer información del periodo desde planificación
    est = True
    SSTab1.TabVisible(1) = False
    vaSpread3.MaxRows = 0
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor 'Dim codSubseg, codreg, codser As Integer
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 9, " & codSubseg & ", " & codreg & ", " & codser & ",0, " & Format(fpDateTime1(0).text, "yyyymm") & ", 0, 0,'" & vg_Indppr & "'")
    Do While Not RS.EOF
       If RS!sub_codigo <> auxsub Or RS!reg_codigo <> auxreg Or RS!ser_codigo <> auxser Or RS!lpr_codigo <> auxlpr Then
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.Row = vaSpread1.MaxRows
          vaSpread1.Col = 1: vaSpread1.text = 0
          vaSpread1.Col = 2: vaSpread1.text = RS!sub_codigo
          vaSpread1.Col = 3: vaSpread1.text = RS!sub_codigo & " - " & Trim(RS!sub_nombre)
          vaSpread1.Col = 4: vaSpread1.text = RS!reg_codigo
          vaSpread1.Col = 5: vaSpread1.text = RS!reg_codigo & " - " & Trim(RS!reg_nombre)
          vaSpread1.Col = 6: vaSpread1.text = RS!ser_codigo
          vaSpread1.Col = 7: vaSpread1.text = RS!ser_codigo & " - " & Trim(RS!ser_nombre)
          vaSpread1.Col = 8: vaSpread1.text = RS!lpr_codigo
          vaSpread1.Col = 9: vaSpread1.text = RS!lpr_codigo & " - " & Trim(RS!lpr_nombre)
          vaSpread1.Col = 10:
          vaSpread2.Row = 1: vaSpread2.Col = IIf(IsNull(RS!min_estact) Or RS!min_estact = "0", 1, IIf(RS!min_estact = "1", 2, 3)): vaSpread1.Col = 10: vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
          vaSpread1.Col = 11: vaSpread1.text = 0
          auxsub = RS!sub_codigo
          auxreg = RS!reg_codigo
          auxser = RS!ser_codigo
          auxlpr = RS!lpr_codigo
        End If
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    vaSpread1.Visible = True
'    If vaSpread1.MaxRows < 1 Then MsgBox "No existe lista de precio para este periodo...", vbCritical, Msgtitulo: Exit Sub
    If vaSpread1.MaxRows < 1 Then MsgBox "No existe minuta para este servicio...", vbCritical, Msgtitulo: Exit Sub
    est = False
'    Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows > 0, 4, 2), modo
End Select
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If est Then Exit Sub
If modo = "" Then modo = "M"
If Toolbar1.Buttons(12).Visible = False Then SSTab1.TabVisible(1) = False: Frame1(1).Enabled = False: Gl_Ac_Botones Me, 1, 0, modo
'vaSpread1.Row = Row: vaSpread1.Col = 10: vaSpread2.Row = 1: vaSpread2.Col = 1: vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
Dim RS As New ADODB.Recordset
Dim i As Long, codsub As Long, codreg As Long, codser As Long, codlpr As Long, anomes As Long, auxing As String, auxpro As String
Dim lisnom As String, liscod As String, lispre As String, codaux As Long, precio As Double
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = IIf(vaSpread1.Value = "1", "0", "1")
If Col = 1 Or Row < 1 Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = 11
'If vaSpread1.text <> 0 Then
   fg_carga ""
   Label1(0).Visible = True
   Label1(0).Caption = ""
   vaSpread1.Col = 3
   Label1(0).Caption = Label1(0).Caption & "" & vaSpread1.text
   vaSpread1.Col = 5
   Label1(0).Caption = Label1(0).Caption & "\" & vaSpread1.text
   vaSpread1.Col = 7
   Label1(0).Caption = Label1(0).Caption & "\" & vaSpread1.text
   vaSpread1.Col = 9
   Label1(0).Caption = Label1(0).Caption & "\" & vaSpread1.text
   vaSpread1.Col = 2: codsub = vaSpread1.text
   vaSpread1.Col = 4: codreg = vaSpread1.text
   vaSpread1.Col = 6: codser = vaSpread1.text
   vaSpread1.Col = 8: codlpr = vaSpread1.text
   anomes = Format(fpDateTime1(0).text, "yyyymm")
   auxing = "": auxpro = ""
   vaSpread3.Visible = False
   vaSpread3.MaxRows = 0
   vaSpread3.Row = -1: vaSpread3.Col = -1
   vaSpread3.BackColor = Shape1(0).FillColor
   
   '-------> Rutina traer ingrediente & producto y precio en cero
   Set RS = vg_db.Execute("sgpadm_s_asocingprodplanif " & codsub & ", " & codreg & ", " & codser & ", " & codlpr & ", " & anomes & ", 0,'" & vg_NUsr & "'")
   Do While Not RS.EOF
      If RS!ing_codigo <> auxing Then
         If Trim(auxing) <> "" Then
            vaSpread3.Col = 4
            codaux = -1
            For i = 0 To vaSpread3.TypeComboBoxCount
                vaSpread3.TypeComboBoxCurSel = i
                If vaSpread3.text = auxpro Then codaux = i: Exit For
                codaux = -1
            Next i
            vaSpread3.Col = 3: vaSpread3.TypeComboBoxCurSel = codaux
            vaSpread3.Col = 5: vaSpread3.TypeComboBoxCurSel = codaux
            precio = IIf(vaSpread3.TypeComboBoxCurSel <> -1, vaSpread3.text, 0)
            vaSpread3.Col = 6: vaSpread3.text = Format(precio, fg_Pict(6, 2))
            If precio = 0 Then vaSpread3.ForeColor = &HFF&
            vaSpread3.Col = 7: vaSpread3.text = 0
         End If
         vaSpread3.MaxRows = vaSpread3.MaxRows + 1
         vaSpread3.Row = vaSpread3.MaxRows
         vaSpread3.Col = 1: vaSpread3.text = RS!ing_codigo
         vaSpread3.Col = 2: vaSpread3.text = RS!ing_nombre
         auxing = IIf(IsNull(RS!ing_codigo), "", RS!ing_codigo)
         lisnom = "": liscod = "": lispre = "": encuentra = False
      End If
      If RS!pri_propre > 0 Then auxpro = RS!pro_codigo
      '-------> Mover producto a lista
      If RS!pro_codigo <> "" Then
         vaSpread3.Col = 3: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & "(" & RS!pro_codigo & ") " & Trim(RS!pro_nombre)
         vaSpread3.Col = 4: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS!pro_codigo
         vaSpread3.Col = 5: lispre = lispre & IIf(lispre <> "", Chr$(9), "") & RS!dlp_precio
         vaSpread3.Col = 3: vaSpread3.TypeComboBoxList = lisnom
         vaSpread3.Col = 4: vaSpread3.TypeComboBoxList = liscod
         vaSpread3.Col = 5: vaSpread3.TypeComboBoxList = lispre
         
         
         
      End If
      RS.MoveNext
   Loop
   If Trim(auxing) <> "" Then
      vaSpread3.Col = 4
      codaux = -1
      For i = 0 To vaSpread3.TypeComboBoxCount
          vaSpread3.TypeComboBoxCurSel = i
          If vaSpread3.text = auxpro Then codaux = i: Exit For
          codaux = -1
      Next i
      vaSpread3.Col = 3: vaSpread3.TypeComboBoxCurSel = codaux
      vaSpread3.Col = 5: vaSpread3.TypeComboBoxCurSel = codaux
      precio = IIf(vaSpread3.TypeComboBoxCurSel <> -1, vaSpread3.text, 0)
      vaSpread3.Col = 6: vaSpread3.text = Format(precio, fg_Pict(6, 2))
      If precio = 0 Then vaSpread3.ForeColor = &HFF&
   End If
   RS.Close: Set RS = Nothing
   vaSpread3.Visible = True
   SSTab1.TabVisible(1) = True
   fg_descarga
'Else
'   Label1(0).Visible = False
'   SSTab1.TabVisible(1) = False
'End If
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
End Sub

Private Sub vaSpread3_EditChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread3.MaxRows < 1 Or Col <> 6 Then Exit Sub
vaSpread3.Row = Row
If modo = "" Then modo = "M"
If Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo
Frame1(1).Enabled = False: SSTab1.TabEnabled(0) = False
vaSpread3.Col = 7: vaSpread3.text = 1
Select Case Col
Case 6
    vaSpread3.Col = 6
    If Val(vaSpread3.text) > 0 Then vaSpread3.ForeColor = &H0& Else vaSpread3.ForeColor = &HFF&
End Select
End Sub

Private Sub vaSpread3_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 3
    Dim indice As Long, precio As Double
    vaSpread3.Row = Row
    vaSpread3.Col = 3: indice = vaSpread3.TypeComboBoxCurSel
    vaSpread3.Col = 4: vaSpread3.TypeComboBoxCurSel = indice
    vaSpread3.Col = 5: vaSpread3.TypeComboBoxCurSel = indice
    precio = IIf(vaSpread3.TypeComboBoxCurSel <> -1, vaSpread3.text, 0)
    vaSpread3.Col = 6: vaSpread3.text = Format(precio, fg_Pict(6, 2))
    If Val(vaSpread3.text) > 0 Then vaSpread3.ForeColor = &H0& Else vaSpread3.ForeColor = &HFF&
    vaSpread3.Col = 7: vaSpread3.text = 1
    If modo = "" Then modo = "M"
    If Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo
    Frame1(1).Enabled = False: SSTab1.TabEnabled(0) = False:
End Select
End Sub
