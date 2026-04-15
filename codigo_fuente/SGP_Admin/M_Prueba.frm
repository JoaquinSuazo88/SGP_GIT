VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Prueba 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenedor Prueba"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   8340
   ScaleWidth      =   12705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame6 
      Height          =   6255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   12255
      Begin VB.Frame Frame7 
         Height          =   435
         Left            =   1650
         TabIndex        =   9
         Top             =   5640
         Width           =   6885
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   6780
         End
      End
      Begin VB.Frame Frame8 
         Height          =   435
         Left            =   240
         TabIndex        =   7
         Top             =   5640
         Width           =   1275
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   1170
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5250
         Left            =   240
         TabIndex        =   11
         Top             =   240
         Width           =   11790
         _Version        =   393216
         _ExtentX        =   20796
         _ExtentY        =   9260
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   3
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Prueba.frx":0000
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame5 
      Height          =   735
      Left            =   2160
      TabIndex        =   1
      Top             =   720
      Width           =   8415
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1245
         _Version        =   196608
         _ExtentX        =   2196
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
         BackColor       =   16777215
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
         Top             =   120
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
               Picture         =   "M_Prueba.frx":199A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3495
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         Index           =   5
         Left            =   120
         TabIndex        =   4
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3465
         TabIndex        =   3
         Top             =   195
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3000
         Picture         =   "M_Prueba.frx":1D34
         Top             =   120
         Width           =   480
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12705
      _ExtentX        =   22410
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Prueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim est As Boolean
Dim strSQL As String


Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Select
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
    Select Case Index
    Case 1
        If Val(fpLongInteger1(1).Value) < 1 Then fpayuda(0).Caption = "": Exit Sub
        Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 3, '" & fpLongInteger1(1).Value & "', '', ''")
        If RS.EOF Then RS.Close: Set RS = Nothing:: fpayuda(0).Caption = "": Exit Sub
        fpayuda(0).Caption = Trim(RS!descripcion)
        RS.Close: Set RS = Nothing
    End Select
End Sub

Private Sub Image1_Click(Index As Integer)
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Mantenedor Centro Costo", "CentCost"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    'fpDateTime1(0).SetFocus

    
End Sub

Private Sub Form_Load()

    Me.Height = 9390
    Me.Width = 13095
    Msgtitulo = "Mantenedor Prueba"
    fg_centra Me
    
    Gl_Mo_Botones Me, 14
    Toolbar1.Buttons.item(15).ButtonMenus(1).Visible = False
    Gl_Ac_Botones Me, 14, 1, modo
    
    'fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
    
'    MoverDatosGrilla
'    MoverDatosListadePrecios
'    MoverDatosListadePreciosCasinoAsignados
        
End Sub


'Sub MoverDatosGrilla()
'    fg_carga ""
'    Dim X As Boolean
'    ' Control displays text tips aligned to pointer with focus
'    vaSpread1.TextTip = 2
'    vaSpread1.TextTipDelay = 250
'    X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
'    vaSpread1.Visible = False
'    vaSpread1.MaxRows = 0
'    vaSpread1.Row = -1
'    vaSpread1.Col = -1
'    vaSpread1.Lock = True
'
'
'
'    strSQL = "SELECT prv_codigo, prv_nombre, pro_codigo, pro_nombre " & _
'             "FROM b_proveedor, b_productos " & _
'             "WHERE prv_activo = 0 AND pro_fecven >= " & Format(Date, "yyyymmdd") & " "
'
'
'
'
'    Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 2, '', '', ''")
'    Do While Not RS.EOF
'       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'       vaSpread1.Row = vaSpread1.MaxRows
'       vaSpread1.Col = 1
'       vaSpread1.text = RS!codigo
'       vaSpread1.Col = 2
'       vaSpread1.text = Trim(RS!descripcion)
'       RS.MoveNext
'    Loop
'    RS.Close: Set RS = Nothing
'    vaSpread1.Visible = True
'    If vaSpread1.MaxRows > 0 Then
'       vaSpread1.Row = 1
'       vaSpread1.Col = 1
'       codigo = ""
'       codigo = Val(vaSpread1.text)
'       vaSpread1.SetActiveCell 1, 1 ': vaSpread1.SetFocus
'    End If
'    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
'    fg_descarga
'End Sub







'Sub MoverDatos()
'    fg_carga ""
'    Est = True
'    Limpia 1
'
'    strSQL = "SELECT cli_codigo, cli_nombre, cli_tipo, cli_activo FROM b_clientes"
'
'    Set RS = vg_dbpedweb.Execute(strSQL)   '("pedweb_s_listaprecios 3, '" & codigo & "', '', ''")
'    If Not RS.EOF Then
'       fpLongInteger1(0).Value = RS!cli_codigo
'       fpText1(0).text = Trim(RS!cli_nombre)
'       Frame3.Caption = RS!cli_codigo & " - " & Trim(RS!cli_nombre)
'    End If
'    RS.Close: Set RS = Nothing
'    fg_descarga
'    Est = False
'End Sub
