VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SsllDxv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descuentos por Volumen"
   ClientHeight    =   8340
   ClientLeft      =   2385
   ClientTop       =   1395
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   14175
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame6 
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   13815
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   12720
         TabIndex        =   25
         Top             =   5640
         Width           =   765
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   660
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   735
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
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
         Left            =   1695
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   735
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   11640
         TabIndex        =   18
         Top             =   5640
         Width           =   1005
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   435
         Left            =   6960
         TabIndex        =   17
         Top             =   5640
         Width           =   4605
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   4500
         End
      End
      Begin VB.Frame Frame1 
         Height          =   435
         Left            =   5520
         TabIndex        =   16
         Top             =   5640
         Width           =   1365
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   1260
         End
      End
      Begin VB.Frame Frame7 
         Height          =   435
         Left            =   1770
         TabIndex        =   15
         Top             =   5640
         Width           =   3645
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   3540
         End
      End
      Begin VB.Frame Frame8 
         Height          =   435
         Left            =   600
         TabIndex        =   14
         Top             =   5640
         Width           =   1035
         Begin VB.TextBox TextCai2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   930
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5250
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   13575
         _Version        =   393216
         _ExtentX        =   23945
         _ExtentY        =   9260
         _StockProps     =   64
         BackColorStyle  =   3
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
         MaxCols         =   6
         SpreadDesigner  =   "M_SsllDxv.frx":0000
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   2160
      TabIndex        =   8
      Top             =   720
      Width           =   8415
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
               Picture         =   "M_SsllDxv.frx":1B6C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3000
         TabIndex        =   9
         Top             =   600
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
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
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3465
         TabIndex        =   1
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3495
         TabIndex        =   12
         Top             =   290
         Width           =   4335
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3000
         Picture         =   "M_SsllDxv.frx":1F06
         Top             =   120
         Width           =   480
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
         TabIndex        =   11
         Top             =   315
         Width           =   1380
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
         Left            =   3480
         TabIndex        =   10
         Top             =   720
         Width           =   75
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   10920
      TabIndex        =   24
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   23
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label4 
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   22
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "M_SsllDxv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim est As Boolean
Dim strSQL
Dim Local_CallForm As String, itop As Long
Dim wvarMaxColumnas As String
Dim varNuevaCol As String

Private Sub Command1_Click()
    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_proveedor", "prv_", "Listado de Proveedores", "SsllListProv"
    B_TabEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
        vaSpread2.Col = 5: vaSpread2.Row = iRow: vaSpread2.SetActiveCell 5, iRow
        vaSpread2.EditMode = True: vaSpread2.EditModeReplace = True: vaSpread2.SetFocus
        Exit Sub
    End If
    
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1
    vaSpread2.text = fg_PintaRut(IIf(vg_Dig = "N", vg_codigo, vg_codigo))
    vaSpread2.Col = 2
    vaSpread2.text = vg_nombre
    
    If modo <> "A" Then modo = "M"
    
    Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Command2_Click()
Dim i As Long
Dim codpro As String
Dim auxpro As String
Dim wvarCodProd As String

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Listado Formato de Compras", "SsllListFormComp"
    B_TabEst.Show 1
    Me.Refresh
    
    If vg_codigo = "" Then
        vaSpread2.Col = 5: vaSpread2.Row = iRow: vaSpread2.SetActiveCell 5, iRow
        vaSpread2.EditMode = True: vaSpread2.EditModeReplace = True: vaSpread2.SetFocus
        Exit Sub
    End If
    
    '-------> Validar si el proveedor ya tiene asignado mismo productos
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        auxpro = vaSpread2.text
        vaSpread2.Col = 3
        If Trim(vaSpread2.text) = Trim(wvarCodProd) And codpro = auxpro And vaSpread2.MaxRows <> i Then
           MsgBox "El producto ya existe en la grilla, para este proveedor...", vbExclamation + vbOKOnly, Msgtitulo
           vaSpread2.SetActiveCell 3, vaSpread2.ActiveRow
           vaSpread2.Col = 3
           vaSpread2.text = ""
           Exit Sub
        End If
    Next i
    
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 3
    vaSpread2.text = vg_codigo
    vaSpread2.Col = 4
    vaSpread2.text = vg_nombre
    Label4(1) = Trim(vaSpread2.text)
    
    If modo <> "A" Then modo = "M"
    
    Gl_Ac_Botones Me, 1, 0, modo
    
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub


Private Sub fpText_Change(Index As Integer)
    If est Then Exit Sub
    fpayuda(0).Caption = ""
    If modo = "" Then modo = "M"
End Sub


Private Sub fpText_LostFocus(Index As Integer)
    Dim codi As Long, Bd As String, Ul As String
    On Error GoTo Man_Error
    If fpText(Index).text = "" Then fpayuda(0).Caption = "": codi = 0: Exit Sub
    
    codi = fpText(Index).text
    Bd = IIf(Index = 0, "b_clientes", "")
    Ul = IIf(Bd = "b_clientes", "cli", "")
    
    Set RS1 = Nothing
    
    strSQL = "SELECT " & Ul & "_codigo, " & Ul & "_nombre FROM " & Bd & " WHERE " & Ul & "_codigo='" & IIf(Ul = "cli", codi, codi) & "'"
    RS1.Open strSQL, vg_db, adOpenStatic
    
    If Not RS1.EOF Then
        fpayuda(0).Caption = IIf(IsNull(Trim(RS1!cli_nombre) = ""), "", RS1!cli_nombre)
        vg_codigo = RS1!cli_codigo
        codi = 0
    Else
        MsgBox "No existe codigo en la tabla..."
        fpayuda(0).Caption = ""
        fpText(Index).text = ""
        codi = 0
        On Error Resume Next: fpText(Index).SetFocus
    End If
    
    RS1.Close: Set RS1 = Nothing
    Exit Sub
    
Man_Error:
    If Err = 3034 Then Exit Sub
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Image1_Click(Index As Integer)
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Mantenedor Centro Costo", "CentCost"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = (vg_codigo)
    fpayuda(0).Caption = vg_nombre
End Sub

Private Sub Form_Load()
    Me.HelpContextID = vg_OpcM
    Me.Height = 9390
    Me.Width = 14160
    Msgtitulo = "Descuentos Por Volumen"
    Image1(0).ToolTipText = "Buscar Centro de costo"
    
    For i = 1 To 5
        TextCai2(i).ToolTipText = "Filtrar"
    Next
    
    fg_centra Me
    
    Label4(0).Caption = "": Label4(0).Visible = False
    Label4(1).Caption = "": Label4(1).Visible = False
    
    modo = ""
    est = True
    
    itop = 1
    CallForm = Local_CallForm
    
    Gl_Mo_Botones Me, 1
    Gl_Ac_Botones Me, 1, 3, modo
        
    varNuevaCol = ""
    OpGr = False
    
    vaSpread2.MaxRows = 0
End Sub

Sub MoverDatosGrilla()
    fg_carga ""
    Dim x As Boolean
    ' Control displays text tips aligned to pointer with focus
    vaSpread2.TextTip = 2
    vaSpread2.TextTipDelay = 250
    x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    vaSpread2.Visible = False
    vaSpread2.MaxRows = 0
    vaSpread2.Row = -1
    vaSpread2.Col = -1
    
    If (fpText(0).text = "") Then
        strSQL = "SELECT prv_codigo, prv_nombre, foc_codsac, foc_nomsac, dxv_fecvig, dxv_pctdxv " & _
                 "From b_proveedor WITH (holdlock), b_formatocompras WITH (holdlock), b_ssll_dxv WITH (holdlock), b_clientes WITH (holdlock) " & _
                 "WHERE prv_activo = 0 AND cli_tipo = 0 AND cli_activo = 1 " & _
                 "AND dxv_codcen = cli_codigo AND dxv_rutpro = prv_codigo " & _
                 "AND dxv_codfmc = foc_codsac ORDER BY prv_nombre"
    Else
        strSQL = "SELECT prv_codigo, prv_nombre, foc_codsac, foc_nomsac, dxv_fecvig, dxv_pctdxv " & _
                 "From b_proveedor WITH (holdlock), b_formatocompras WITH (holdlock), b_ssll_dxv WITH (holdlock), b_clientes WITH (holdlock) " & _
                 "WHERE prv_activo = 0 AND cli_tipo = 0 AND cli_activo = 1 " & _
                 "AND dxv_codcen = cli_codigo AND dxv_rutpro = prv_codigo " & _
                 "AND dxv_codfmc = foc_codsac AND cli_codigo = '" & fpText(0).text & "' " & _
                 "ORDER BY prv_nombre"
    End If

    Set RS = vg_db.Execute(strSQL)
    
    Do While Not RS.EOF
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 1
        vaSpread2.text = fg_PintaRut(RS!prv_codigo)
       
        vaSpread2.Col = 2
        vaSpread2.text = Trim(RS!prv_nombre)
        
       
        vaSpread2.Col = 3
        vaSpread2.text = Trim(RS!foc_codsac)
       
        vaSpread2.Col = 4
        vaSpread2.text = Trim(RS!foc_nomsac)
        'Label4(1) = Trim(vaSpread2.text)
        
        vaSpread2.Col = 5
        vaSpread2.CellType = CellTypeDate
        vaSpread2.TypeDateFormat = TypeDateFormatDDMMYY
        vaSpread2.TypeDateMin = "01011900"
        vaSpread2.TypeDateMax = "31125000"
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.TypeDateCentury = True
        vaSpread2.ForeColor = &HFF0000
        vaSpread2.text = Trim(IIf(IsNull(RS!dxv_fecvig), "", Format(RS!dxv_fecvig, "dd/mm/yyyy")))
        
        vaSpread2.Col = 6
        vaSpread2.text = Trim(IIf(IsNull(RS!dxv_pctdxv), "0", RS!dxv_pctdxv))
                
        RS.MoveNext
    Loop
    
    RS.Close: Set RS = Nothing
    
    vaSpread2.Visible = True
    
    If vaSpread2.MaxRows > 0 Then
       vaSpread2.Row = 1
       vaSpread2.Col = 1
       codigo = ""
       codigo = Val(vaSpread2.text)
       vaSpread2.SetActiveCell 1, 1
    End If
    
    Label2.Caption = Format(vaSpread2.MaxRows, fg_Pict(7, 0)) & " Registro"
    fg_descarga
End Sub

Sub MoverDatos()
    fg_carga ""
    est = True

    strSQL = "SELECT cli_codigo, cli_nombre, cli_tipo, cli_activo FROM b_clientes WITH (holdlock)"

    Set RS = vg_dbpedweb.Execute(strSQL)
    
    If Not RS.EOF Then
        fpText(0).text = RS!cli_codigo
'        fpText(0).text = Trim(RS!cli_nombre)
        Frame3.Caption = RS!cli_codigo & " - " & Trim(RS!cli_nombre)
    End If
    
    RS.Close: Set RS = Nothing
    fg_descarga
    est = False
End Sub

Private Sub TextCai2_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 0, 1, 2, 3, 4, 5
    vaSpread2.Visible = False
    
    If Trim(TextCai2(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai2(Index).text) & "*"
           vaSpread2.Col = Index
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index, 1
    End If
    
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(TextCai2(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(TextCai2(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    
    If Trim(TextCai2(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextCai2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    
    vaSpread2.Visible = True
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim codigo As Long, Nombre As String, NomCor As String

Dim vartxtCol1, vartxtCol3, vartxtCol5, vartxtCol6 As String

On Error GoTo Man_Error

Command1.Visible = False
Command2.Visible = False

Select Case Button.Index

Case 1
    'If (vaSpread2.Col <> -1 And fpText(0).text <> "") Then
    If (fpText(0).text <> "") Then
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        lisnom = "": liscod = ""
        
        vaSpread2.Row = vaSpread2.MaxRows: vaSpread2.Col = 2: vaSpread2.SetActiveCell 1, vaSpread2.MaxRows: vaSpread2.SetFocus
        
        Call AgregarBotones(1, vaSpread2.Row)
        Call AgregarBotones(3, vaSpread2.Row)
        
        vaSpread2.Col = 5
        vaSpread2.SetFocus
        
        vaSpread2.Col = 1
        vaSpread2.CellType = CellTypeEdit
        vaSpread2.SetFocus
        
        vaSpread2.Col = 3
        vaSpread2.CellType = CellTypeEdit
        
        vaSpread2.Col = 5
        vaSpread2.CellType = CellTypeDate
        
        varNuevaCol = 1
    End If
    
Case 3
    'Boton Alterar
   
Case 10
    'NO guarda la asociacion nueva
    
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        modo = "Cancel"
        
        If modo = "Cancel" Then
            modo = ""
            Cancela
        Else
            Cancela
        End If
    
Case 12
    'SI guarda la asociacion nueva
    
    If Trim(Label4(0).Caption) <> "" And Trim(Label4(1).Caption) = "" Then
        vaSpread2.Col = 1
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            If vaSpread2.text = Label4(1).Caption Then
                vaSpread2.SetActiveCell 1, vaSpread2.Row
                Exit For
            End If
        Next i
    End If
    
    GrabaRegistro vaSpread2.ActiveRow
    
Case 13
    
'    vaSpread2.Row = vaSpread2.MaxRows
'
'    vaSpread2.Col = 1: vartxtCol1 = fg_DespintaRut(Trim(vaSpread2.text))
'    vaSpread2.Col = 3: vartxtCol3 = Trim(vaSpread2.text)
'    vaSpread2.Col = 5: vartxtCol5 = Format(CDate(Trim(vaSpread2.text)), "yyyy/mm/dd")
'    vaSpread2.Col = 6: vartxtCol6 = fg_Quitachar(Trim(vaSpread2.text), "%")
'
'    If (fpText(0).text = "" Or vartxtCol1 = "" Or vartxtCol3 = "") Then
'
'        MsgBox "Falta Información Por Ingresar", vbCritical + vbOKOnly, Msgtitulo
'    Else
'        strSQL = "INSERT INTO b_ssll_dxv(dxv_codcen, dxv_rutpro, dxv_codfmc, dxv_fecvig, dxv_pctdxv) " & _
'                 "VALUES('" & fpText(0).text & "','" & vartxtCol1 & "','" & vartxtCol3 & "','" & _
'                 "" & vartxtCol5 & "'," & vartxtCol6 & ")"
'
'        vg_db.Execute strSQL
'
'    End If
Case 15
    If vaSpread2.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_DsctoxVolumen
Case 4
    
Case 18
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If Not (fpText(0).text = "") Then
            Command1.Visible = False
            Command2.Visible = False
            MoverDatosGrilla
            Gl_Ac_Botones Me, 1, 14, modo
        End If
    End Select
End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)
    wvarMaxColumnas = vaSpread2.MaxRows
End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread2.MaxRows < 1 Then Exit Sub
vaSpread2.Row = Row
Select Case Col
Case 1 And Not ChangeMade
    vaSpread2.Col = Col
    If Trim(vaSpread2.Value) = "" Or vg_Dig = "N" Then Exit Sub
    vaSpread2.Value = fg_DespintaRut(vaSpread2.Value)
    vaSpread2.Value = Mid(vaSpread2.Value, 1, Len(Trim(vaSpread2.Value)) - 1)
End Select
End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread2.MaxRows < 1 Then Exit Sub
vaSpread2.Row = NewRow 'vaSpread2.MaxRows
Select Case Col
Case 1
    vaSpread2.Col = Col
    If Trim(vaSpread2.text) <> "" And vaSpread2.CellType = CellTypeEdit Then
'       vaSpread2.Row = vaSpread2.MaxRows
       vaSpread2.Col = Col
       vaSpread2.Value = UCase(vaSpread2.Value)
       If Trim(vaSpread2.Value) = "" Or vg_Dig = "N" Then Exit Sub
       vaSpread2.Value = fg_DespintaRut(vaSpread2.Value)
       vaSpread2.Value = fg_PintaRut(vaSpread2.Value)
'       vaSpread2.Value = fg_RutDig(Trim(vaSpread2.Value))

    End If
End Select
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread2.MaxRows < 1 Then Exit Sub
Dim i As Long
Dim codpro As String
Dim auxpro As String
    If modo = "" Then modo = "M"
    
    Dim wvarRutProv, wvarCodProd As String
    
    If (Col = 1 And Row = vaSpread2.MaxRows) Then
        vaSpread2.Col = 1

        vaSpread2.Row = Row 'vaSpread2.MaxRows
        wvarRutProv = fg_RutDig(vaSpread2.text)

        strSQL = "SELECT prv_codigo, prv_nombre FROM b_proveedor WITH (holdlock) WHERE prv_codigo = '" & wvarRutProv & "'"

        Set RS = vg_db.Execute(strSQL)

        If RS.EOF Then
            vaSpread2.Col = 2
            vaSpread2.text = ""
        Else
            Do While Not RS.EOF
                vaSpread2.Col = 2
                vaSpread2.text = ""
                vaSpread2.text = Trim(RS!prv_nombre)
                
                RS.MoveNext
            Loop

            RS.Close: Set RS = Nothing
        End If
    End If

    If (Col = 3 And Row = vaSpread2.MaxRows) Then
        vaSpread2.Row = vaSpread2.MaxRows
        vaSpread2.Col = 1
        codpro = vaSpread2.text
        vaSpread2.Col = 3
        wvarCodProd = vaSpread2.text
        '-------> Validar si el proveedor ya tiene asignado mismo productos
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1
            auxpro = vaSpread2.text
            vaSpread2.Col = 3
            If Trim(vaSpread2.text) = Trim(wvarCodProd) And codpro = auxpro And vaSpread2.MaxRows <> i Then
               MsgBox "El producto ya existe en la grilla, para este proveedor...", vbExclamation + vbOKOnly, Msgtitulo
               vaSpread2.SetActiveCell 3, vaSpread2.ActiveRow
               vaSpread2.Col = 3
               vaSpread2.text = ""
               Exit Sub
            End If
        Next i

        strSQL = "SELECT foc_codsac, foc_nomsac FROM b_formatocompras WITH (holdlock) WHERE foc_codsac = '" & wvarCodProd & "'"

        Set RS = vg_db.Execute(strSQL)

        If RS.EOF Then
            vaSpread2.Col = 4
            vaSpread2.text = ""
        Else
            Do While Not RS.EOF
                vaSpread2.Col = 4
                vaSpread2.text = ""
                vaSpread2.text = Trim(RS!foc_nomsac)
                RS.MoveNext
            Loop

            RS.Close: Set RS = Nothing
        End If
    End If
    
End Sub

Private Sub vaSpread2_Change(ByVal Col As Long, ByVal Row As Long)
If vaSpread2.MaxRows < 1 Then Exit Sub
    Dim Fecha As Date
    Dim varProvRut, varProvNom As String
    Dim varProdCod, varProdNom As String
    Dim varFecVig, varDxvPctje As String
    
    On Error GoTo Man_Error
    vaSpread2.Row = Row
    
    If (Col = 5 Or Col = 6) Then
        If Col = 5 Then
            vaSpread2.Col = 5
            Fecha = vaSpread2.text
            vaSpread2.text = Fecha
        End If
        
        vaSpread2.Col = 1
        varProvRut = fg_DespintaRut(Trim(vaSpread2.text))
        
        vaSpread2.Col = 2
        varProvNom = Trim(vaSpread2.text)
        
        vaSpread2.Col = 3
        varProdCod = Trim(vaSpread2.text)
        
        vaSpread2.Col = 4
        varProdNom = Trim(vaSpread2.text)
        
        vaSpread2.Col = 5
        varFecVig = CDate(Trim(vaSpread2.text))
        varFecVig = Format(varFecVig, "yyyy/mm/dd")
        
        vaSpread2.Col = 6
        varDxvPctje = Mid(Trim(vaSpread2.text), 1, Len(vaSpread2.text) - 1)
           
        strSQL = "UPDATE b_ssll_dxv SET dxv_fecvig = '" & varFecVig & "', dxv_pctdxv = '" & varDxvPctje & "' " & _
                 "WHERE dxv_codcen = '" & fpText(0).text & "' AND dxv_rutpro = '" & varProvRut & "' " & _
                 "AND dxv_codfmc = '" & varProdCod & "'"
        
        vg_db.Execute strSQL
        
        vaSpread2.Col = Col + 1
    End If
    
Man_Error:
    If Err.Number = 13 Then vaSpread2.text = "": vaSpread2.SetActiveCell 5, vaSpread2.ActiveRow
End Sub

Private Sub vaSpread2_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread2.MaxRows < 1 Then Exit Sub
Dim varProvRut As String
Dim varProdCod As String
Dim varFecVig  As String

vaSpread2.Row = Row

vaSpread2.Col = 1
varProvRut = fg_DespintaRut(Trim(vaSpread2.text))

vaSpread2.Col = 3
varProdCod = Trim(vaSpread2.text)

Select Case Col
    Case 5
        vg_AuxCentCost = fpText(0).text
        
        M_SsllFecVig.Show 1

        vaSpread2.Col = Col
        vaSpread2.Row = Row

        If (vg_FecVig <> "") Then
            vaSpread2.text = vg_FecVig
            varFecVig = CDate(vg_FecVig)
            varFecVig = Format(varFecVig, "yyyy/mm/dd")
            Me.Refresh
            
            strSQL = "UPDATE b_ssll_dxv SET dxv_fecvig = '" & varFecVig & "' " & _
                     "WHERE dxv_codcen = '" & fpText(0).text & "'" & _
                     "AND dxv_rutpro = '" & varProvRut & "' AND dxv_codfmc = '" & varProdCod & "'"
            
            vg_db.Execute strSQL
        End If

        vaSpread2.EditMode = False
End Select
End Sub

Private Sub Cancela()
    OpGr = True
    vaSpread2.Row = vaSpread2.ActiveRow
    
    MoverDatosGrilla
    
    OpGr = False
   
    modo = "": Gl_Ac_Botones Me, 1, 14, modo
   
End Sub

Private Sub vaSpread2_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
    itop = NewTop
    Command1.Visible = False
    Command2.Visible = False
End Sub

Sub AgregarBotones(ByVal Col As Long, ByVal Row As Long)
    If Not (fpText(0).text = "") Then
        
        If Col <> 1 Then
            If Col <> 3 Then
                Command1.Visible = False
                Command2.Visible = False
            End If
        End If
        
        
        If Col = 1 And Row > 0 Then
            Command1.Top = IIf(Row <= 1, 735, 735 + (240 * (Row - itop)))
            Command1.Visible = True
            vaSpread2.EditMode = True
            vaSpread2.EditModeReplace = True
            vaSpread2.Row = Row
            iRow = Row
            vaSpread2.Col = 4
            vaSpread2.TypeHAlign = TypeHAlignLeft
        End If
        
        
        If Col = 3 And Row > 0 Then
            Command2.Top = IIf(Row <= 1, 735, 735 + (240 * (Row - itop)))
            Command2.Visible = True
            vaSpread2.EditMode = True
            vaSpread2.EditModeReplace = True
            vaSpread2.Row = Row
            iRow = Row
            vaSpread2.Col = 4
            vaSpread2.TypeHAlign = TypeHAlignLeft
        End If
        
    End If
End Sub

Private Sub GrabaRegistro(Fila)
    Dim PrvRut, PrvNombre As String
    Dim ProdCod, ProdNombre As String
    Dim FechaVigencia As String
    Dim Porcentaje, CodCentroCosto As String
    
    On Error GoTo Man_Error
    
    OpGr = True
    vaSpread2.Row = Fila
    
    CodCentroCosto = Trim(fpText(0).text)
    
    vaSpread2.Col = 1: PrvRut = fg_DespintaRut(Trim(vaSpread2.Value))
    vaSpread2.Col = 2: PrvNombre = Trim(vaSpread2.Value)
    
    vaSpread2.Col = 3: ProdCod = Trim(vaSpread2.Value)
    vaSpread2.Col = 4: ProdNombre = Trim(vaSpread2.Value)
    
    vaSpread2.Col = 5
    If Trim(vaSpread2.text) = "" Then
        FechaVigencia = Format(Date, "yyyy/mm/dd")
    Else
        FechaVigencia = Format(CDate(Trim(vaSpread2.text)), "yyyy/mm/dd")
    End If
    
    vaSpread2.Col = 6
    If Trim(vaSpread2.text) = "" Then
        Porcentaje = 0
    Else
        Porcentaje = fg_Quitachar(Trim(vaSpread2.text), "%")
    End If
    
    
    If PrvNombre = "" Or ProdNombre = "" Or CodCentroCosto = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread2.Row = Fila: vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow: vaSpread2.SetFocus: OpGr = False: Cancela: Exit Sub
    
    If modo = "A" Then
    
        strSQL = "INSERT INTO b_ssll_dxv (dxv_codcen, dxv_rutpro, dxv_codfmc, dxv_fecvig, dxv_pctdxv) " & _
                 "VALUES('" & CodCentroCosto & "','" & PrvRut & "','" & ProdCod & "','" & _
                 "" & FechaVigencia & "'," & Porcentaje & ")"
            
        vg_db.Execute strSQL
    
    modo = "": Gl_Ac_Botones Me, 1, 14, modo
    OpGr = False

    End If
    
Man_Error:
    If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
    If Err = 3034 Then Exit Sub
    fg_descarga
    'MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
