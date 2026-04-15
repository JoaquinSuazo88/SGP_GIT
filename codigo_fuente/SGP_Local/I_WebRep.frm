VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_WebRep 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7575
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   480
         TabIndex        =   5
         Top             =   6960
         Width           =   1005
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   900
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   1545
         TabIndex        =   3
         Top             =   6960
         Width           =   3765
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   3660
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6645
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7230
         _Version        =   393216
         _ExtentX        =   12753
         _ExtentY        =   11721
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
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
         MaxCols         =   4
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "I_WebRep.frx":0000
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   165
         Left            =   120
         TabIndex        =   7
         Top             =   7440
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_WebRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim opimp As String
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Pedidos Web"
End Sub

Sub LlenaDatos(titgen As String, op As String)
'-------> cargar familia productos
Me.Caption = titgen
Msgtitulo = titgen
opimp = op
fg_carga ""
Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
If op = "rutprod" Or op = "rutcale" Then
   vaSpread1.Col = 4
   vaSpread1.ColHidden = True
   vaSpread1.ColWidth(3) = 46.25
   Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 4, '', ''")
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
   
      vaSpread1.Col = 1
      vaSpread1.Text = "0"
      
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Text = RS!recorrido
   
      vaSpread1.Col = 3
      vaSpread1.Text = Trim(RS!descripcion)
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
ElseIf op = "regnegfam" Or op = "regnegpro" Or op = "regnegcas" Then
   Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 1, '', ''")
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1
      vaSpread1.Text = "0"
      
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Text = RS!rn_codigo
      
      vaSpread1.Col = 3
      vaSpread1.Text = Trim(RS!rn_nombre)
      
      vaSpread1.Col = 4
      vaSpread1.Text = IIf(Trim(RS!rn_tipo_ruta) = "1", "Con Rutas", IIf(Trim(RS!rn_tipo_ruta) = "2", "Sin Rutas", "Con y Sin Rutas"))
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
ElseIf op = "lisprecas" Then
   vaSpread1.Col = 4
   vaSpread1.ColHidden = True
   vaSpread1.ColWidth(3) = 46.25
   Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 2, '', '', ''")
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.Text = "0"
       
      vaSpread1.Col = 2
      vaSpread1.Text = RS!codigo
       
      vaSpread1.Col = 3
      vaSpread1.Text = Trim(RS!descripcion)
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
ElseIf op = "lisprecio" Then
'   vaSpread1.Col = 4
'   vaSpread1.ColHidden = False
'   vaSpread1.ColWidth(3) = 46.25
   vaSpread1.Row = 0
   vaSpread1.Col = 4
   vaSpread1.Text = "Fecha"
   Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 7, '', '', ''")
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.Text = "0"
       
      vaSpread1.Col = 2
      vaSpread1.Text = Trim(RS!CodigoLista)
       
      vaSpread1.Col = 3
      vaSpread1.Text = Trim(RS!descripcion)
      
      vaSpread1.Col = 4
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Text = Format(RS!FechaValidez, "dd/mm/yyyy")
      
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
End If
vaSpread1.Visible = True
fg_descarga
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 2, 3
    vaSpread1.Visible = False
    If Trim(Text1(Index).Text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index: nom = UCase(Trim(vaSpread1.Text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Text1(Index).Text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.Text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).Text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).Text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    If Trim(Text1(Index).Text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).Text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Dim i As Long, estsel As Boolean, cuenta As Long
    estsel = False
    cuenta = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.Text = "1" Then estsel = True: cuenta = cuenta + 1
    Next i
    If Not estsel Then If iselecc = 0 Then MsgBox "Debe Seleccionar A lo menor un ítem", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Toolbar1.Enabled = False
    If opimp = "rutprod" Then
       I_RutaProductos cuenta 'CStr(codigo)
    ElseIf opimp = "regnegfam" Then
       I_ReglasdeNegociosFamilia cuenta 'CStr(Codigo)
    ElseIf opimp = "regnegpro" Then
       I_ReglasdeNegociosProducto cuenta 'CStr(codigo)
    ElseIf opimp = "regnegcas" Then
       I_ReglasdeNegociosCasino cuenta 'CStr(codigo)
    ElseIf opimp = "lisprecas" Then
       I_ListadePreciosCasinoAsignados cuenta 'CStr(codigo)
    ElseIf opimp = "lisprecio" Then
       I_ListadePrecioss cuenta 'CStr(codigo)
    End If
    Toolbar1.Enabled = True
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread1.Col = 1
For i = BlockRow To BlockRow2
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
Next i
End Sub
