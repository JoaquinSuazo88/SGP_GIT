VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_ACProd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Cantidad y Lista Cantidad Productos"
   ClientHeight    =   8730
   ClientLeft      =   150
   ClientTop       =   930
   ClientWidth     =   13245
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   13245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   8055
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12975
      Begin TabDlg.SSTab SSTab1 
         Height          =   7695
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   13573
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Agregar Cantidad Productos"
         TabPicture(0)   =   "M_ACProd.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame2"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Listar Cantidad Productos"
         TabPicture(1)   =   "M_ACProd.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   6975
            Left            =   -74880
            TabIndex        =   3
            Top             =   600
            Width           =   12255
            Begin VB.Frame Frame10 
               Height          =   435
               Left            =   9960
               TabIndex        =   19
               Top             =   6480
               Width           =   915
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   7
                  Left            =   45
                  TabIndex        =   20
                  Top             =   135
                  Width           =   810
               End
            End
            Begin VB.Frame Frame9 
               Height          =   435
               Left            =   9000
               TabIndex        =   17
               Top             =   6480
               Width           =   915
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   6
                  Left            =   45
                  TabIndex        =   18
                  Top             =   135
                  Width           =   810
               End
            End
            Begin VB.Frame Frame8 
               Height          =   435
               Left            =   7920
               TabIndex        =   15
               Top             =   6480
               Width           =   1035
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   5
                  Left            =   45
                  TabIndex        =   16
                  Top             =   135
                  Width           =   930
               End
            End
            Begin VB.Frame Frame7 
               Height          =   435
               Left            =   2700
               TabIndex        =   13
               Top             =   6480
               Width           =   1485
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   14
                  Top             =   135
                  Width           =   1380
               End
            End
            Begin VB.Frame Frame5 
               Height          =   435
               Left            =   600
               TabIndex        =   11
               Top             =   6480
               Width           =   1155
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   1
                  Left            =   45
                  TabIndex        =   12
                  Top             =   135
                  Width           =   1050
               End
            End
            Begin VB.Frame Frame4 
               Height          =   435
               Left            =   1770
               TabIndex        =   9
               Top             =   6480
               Width           =   885
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   10
                  Top             =   135
                  Width           =   780
               End
            End
            Begin VB.Frame Frame6 
               Height          =   435
               Left            =   4245
               TabIndex        =   7
               Top             =   6480
               Width           =   3645
               Begin VB.TextBox Text1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   4
                  Left            =   45
                  TabIndex        =   8
                  Top             =   135
                  Width           =   3540
               End
            End
            Begin FPSpread.vaSpread vaSpread2 
               Height          =   6015
               Left            =   120
               TabIndex        =   5
               Top             =   360
               Width           =   12015
               _Version        =   393216
               _ExtentX        =   21193
               _ExtentY        =   10610
               _StockProps     =   64
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
               MaxCols         =   8
               SpreadDesigner  =   "M_ACProd.frx":0038
            End
         End
         Begin VB.Frame Frame2 
            Height          =   6975
            Left            =   120
            TabIndex        =   2
            Top             =   600
            Width           =   12255
            Begin VB.Frame Frame16 
               Height          =   435
               Left            =   8880
               TabIndex        =   31
               Top             =   6480
               Width           =   915
               Begin VB.TextBox Texta1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   6
                  Left            =   45
                  TabIndex        =   32
                  Top             =   135
                  Width           =   810
               End
            End
            Begin VB.Frame Frame15 
               Height          =   435
               Left            =   7800
               TabIndex        =   29
               Top             =   6480
               Width           =   1035
               Begin VB.TextBox Texta1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   5
                  Left            =   45
                  TabIndex        =   30
                  Top             =   135
                  Width           =   930
               End
            End
            Begin VB.Frame Frame14 
               Height          =   435
               Left            =   4080
               TabIndex        =   27
               Top             =   6480
               Width           =   3675
               Begin VB.TextBox Texta1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   4
                  Left            =   45
                  TabIndex        =   28
                  Top             =   135
                  Width           =   3570
               End
            End
            Begin VB.Frame Frame13 
               Height          =   435
               Left            =   2760
               TabIndex        =   25
               Top             =   6480
               Width           =   1275
               Begin VB.TextBox Texta1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   3
                  Left            =   45
                  TabIndex        =   26
                  Top             =   135
                  Width           =   1170
               End
            End
            Begin VB.Frame Frame12 
               Height          =   435
               Left            =   1680
               TabIndex        =   23
               Top             =   6480
               Width           =   1035
               Begin VB.TextBox Texta1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   2
                  Left            =   45
                  TabIndex        =   24
                  Top             =   135
                  Width           =   930
               End
            End
            Begin VB.Frame Frame11 
               Height          =   435
               Left            =   720
               TabIndex        =   21
               Top             =   6480
               Width           =   915
               Begin VB.TextBox Texta1 
                  Appearance      =   0  'Flat
                  BorderStyle     =   0  'None
                  Height          =   240
                  Index           =   1
                  Left            =   45
                  TabIndex        =   22
                  Top             =   135
                  Width           =   810
               End
            End
            Begin FPSpread.vaSpread vaSpread1 
               Height          =   6015
               Left            =   120
               TabIndex        =   4
               Top             =   360
               Width           =   11895
               _Version        =   393216
               _ExtentX        =   20981
               _ExtentY        =   10610
               _StockProps     =   64
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
               MaxCols         =   8
               SpreadDesigner  =   "M_ACProd.frx":1A95
            End
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ACProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String, Msgtitulo As String
Dim est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

'DECLARE @nreg AS INT, @fechoy DATETIME
'SELECT  @nreg   = 0
'SET     @fechoy = Getdate()
'
'SELECT  a.producto, a.CentralDeCompra, b.descripcion, a.cantidad, a.FechaInicio, c.Descripcioncompleta
'FROM    dbo.s_productosCantidad a, s_productos AS b, s_familiasDeProductos c
'Where a.producto = b.codigo
'--AND     a.producto = d.pce_codpro
'--AND     a.CentralDeCompra = d.pce_codcen
'AND     b.categoria  = c.codigo
'AND     a.FechaInicio IS not NULL
'AND     (a.cantidad > -1 )
'AND     b.institucional = 1
'AND     b.activo = 0
'AND    (b.vigencia > Convert(Char(10), @fechoy,101) OR b.vigencia <= 0)
'--and     a.producto = '002110'
'order by a.producto, a.CentralDeCompra

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9210
Me.Width = 13335
Msgtitulo = "Agregar Cantidad y Lista Cantidad Productos"
fg_centra Me
SSTab1.Tab = 0
modo = ""
est = True
Gl_Mo_Botones Me, 14
Toolbar1.Buttons.item(15).ButtonMenus(1).Visible = False
Gl_Ac_Botones Me, 14, 5, modo
MoverCantidadProductos
MoverListaProductos
'Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverCantidadProductos()
Dim i As Long
fg_carga ""
For i = 1 To 6
    Texta1(i).text = ""
Next i
Dim x As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 250
x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1.MaxRows = 0
vaSpread1.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_listaproductossincantidad 1, '', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = Trim(IIf(IsNull(RS!codigo), "", RS!codigo))
   vaSpread1.Col = 2
   vaSpread1.text = Trim(IIf(IsNull(RS!CentralCompra), "", RS!CentralCompra))
   vaSpread1.Col = 3
   vaSpread1.text = Trim(IIf(IsNull(RS!descripcioncompleta), "", RS!descripcioncompleta))
   vaSpread1.Col = 4
   vaSpread1.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
   vaSpread1.Col = 5
   vaSpread1.text = Trim(IIf(IsNull(RS!unidad), "", RS!unidad))
   '-------> definir formato numerico
   vaSpread1.Col = 6
   vaSpread1.CellType = CellTypeNumber
   vaSpread1.TypeNumberDecPlaces = 0
   vaSpread1.TypeIntegerMin = 1
   vaSpread1.TypeIntegerMax = 9999
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.TypeSpin = False
   vaSpread1.TypeIntegerSpinInc = 1
   vaSpread1.TypeIntegerSpinWrap = False
   vaSpread1.ForeColor = &HFF0000
   vaSpread1.text = 0
   '-------> definir formato fecha
   
'      vaSpread1.Col = 3
'      vaSpread1.CellType = CellTypeDate
'      vaSpread1.TypeDateCentury = False
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'      vaSpread1.TypeSpin = False
'      vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
'      vaSpread1.TypeDateMin = "01011973":  vaSpread1.TypeDateMax = "31125000"
'   vaSpread1.TypeDateCentury = True
   
   vaSpread1.Col = 7
   vaSpread1.CellType = CellTypeDate
   vaSpread1.TypeSpin = False
   vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
   vaSpread1.TypeDateMin = "01011973"
   vaSpread1.TypeDateMax = "31125000"
   vaSpread1.TypeHAlign = TypeHAlignCenter
   vaSpread1.ForeColor = &HFF0000
   vaSpread1.TypeDateCentury = True
   
   vaSpread1.Col = 8
   vaSpread1.text = Trim(RS!vigencia)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fg_descarga
vaSpread1.Visible = True
End Sub

Sub MoverListaProductos()
Dim i As Long
fg_carga ""
For i = 1 To 7
    Text1(i).text = ""
Next i
Dim x As Boolean
vaSpread2.TextTip = 2
vaSpread2.TextTipDelay = 250
x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread2.MaxRows = 0
vaSpread2.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_listaproductoscantidad 1")
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.text = Trim(IIf(IsNull(RS!codigo), "", RS!codigo))
   vaSpread2.Col = 2
   vaSpread2.text = Trim(IIf(IsNull(RS!CentralCompra), "", RS!CentralCompra))
   vaSpread2.Col = 3
   vaSpread2.text = Trim(IIf(IsNull(RS!descripcioncompleta), "", RS!descripcioncompleta))
   vaSpread2.Col = 4
   vaSpread2.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
   vaSpread2.Col = 5
   vaSpread2.text = Trim(IIf(IsNull(RS!unidad), "", RS!unidad))
   '-------> definir formato numerico
   vaSpread2.Col = 6
   vaSpread2.CellType = CellTypeNumber
   vaSpread2.TypeNumberDecPlaces = 0
   vaSpread2.TypeIntegerMin = 1
   vaSpread2.TypeIntegerMax = 9999
   vaSpread2.TypeHAlign = TypeHAlignRight
   vaSpread2.TypeSpin = False
   vaSpread2.TypeIntegerSpinInc = 1
   vaSpread2.TypeIntegerSpinWrap = False
   vaSpread2.ForeColor = &HFF0000
   vaSpread2.text = Trim(IIf(IsNull(RS!cantidad), "", RS!cantidad))
   '-------> definir formato fecha
   vaSpread2.Col = 7
   vaSpread2.CellType = CellTypeDate
   vaSpread2.TypeDateFormat = TypeDateFormatDDMMYY
   vaSpread2.TypeDateMin = "01011900"
   vaSpread2.TypeDateMax = "31125000"
   vaSpread2.TypeHAlign = TypeHAlignCenter
   vaSpread2.TypeDateCentury = True
   vaSpread2.ForeColor = &HFF0000
   vaSpread2.text = Trim(IIf(IsNull(RS!FechaInicio), "", Format(RS!FechaInicio, "dd/mm/yyyy")))
   '-------> Vigencia
   vaSpread2.Col = 8
   vaSpread2.text = Trim(RS!vigencia)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
fg_descarga
vaSpread2.Visible = True
End Sub

Private Sub Text1_Change(Index As Integer)
Select Case Index
Case 1, 2, 3, 4, 5, 6, 7
    vaSpread2.Visible = False
    If Trim(Text1(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread2.Col = 1
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(Text1(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select
End Sub

Private Sub Texta1_Change(Index As Integer)
Select Case Index
Case 1, 2, 3, 4, 5, 6
    vaSpread1.Visible = False
    If Trim(Texta1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index: nom = UCase(Trim(vaSpread1.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(Texta1(Index).text) & "*"
           vaSpread1.Col = 1
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(Texta1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Texta1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
    If Trim(Texta1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Texta1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    End If
    vaSpread1.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, cantidad As Long, Fecha As String, codpro As String, codcco As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '-------> Agregar
Case 3 '-------> Alterar
    Select Case SSTab1.Tab
    Case 0
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
    Case 1
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = True
    End Select
    modo = "M"
    Gl_Ac_Botones Me, 14, 0, modo
Case 5 '------> Eliminar
Case 7 '------> Actualizar
    Select Case SSTab1.Tab
    Case 0 '-------> Agregar Cantidad Producto
        MoverCantidadProductos
    Case 1 '------> Listar Cantidad Producto
        MoverListaProductos
    End Select
Case 10 '-------> Calcelar Proceso
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    Select Case SSTab1.Tab
    Case 0 '-------> Agregar Cantidad Producto
        MoverCantidadProductos
    Case 1 '------> Listar Cantidad Producto
        MoverListaProductos
    End Select
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 14, 5, modo
Case 12 '-------> Grabar
    Select Case SSTab1.Tab
    Case 0
        fg_carga ""
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = 6
            cantidad = Val(vaSpread1.text)
            vaSpread1.Col = 1
            If vaSpread1.BackColor = &H80000013 And cantidad <> 0 Then
               codpro = vaSpread1.text
               vaSpread1.Col = 2
               codcco = vaSpread1.text
               vaSpread1.Col = 6
               cantidad = Val(vaSpread1.text)
               vaSpread1.Col = 7
               Fecha = vaSpread1.text
               Set RS = vg_dbpedweb.Execute("SELECT DISTINCT producto FROM s_productosCantidad WHERE producto = '" & codpro & "' AND CentralDeCompra = '" & codcco & "'")
               If RS.EOF Then
                  vg_dbpedweb.Execute ("INSERT INTO s_productosCantidad VALUES ('" & codpro & "', '" & codcco & "', " & cantidad & ", '" & Format(Fecha, "yyyymmdd") & "') ")
               End If
               RS.Close: Set RS = Nothing
               vaSpread1.DeleteRows vaSpread1.Row, 1
               vaSpread1.MaxRows = vaSpread1.MaxRows - 1
'               vg_dbpedweb.Execute ("UPDATE s_productosCantidad SET cantidad = " & cantidad & ", FechaInicio = '" & Format(fecha, "yyyymmdd") & "' WHERE producto = '" & codpro & "' AND CentralDeCompra = '" & codcco & "'")
            End If
        Next i
'        MoverCantidadProductos
        MoverListaProductos
        fg_descarga
    Case 1
        fg_carga ""
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1
            If vaSpread2.BackColor = &H80000013 Then
               codpro = vaSpread2.text
               vaSpread2.Col = 2
               codcco = vaSpread2.text
               vaSpread2.Col = 6
               cantidad = Val(vaSpread2.text)
               vaSpread2.Col = 7
               Fecha = vaSpread2.text
               Set RS = vg_dbpedweb.Execute("SELECT DISTINCT producto FROM s_productosCantidad WHERE producto = '" & codpro & "' AND CentralDeCompra = '" & codcco & "'")
               If Not RS.EOF Then
                  vg_dbpedweb.Execute ("UPDATE s_productosCantidad SET cantidad = " & cantidad & ", FechaInicio = '" & Format(Fecha, "yyyymmdd") & "' WHERE producto = '" & codpro & "' AND CentralDeCompra = '" & codcco & "'")
               End If
               RS.Close: Set RS = Nothing
            End If
        Next i
        MoverListaProductos
        fg_descarga
    End Select
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 14, 5, modo
Case 19 '-------> Imprimir
    vg_opimp = 1
    Select Case SSTab1.Tab
    Case 0
        I_AgregarCantidadProductos
    Case 1
        I_ListarCantidadProductos
    End Select
Case 22 '-------> Salir
    Me.Hide
    Unload Me
    vg_opimp = 0
End Select
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then
   MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Select Case ButtonMenu
Case "Importar Datos"
    vg_codigo = ""
    If vaSpread1.MaxRows < 1 Then Exit Sub
    P_ImpRut.LlenaDatos "Importar Agregar Cantidad y Lista Cantidad Productos", "acprod"
    P_ImpRut.Show 1
    If Trim(vg_codigo) <> "" Then MoverListaProductos
End Select
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
Dim Fecha As Date, cantidad As Long
On Error GoTo Man_Error
vaSpread1.Row = Row
If Col = 7 Then
   vaSpread1.Col = 7
   Fecha = vaSpread1.text
   vaSpread1.text = Fecha
End If
'vaSpread1.Col = 6
'cantidad = Val(vaSpread1.Text)
'If cantidad = 0 Then
'   vaSpread1.Col = 1
'   vaSpread1.BackColor = &H80000018
'   vaSpread1.SetCellBorder 1, vaSpread1.Row, MaxCols, vaSpread1.Row, 16, &H800000, CellBorderStyleFineDot
'   Exit Sub
'End If
vaSpread1.Col = -1
vaSpread1.BackColor = &H80000013
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = False
Exit Sub
Man_Error:
If Err.Number = 13 Then vaSpread1.text = "": vaSpread1.SetActiveCell 7, vaSpread1.ActiveRow
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
'Dim fecha As Date, cantidad As Long
'On Error GoTo Man_Error
'vaSpread1.Row = Row
'If Col = 7 Then
'   vaSpread1.Col = 7
'   fecha = vaSpread1.Text
'   vaSpread1.Text = fecha
'End If
''vaSpread1.Col = 6
''cantidad = Val(vaSpread1.Text)
''If cantidad = 0 Then
''   vaSpread1.Col = -1
''   vaSpread1.BackColor = &H80000018
''   vaSpread1.SetCellBorder 1, vaSpread1.Row, vaSpread1.MaxCols, vaSpread1.Row, 16, &H800000, CellBorderStyleFineDot
''   Exit Sub
''End If
'vaSpread1.Col = -1
'vaSpread1.BackColor = &H80000013
'If Toolbar1.Buttons(12).Visible = True Then Exit Sub
'If modo = "" Then modo = "M"
'Gl_Ac_Botones Me, 14, 0, modo
'SSTab1.TabEnabled(0) = True
'SSTab1.TabEnabled(1) = False
'Exit Sub
'Man_Error:
'If Err.Number = 13 Then vaSpread1.Text = "": vaSpread1.SetActiveCell 7, vaSpread1.ActiveRow
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread1.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 1
    vaSpread1.Col = Col
    TipText = "Código Producto : " & vaSpread1.text
Case 2
    vaSpread1.Col = Col
    TipText = "Central de Compras : " & Trim(vaSpread1.text)
Case 3
    vaSpread1.Col = Col
    TipText = "Familia : " & Trim(vaSpread1.text)
Case 4
    vaSpread1.Col = Col
    TipText = "Descripción : " & Trim(vaSpread1.text)
End Select
End Sub

Private Sub vaSpread2_Change(ByVal Col As Long, ByVal Row As Long)
Dim Fecha As Date
On Error GoTo Man_Error
vaSpread2.Row = Row
On Error GoTo Man_Error
vaSpread2.Row = Row
If Col = 7 Then
   vaSpread2.Col = 7
   Fecha = vaSpread2.text
   vaSpread2.text = Fecha
End If
vaSpread2.Col = -1
vaSpread2.BackColor = &H80000013
If Toolbar1.Buttons(12).Visible = True Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = True
Exit Sub
Man_Error:
If Err.Number = 13 Then vaSpread2.text = "": vaSpread2.SetActiveCell 7, vaSpread2.ActiveRow
End Sub

Private Sub vaSpread2_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread2.MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread2.Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 1
    vaSpread2.Col = Col
    TipText = "Código Producto : " & vaSpread2.text
Case 2
    vaSpread2.Col = Col
    TipText = "Central de Compras : " & Trim(vaSpread2.text)
Case 3
    vaSpread2.Col = Col
    TipText = "Familia : " & Trim(vaSpread2.text)
Case 4
    vaSpread2.Col = Col
    TipText = "Descripción : " & Trim(vaSpread2.text)
End Select

End Sub
