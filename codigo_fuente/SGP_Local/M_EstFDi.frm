VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_EstFDi 
   Caption         =   "Estructura Fija x Día"
   ClientHeight    =   7545
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   14175
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   14175
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3900
      ScaleHeight     =   1035
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   210
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar gauge1 
         Height          =   210
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Día"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   195
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   5055
      Left            =   120
      TabIndex        =   11
      Top             =   1320
      Width           =   10935
      _Version        =   393216
      _ExtentX        =   19288
      _ExtentY        =   8916
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
      SpreadDesigner  =   "M_EstFDi.frx":0000
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   45
      TabIndex        =   5
      Top             =   840
      Width           =   10905
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Bloqueada"
         Height          =   195
         Index           =   1
         Left            =   7650
         TabIndex        =   8
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   7305
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         Height          =   195
         Index           =   0
         Left            =   9405
         TabIndex        =   7
         Top             =   135
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   9060
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Semana Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   14175
      _ExtentX        =   25003
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estructura Fija x Día Teórica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   15
      TabIndex        =   10
      Top             =   345
      Width           =   10905
   End
   Begin VB.Menu Main 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Index           =   0
   End
   Begin VB.Menu Main 
      Caption         =   "&Producto Menú"
      Index           =   10
      Begin VB.Menu Plato 
         Caption         =   "Cambia Producto"
         Index           =   0
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu Plato 
         Caption         =   "Insertar Línea"
         Index           =   20
      End
      Begin VB.Menu Plato 
         Caption         =   "Eliminar Línea"
         Index           =   30
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu Plato 
         Caption         =   "Cortar"
         Index           =   50
      End
      Begin VB.Menu Plato 
         Caption         =   "Copiar"
         Index           =   60
      End
      Begin VB.Menu Plato 
         Caption         =   "Pegar"
         Index           =   70
      End
      Begin VB.Menu Plato 
         Caption         =   "Pegado Especial"
         Index           =   80
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Salir"
      Index           =   20
   End
   Begin VB.Menu MenuDetalle 
      Caption         =   ""
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu OpGrilla 
         Caption         =   "Cambia Producto"
         Index           =   0
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Insertar Línea"
         Index           =   20
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Eliminar Línea"
         Index           =   30
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Cortar"
         Index           =   50
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Copiar"
         Index           =   60
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Pegar"
         Index           =   70
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Pegado Especial"
         Index           =   80
      End
   End
End
Attribute VB_Name = "M_EstFDi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim IndCortarPegar As Long, Fecha As Long, MaxColumna As Long, MaxFila As Long, AddReceta As Long
Dim IblockRow As Integer, IblockRow2 As Integer, IblockCol As Integer, iblockcol2 As Integer, SwSalir As Integer
Dim AiBlockRow As Integer, AiBlockRow2 As Integer, AiBlockCol As Integer, AiBlockCol2 As Integer, IndActivo As Integer
Dim VecCos() As Variant
Dim VecCosenc() As Variant
Dim VectorCol() As Long
Dim Msgtitulo As String, tipmin As String
Public lc_Aux1 As String

Private Sub Form_Activate()
fg_descarga
'-------> Traer fecha cierre día
TraerFechaCierre
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6765
Me.Width = 11055
fg_centra Me
fg_carga ""
If lc_Aux1 = "EstTeo" Then
    Msgtitulo = "Estructura Fija x Día Teórica"
    Me.Caption = "Estructura Fija x Día Teórica"
    tipmin = "1"
ElseIf lc_Aux1 = "EstRea" Then
    Msgtitulo = "Estructura Fija x Día Real"
    Me.Caption = "Estructura Fija x Día Real"
    tipmin = "2"
End If
Dim nomser As String, nomreg As String
'-------> Traer nombre regimen
RS.Open "SELECT reg_nombre FROM a_regimen WHERE reg_codigo=" & vg_codregimen & "", vg_db, adOpenStatic
If Not RS.EOF Then nomreg = Trim(RS!reg_nombre)
RS.Close: Set RS = Nothing
'-------> Traer nombre servicio
RS.Open "SELECT ser_nombre FROM a_servicio WHERE ser_codigo=" & vg_codservicio & "", vg_db, adOpenStatic
If Not RS.EOF Then nomser = Trim(RS!ser_nombre)
RS.Close: Set RS = Nothing
Label4.Caption = M_Plami1.fpayuda(0).Caption & "(" & M_Plami1.fpText.text & ")" & " - " & nomreg & " - " & nomser
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = " "
Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = "Grabar Datos": BtnX.Enabled = IIf(Mid(ValidarUsuario(M_Plami1), 2, 2) = "0", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Cortar", , tbrDefault, "A_Cortar"): BtnX.Visible = True: BtnX.ToolTipText = "Cortar"
Set BtnX = Toolbar1.Buttons.Add(, "A_Copiar", , tbrDefault, "A_Copiar"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar"
Set BtnX = Toolbar1.Buttons.Add(, "I_Pegar", , tbrDefault, "I_Pegar"): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, "A_Pegar", , tbrDefault, "A_Pegar"): BtnX.Visible = False: BtnX.ToolTipText = "Pegar"
Set BtnX = Toolbar1.Buttons.Add(, "I_PegadoEspecial", , tbrDefault, "I_PegadoEspecial"): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, "A_PegadoEspecial", , tbrDefault, "A_PegadoEspecial"): BtnX.Visible = False: BtnX.ToolTipText = "Pegado Especial"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): BtnX.Visible = True: BtnX.ToolTipText = "Insertar"
Set BtnX = Toolbar1.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): BtnX.Visible = True: BtnX.ToolTipText = "Eliminar"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
DetallePlantilla
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435
If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 445
If Me.WindowState <> 1 Then vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380
End Sub

Private Sub Form_Unload(Cancel As Integer)
If SwSalir <> 0 Then Exit Sub
If Toolbar1.Buttons(1).Visible = True Then Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
If MsgBox(" Actualiza estructura fija día " & IIf(tipmin = "1", "Teórica...", "Real..."), vbQuestion + vbYesNo, Msgtitulo) = vbNo Then
   'Cancel = -1
   Me.Hide
   Unload Me
   M_Plami1.WindowState = 0
   Exit Sub
End If
If Toolbar1.Buttons(2).Visible = True Then GrabarPlantilla
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
SwSalir = 1
Me.Hide
Unload Me
M_Plami1.WindowState = 0
End Sub

Private Sub Main_Click(Index As Integer)
Select Case Index
Case 0
    If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
    If MsgBox(" Actualiza estructura fija día " & IIf(tipmin = "1", "Teórica...", "Real..."), vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Cancel = -1: Exit Sub
    If Toolbar1.Buttons(2).Visible = True Then GrabarPlantilla
    Main(0).Enabled = False
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
Case 20
    SwSalir = 0
    If Toolbar1.Buttons(1).Visible = True Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
    If MsgBox(" Actualiza estructura fija día " & IIf(tipmin = "1", "Teórica...", "Real..."), vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Toolbar1.Buttons(2).Visible = False
    If Toolbar1.Buttons(2).Visible = True Then GrabarPlantilla
    SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
End Select
End Sub

Private Sub Plato_Click(Index As Integer)
Dim Del_Row As Integer, indcol As Integer, indrow As Integer, IndCol2 As Integer, IndRow2 As Integer, indrow3 As Long, XX As Long
Dim Col As Long, fil As Long, AddRec As Long
If Main(10).Enabled = False Then Exit Sub
With vaSpread1
    Select Case Index
    Case 0 '-------> Ingreso producto
        IblockCol = .ActiveCol: AiBlockCol = .ActiveCol
        iblockcol2 = .ActiveCol: AiBlockCol2 = .ActiveCol
        IblockRow = .ActiveRow: AiBlockRow = .ActiveRow
        IblockRow2 = .ActiveRow: AiBlockRow2 = .ActiveRow
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        j = 0
        For i = 1 To MaxColumna
            If .Col = VectorCol(i) Then j = VectorCol(i): Exit For
        Next i
        If j = 0 Then Exit Sub
        .Col = j - 1
        .Row = .ActiveRow
        vg_nombre = "": vg_codigo = ""
        vg_left = 2300
        B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
        B_TabEst.Show 1
        If vg_codigo = "" Or Trim(vg - Nombre) = "" Then Exit Sub
        If ValidarProducto(vg_codigo, .Row, j - 1) Then MsgBox "Producto ya existe...", vbExclamation + vbOKOnly, Msgtitulo: .SetActiveCell j, .ActiveRow: Exit Sub
        .Row = .ActiveRow
        '-------> Mover código producto
        .Col = j - 1
        .CellType = CellTypeEdit
        .TypeHAlign = TypeHAlignRight
        .text = Trim(vg_codigo)
        '-------> Mover descripción del producto
        .Col = j
        '-------> Limpiar Datos y Formato Celda
        .Action = 3
        '-------> Retorna Modo de la columna
        .BlockMode = False
        .Font.Bold = False
        .Font.Size = 8
        .text = vg_nombre
        '-------> Mover valor cero a cantidad del producto
        .Col = j + 1
        If Trim(.text) = "" Then
           .Row = .ActiveRow
           .Col = j + 1
           .CellType = CellTypeNumber
           .TypeNumberDecPlaces = 2
           .TypeIntegerMin = 1
           .TypeIntegerMax = 9999999
           .TypeHAlign = 1
           .TypeSpin = False
           .TypeIntegerSpinInc = 1
           .TypeIntegerSpinWrap = False
           .text = 0
           .ForeColor = &HFF0000
        End If
        '-------> Mover descripción unidad medida
        .Col = j + 2
        .CellType = CellTypeStaticText
        .TypeHAlign = TypeHAlignLeft
        .text = ""
        '-------> Traer descripción unidad medida
    '    RS.Open "SELECT c.ppd_propon, b.uni_nomcor  FROM b_productos a, a_unidad b, b_productospmpdia c " & _
    '            "WHERE a.pro_codigo = c.ppd_codpro " & _
    '            "AND   a.pro_coduni = b.uni_codigo " & _
    '            "AND   a.pro_codigo = '" & vg_codigo & "' " & _
    '            "AND   c.ppd_cencos = '" & MuestraCasino(1) & "' " & _
    '            "AND   c.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & "", vg_db, adOpenStatic
        RS.Open "SELECT DISTINCT b.uni_nomcor  FROM b_productos a, a_unidad b " & _
                "WHERE a.pro_coduni = b.uni_codigo " & _
                "AND   a.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
        If Not RS.EOF Then
           .text = Trim(RS!uni_nomcor)
           Dim propon As Double
           propon = 0
           RS1.Open "SELECT ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                    "FROM b_productospmpdia " & _
                    "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                    "AND   ppd_codpro = '" & vg_codigo & "' " & _
                    "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                    "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
                    "HAVING (ppd_propon)>0", vg_db, adOpenStatic
           If Not RS1.EOF Then propon = RS1!ppd_propon
           RS1.Close: Set RS1 = Nothing
           .Col = j + 3
           .CellType = CellTypeStaticText
           .TypeHAlign = TypeHAlignLeft
           .text = propon
        End If
        RS.Close: Set RS = Nothing
        .Row = .ActiveRow
        Main(0).Enabled = True
        Plato(70).Enabled = False: Plato(80).Enabled = False
        OpGrilla(70).Enabled = False: OpGrilla(80).Enabled = False
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        .SetActiveCell j + 1, .ActiveRow
    Case 20 '-------> Insertar línea
        indcol = IblockCol
        IblockCol = 1: iblockcol2 = .MaxCols
        .MaxRows = .MaxRows + ((IblockRow2 - IblockRow) + 1)
        .InsertRows IblockRow, ((IblockRow2 - IblockRow) + 1)
        For i = 1 To (.MaxCols - MaxColumna) Step 5
            .Row = SpreadHeader: .Col = i
            If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
                Dim f As Long, c As Long
                For c = i To i + 4
                    .Row = IblockRow: .Col = c
                    .BackColor = Shape1(1).FillColor
                Next c
            End If
        Next i
        '-------> Validar días modificados
        For j = IblockRow To ((.MaxRows - 1) - ((IblockRow2 - IblockRow) + 1))
            For i = 1 To (.MaxCols - MaxColumna) Step 5
                .Row = j
                .Col = (MaxColumna * 5 + 1) + ((i + 3) / 5)
                If Trim(.text) = "" Then .text = 2
            Next i
        Next j
        IblockCol = indcol
        Main(0).Enabled = True
        Plato(70).Enabled = False: Plato(80).Enabled = False
        OpGrilla(70).Enabled = False: OpGrilla(80).Enabled = False
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    Case 30 '-------> Eliminar línea
        indcol = IblockCol
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .BackColor = Shape1(1).FillColor And Trim(.text) <> "" Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        If .BackColor = Shape1(1).FillColor Or Trim(.text) = "" Then GoTo paso
        j = 0
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = .Col Or VectorCol(i) = .Col Then j = (VectorCol(i) - 1): Exit For
        Next i
        If j = 0 Then Exit Sub
        If IndActivo = 0 Then IblockCol = .ActiveCol: iblockcol2 = .ActiveCol: IblockRow = .ActiveRow: IblockRow2 = .ActiveRow
        AiBlockCol = IblockCol
        AiBlockRow = IblockRow
        AiBlockCol2 = iblockcol2
        AiBlockRow2 = IblockRow2
        If IblockCol < 0 Then IblockCol = 1: iblockcol2 = .MaxCols
        AiBlockCol = IblockCol
        AiBlockRow = IblockRow
        AiBlockCol2 = iblockcol2
        AiBlockRow2 = IblockRow2
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 4)): Exit For
            If VectorCol(i) = iblockcol2 Then iblockcol2 = (VectorCol(i) + 3): Exit For
        Next i
        indcol = AiBlockCol: IndCol2 = iblockcol2
        indrow = AiBlockRow: IndRow2 = AiBlockRow2
        '-------> Validar días modificados
        For j = IblockRow To ((.MaxRows - 1) - ((IblockRow2 - IblockRow) + 1))
            For i = 1 To (.MaxCols - MaxColumna) Step 5
                .Row = j
                .Col = i + 1
                If Trim(.text) <> "" Then
                   .Col = (MaxColumna * 5 + 1) + ((i + 3) / 5)
                   If Trim(.text) = "" Then .text = 2
                End If
            Next i
        Next j
        '-------> Fin validar días modificados
        IblockCol = AuxCol
        .BlockMode = False
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        IndActivo = 0
paso:
        .Row = .ActiveRow
        For i = IblockCol To iblockcol2
            .Col = i
            For j = IblockRow To IblockRow2
                .Row = j
                If .BackColor = Shape1(1).FillColor Then MsgBox "Existen días bloqueado, no puede eliminar fila", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
            Next j
        Next i
        .Row = IblockRow2
        .Col = IblockCol
        .Visible = False
        .DeleteRows IblockRow, 1
        .MaxRows = .MaxRows - 1
        .Visible = True
        IblockCol = indcol
        For i = 1 To (.MaxCols - MaxColumna) Step 5
            .Row = SpreadHeader: .Col = i
            If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
                For Col = 0 To i - 4
                    .Row = (.MaxRows - 1): .Col = Col + 2
                    .BackColor = Shape1(1).FillColor
                Next Col
            End If
        Next i
        '-------> Validar días modificados
        For j = IblockRow To ((.MaxRows - 1) - ((IblockRow2 - IblockRow) + 1))
            For i = 1 To (.MaxCols - MaxColumna) Step 5
                .Row = j
                .Col = i + 1
                If Trim(.text) <> "" Then
                   .Col = (MaxColumna * 5 + 1) + ((i + 3) / 5)
                   If Trim(.text) = "" Then .text = 2
                End If
            Next i
        Next j
        '-------> Fin validar días modificados
        Main(0).Enabled = True
        Plato(70).Enabled = False: Plato(80).Enabled = False
        OpGrilla(70).Enabled = False: OpGrilla(80).Enabled = False
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    Case 50, 60 '-------> Copiar y pegar línea
        If .ActiveRow = .MaxRows Then Exit Sub
        If Index = 50 Then
           If IblockCol < 1 Then
              For i = 1 To MaxColumna
                  .Col = VectorCol(i)
                  .Row = 1
                  If .BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede usar cortar", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
              Next i
           Else
              For i = IblockCol To iblockcol2
                  .Col = i
                  For j = IblockRow To IblockRow2
                     .Row = j
                     If .BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede usar cortar", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
                  Next j
              Next i
           End If
           '-------> Validar recetas 5 etapas
           j = 0
           For i = 1 To MaxColumna
               If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Then j = (VectorCol(i) - 1): Exit For
           Next i
           If j = 0 Then Exit Sub
           If etapa5 And AddReceta > 0 Then
              For j = j To iblockcol2 Step 5
                  .Col = j
                  For i = IblockRow To (IblockRow2)
                      .Row = i
                      If .BackColor = &H80FF80 Then MsgBox "No puede cortar receta, corresponde 5 etapas", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
                  Next i
              Next j
           End If
        End If
        .Row = .ActiveRow
        .Col = .ActiveCol
        AiBlockRow = IblockRow: AiBlockRow2 = IblockRow2
        AiBlockCol = IblockCol: AiBlockCol2 = iblockcol2
        If .Col = 1 Then Exit Sub
        Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(7).Visible = True
        Plato(70).Enabled = True: OpGrilla(70).Enabled = True
        Plato(80).Enabled = True: OpGrilla(80).Enabled = True
        If IblockCol < 1 Then AiBlockCol = 1: AiBlockCol2 = .MaxCols
        IndCortarPegar = 1
        If Index = 50 Then IndCortarPegar = 0: Toolbar1.Buttons(8).Visible = True: Toolbar1.Buttons(9).Visible = False: 'Plato(14).Enabled = False: OpGrilla(14).Enabled = False Else Toolbar1.Buttons(8).Visible = False: Toolbar1.Buttons(9).Visible = True: Plato(14).Enabled = True: OpGrilla(14).Enabled = True
    Case 70, 80 '-------> Copiar ó Pegar
        If IndCortarPegar = 0 Then
           If (iblockcol2 - IblockCol) > (AiBlockCol2 - AiBlockCol) Or (IblockRow2 - IblockRow) > (AiBlockRow2 - AiBlockRow) Then MsgBox "Imposible Pegar la infomación ya que el área de Cortar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    '      If IBlockCol2 > AIBlockCol2 Then
    '         MsgBox "Imposible Cortar la infomación ya que el área de Cortar y el área de Pegado tienen formas distintas", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
    '         Exit Sub
     '     End If
           IndCortarPegar = 0
        Else
           If (IblockRow2 - IblockRow) > (AiBlockRow2 - AiBlockRow) Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
           If AiBlockCol <> iblockcol2 And AiBlockCol = 1 Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
        End If
        If IblockCol < 1 Then
           For i = 1 To MaxColumna
               .Col = VectorCol(i)
               .Row = 1
               If .BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
           Next i
        Else
           For i = IblockCol To iblockcol2
               .Col = i
               For j = IblockRow To IblockRow2
                  .Row = j
                  If .BackColor = Shape1(1).FillColor And Index <> 80 Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
               Next j
           Next i
        End If
        Dim codigo As String
        For jx = AiBlockRow To AiBlockRow2
            .Row = jx
            For ix = AiBlockCol To AiBlockCol2 Step 5
                j = 0
                For X = 1 To MaxColumna
                    If ix = VectorCol(X) Or (ix + 1) = VectorCol(X) Or (ix - 1) = VectorCol(X) Or (ix - 2) = VectorCol(X) Then j = VectorCol(X): Exit For
                Next X
                If j = 0 Then Exit Sub
                .Col = j - 1: codigo = .text
                For i = IblockCol To iblockcol2 Step 5
                    j = 0
                    For X = 1 To MaxColumna
                        If i = VectorCol(X) Then j = VectorCol(X): Exit For
                    Next X
                    If j = 0 Then Exit Sub
                    If ValidarProducto(codigo, .Row, j - 1) Then MsgBox "Producto ya existe...", vbExclamation + vbOKOnly, Msgtitulo: .SetActiveCell j, .ActiveRow: Exit Sub
                Next i
            Next ix
        Next jx
        .Col = .ActiveCol
        If .Col = 1 Then Exit Sub
        If IndCortarPegar = 0 Then Toolbar1.Buttons(6).Visible = True: Toolbar1.Buttons(7).Visible = False
        '-------> destinacion de copiar y pegar datos
        If IblockCol < 1 Then IblockCol = 1: iblockcol2 = .MaxCols
        If AiBlockCol2 = .MaxCols Then AiBlockCol2 = .MaxCols - 1
        .Row = 0: .Col = IblockCol
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Or (VectorCol(i) + 1) = IblockCol Or (VectorCol(i) + 2) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = AiBlockCol Or VectorCol(i) = AiBlockCol Or (VectorCol(i) + 1) = AiBlockCol Or (VectorCol(i) + 2) = AiBlockCol Then AiBlockCol = (VectorCol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = iblockcol2 Or VectorCol(i) = iblockcol2 Or (VectorCol(i) + 1) = iblockcol2 Or (VectorCol(i) + 2) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 3)): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = AiBlockCol2 Or VectorCol(i) = AiBlockCol2 Or (VectorCol(i) + 1) = AiBlockCol2 Or (VectorCol(i) + 2) = AiBlockCol2 Then AiBlockCol2 = (VectorCol(i) + 3): Exit For
        Next i
        indcol = AiBlockCol: IndCol2 = iblockcol2
        indrow = AiBlockRow: IndRow2 = AiBlockRow2
        If Index = 80 And IndCortarPegar = 1 Then
           If (AiBlockRow2 - AiBlockRow) <> 0 Or (AiBlockCol2 - AiBlockCol) <> 4 Then MsgBox "Por esta opción solamente puede copiar una producto", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
           '-------> Rutina pegado especial
           Dim nrodia As String
           .Row = SpreadHeader: nrodia = ""
           For i = AiBlockCol To AiBlockCol2 Step 5
               .Col = i '+ 1
               nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd"), 7, 2)) & ";"
           Next i
           For i = 1 To MaxColumna
               .Col = VectorCol(i) - 1
               .Row = 1
               If .BackColor = Shape1(1).FillColor Then .Row = SpreadHeader: nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd"), 7, 2)) & ";"
           Next i
           vg_codigo = ""
           Call M_CpRPla.Inicio("Copia Especial Recetas en Planificación Real", "PLAREA", Vg_FechaDesde, Vg_FechaHasta, nrodia, 1)
           M_CpRPla.Show 1
           If Trim(vg_codigo) = "" Then Exit Sub
           Dim vecdia() As String
           Dim xSer As Long, iSer As Long
           'mover días no permitidos
           ReDim Preserve vecdia(0)
           ValLcntH = "": i = 0
           For j = 1 To Len(vg_codigo)
               If Asc(Mid(vg_codigo, j, 1)) <> 59 Then
                  ValLcntH = ValLcntH + Mid(vg_codigo, j, 1)
               Else
                  ReDim Preserve vecdia(i): vecdia(i) = ValLcntH: ValLcntH = "": i = i + 1
               End If
           Next j
           If Trim(ValLcntH) <> "" Then ReDim Preserve vecdia(i): vecdia(i) = ValLcntH
           
           For i = 1 To (.MaxCols - MaxColumna) Step 5
               .Row = AiBlockRow
               .Col = .MaxCols
               iSer = Val(.text)
               .Row = SpreadHeader
               .Col = i
               L = 0
               nrodia = Val(Mid(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd"), 7, 2))
               For j = 0 To UBound(vecdia)
                   If nrodia = vecdia(j) Then
                      .Row = AiBlockRow: .Col = i ' - 1
                      If Trim(.text) <> "" Then
                         For X = AiBlockRow + 1 To .MaxRows
                             .Row = X: .Col = .MaxCols: xSer = Val(.text)
                             .Col = i + 1
                             If .Row = .MaxRows Then .MaxRows = .MaxRows + 1: .InsertRows X, 1: L = X: Exit For
                             If xSer <> iSer And xSer > 0 Then
                                .MaxRows = .MaxRows + 1: .InsertRows X, 1: L = X: Exit For
                             ElseIf Trim(.text) <> "" And xSer > 0 Then
                                .MaxRows = .MaxRows + 1: .InsertRows X + 1, 1: X = X + 1: L = X: Exit For
                             ElseIf Trim(.text) = "" Then
                                Exit For
                             End If
                         Next X
    '                     .CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, X
                         .CopyRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, X
                         .Row = X
                      Else
                         .CopyRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, AiBlockRow
                         .Row = AiBlockRow
                      End If
                      '-------> Asignar colores
                      For X = (i - 1) To (i - 1) + 4
                          .Col = X
                          .BackColor = Shape1(0).FillColor
                          For XX = 1 To MaxColumna
                              If (VectorCol(XX) - 1) = .Col Then
                                  .Col = X + 2
                                  .CellType = CellTypeNumber
                                  .TypeNumberDecPlaces = 2
                                  .TypeIntegerMin = 1
                                  .TypeIntegerMax = 9999999
                                  .TypeHAlign = TypeHAlignRight
                                  .TypeSpin = False
                                  .TypeIntegerSpinInc = 1
                                  .TypeIntegerSpinWrap = False
                                  Exit For
                              End If
                          Next XX
                          .Col = X
    '                      If X = (i - 1) Then .ForeColor = &HFF&: .BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
                      Next X
                      If L > 0 Then
                         z = L
                         For L = 2 To (.MaxCols - MaxColumna) Step 5
                             .Row = 1: .Col = L
                             If .BackColor = Shape1(1).FillColor Then
                                .Row = z
                                For X = (L - 1) To (L - 1) + 4
                                    .Col = X
                                    .BackColor = Shape1(1).FillColor
                                Next X
                             End If
                         Next L
                      End If
                      '-------> Fin asignar colores
                      
                      '-------> Validar días modificados
                      For z = AiBlockRow To .ActiveRow + (AiBlockRow2 - AiBlockRow)
                          .Row = z
                          .Col = i
                          If Trim(.text) <> "" Then
                             .Col = (MaxColumna * 5 + 1) + ((i + 2) / 5)
                             .text = 1
                          End If
                      Next z
                      '-------> Fin validar días modificados
                      Exit For
                   End If
               Next j
           Next i
        Else
           indrow3 = .MaxRows
           For i = IblockCol To iblockcol2 Step 5
               If IndCortarPegar = 1 Then
                  .Row = AiBlockRow: .Col = AiBlockCol
                     .MaxRows = .MaxRows + (AiBlockRow2 - AiBlockRow) + 1
                     .CopyRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, .MaxRows - (AiBlockRow2 - AiBlockRow)
                     '-------> Asignar colores
                     For j = .MaxRows - (AiBlockRow2 - AiBlockRow) To .MaxRows
                         .Row = j
                         For X = (i) To (i) + 4
                             .Col = X
                             .BackColor = Shape1(0).FillColor
                             For XX = 1 To MaxColumna
                                 If (VectorCol(XX) - 1) = .Col Then
                                    .Col = X + 2
                                    .CellType = CellTypeNumber
                                    .TypeNumberDecPlaces = 2
                                    .TypeIntegerMin = 1
                                    .TypeIntegerMax = 9999999
                                    .TypeHAlign = TypeHAlignRight
                                    .TypeSpin = False
                                    .TypeIntegerSpinInc = 1
                                    .TypeIntegerSpinWrap = False
                                    Exit For
                                 End If
                             Next XX
                             .Col = X
    '                         If X = (i) And Trim(.Text) <> "" Then .ForeColor = &HFF&: .BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
                         Next X
                     Next j
                     '-------> Fin asignar colores
                     .CopyRange IblockCol, .MaxRows - (AiBlockRow2 - AiBlockRow), iblockcol2, .MaxRows, i, .ActiveRow
                     .MaxRows = indrow3
               ElseIf IndCortarPegar = 0 Then
                  .Row = AiBlockRow: .Col = AiBlockCol
                  If .BackColor = Shape1(1).FillColor Then
                     .MaxRows = .MaxRows + (AiBlockRow2 - AiBlockRow) + 1
                     .MoveRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, .MaxRows - (AiBlockRow2 - AiBlockRow)
                     '-------> Asignar colores
                     For j = .MaxRows - (AiBlockRow2 - AiBlockRow) To .MaxRows
                         .Row = j
                         For X = (i) To (i) + 4
                             .Col = X
                             .BackColor = Shape1(0).FillColor
                             For XX = 1 To MaxColumna
                                 If (VectorCol(XX) - 1) = .Col Then
                                    .Col = X + 2
                                    .CellType = CellTypeNumber
                                    .TypeNumberDecPlaces = 0
                                    .TypeIntegerMin = 1
                                    .TypeIntegerMax = 9999999
                                    .TypeHAlign = TypeHAlignRight
                                    .TypeSpin = False
                                    .TypeIntegerSpinInc = 1
                                    .TypeIntegerSpinWrap = False
                                    Exit For
                                 End If
                             Next XX
                             .Col = X
    '                         If X = (i) And Trim(.Text) <> "" Then .ForeColor = &HFF&: .BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
                         Next X
                     Next j
                     '-------> Fin asignar colores
                     .MoveRange IblockCol, .MaxRows - (AiBlockRow2 - AiBlockRow), iblockcol2, .MaxRows, i, .ActiveRow
                     .MaxRows = indrow3
                  Else
                     .MoveRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, .ActiveRow
                  End If
               End If
               For j = .ActiveRow To .ActiveRow + (AiBlockRow2 - AiBlockRow)
                   .Row = j
                   For X = AiBlockCol To AiBlockCol2 Step 5
                       .Col = X + 1
                       If Trim(.text) <> "" Then
                          .Col = (MaxColumna * 5 + 1) + ((X + 3) / 5)
                          .text = 1
                       End If
                   Next X
               Next j
               '-------> Fin validar días modificados
           Next i
        End If
        AiBlockCol = indcol: iblockcol2 = IndCol2
        AiBlockRow = indrow: AiBlockRow2 = IndRow2
        Main(0).Enabled = True
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
    End Select
End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2 '-------> Grabar
    Main_Click (0)
Case 4 '-------> Cortar
    Plato_Click (50)
Case 5 '-------> Copiar
    Plato_Click (60)
Case 7 '-------> Pegar
    Plato_Click (70)
Case 9 '-------> Pegado especial
    Plato_Click (80)
Case 11 '-------> Insertar línea
    Plato_Click (20)
Case 12 '-------> Eliminar línea
    Plato_Click (30)
Case 14 '-------> Salir
    Main_Click (20)
End Select
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
IndActivo = 1
IblockRow = BlockRow
IblockRow2 = BlockRow2
IblockCol = BlockCol
iblockcol2 = BlockCol2
If BlockRow < 0 Then IblockRow = 1
If BlockRow2 < 0 Then IblockRow2 = (vaSpread1.MaxRows - 1)
If BlockRow2 >= vaSpread1.MaxRows Then IblockRow2 = (vaSpread1.MaxRows - 1)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Then Exit Sub
With vaSpread1
    IndActivo = 1
    IblockRow = .ActiveRow
    IblockRow2 = .ActiveRow
    IblockCol = .ActiveCol
    iblockcol2 = .ActiveCol
    .Row = .ActiveRow
    .Col = .ActiveCol
End With
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Or Col = 1 Then Exit Sub
Plato_Click (0)
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim i As Long, j As Long, codigo As String
With vaSpread1
    If .MaxRows < 1 Or ChangeMade = False Then Exit Sub
    .Row = Row
    .Col = Col
    'If Trim(.Text) = "" Or .CellType = CellTypeStaticText Or .CellType = CellTypeNumber Then Exit Sub
    If Trim(.text) = "" Or .CellType = CellTypeStaticText Then Exit Sub
    j = 0
    For i = 1 To MaxColumna
        If Col = (VectorCol(i) - 1) Or Col = (VectorCol(i) + 1) Then j = VectorCol(i): Exit For
    Next i
    If j = 0 Then Exit Sub
    .Row = Row
    Select Case Col
    Case 1, 6, 11, 16, 21, 26, 31, 36, 41, 46, 51, 56, 61, 66, 71, 76, 81, 86, 91, 96, 101, 106, 111, 116, 121, 126, 131, 136, 141, 146, 151
       .Col = j - 1
       codigo = Trim(.text)
       If ValidarProducto1(codigo, .Row, j - 1) Then
          MsgBox "Producto ya existe...", vbExclamation + vbOKOnly, Msgtitulo
          .Row = Row
          .Col = j - 1
          .text = ""
          .Col = j
          .text = ""
          .Col = j + 1
          .text = ""
          .Col = j + 2
          .text = ""
          .Col = j + 3
          .text = ""
          .SetActiveCell j - 1, .ActiveRow
          Exit Sub
       End If
    '   RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_ctrsto, b.uni_nomcor, c.ppd_propon " & _
    '           "FROM b_productos a, a_unidad b, b_productospmpdia c, a_tiposervicio d, b_clientes e " & _
    '           "WHERE (d.tis_codigo = e.cli_codtis OR a.pro_maepro < 1) " & _
    '           "AND    e.cli_codigo = '" & MuestraCasino(1) & "' " & _
    '           "AND   (d.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
    '           "AND    a.pro_codigo = c.ppd_codpro " & _
    '           "AND    a.pro_coduni = b.uni_codigo " & _
    '           "AND    a.pro_codigo = '" & codigo & "' " & _
    '           "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
    '           "AND    a.pro_ctrsto = 1 " & _
    '           "AND    c.ppd_cencos = '" & MuestraCasino(1) & "' " & _
    '           "AND    c.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & "", vg_db, adOpenStatic
       RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_ctrsto, b.uni_nomcor " & _
               "FROM b_productos a, a_unidad b, a_tiposervicio d, b_clientes e " & _
               "WHERE (d.tis_codigo = e.cli_codtis OR a.pro_maepro < 1) " & _
               "AND    e.cli_codigo = '" & MuestraCasino(1) & "' " & _
               "AND   (d.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
               "AND    a.pro_coduni = b.uni_codigo " & _
               "AND    a.pro_codigo = '" & codigo & "' " & _
               "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) " & _
               "AND    a.pro_ctrsto = 1", vg_db, adOpenStatic
       If RS.EOF Then
          RS.Close: Set RS = Nothing
          .Row = Row
          .Col = j - 1
          .text = ""
          .Col = j
          .text = ""
          .Col = j + 1
          .text = ""
          .Col = j + 2
          .text = ""
          .Col = j + 3
          .text = ""
          .SetActiveCell j - 1, .ActiveRow
          Exit Sub
       End If
       .Row = Row
       '-------> Mover código producto
       .Col = j - 1
       .CellType = CellTypeEdit
       .TypeHAlign = TypeHAlignRight
       .text = Trim(RS!pro_codigo)
    
       '-------> Mover descripción del producto
       .Col = j
       '-------> Limpiar Datos y Formato Celda
       .Action = 3
       '-------> Retorna Modo de la columna
       .BlockMode = False
       .Font.Bold = False
       .Font.Size = 8
       .text = Trim(RS!pro_nombre)
       '-------> Mover a cero cantidad del producto
       .Col = j + 1
       If Trim(.text) = "" Then
          .Row = .ActiveRow
          .Col = j + 1
          .CellType = CellTypeNumber
          .TypeNumberDecPlaces = 2
          .TypeIntegerMin = 1
          .TypeIntegerMax = 9999999
          .TypeHAlign = 1
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
          .text = 0
          .ForeColor = &HFF0000
       End If
    
       '-------> Mover descripción unidad medida
       .Col = j + 2
       .CellType = CellTypeStaticText
       .TypeHAlign = TypeHAlignLeft
       .text = Trim(RS!uni_nomcor)
       RS.Close: Set RS = Nothing
       
       '-------> Traer precio promedio ponderado
       Dim propon As Double
       propon = 0
       RS.Open "SELECT ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
               "FROM b_productospmpdia " & _
               "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
               "AND   ppd_codpro = '" & codigo & "' " & _
               "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
               "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
               "HAVING (ppd_propon)>0", vg_db, adOpenStatic
       If Not RS.EOF Then propon = RS!ppd_propon
       RS.Close: Set RS = Nothing
       '-------> Mover precio ponderado
       .Col = j + 3
       .CellType = CellTypeStaticText
       .TypeHAlign = TypeHAlignLeft
       .text = propon
    
    
    '   .EditEnterAction = 0
    '   .SetActiveCell j + 1, .Row
    '   .EditEnterAction = 2
    End Select
    Main(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    'If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
    '.Row = Row: .Col = Col
    'If .BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    'If .ChangeMade = False Or Col = 1 Or Mode = 1 Then i = .Text: Exit Sub
    'If .ChangeMade = True Then
    '   .Col = (maxcolumna * 5 + 1) + (.Col / 5): .Text = 1
    'End If
    '.Row = Row
    'Main(0).Enabled = True
    'Toolbar1.Buttons(1).Visible = False
    'Toolbar1.Buttons(2).Visible = True
End With
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
If Toolbar1.Buttons(2).Enabled = False Or (vg_codregimen > 9999 And etapa5 And AddReceta = 0) Then Exit Sub
Dim DelRow As Integer, indcol As Integer, indrow As Integer, IndCol2 As Integer, IndRow2 As Integer
With vaSpread1
    Select Case KeyCode
    Case 65 To 90
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        ws_respuesta = ""
        ws_respuesta = Chr(KeyCode)
        Plato_Click (0)
    Case 86
        Exit Sub
    Case 46
    '    If .MaxRows = .ActiveRow Then Exit Sub
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
        j = 0
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = .Col Or VectorCol(i) = .Col Then j = (VectorCol(i) - 1): Exit For
        Next i
        If j = 0 Then Exit Sub
        Plato(0).Enabled = True
        OpGrilla(0).Enabled = True
        Plato(70).Enabled = False
        OpGrilla(70).Enabled = False
        If IndActivo = 0 Then IblockCol = .ActiveCol: iblockcol2 = .ActiveCol: IblockRow = .ActiveRow: IblockRow2 = .ActiveRow
        AiBlockCol = IblockCol
        AiBlockRow = IblockRow
        AiBlockCol2 = iblockcol2
        AiBlockRow2 = IblockRow2
        If IblockCol < 0 Then IblockCol = 1: iblockcol2 = .MaxCols
        AiBlockCol = IblockCol
        AiBlockRow = IblockRow
        AiBlockCol2 = iblockcol2
        AiBlockRow2 = IblockRow2
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 3)): Exit For
            If VectorCol(i) = iblockcol2 Then iblockcol2 = (VectorCol(i) + 3): Exit For
        Next i
        indcol = AiBlockCol: IndCol2 = iblockcol2
        indrow = AiBlockRow: IndRow2 = AiBlockRow2
        .ClearRange IblockCol, IblockRow, iblockcol2, IblockRow2, False
        IblockCol = AuxCol
        .BlockMode = False
        Main(0).Enabled = True
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        IndActivo = 0
    End Select
End With
End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case Button
Case 2
    If vaSpread1.Visible <> True Or Main(10).Enabled = False Then Exit Sub
    Indvaspread1 = 0
    PopupMenu MenuDetalle
End Select
End Sub

Private Sub Opgrilla_Click(Index As Integer)
Select Case Index
Case 0 'Insertar producto
    Plato_Click (0)
Case 20 'Insertar línea
    Plato_Click (20)
Case 30 'Eliminar línea
    Plato_Click (30)
Case 50 'Cortar
    Plato_Click (50)
Case 60 'Copiar
    Plato_Click (60)
Case 70 'Pegar
    Plato_Click (70)
Case 80 'Pegado especial
    Plato_Click (80)
End Select
End Sub

Private Sub GrabarPlantilla()
Dim codpro As String, canpro As Double, Fecha As Long, conregdet As Long, ExisteDat As Long, inddia As Long, cospro As Double
On Error GoTo Man_Error
inddia = 1: conregdet = 0: gauge1.Value = 0: gauge.Value = 0: Fecha = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh
fg_carga ""
'Grabar datos
With vaSpread1
    vg_db.BeginTrans
    For i = 1 To (.MaxCols) Step 5
        gauge1.Value = Val((inddia / MaxColumna) * 100)
        Label3.Caption = "": Label3.Caption = "Día : " & inddia
        ExisteDat = 0: .Row = 1: .Col = i
        Fecha = Val(vg_fecha) & fg_pone_cero(inddia, 2)
        For j = 1 To (.MaxRows - 1)
            .Row = j
            .Col = i + 1
            If Trim(.text) <> "" Then ExisteDat = 1: Exit For
        Next j
        .Row = .MaxRows: .Col = i + 2: totrac = Val(.text)
        RS.Open "SELECT DISTINCT mfd_cencos FROM  b_minutafijadia " & _
                "WHERE  mfd_cencos='" & vg_codcasino & "' AND mfd_codreg=" & vg_codregimen & " " & _
                "AND    mfd_codser=" & vg_codservicio & " AND mfd_fecha=" & Val(Fecha) & " " & _
                "AND    mfd_tipmin='" & tipmin & "'", vg_db, adOpenStatic
        If Not RS.EOF Then
           If ExisteDat = 0 Then
              vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE  mfd_cencos='" & vg_codcasino & "' AND mfd_codreg=" & vg_codregimen & " " & _
                            "AND    mfd_codser=" & vg_codservicio & " AND mfd_fecha=" & Val(Fecha) & " " & _
                            "AND    mfd_tipmin='" & tipmin & "'"
           End If
        End If
        RS.Close: Set RS = Nothing
        gauge.Value = 0: conregdet = 0: estser = 0
        If ExisteDat > 0 Then
           vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE  mfd_cencos='" & vg_codcasino & "' AND mfd_codreg=" & vg_codregimen & " " & _
                         "AND    mfd_codser=" & vg_codservicio & " AND mfd_fecha=" & Val(Fecha) & " " & _
                         "AND    mfd_tipmin='" & tipmin & "'"
           'Actualizar detalle
           For j = 1 To .MaxRows
               conregdet = conregdet + 1
               gauge.Value = Val((conregdet / (.MaxRows)) * 100)
               codpro = "": canpro = 0: cospro = 0
               .Row = j
               .Col = i: codpro = .text
               If Trim(codpro) <> "" Then
                  .Col = i + 2
                  If Val(Trim(.text)) > 0 Then: canpro = .text
                  .Col = i + 4
                  If Val(Trim(.text)) > 0 Then cospro = .text
                  vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) " & _
                                "VALUES ('" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & Fecha & ", '" & codpro & "', '" & tipmin & "', " & canpro & ", " & cospro & ")"
               End If
           Next j
        End If
        inddia = inddia + 1
    Next i
    vg_db.CommitTrans
    Picture1.Visible = False: gauge.Visible = False
    .Refresh
End With
fg_descarga

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub DetallePlantilla()
fg_carga ""
Dim iRow As Long, i As Long, j As Long, inddia As Long, Fecha As String
SwSalir = 0: MaxColumna = 0: IndActivo = 0: vCtoPis = 0: vCtoTec = 0
IblockRow = 0: IblockRow2 = 0: IblockCol = 0: iblockcol2 = 0: SwSalir = 0
AiBlockRow = 0: AiBlockRow2 = 0: AiBlockCol = 0: AiBlockCol2 = 0
With vaSpread1
    '------- formatear columna
    MaxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
    'Defenir vector costo encabezado
    .MaxRows = 1000
    .MaxCols = 0: .MaxCols = 5 * MaxColumna: .Row = 0
    'turn off display of row headers
    '.RowHeadersShow = False
    'Set up column headers
    .ColHeaderRows = 2
    .ShadowColor = &H8000000F
    .ShadowText = &H800000
    For i = 1 To .MaxCols Step 5
        .AddCellSpan i, SpreadHeader, 5, 1
        .Col = i
        .Row = SpreadHeader
        .TypeHAlign = TypeHAlignCenter
        .text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & CLng((i / 5) + 1), 2), 1), 1, 3) & " " & CLng((i / 5) + 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
    Next i
    'Ampliar ancho de la columna
    .RowHeight(SpreadHeader + 1) = 15.5
    ReDim Preserve VectorCol(0)
    For i = 1 To .MaxCols Step 5
        .Row = SpreadHeader + 1
        .Col = i
        .ColWidth(i) = 5.5
        .text = "Codigo"
        .ColHidden = False
    
        If i = 2 Then
           ReDim Preserve VectorCol(1)
           VectorCol(1) = 2
        Else
           ReDim Preserve VectorCol(CLng((i / 5) + 1))
           VectorCol(CLng((i / 5) + 1)) = i + 1
        End If
        
        .Col = i + 1
        .ColWidth(i + 1) = 21
        .text = "Descripción Producto"
        .ColHidden = False
    
        .Col = i + 2
        .ColWidth(i + 2) = 7
        .text = "Cantidad"
        .ColHidden = False
    
        .Col = i + 3
        .ColWidth(i + 3) = 5
        .text = "U.M."
        .ColHidden = False
    
        .Col = i + 4
        .text = "Cod. Receta"
        .ColHidden = True
        For j = 1 To .MaxRows
            .Row = j
            .Col = i
            .CellType = CellTypeEdit '= CellTypeStaticText
            .TypeHAlign = TypeHAlignRight
            .text = ""
    
            .Col = i + 1
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 2
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 3
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 4
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
    '        .Col = i + 5
    '        .CellType = CellTypeDate '= 1
    '        .TypeHAlign = TypeHAlignLeft
    '        .Text = " "
        Next j
    Next i
    .Row = -1: .Col = -1: .BackColor = Shape1(0).FillColor  'Amarillo
    'RS.Open "SELECT DISTINCT b.mid_fecval FROM b_minuta a, b_minutadet b WHERE a.min_codigo=b.mid_codigo AND a.min_cencos='" & vg_codcasino & "' AND val(mid(a.min_fecmin,1,6))=" & Val(vg_fecha) & " AND b.mid_fecval>0", vg_db, adOpenStatic
    'If RS.EOF Then .BackColor = Shape1(0).FillColor: indblo = True
    'RS.Close: Set RS = Nothing
    j = 0: i = 0: iRow = 0
    '-------> Validar si existe estructura fija si no existe crear
    If vg_tipbase = "1" Then
       RS.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
               "WHERE mfd_cencos='" & vg_codcasino & "' " & _
               "AND   mfd_codreg=" & vg_codregimen & " " & _
               "AND   mfd_codser=" & vg_codservicio & " " & _
               "AND mid(mfd_fecha,1,6)=" & Val(vg_fecha) & " " & _
               "AND   mfd_tipmin='" & tipmin & "'", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
               "WHERE mfd_cencos = '" & vg_codcasino & "' " & _
               "AND   mfd_codreg = " & vg_codregimen & " " & _
               "AND   mfd_codser = " & vg_codservicio & " " & _
               "AND   convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) = " & Val(vg_fecha) & " " & _
               "AND   mfd_tipmin = '" & tipmin & "'", vg_db, adOpenStatic
    End If
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       '------- Buscar fecha mayor de estructura fija
       Dim fecval As Long
       fecval = 0
       RS.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & vg_codcasino & "' AND mif_codreg = " & vg_codregimen & " AND mif_codser = " & vg_codservicio & "", vg_db, adOpenStatic
       If Not RS.EOF Then fecval = IIf(IsNull(RS!fecval), 0, RS!fecval)
       RS.Close: Set RS = Nothing
       If fecval > 0 Then
          Dim aAp As String
          If vg_tipbase = "1" Then
             '-------> Insert tabla productospmpdia
             aAp = Trim(vg_NUsr) & "_tmp_ProductoEstFDiPMP"
             fg_CheckTmp aAp
             vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                           "INTO " & aAp & " " & _
                           "FROM b_productospmpdia " & _
                           "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                           "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                           "AND   ppd_propon>0 " & _
                           "GROUP BY ppd_cencos, ppd_codpro"
             vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
             vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
             vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
          End If
          '------- Traer estructura fija
          For i = 1 To Val(Mid(dEoM("26/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)), 1, 2))
              If fecval <= vg_fecha & fg_pone_cero(Str(i), 2) Then
                 '------- Grabar estructura fija día teorica
                 If vg_tipbase = "1" Then
                    vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & vg_fecha & fg_pone_cero(i, 2) & ", b.pro_codigo, '" & tipmin & "', a.mif_canpro, c.ppd_propon " & _
                                  "FROM b_minutafija a, b_productos b, " & aAp & " c " & _
                                  "WHERE a.mif_codpro = b.pro_codigo AND b.pro_codigo = c.ppd_codpro AND c.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                                  "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                  "AND   a.mif_cencos = '" & vg_codcasino & "' " & _
                                  "AND   a.mif_codreg = " & vg_codregimen & " " & _
                                  "AND   a.mif_codser = " & vg_codservicio & " " & _
                                  "AND   a.mif_fecval = " & fecval & " " & _
                                  "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                  "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                 Else
                    vg_db.Execute "INSERT INTO b_minutafijadia (mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro) SELECT a.mif_cencos, a.mif_codreg, a.mif_codser, " & vg_fecha & fg_pone_cero(i, 2) & ", b.pro_codigo, '" & tipmin & "', a.mif_canpro, c.ppd_propon " & _
                                  "FROM b_minutafija a, b_productos b, b_productospmpdia c " & _
                                  "WHERE a.mif_codpro = b.pro_codigo AND b.pro_codigo = c.ppd_codpro AND c.ppd_cencos = '" & MuestraCasino(1) & "' AND c.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
                                  "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0) " & _
                                  "AND   a.mif_cencos = '" & vg_codcasino & "' " & _
                                  "AND   a.mif_codreg = " & vg_codregimen & " " & _
                                  "AND   a.mif_codser = " & vg_codservicio & " " & _
                                  "AND   a.mif_fecval = " & fecval & " " & _
                                  "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & i, 2), 2)) - 2))) & " " & _
                                  "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven <= 0)"
                 End If
              End If
              Set RS = Nothing
          Next i
          '-------> Borrar tablas temporales
          If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
       End If
    Else
       RS.Close: Set RS = Nothing
    End If
    If vg_tipbase = "1" Then
        RS.Open "SELECT a.mfd_fecha, a.mfd_codpro, b.pro_nombre, a.mfd_canpro, a.mfd_cospro, c.uni_nomcor " & _
                "FROM b_minutafijadia a, b_productos b, a_unidad c " & _
                "WHERE a.mfd_codpro=b.pro_codigo " & _
                "AND   b.pro_coduni=c.uni_codigo " & _
                "AND   a.mfd_cencos='" & vg_codcasino & "' " & _
                "AND   a.mfd_codreg=" & vg_codregimen & " " & _
                "AND   a.mfd_codser=" & vg_codservicio & " " & _
                "AND mid(a.mfd_fecha,1,6)=" & Val(vg_fecha) & " " & _
                "AND   a.mfd_tipmin='" & tipmin & "' ORDER BY a.mfd_fecha, b.pro_nombre", vg_db, adOpenStatic
    Else
        RS.Open "SELECT a.mfd_fecha, a.mfd_codpro, b.pro_nombre, a.mfd_canpro, a.mfd_cospro, c.uni_nomcor " & _
                "FROM b_minutafijadia a, b_productos b, a_unidad c " & _
                "WHERE a.mfd_codpro = b.pro_codigo " & _
                "AND   b.pro_coduni = c.uni_codigo " & _
                "AND   a.mfd_cencos = '" & vg_codcasino & "' " & _
                "AND   a.mfd_codreg = " & vg_codregimen & " " & _
                "AND   a.mfd_codser = " & vg_codservicio & " " & _
                "AND   convert(int,substring(convert(varchar(8),a.mfd_fecha),1,6)) = " & Val(vg_fecha) & " " & _
                "AND   a.mfd_tipmin = '" & tipmin & "' ORDER BY a.mfd_fecha, b.pro_nombre", vg_db, adOpenStatic
    End If
    If Not RS.EOF Then
       Do While Not RS.EOF
          If auxfec <> RS!mfd_fecha Then i = 1: auxfec = RS!mfd_fecha
          j = (((Val(Mid(RS!mfd_fecha, 7, 2)) * 5) - 5) + 1) ' + 1
          .Row = i 'RS!mid_numlin
          If iRow < .Row Then iRow = .Row
          .Col = j
          .CellType = CellTypeEdit '= CellTypeStaticText
          .TypeHAlign = TypeHAlignRight
          .text = IIf(IsNull(RS!mfd_codpro), "", Trim(RS!mfd_codpro))
               
          .Col = j + 1
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!pro_nombre), "", Trim(RS!pro_nombre))
                             
          .Col = j + 2
          .CellType = CellTypeNumber
          .TypeNumberDecPlaces = vg_RDCa
          .TypeIntegerMin = 1
          .TypeIntegerMax = 9999999
          .TypeHAlign = TypeHAlignRight
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
          .Value = IIf(IsNull(RS!mfd_canpro), 0, RS!mfd_canpro)
          .ForeColor = &HFF0000
                           
          .Col = j + 3
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!uni_nomcor), "", Trim(RS!uni_nomcor))
          
          .Col = j + 4
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!mfd_cospro), 0, RS!mfd_cospro)
          RS.MoveNext: i = i + 1
       Loop
    End If
    RS.Close: Set RS = Nothing
    .MaxRows = IIf(iRow = 0, 0, iRow + 1)
    If iRow = 0 Then .Row = -1: .Col = -1: Main(10).Enabled = False
    .Row = .MaxRows
    MaxFila = .MaxRows
    If tipmin = "1" Then
       .Row = -1: .Col = -1
       If vg_tipbase = "1" Then
          RS.Open "SELECT DISTINCT b.mid_fecval FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & vg_codcasino & "' AND val(mid(a.min_fecmin,1,6)) = " & Val(vg_fecha) & " AND b.mid_fecval > 0", vg_db, adOpenStatic
       Else
          RS.Open "SELECT DISTINCT b.mid_fecval FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & vg_codcasino & "' AND convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) = " & Val(vg_fecha) & " AND b.mid_fecval > 0", vg_db, adOpenStatic
       End If
       If Not RS.EOF Then .Lock = True: .BackColor = Shape1(1).FillColor
       RS.Close: Set RS = Nothing
    Else
       For i = 1 To (.MaxCols - MaxColumna) Step 5
           .Row = SpreadHeader: .Col = i
           If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Or CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < CDate(vg_ciedia) Then
              Dim fil As Long, Col As Long
              For fil = 1 To (.MaxRows)
                  For Col = i To i + 4
                      .Row = fil: .Col = Col
                      If .CellType = CellTypeNumber Or .CellType = CellTypeEdit Then .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight
                      .BackColor = Shape1(1).FillColor
                  Next Col
              Next fil
           End If
       Next i
    End If
    fg_descarga
    .Row = 1: .Col = 1
    IblockRow = .Row: AiBlockRow = .Row
    IblockRow2 = .Row: AiBlockRow2 = .Row
    IblockCol = .Col: AiBlockCol = .Col
    iblockcol2 = .Col: AiBlockCol2 = .Col
    .SetActiveCell 1, 1
End With
End Sub

Function ValidarProducto(codpro As String, Row As Long, Col As Long) As Boolean
ValidarProducto = False
'-------> Validar si existe codigo producto
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Col = Col: vaSpread1.Row = i
    If Trim(vaSpread1.text) = Trim(codpro) And Trim(vaSpread1.text) <> "" Then ValidarProducto = True
Next i
End Function

Function ValidarProducto1(codpro As String, Row As Long, Col As Long) As Boolean
ValidarProducto1 = False
'-------> Validar si existe codigo producto
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Col = Col: vaSpread1.Row = i
    If Trim(vaSpread1.text) = Trim(codpro) And Row <> i And Trim(vaSpread1.text) <> "" Then ValidarProducto1 = True
Next i
End Function
