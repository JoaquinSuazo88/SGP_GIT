VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form C_MReCas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minuta Real Casino"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   30
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      DragIcon        =   "C_MReCas.frx":0000
      Height          =   3600
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   11055
      _Version        =   393216
      _ExtentX        =   19500
      _ExtentY        =   6350
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
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
      MaxCols         =   250
      MaxRows         =   100
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RestrictRows    =   -1  'True
      RowsFrozen      =   1
      SpreadDesigner  =   "C_MReCas.frx":0442
      UserResize      =   1
      VisibleCols     =   1
      VisibleRows     =   100
      TextTip         =   2
      TextTipDelay    =   0
      ScrollBarTrack  =   3
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00DEFEDE&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   0
      Top             =   0
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H80000018&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   4020
      Top             =   0
      Width           =   300
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planificación Minutas Teórica"
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
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   10905
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planificación Minutas Teórica"
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
      Left            =   0
      TabIndex        =   3
      Top             =   -600
      Width           =   10905
   End
End
Attribute VB_Name = "C_MReCas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6765
Me.Width = 11055
fg_centra Me
Msgtitulo = "Miuta Real Casino"
fg_carga ""
Label1.Caption = Trim(C_IMiRCa.fpayuda(0).Caption) & "(" & Trim(C_IMiRCa.fpText.text) & ")" & " - " & Trim(C_IMiRCa.fpayuda(1).Caption) & " - " & Trim(C_IMiRCa.fpayuda(2).Caption) & " - " & IIf(vg_tmisgp = "1", "Planificación Teórica", "Planificación Real")
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): btnX.Visible = True: btnX.ToolTipText = "Exporta Minuta Real Casino a Excel "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
DetallePlantillaMinuta
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then Label1.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
If Me.WindowState <> 1 Then vaSpread1.Move 0, 840, ScaleWidth, ScaleHeight - 840
End Sub

Sub DetallePlantillaMinuta()
fg_carga ""
Dim RS As New ADODB.Recordset
Dim indrow3 As Long, inddia As Long, fecha As String
Dim maxColumna As Long, i As Long, ii As Long, j As Long, j1 As Long
Dim vectorrac() As String, vectorpon() As String
maxColumna = 0
'-------> Formatear columna
maxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
vaSpread1.MaxRows = 1000
vaSpread1.MaxCols = 0: vaSpread1.MaxCols = (4 * maxColumna) + 1: vaSpread1.Row = 1
vaSpread1.Col = 1
vaSpread1.ColsFrozen = 1
vaSpread1.VisibleCols = 1
vaSpread1.ColWidth(1) = 15
vaSpread1.text = "Estructura Servicio"
inddia = 1
'-------> Mover vector raciones
ReDim Preserve vectorrac(1): vectorrac(1) = "D"
ReDim Preserve vectorrac(2): vectorrac(2) = "H"
ReDim Preserve vectorrac(3): vectorrac(3) = "L"
ReDim Preserve vectorrac(4): vectorrac(4) = "P"
ReDim Preserve vectorrac(5): vectorrac(5) = "T"
ReDim Preserve vectorrac(6): vectorrac(6) = "X"
ReDim Preserve vectorrac(7): vectorrac(7) = "AB"
ReDim Preserve vectorrac(8): vectorrac(8) = "AF"
ReDim Preserve vectorrac(9): vectorrac(9) = "AJ"
ReDim Preserve vectorrac(10): vectorrac(10) = "AN"
ReDim Preserve vectorrac(11): vectorrac(11) = "AR"
ReDim Preserve vectorrac(12): vectorrac(12) = "AV"
ReDim Preserve vectorrac(13): vectorrac(13) = "AZ"
ReDim Preserve vectorrac(14): vectorrac(14) = "BD"
ReDim Preserve vectorrac(15): vectorrac(15) = "BH"
ReDim Preserve vectorrac(16): vectorrac(16) = "BL"
ReDim Preserve vectorrac(17): vectorrac(17) = "BP"
ReDim Preserve vectorrac(18): vectorrac(18) = "BT"
ReDim Preserve vectorrac(19): vectorrac(19) = "BX"
ReDim Preserve vectorrac(20): vectorrac(20) = "CB"
ReDim Preserve vectorrac(21): vectorrac(21) = "CF"
ReDim Preserve vectorrac(22): vectorrac(22) = "CJ"
ReDim Preserve vectorrac(23): vectorrac(23) = "CN"
ReDim Preserve vectorrac(24): vectorrac(24) = "CR"
ReDim Preserve vectorrac(25): vectorrac(25) = "CV"
ReDim Preserve vectorrac(26): vectorrac(26) = "CZ"
ReDim Preserve vectorrac(27): vectorrac(27) = "DD"
ReDim Preserve vectorrac(28): vectorrac(28) = "DH"
ReDim Preserve vectorrac(29): vectorrac(29) = "DL"
ReDim Preserve vectorrac(30): vectorrac(30) = "DP"
ReDim Preserve vectorrac(31): vectorrac(31) = "DT"
'-------> Mover vector ponderaciones
ReDim Preserve vectorpon(1): vectorpon(1) = "E"
ReDim Preserve vectorpon(2): vectorpon(2) = "I"
ReDim Preserve vectorpon(3): vectorpon(3) = "M"
ReDim Preserve vectorpon(4): vectorpon(4) = "Q"
ReDim Preserve vectorpon(5): vectorpon(5) = "U"
ReDim Preserve vectorpon(6): vectorpon(6) = "Y"
ReDim Preserve vectorpon(7): vectorpon(7) = "AC"
ReDim Preserve vectorpon(8): vectorpon(8) = "AG"
ReDim Preserve vectorpon(9): vectorpon(9) = "AK"
ReDim Preserve vectorpon(10): vectorpon(10) = "AO"
ReDim Preserve vectorpon(11): vectorpon(11) = "AS"
ReDim Preserve vectorpon(12): vectorpon(12) = "AW"
ReDim Preserve vectorpon(13): vectorpon(13) = "BA"
ReDim Preserve vectorpon(14): vectorpon(14) = "BE"
ReDim Preserve vectorpon(15): vectorpon(15) = "BI"
ReDim Preserve vectorpon(16): vectorpon(16) = "BM"
ReDim Preserve vectorpon(17): vectorpon(17) = "BQ"
ReDim Preserve vectorpon(18): vectorpon(18) = "BU"
ReDim Preserve vectorpon(19): vectorpon(19) = "BY"
ReDim Preserve vectorpon(20): vectorpon(20) = "CC"
ReDim Preserve vectorpon(21): vectorpon(21) = "CG"
ReDim Preserve vectorpon(22): vectorpon(22) = "CK"
ReDim Preserve vectorpon(23): vectorpon(23) = "CO"
ReDim Preserve vectorpon(24): vectorpon(24) = "CS"
ReDim Preserve vectorpon(25): vectorpon(25) = "CW"
ReDim Preserve vectorpon(26): vectorpon(26) = "DA"
ReDim Preserve vectorpon(27): vectorpon(27) = "DE"
ReDim Preserve vectorpon(28): vectorpon(28) = "DI"
ReDim Preserve vectorpon(29): vectorpon(29) = "DM"
ReDim Preserve vectorpon(30): vectorpon(30) = "DQ"
ReDim Preserve vectorpon(31): vectorpon(31) = "DU"

For i = 2 To vaSpread1.MaxCols Step 4
    vaSpread1.Col = i
    vaSpread1.ColWidth(i) = 2.5
    vaSpread1.text = "  "
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 1
    vaSpread1.ColWidth(i + 1) = 21
    vaSpread1.text = " " & fg_Fecha_Dia(Mid(vg_fecha, 1, 4) & Mid(vg_fecha, 5, 2) & fg_pone_cero(inddia, 2), 2) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
    vaSpread1.ColHidden = False
    
   
    vaSpread1.Col = i + 2
    vaSpread1.ColWidth(i + 2) = 6
    
    vaSpread1.text = "N.Rac."
    vaSpread1.ColHidden = False
   
    vaSpread1.Col = i + 3
    vaSpread1.ColWidth(i + 3) = 6
    vaSpread1.text = "Pond "
    vaSpread1.ColHidden = False
    
'    For j = 1 To vaSpread1.MaxRows
'        vaSpread1.Row = j
'
'        vaSpread1.Col = i
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignLeft
'        vaSpread1.text = ""
'
'        vaSpread1.Col = i + 1
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignLeft
'        vaSpread1.text = ""
'
'        vaSpread1.Col = i + 2
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignLeft
'        vaSpread1.text = ""
'
'        vaSpread1.Col = i + 3
'        vaSpread1.CellType = CellTypeStaticText
'        vaSpread1.TypeHAlign = TypeHAlignLeft
'        vaSpread1.text = ""
'
'     Next j
    vaSpread1.Row = 1
    inddia = inddia + 1
Next i

vaSpread1.Row = 1
vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
vaSpread1.Row = 1: vaSpread1.Col = -1: vaSpread1.BackColor = &HE0E0E0
vaSpread1.Row = -1: vaSpread1.Col = 1
vaSpread1.Font.Bold = True
vaSpread1.Font.Size = 9
vaSpread1.BackColor = Shape1(2).FillColor 'Verde
j = 0: i = 0: indrow3 = 0
Set RS = vg_db.Execute("sgpadm_s_minutarealcasino 1, " & vg_cencos & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Val(vg_fecha) & ", 0, 0, '" & vg_tmisgp & "'")
DoEvents
If Not RS.EOF Then
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 4) - 4) + 1) + 1
      vaSpread1.Row = RS!mid_numlin + 1
      If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
      If RS!ess_codigo <> i Then
         vaSpread1.Col = 1
         vaSpread1.text = RS!ess_codigo & " - " & Trim(RS!ess_nombre)
         i = RS!ess_codigo
      End If

      vaSpread1.Col = j
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Value = IIf(IsNull(RS!mid_rec5eta) Or RS!mid_rec5eta = "0", "RA", "R")
      vaSpread1.ForeColor = &HFF&
      vaSpread1.BackColor = &H80FF80

      vaSpread1.Col = j + 1
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS!rec_nombre)

      vaSpread1.Col = j + 2
      vaSpread1.CellType = CellTypeCurrency
      vaSpread1.TypeCurrencyDecPlaces = 0
      vaSpread1.TypeCurrencyShowSymbol = False
      vaSpread1.Lock = True
      vaSpread1.Value = IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac)
      vaSpread1.ForeColor = &HFF0000

      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing: fg_descarga
End If
vaSpread1.MaxRows = indrow3 + 1
vaSpread1.Row = vaSpread1.MaxRows
vaSpread1.Col = 1
vaSpread1.text = "Comensales"
vaSpread1.Col = -1: vaSpread1.BackColor = &HE0E0E0
For i = 2 To (vaSpread1.MaxCols - maxColumna) Step 4
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = i + 2
    vaSpread1.CellType = CellTypeCurrency
    vaSpread1.TypeCurrencyDecPlaces = 0
    vaSpread1.TypeCurrencyShowSymbol = False
    vaSpread1.Lock = True
    vaSpread1.Value = Format(0, fg_Pict(6, 0))
    vaSpread1.ForeColor = &HFF0000
Next i
'-------> Mover raciones totales
Set RS = vg_db.Execute("sgpadm_s_minutarealcasino 2, " & vg_cencos & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Val(vg_fecha) & ", 0, 0, '" & vg_tmisgp & "'")
DoEvents
If Not RS.EOF Then
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 4) - 4) + 1) + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = j + 2
      vaSpread1.CellType = CellTypeNumber
      vaSpread1.TypeNumberDecPlaces = 0
      vaSpread1.TypeIntegerMin = 1
      vaSpread1.TypeIntegerMax = 9999999
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.TypeSpin = False
      vaSpread1.TypeIntegerSpinInc = 1
      vaSpread1.TypeIntegerSpinWrap = False
      vaSpread1.Value = IIf(IsNull(RS!min_racrea), 0, RS!min_racrea)
      vaSpread1.ForeColor = &HFF0000
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
End If
'-------> Calcular porcentaje Vertical
Dim letra As String
Dim estcer As Boolean
estcer = False
j1 = 1
For i = 2 To (vaSpread1.MaxCols) Step 4
   For ii = 2 To vaSpread1.MaxRows - 1
    vaSpread1.Row = ii
    vaSpread1.Col = i + 2
    estcer = False
    If Trim(vaSpread1.text) <> "" Then
       letra = vectorrac(j1)
       vaSpread1.Row = vaSpread1.MaxRows
       If vaSpread1.text = "0" Then estcer = True
       vaSpread1.Row = ii
       vaSpread1.Col = i + 3
'       vaSpread1.CellType = CellTypeCurrency
'       vaSpread1.TypeCurrencyDecPlaces = 0
'       vaSpread1.TypeCurrencyShowSymbol = False
       vaSpread1.CellType = CellTypePercent
       vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
       vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
       vaSpread1.TypePercentDecPlaces = 0
       vaSpread1.TypePercentMax = 100
       vaSpread1.Lock = True
'       vaSpread1.Formula = Fg_Sacacremilla("SUM('" & letra & (ii) & "'/'" & IIf(Not estcer, letra & (vaSpread1.MaxRows), 1) & "')*100")
       vaSpread1.Formula = Fg_Sacacremilla("SUM('" & letra & (ii) & "'/'" & IIf(Not estcer, letra & (vaSpread1.MaxRows), 1) & "')")
    End If
   Next ii
   j1 = j1 + 1
Next i

'-------> Calcular comensales Totales
vaSpread1.MaxCols = vaSpread1.MaxCols + 1
estcer = False
letra = ""
j1 = 1
For i = 2 To (vaSpread1.MaxCols - 1) Step 4
    vaSpread1.Row = 1
    vaSpread1.Col = i + 1
    If Mid(Trim(vaSpread1.text), 1, 3) <> "Sáb" And Mid(Trim(vaSpread1.text), 1, 3) <> "Dom" Then
       vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = i + 2
       estcer = False
       If Trim(vaSpread1.text) <> "" Then
          letra = letra & "'" & vectorrac(j1) & vaSpread1.Row & "'+"
       End If
    End If
    j1 = j1 + 1
Next i
If letra <> "" Then
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = vaSpread1.MaxCols
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
   vaSpread1.Formula = Fg_Sacacremilla("(" & Mid(letra, 1, Len(letra) - 1) & ")")
End If

'-------> Calcular porcentaje Horizontal de las raciones
Dim numdia As Double
Dim codest As String, datfinal As String
Dim irow As Long
datfinal = IIf(maxColumna = 28, "DJ", IIf(maxColumna = 29, "DN", IIf(maxColumna = 30, "DR", "DV"))) & vaSpread1.MaxRows
For i = 2 To vaSpread1.MaxRows - 1
    vaSpread1.Row = i
    j1 = 1
    vaSpread1.Col = 1
    If Trim(vaSpread1.text) <> codest And Trim(vaSpread1.text) <> "" Then
       codest = Trim(vaSpread1.text)
       irow = i
       letra = ""
       numdia = 0
    End If
    For ii = 2 To (vaSpread1.MaxCols - 1) Step 4
        vaSpread1.Col = ii + 1
        vaSpread1.Row = 1
        If Mid(Trim(vaSpread1.text), 1, 3) <> "Sáb" And Mid(Trim(vaSpread1.text), 1, 3) <> "Dom" Then
           vaSpread1.Row = i
           If vaSpread1.text <> "" Then
'              letra = letra & "'" & vectorpon(j1) & vaSpread1.Row & "'+"
              letra = letra & "'" & vectorrac(j1) & vaSpread1.Row & "'+"
              numdia = numdia + 1
           End If
        End If
        j1 = j1 + 1
    Next ii
    vaSpread1.Row = i + 1
    vaSpread1.Col = 1
    If Trim(vaSpread1.text) <> codest And Trim(vaSpread1.text) <> "" Then
    If letra <> "" Then
       vaSpread1.Row = irow
       vaSpread1.Col = vaSpread1.MaxCols
       vaSpread1.Font.Bold = True
       vaSpread1.Font.Size = 9
       
'       vaSpread1.CellType = CellTypePercent '
'       vaSpread1.TypeCurrencySymbol = "%"
'       vaSpread1.TypeCurrencyDecPlaces = 0
'       vaSpread1.TypeCurrencyShowSymbol = True
       vaSpread1.CellType = CellTypePercent
       vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
       vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
       vaSpread1.TypePercentDecPlaces = 0
       vaSpread1.TypePercentMax = 100
       vaSpread1.Lock = True
''      vaSpread1.Formula = Fg_Sacacremilla("SUM(" & Mid(letra, 1, Len(letra) - 1) & ")/" & numdia & "")
'       vaSpread1.Formula = Fg_Sacacremilla("(" & Mid(letra, 1, Len(letra) - 1) & ")" & "/" & "SUM('" & datfinal & "')*100")  '/" & 1 & "")
       vaSpread1.Formula = Fg_Sacacremilla("(" & Mid(letra, 1, Len(letra) - 1) & ")" & "/" & "SUM('" & datfinal & "')")  '/" & 1 & "")
    End If
    End If
Next i

'-------> formatear ultima columna
Dim conrec As Double, con5et As Double, concer As Double
Dim rac As Double
vaSpread1.MaxRows = vaSpread1.MaxRows + 5
'-------> Calcular Total de recetas
For i = 2 To (vaSpread1.MaxCols - 1) Step 4
   con5et = 0: conrec = 0: concer = 0
   For ii = 2 To vaSpread1.MaxRows - 6
       vaSpread1.Row = ii
       vaSpread1.Col = i ' + 2
       If Trim(vaSpread1.text) <> "" Then
          vaSpread1.Col = i + 2
          rac = vaSpread1.text
          vaSpread1.Col = i
          If Trim(vaSpread1.text) = "R" Then
             conrec = conrec + 1
             If rac = 0 Then concer = concer + 1
          ElseIf Trim(vaSpread1.text) = "RA" Then
             con5et = con5et + 1
             If rac = 0 Then concer = concer + 1
          End If
       End If
   Next ii
   vaSpread1.Row = vaSpread1.MaxRows - 3
   vaSpread1.Col = i + 1
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = "Numero recetas (R)"
   vaSpread1.Col = i + 2
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
   vaSpread1.text = conrec

   vaSpread1.Row = vaSpread1.MaxRows - 2
   vaSpread1.Col = i + 1
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = "Numero recetas (RA)"
   vaSpread1.Col = i + 2
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
   vaSpread1.text = con5et

   vaSpread1.Row = vaSpread1.MaxRows - 1
   vaSpread1.Col = i + 1
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = "N recetas ponderadas en cero"
   vaSpread1.Col = i + 2
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
   vaSpread1.text = concer
Next i
'-------> Mover titulo Total
vaSpread1.Row = vaSpread1.MaxRows - 4
vaSpread1.Col = vaSpread1.MaxCols
vaSpread1.Font.Bold = True
vaSpread1.Font.Size = 9
vaSpread1.CellType = CellTypeStaticText
vaSpread1.TypeHAlign = TypeHAlignCenter
vaSpread1.text = "Total"

'-------> Sumar numero recetas
vaSpread1.Row = vaSpread1.MaxRows - 3
letra = ""
numdia = 0
j1 = 1
For ii = 2 To (vaSpread1.MaxCols - 1) Step 4  '- maxColumna
    vaSpread1.Col = ii + 1
    If vaSpread1.text <> "" Then
       letra = letra & "'" & vectorrac(j1) & vaSpread1.Row & "'+"
    End If
    j1 = j1 + 1
Next ii
If letra <> "" Then
   vaSpread1.Row = vaSpread1.MaxRows - 3
   vaSpread1.Col = vaSpread1.MaxCols
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
'  vaSpread1.Formula = Fg_Sacacremilla("SUM(" & Mid(letra, 1, Len(letra) - 1) & ")/" & numdia & "")
   vaSpread1.Formula = Fg_Sacacremilla("(" & Mid(letra, 1, Len(letra) - 1) & ")")
End If

'-------> Sumar numero recetas casino
vaSpread1.Row = vaSpread1.MaxRows - 2
letra = ""
numdia = 0
j1 = 1
For ii = 2 To (vaSpread1.MaxCols - 1) Step 4  '- maxColumna
    vaSpread1.Col = ii + 1
    If vaSpread1.text <> "" Then
       letra = letra & "'" & vectorrac(j1) & vaSpread1.Row & "'+"
    End If
    j1 = j1 + 1
Next ii
If letra <> "" Then
   vaSpread1.Row = vaSpread1.MaxRows - 2
   vaSpread1.Col = vaSpread1.MaxCols
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
'  vaSpread1.Formula = Fg_Sacacremilla("SUM(" & Mid(letra, 1, Len(letra) - 1) & ")/" & numdia & "")
   vaSpread1.Formula = Fg_Sacacremilla("(" & Mid(letra, 1, Len(letra) - 1) & ")")
End If

'-------> Sumar numero recetas no planificada
vaSpread1.Row = vaSpread1.MaxRows - 1
letra = ""
numdia = 0
j1 = 1
For ii = 2 To (vaSpread1.MaxCols - 1) Step 4 '- maxColumna
    vaSpread1.Col = ii + 1
    If vaSpread1.text <> "" Then
       letra = letra & "'" & vectorrac(j1) & vaSpread1.Row & "'+"
    End If
    j1 = j1 + 1
Next ii
If letra <> "" Then
   vaSpread1.Row = vaSpread1.MaxRows - 1
   vaSpread1.Col = vaSpread1.MaxCols
   vaSpread1.Font.Bold = True
   vaSpread1.Font.Size = 9
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = 0
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.Lock = True
'  vaSpread1.Formula = Fg_Sacacremilla("SUM(" & Mid(letra, 1, Len(letra) - 1) & ")/" & numdia & "")
   vaSpread1.Formula = Fg_Sacacremilla("(" & Mid(letra, 1, Len(letra) - 1) & ")")
End If

vaSpread1.Col = -1: vaSpread1.Row = -1
vaSpread1.Lock = True
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 2
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Dim x As Boolean
    ' Export Excel file and set result to x
    If Dir(dir_trabajo & "Casino Real Casino.XLS") <> "" Then Kill dir_trabajo & "Casino Real Casino.XLS"
    x = vaSpread1.ExportToExcel(dir_trabajo & "Casino Real Casino.XLS", "Test Sheet 1", dir_trabajo & "LOGFILE.TXT")
    ' Display result to user based on T/F value of x
    If x = True Then
'        MsgBox "Export complete.", , "Result"
        Dim XL As Excel.Application
        Set XL = CreateObject("Excel.application")
        XL.Workbooks.Open FileName:=dir_trabajo & "Casino Real Casino.XLS"
'        XL.Cells.Select ''-------> Desactivar proteción
'        XL.ActiveSheet.Unprotect
        XL.Rows("1:1").Select '------> Insert Fila
        XL.Selection.Insert 'Shift:=xlDown
        XL.Range("B1").Select
        XL.ActiveCell.FormulaR1C1 = Label1.Caption
'        XL.Range("B1").Select
'        XL.ActiveCell.FormulaR1C1 = "Código Ingrediente"
'        XL.Range("C1").Select
'        XL.ActiveCell.FormulaR1C1 = "Descripción"
'        XL.Range("D1").Select
'        XL.ActiveCell.FormulaR1C1 = "Unidad Ingrediente"
'        XL.Range("E1").Select
'        XL.ActiveCell.FormulaR1C1 = "Valor Unidad"
'        XL.Range("F1").Select
'        XL.ActiveCell.FormulaR1C1 = "Tipo Ingrediente"
'        XL.Range("G1").Select
'        XL.ActiveCell.FormulaR1C1 = "Frecuencia Ingrediente"
'        XL.Range("H1").Select
'        XL.ActiveCell.FormulaR1C1 = "Código Productos"
'        XL.Range("I1").Select
'        XL.ActiveCell.FormulaR1C1 = "Descripción"
'        XL.Range("J1").Select
'        XL.ActiveCell.FormulaR1C1 = "Tipo Productos"
'        XL.Range("K1").Select
'        XL.ActiveCell.FormulaR1C1 = "Precio"
'        XL.Range("L1").Select
'        XL.ActiveCell.FormulaR1C1 = "Precio Calculado"
'        XL.Range("M1").Select
'        XL.ActiveCell.FormulaR1C1 = "Gramaje"
'        XL.Range("N1").Select
'        XL.ActiveCell.FormulaR1C1 = "Total"
'        XL.ActiveWindow.SplitRow = 0.625
'        XL.ActiveWindow.SplitRow = 0.6875
'        XL.Cells.Select '-------> Activar proteción
'        XL.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
        XL.Visible = True '------->Visualizar
    Else
        MsgBox "Archivo esta abierto, grabe con otro nombre y luego cierre libro", , "Result"
    End If
Case 4
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 70 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub
