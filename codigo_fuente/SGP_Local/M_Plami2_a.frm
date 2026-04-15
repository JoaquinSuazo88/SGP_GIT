VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_Plami2_a 
   Caption         =   "Planificación Teórica"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12150
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   12150
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   10905
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estructura de Servicio"
         Height          =   195
         Index           =   2
         Left            =   5520
         TabIndex        =   6
         Top             =   135
         Width           =   1560
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   5160
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Bloqueada"
         Height          =   195
         Index           =   1
         Left            =   7725
         TabIndex        =   5
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   7365
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         Height          =   195
         Index           =   0
         Left            =   9540
         TabIndex        =   4
         Top             =   135
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   9180
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
         TabIndex        =   3
         Top             =   150
         Width           =   1215
      End
   End
   Begin FPSpread.vaSpread vaSpread2 
      DragIcon        =   "M_Plami2_a.frx":0000
      Height          =   3840
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   11055
      _Version        =   393216
      _ExtentX        =   19500
      _ExtentY        =   6773
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
      MaxCols         =   249
      MaxRows         =   100
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RestrictRows    =   -1  'True
      SpreadDesigner  =   "M_Plami2_a.frx":0442
      UserResize      =   1
      VisibleCols     =   1
      VisibleRows     =   100
      TextTip         =   2
      TextTipDelay    =   0
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12150
      _ExtentX        =   21431
      _ExtentY        =   1005
      ButtonWidth     =   714
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
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
      Height          =   675
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   10905
   End
End
Attribute VB_Name = "M_Plami2_a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim i As Long, j As Long, indcortarpegar As Long, fecha As Long, maxColumna As Long, maxfila As Long
Dim iblockrow As Integer, iblockrow2 As Integer, iblockcol As Integer, iblockcol2 As Integer, SwSalir As Integer
Dim aiblockrow As Integer, aiblockrow2 As Integer, aiblockcol As Integer, aiblockcol2 As Integer, indactivo As Integer
Dim indcos As Boolean
Dim veccos() As Variant
Dim vectorcol() As Long
Dim Msgtitulo As String
Dim TipoCopia As String

Sub DetallePlantillaMinuta_Calorias()

vaSpread2.Lock = True
If vaSpread2.Visible = True Then vaSpread2.Visible = False: Exit Sub
fg_carga ""
Dim indrow3 As Long, inddia As Long, fecha As String, spid As Long
Dim sw As Boolean: sw = False

SwSalir = 0: maxColumna = 0: indactivo = 0
iblockrow = 0: iblockrow2 = 0: iblockcol = 0: iblockcol2 = 0: SwSalir = 0
aiblockrow = 0: aiblockrow2 = 0: aiblockcol = 0: aiblockcol2 = 0

vg_db.Execute "DELETE paso_servicio WHERE ser_spid=@@spid and ser_usr='" & vg_NUsr & "'"
'--isel = 0
'-------> Buscar spid
Set RS = vg_db.Execute("SELECT @@spid spid")
If Not RS.EOF Then spid = RS!spid: vg_db.Execute "INSERT INTO paso_servicio VALUES (" & spid & ", '" & vg_NUsr & "', " & Val(vg_codservicio) & ")"
RS.Close: Set RS = Nothing

'-------> Formatear columna
maxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
vaSpread2.MaxRows = 100
vaSpread2.MaxCols = 0: vaSpread2.MaxCols = 5 * maxColumna + 1: vaSpread2.Row = 0
vaSpread2.Col = 1
vaSpread2.ColsFrozen = 1
vaSpread2.VisibleCols = 1
vaSpread2.ColWidth(1) = 15
vaSpread2.text = "Estructura Servicio"
ReDim Preserve vectorcol(0)
For i = 2 To vaSpread2.MaxCols Step 5
    
    vaSpread2.Col = i
    vaSpread2.ColWidth(i) = 1.5
    vaSpread2.text = " "
    vaSpread2.ColHidden = False
    
    vaSpread2.Col = i + 1
    vaSpread2.ColWidth(i + 1) = 21
    If i = 2 Then
       ReDim Preserve vectorcol(1)
       vectorcol(1) = 3
       vaSpread2.text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & (i - 1), 2), 1), 1, 3) & " " & (i - 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
    Else
       vaSpread2.text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & CLng((i / 5) + 1), 2), 1), 1, 3) & " " & CLng((i / 5) + 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
       ReDim Preserve vectorcol(CLng((i / 5) + 1))
       vectorcol(CLng((i / 5) + 1)) = i + 1
    End If
    vaSpread2.ColHidden = False
    
    vaSpread2.Col = i + 2
    vaSpread2.ColWidth(i + 2) = 6
    vaSpread2.text = "N.Rac."
    vaSpread2.ColHidden = False
   
    vaSpread2.Col = i + 3
    vaSpread2.ColWidth(i + 3) = 9
    vaSpread2.text = "Costo"
    vaSpread2.ColHidden = False
    
    vaSpread2.Col = i + 4
    vaSpread2.text = "Calorias"
    vaSpread2.ColHidden = False
    
'    vaSpread2.Col = i + 5
'    vaSpread2.ColWidth(i + 2) = 6
'    vaSpread2.text = "Calorias"
'    vaSpread2.ColHidden = False
'
    For j = 1 To vaSpread2.MaxRows
        vaSpread2.Row = j

        vaSpread2.Col = i
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignLeft
        vaSpread2.text = ""

        vaSpread2.Col = i + 1
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignLeft
        vaSpread2.text = " "

        vaSpread2.Col = i + 2
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignLeft
        vaSpread2.text = " "

        vaSpread2.Col = i + 3
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignLeft
        vaSpread2.text = " "

        vaSpread2.Col = i + 4
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignLeft
        vaSpread2.text = " "

    Next j
    vaSpread2.Row = 0
Next i

vaSpread2.Row = 0
For i = 1 To maxColumna
   vaSpread2.MaxCols = vaSpread2.MaxCols + 1
   vaSpread2.Col = vaSpread2.MaxCols
   vaSpread2.text = "Estado"
   vaSpread2.ColHidden = True
Next i
vaSpread2.MaxCols = vaSpread2.MaxCols + 1
vaSpread2.Col = vaSpread2.MaxCols
vaSpread2.ColWidth(vaSpread2.MaxCols) = 5
vaSpread2.text = "Còd. Est."
vaSpread2.ColHidden = True

vaSpread2.Row = -1: vaSpread2.Col = -1: vaSpread2.BackColor = Shape1(0).FillColor  'Amarillo
vaSpread2.Row = -1: vaSpread2.Col = 1
vaSpread2.Font.Bold = True
vaSpread2.Font.Size = 9
vaSpread2.BackColor = Shape1(2).FillColor 'Verde

j = 0: i = 0: indrow3 = 0
Set RS = vg_db.Execute("sgpadm_s_PlanMinutaDetreal " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(vg_fecha) & ", " & vg_codlpr & ",'" & vg_NUsr & "','" & spid & "','" & vg_IndpprSelec & "'")
DoEvents
If Not RS.EOF Then
  sw = True
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 5) - 5) + 1) + 1
      vaSpread2.Row = RS!mid_numlin
      If indrow3 < vaSpread2.Row Then indrow3 = vaSpread2.Row
      If RS!ess_codigo <> i Then
         vaSpread2.Col = 1
         vaSpread2.text = RS!ess_nombre
         
         vaSpread2.Col = vaSpread2.MaxCols
         vaSpread2.CellType = CellTypeStaticText
         vaSpread2.TypeHAlign = TypeHAlignCenter
         vaSpread2.text = RS!ess_codigo
         i = RS!ess_codigo
      End If
      vaSpread2.Col = j
      vaSpread2.CellType = CellTypeStaticText
      vaSpread2.TypeHAlign = TypeHAlignCenter
      vaSpread2.Value = "R"
      vaSpread2.ForeColor = &HFF&
      vaSpread2.BackColor = &H80FF80
           
      vaSpread2.Col = j + 1
      vaSpread2.CellType = CellTypeStaticText
      vaSpread2.TypeHAlign = TypeHAlignLeft
      vaSpread2.text = Trim(RS!pas_nombre)
                         
      vaSpread2.Col = j + 2
      vaSpread2.CellType = CellTypeNumber
      'vaSpread2.TypeNumberDecPlaces = 0
      'vaSpread2.TypeIntegerMin = 1
      'vaSpread2.TypeIntegerMax = 9999999
      'vaSpread2.TypeHAlign = TypeHAlignRight
      'vaSpread2.TypeSpin = False
      'vaSpread2.TypeIntegerSpinInc = 1
      'vaSpread2.TypeIntegerSpinWrap = False
      vaSpread2.Value = RS!mid_numrac
      'vaSpread2.ForeColor = &HFF0000
                       
      vaSpread2.Col = j + 3
      vaSpread2.CellType = CellTypeStaticText
      vaSpread2.TypeHAlign = TypeHAlignRight
      precio = Format(IIf(IsNull(RS!pas_prerec) Or Trim(RS!pas_prerec) = 0, 0, RS!pas_prerec), fg_Pict(6, 2))
      vaSpread2.text = precio
      
      vaSpread2.Col = j + 4: vaSpread2.text = Format(RS!candiet, fg_Pict(6, 2)) 'RS!pas_codrec & "&" & RS!mid_tiprec & "&;"
      'vaSpread2.Col = j + 5: vaSpread2.text = RS!candiet
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing: fg_descarga
Else
    'Retorna minuta sin precio
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 1, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Val(vg_fecha) & ", 0,0,'" & vg_IndpprSelec & "'")
    DoEvents
    If Not RS.EOF Then 'Consulta trae productos sin costo
      sw = True
        Do While Not RS.EOF
              DoEvents
              j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 5) - 5) + 1) + 1
              vaSpread2.Row = RS!mid_numlin
              If indrow3 < vaSpread2.Row Then indrow3 = vaSpread2.Row
              If RS!mid_estser <> i Then
                 vaSpread2.Col = 1
                 vaSpread2.text = RS!ess_nombre
                 vaSpread2.Col = vaSpread2.MaxCols
                 vaSpread2.CellType = CellTypeStaticText
                 vaSpread2.TypeHAlign = TypeHAlignCenter
                 vaSpread2.text = RS!mid_estser
                 i = RS!mid_estser
              End If
              vaSpread2.Col = j
              vaSpread2.CellType = CellTypeStaticText
              vaSpread2.TypeHAlign = TypeHAlignCenter
              vaSpread2.Value = "R"
              vaSpread2.ForeColor = &HFF&
              vaSpread2.BackColor = &H80FF80
                   
              vaSpread2.Col = j + 1
              vaSpread2.CellType = CellTypeStaticText
              vaSpread2.TypeHAlign = TypeHAlignLeft
              vaSpread2.text = Trim(RS!mid_descri)
                                 
              vaSpread2.Col = j + 2
              vaSpread2.CellType = CellTypeNumber
'              vaSpread2.TypeNumberDecPlaces = 0
'              vaSpread2.TypeIntegerMin = 1
'              vaSpread2.TypeIntegerMax = 9999999
'              vaSpread2.TypeHAlign = TypeHAlignRight
'              vaSpread2.TypeSpin = False
'              vaSpread2.TypeIntegerSpinInc = 1
'              vaSpread2.TypeIntegerSpinWrap = False
              vaSpread2.Value = RS!mid_numrac
'              vaSpread2.ForeColor = &HFF0000
                               
              vaSpread2.Col = j + 3
              vaSpread2.CellType = CellTypeStaticText
              vaSpread2.TypeHAlign = TypeHAlignRight
              vaSpread2.text = Format((IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec) + IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec)), fg_Pict(6, 2))
              
              vaSpread2.Col = j + 4: vaSpread2.text = Format(RS!candiet, fg_Pict(6, 2))
              vaSpread2.ForeColor = &HFF&
              vaSpread2.BackColor = &H80FF80
              'vaSpread2.Col = j + 5: vaSpread2.text = ""
              
              'If RS!min_indblo > 0 Then VaSpread2.Row = -1: VaSpread2.Col = j: VaSpread2.BackColor = Shape1(1).FillColor: VaSpread2.Col = j + 1: VaSpread2.BackColor = Shape1(1).FillColor: VaSpread2.Col = j + 2: VaSpread2.CellType = 5: VaSpread2.TypeHAlign = 1: VaSpread2.BackColor = Shape1(1).FillColor: VaSpread2.Col = j + 3: VaSpread2.BackColor = Shape1(1).FillColor
              RS.MoveNext
           Loop
        End If
   RS.Close: Set RS = Nothing: fg_descarga
End If

If sw = False And vg_IndpprSelec = 1 Then   '--->Trae estructura completa si no hay registros de minuta.
   Set RS = vg_db.Execute("sgpadm_s_estservicio 1, " & vg_codservicio & ",''")
   If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
   Do While Not RS.EOF
      vaSpread2.Row = RS!ess_orden
      If indrow3 < vaSpread2.Row Then indrow3 = vaSpread2.Row
      vaSpread2.Col = 1
      vaSpread2.text = RS!ess_nombre
      For i = 2 To vaSpread2.MaxCols Step 5
          vaSpread2.Col = vaSpread2.MaxCols
          vaSpread2.text = RS!ess_codigo
      Next i
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
Else
   indrow3 = 20
End If

For i = 3 To (vaSpread2.MaxCols - maxColumna) Step 5
    vaSpread2.Row = 0: vaSpread2.Col = i
    If CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
        Dim fil As Long, Col As Long
        For fil = 1 To (vaSpread2.MaxRows - 1)
            For Col = i - 1 To i + 2
                vaSpread2.Row = fil: vaSpread2.Col = Col
                If vaSpread2.CellType = CellTypeNumber Then vaSpread2.CellType = CellTypeStaticText: vaSpread2.TypeHAlign = TypeHAlignRight
                vaSpread2.BackColor = Shape1(1).FillColor
            Next Col
        Next fil
    End If
Next i


For i = 3 To (vaSpread2.MaxCols - maxColumna) Step 5
    vaSpread2.Row = 0: vaSpread2.Col = i
    If CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
        
        For fil = 1 To (vaSpread2.MaxRows - 1)
            For Col = i - 1 To i + 4
                vaSpread2.Row = fil: vaSpread2.Col = Col
                If vaSpread2.CellType = CellTypeNumber Then vaSpread2.CellType = CellTypeStaticText: vaSpread2.TypeHAlign = TypeHAlignRight
                vaSpread2.BackColor = Shape1(1).FillColor
            Next Col
        Next fil
    End If
Next i

vaSpread2.MaxRows = indrow3 + 1
vaSpread2.Row = vaSpread2.MaxRows
maxfila = vaSpread2.MaxRows
vaSpread2.Col = 1
vaSpread2.text = "Comensales"
vaSpread2.Col = -1: vaSpread2.BackColor = &HE0E0E0
'formatear ultima columna
For i = 2 To (vaSpread2.MaxCols - maxColumna) Step 5
    vaSpread2.Row = vaSpread2.MaxRows
    vaSpread2.Col = i + 2
    vaSpread2.CellType = CellTypeNumber
    vaSpread2.TypeNumberDecPlaces = 0
    vaSpread2.TypeIntegerMin = 1
    vaSpread2.TypeIntegerMax = 9999999
    vaSpread2.TypeHAlign = TypeHAlignRight
    vaSpread2.TypeSpin = False
    vaSpread2.TypeIntegerSpinInc = 1
    vaSpread2.TypeIntegerSpinWrap = False
    vaSpread2.Value = Format(0, fg_Pict(6, 0))
    vaSpread2.ForeColor = &HFF0000
Next i
'Mover comensales
'RS.Open "SELECT min_racteo, min_fecmin FROM b_minuta " & _
'        "WHERE  min_subseg=" & vg_codsubseg & " AND min_codreg=" & vg_codregimen & " AND min_codser=" & vg_codservicio & " " & _
'        "AND    substring(convert(char(8),b_minuta.min_fecmin),1,6)=" & Val(vg_fecha) & " ORDER BY min_fecmin", vg_db, adOpenForwardOnly

Set RS = vg_db.Execute("sgpadm_s_planifminuta 2, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Val(vg_fecha) & ", 0, 0,'" & vg_IndpprSelec & "'")
DoEvents
If Not RS.EOF Then
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 5) - 5) + 1) + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = j + 2
      vaSpread2.CellType = CellTypeNumber
      vaSpread2.TypeNumberDecPlaces = 0
      vaSpread2.TypeIntegerMin = 1
      vaSpread2.TypeIntegerMax = 9999999
      vaSpread2.TypeHAlign = TypeHAlignRight
      vaSpread2.TypeSpin = False
      vaSpread2.TypeIntegerSpinInc = 1
      vaSpread2.TypeIntegerSpinWrap = False
      vaSpread2.Value = IIf(IsNull(RS!min_racteo), 0, RS!min_racteo)
      vaSpread2.ForeColor = &HFF0000
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
Else
   RS.Close: Set RS = Nothing
   Set RS = vg_db.Execute("sgpadm_s_servraciones " & vg_codservicio & "")
   DoEvents
   If Not RS.EOF Then
      Do While Not RS.EOF
         inddia = 1
         For i = 2 To (vaSpread2.MaxCols - maxColumna - 1) Step 5
             If RS!sra_serdia = IIf(fg_Dia(vg_fecha & fg_pone_cero(inddia, 2)) = 1, 7, Val(fg_Dia(vg_fecha & fg_pone_cero(inddia, 2)) - 1)) Then
                vaSpread2.Col = i + 2
                vaSpread2.CellType = CellTypeNumber
                vaSpread2.TypeNumberDecPlaces = 0
                vaSpread2.TypeIntegerMin = 1
                vaSpread2.TypeIntegerMax = 9999999
                vaSpread2.TypeHAlign = TypeHAlignRight
                vaSpread2.TypeSpin = False
                vaSpread2.TypeIntegerSpinInc = 1
                vaSpread2.TypeIntegerSpinWrap = False
                vaSpread2.Value = RS!Raciones
                vaSpread2.ForeColor = &HFF0000
             End If
             inddia = inddia + 1
         Next i
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
End If

For i = 3 To (vaSpread2.MaxCols - maxColumna) Step 5
    vaSpread2.Row = 0: vaSpread2.Col = i
    If CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread2.text), 5, Len(Trim(vaSpread2.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
        For fil = 1 To (vaSpread2.MaxRows - 1)
            For Col = i - 1 To i + 2
                vaSpread2.Row = vaSpread2.MaxRows: vaSpread2.Col = Col
                If vaSpread2.CellType = CellTypeNumber Then vaSpread2.CellType = CellTypeStaticText: vaSpread2.TypeHAlign = TypeHAlignRight
            Next Col
        Next fil
    End If
Next i

vaSpread2.Row = 1: vaSpread2.Col = 1
iblockrow = vaSpread2.Row: aiblockrow = vaSpread2.Row
iblockrow2 = vaSpread2.Row: aiblockrow2 = vaSpread2.Row
iblockcol = vaSpread2.Col: aiblockcol = vaSpread2.Col
iblockcol2 = vaSpread2.Col: aiblockcol2 = vaSpread2.Col
If vaSpread2.Visible = False Then vaSpread2.Visible = False: vaSpread2.Visible = True
End Sub

Sub ExportarExcel()
Dim NashXl As Excel.Application
Dim irow As Long, irow2 As Long
Dim NColumnas As Integer

fg_carga ""
Set NashXl = CreateObject("excel.application")
Set NashXl = New Excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add
NashXl.Range("A1").Select
NashXl.ActiveCell.FormulaR1C1 = "Sub-Segmento : " & vg_codsubseg & "-" & vg_nomsubseg
NashXl.Range("A2").Select
NashXl.ActiveCell.FormulaR1C1 = "Regimen      : " & vg_codregimen & "-" & vg_nomreg
NashXl.Range("A3").Select
NashXl.ActiveCell.FormulaR1C1 = "Servicio     : " & vg_codservicio & "-" & vg_nomser
NashXl.Range("A4").Select
NashXl.ActiveCell.FormulaR1C1 = "Fecha        : " & vg_fecha

maxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
NColumnas = (maxColumna * 5) + 1
vaSpread2.AllowMultiBlocks = True
'vaSpread2.SetSelection 1, -1, vaSpread2.MaxCols, vaSpread2.MaxRows + 3
vaSpread2.SetSelection 1, -1, NColumnas, vaSpread2.MaxRows + 3
vaSpread2.ClipboardCopy

irow = vaSpread2.MaxRows + 5
'------- Pegar vaSpread2(0) - Planilla Excel
NashXl.Range("A5").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'------- Colorear titulo
NashXl.Range("A5:EZ5").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A5:EZ" & irow).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Range("A2" & ":" & "A" & irow).Select
NashXl.Selection.NumberFormat = "#,##0.00"

'------- Asigna Colores a Estructura de Servicio
NashXl.Range("A6:" & "A" & irow).Select
With NashXl.Selection.Interior
     .ColorIndex = 10
     .Pattern = xlSolid
End With
'------- Aplicar totales

NashXl.Selection.Font.Bold = True

NashXl.Range("B" & irow & ":" & "B" & 2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread2.AllowMultiBlocks = False: vaSpread2.SetSelection 1, 0, vaSpread2.MaxCols, vaSpread2.MaxRows
'
'NashXl.Cells.Replace What:="&0&;", Replacement:="", LookAt:=xlPart, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'NashXl.Cells.Replace What:="&-1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
'        :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
fg_descarga
NashXl.Visible = True
End Sub


Private Sub Form_Load()
Me.Height = 6765
Me.Width = 11055
fg_centra Me
Msgtitulo = "Planificación Teórica"
fg_carga (ss)
Label4.Caption = M_Plami1.fpayuda(0).Caption & "(" & M_Plami1.fpLongInteger1(0).Value & ")" & " - " & M_Plami1.fpayuda(1).Caption & " - " & M_Plami1.fpayuda(2).Caption & " - " & " Tipo: " & IIf(vg_IndpprSelec = "1", "Real", "Propuesta")
Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "

Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): btnX.Visible = True: btnX.ToolTipText = "Exporta Excel "
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
DetallePlantillaMinuta_Calorias


'Dim x As Long
'Set RS = vg_db.Execute("sgpadm_s_estservicio 1, " & vg_codservicio & ",''")
'If Not RS.EOF Then
'    x = 1
'    Do While Not RS.EOF
'        Load Estructura1(x): Load Estructura2(x)
'        Estructura1(x).Caption = Trim(RS!ess_nombre): Estructura2(x).Caption = Trim(RS!ess_nombre)
'        Estructura1(x).HelpContextID = RS!ess_codigo: Estructura2(x).HelpContextID = RS!ess_codigo
'        Estructura1(x).Enabled = True: Estructura2(x).Enabled = True
'        For i = 1 To vaSpread1.MaxRows
'            vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.Row = i
'            If Trim(vaSpread1.text) <> "" Then
'                If Val(vaSpread1.text) = RS!ess_codigo Then Estructura1(x).Enabled = False: Estructura2(x).Enabled = False
'            End If
'        Next
'        x = x + 1
'        RS.MoveNext
'    Loop
'End If
'RS.Close: Set RS = Nothing
'Estructura1(0).Visible = False: Estructura2(0).Visible = False

End Sub

Private Sub Estructura1_Click(Index As Integer)
'LlenaSubMenu Estructura1, Index
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 445
If Me.WindowState <> 1 Then vaSpread2.Move 0, 1380, ScaleWidth, ScaleHeight - 1380
End Sub

Private Sub Form_Unload(Cancel As Integer)
If SwSalir <> 0 Then Exit Sub
If Toolbar1.Buttons(2).Visible = False Then Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
'If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Cancel = -1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
'If Toolbar1.Buttons(2).Visible = True Then GrabarPlantillaMinuta
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
SwSalir = 1
vg_PartePlani = False
Me.Hide
Unload Me
M_Plami1.WindowState = 0
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Index
   Case 1
     ExportarExcel
   Case 2
     M_Plami2_a.Hide
     '
  End Select
End Sub
