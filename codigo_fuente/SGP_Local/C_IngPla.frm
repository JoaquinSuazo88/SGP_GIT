VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form C_IngPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frecuencia Ingrediente"
   ClientHeight    =   7485
   ClientLeft      =   510
   ClientTop       =   1200
   ClientWidth     =   16080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   16080
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   15255
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4980
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   15015
         _Version        =   393216
         _ExtentX        =   26485
         _ExtentY        =   8784
         _StockProps     =   64
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
         MaxCols         =   14
         MaxRows         =   18
         SpreadDesigner  =   "C_IngPla.frx":0000
         VisibleCols     =   8
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   11
         Left            =   11400
         TabIndex        =   9
         Top             =   5535
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   11400
         TabIndex        =   8
         Top             =   5235
         Width           =   975
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Costo Promedio Diario"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   9360
         TabIndex        =   7
         Top             =   5535
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Ingredientes Listados"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   9360
         TabIndex        =   6
         Top             =   5235
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   2
         Left            =   1875
         TabIndex        =   3
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   1875
         TabIndex        =   2
         Top             =   645
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   1875
         TabIndex        =   1
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3240
         TabIndex        =   13
         Top             =   240
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3240
         TabIndex        =   12
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3240
         TabIndex        =   11
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3285
         TabIndex        =   14
         Top             =   285
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3285
         TabIndex        =   15
         Top             =   645
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3285
         TabIndex        =   16
         Top             =   1005
         Width           =   5535
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7485
      Left            =   15450
      TabIndex        =   10
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13203
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_IngPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------------------
' --------------- Formulario : C_IngPla Samuel Melendez 03/09/09 ----------
' --------------- Creado     : Samuel Melendez                   ----------
' --------------- Fecha      : 08/09/09                          ----------
'--------------------------------------------------------------------------

Dim RS1 As New ADODB.Recordset


Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Sub LlenarFrecIng(tfor As String, subseg As Long, codreg As Long, codser As Long, anomes As Long, TipMin As String)
Dim codreceta As Long, irow As Long, i As Long, condia As Long, auxest As Long
Dim cosreceta As Double, canreceta As Double, totgralreceta As Double
fg_carga ""
'-------> Rutina frecuencia de recetas
Me.Caption = tfor
Msgtitulo = tfor
Set RS1 = vg_db.Execute("SELECT sub_codigo, sub_nombre FROM a_subsegmento WHERE sub_codigo = " & subseg & "")
If Not RS1.EOF Then fpayuda(0).Caption = RS1!sub_nombre
RS1.Close: Set RS1 = Nothing
Set RS1 = vg_db.Execute("SELECT reg_nombre FROM a_regimen WHERE reg_codigo = " & codreg & "")
If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
RS1.Close: Set RS1 = Nothing
Set RS1 = vg_db.Execute("SELECT ser_nombre FROM a_servicio WHERE ser_codigo = " & codser & "")
If Not RS1.EOF Then fpayuda(3).Caption = RS1!ser_nombre
RS1.Close: Set RS1 = Nothing

Dim x As Boolean
' Control displays text tips aligned to pointer with focus
vaSpread1(0).TextTip = 2
vaSpread1(0).TextTipDelay = 250
x = vaSpread1(0).SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
vaSpread1(0).MaxRows = 0

DoEvents
Set RS1 = vg_db.Execute("sgpadm_s_frecuenciaing " & subseg & ", " & codreg & ", " & codser & ", " & vg_codlpr & ", " & anomes & ", " & ExraeCodCombo(M_Plami1.Combo2(0)) & " ")
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
'definir largo del vector
Dim preklu As Double
Dim frecest As Long
Dim indfin As Long
Dim vecTipoPla() As Variant
ReDim Preserve vecTipoPla(1000, 3)
codreceta = 0: cosreceta = 0: canreceta = 0: totgralreceta = 0: condia = 0
indini = 1: indfin = 0
irow = 1: auxest = 0
Do While Not RS1.EOF
   DoEvents
   vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
   vaSpread1(0).Row = vaSpread1(0).MaxRows
    
   If auxest <> RS1!mid_estser Then
      If auxest <> 0 Then
         vaSpread1(0).Col = 1 '-------> Glosa Días Planificados
         vaSpread1(0).CellType = CellTypeStaticText
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignLeft
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
         vaSpread1(0).text = "'Días Planificados"
         
         vaSpread1(0).Col = 7 '-------> Días Planificados
         vaSpread1(0).CellType = CellTypeCurrency
         vaSpread1(0).TypeCurrencyDecPlaces = 0
         vaSpread1(0).TypeCurrencyShowSymbol = False
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).text = Format(frecest, fg_Pict(6, 0))
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
         
         vaSpread1(0).Col = 14 '-------> Total
         vaSpread1(0).CellType = CellTypeNumber
         vaSpread1(0).Lock = True
         vaSpread1(0).TypeHAlign = TypeHAlignRight
         vaSpread1(0).Formula = Fg_Sacacremilla("SUM('" & "N" & indini & "':'" & "N" & vaSpread1(0).Row - 1 & "')" & "/" & "SUM('" & "G" & vaSpread1(0).Row & "')")
         vaSpread1(0).Font.Bold = True
         vaSpread1(0).Font.Size = 9
         
         vecTipoPla(irow - 1, 3) = vaSpread1(0).Row
         
         vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
         vaSpread1(0).Row = vaSpread1(0).MaxRows
      End If
      vaSpread1(0).Col = 1 '-------> Descripción Estructura Servicio
      vaSpread1(0).CellType = CellTypeStaticText
      vaSpread1(0).Lock = True
      vaSpread1(0).TypeHAlign = TypeHAlignLeft
      vaSpread1(0).Font.Bold = True
      vaSpread1(0).Font.Size = 9
      vaSpread1(0).text = " " & Trim(RS1!mid_desest)
      
      auxest = RS1!mid_estser
      vecTipoPla(irow, 1) = RS1!mid_estser
      vecTipoPla(irow, 2) = " " & Trim(RS1!mid_desest)
      frecest = RS1!frecest
      indini = vaSpread1(0).Row
      irow = irow + 1
   End If
   
   vaSpread1(0).Col = 2 '-------> Codigo Ingrediente
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = RS1!ing_codigo
         
   vaSpread1(0).Col = 3 '-------> Nombre Ingrediente
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(IsNull(RS1!ing_nombre), "", " " & Trim(RS1!ing_nombre))
   
   vaSpread1(0).Col = 4 '-------> Nombre Unidad Medida
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(IsNull(RS1!unm_nomcor), "", " " & Trim(RS1!unm_nomcor))
   
   vaSpread1(0).Col = 5 '-------> Valor Ingrediente
   vaSpread1(0).CellType = CellTypeCurrency
   vaSpread1(0).TypeCurrencyDecPlaces = 0
   vaSpread1(0).TypeCurrencyShowSymbol = False
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = IIf(IsNull(RS1!pro_facing), 0, RS1!pro_facing) 'Format(IIf(Trim(RS1!unm_nomcor) = "Gr" Or Trim(RS1!unm_nomcor) = "Cc", 1000, 1), fg_Pict(6, 0))
   
   vaSpread1(0).Col = 6 '-------> Tipo Ingrediente
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(RS1!ing_indppr = "1", "Real", "Propuesta")
   
   vaSpread1(0).Col = 7 '-------> Frecuencia
   vaSpread1(0).CellType = CellTypeCurrency
   vaSpread1(0).TypeCurrencyDecPlaces = 0
   vaSpread1(0).TypeCurrencyShowSymbol = False
   vaSpread1(0).Lock = False
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = Format(RS1!FrecIng, fg_Pict(6, 0))

   vaSpread1(0).Col = 8 '-------> Código Producto
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = Trim(RS1!pro_codigo)
   
   vaSpread1(0).Col = 9 '-------> Nombre Producto
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(IsNull(RS1!pro_nombre), "No existe Productos", " " & Trim(RS1!pro_nombre))

   vaSpread1(0).Col = 10 '-------> Tipo Producto
   vaSpread1(0).CellType = CellTypeStaticText
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignLeft
   vaSpread1(0).text = IIf(RS1!pro_indppr = "1", "Real", "Propuesta")
         
   vaSpread1(0).Col = 11 '-------> Precio
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = Format(RS1!dlp_precio, fg_Pict(6, 0))
   vaSpread1(0).ForeColor = &HFF0000
         
   vaSpread1(0).Col = 12 '------> Conversión Precio
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   If Trim(RS1!unm_nomcor) = "Gr" Or Trim(RS1!unm_nomcor) = "Cc" Then
      vaSpread1(0).Formula = "SUM((K#/E#)*1000)" '& "/" & "SUM(E#)"
   Else
      vaSpread1(0).Formula = "SUM((K#/E#)*1)"
   End If
   'Format(IIf(Trim(RS1!unm_nomcor) = "Gr" Or Trim(RS1!unm_nomcor) = "Cc", (RS1!dlp_precio / 1000) * 1000, RS1!dlp_precio), fg_Pict(6, 0))
   vaSpread1(0).ForeColor = &HFF0000
         
   vaSpread1(0).Col = 13 '-------> Gramaje
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = False
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).text = Format(RS1!red_canpro, fg_Pict(6, vg_RDCa))
   
   vaSpread1(0).Col = 14 '-------> Total
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   If Trim(RS1!unm_nomcor) = "Gr" Or Trim(RS1!unm_nomcor) = "Cc" Then
      vaSpread1(0).Formula = "SUM((G#*L#*M#)/1000)"
   Else
      vaSpread1(0).Formula = "SUM((G#*L#*M#)/1)"
   End If
   RS1.MoveNext
Loop
   vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
   vaSpread1(0).Row = vaSpread1(0).MaxRows
         
    vaSpread1(0).Col = 1 '-------> Glosa Días Planificados
    vaSpread1(0).CellType = CellTypeStaticText
    vaSpread1(0).Lock = True
    vaSpread1(0).TypeHAlign = TypeHAlignLeft
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9
    vaSpread1(0).text = "Días Planificados"
    
    vaSpread1(0).Col = 7 '-------> Días Planificados
    vaSpread1(0).CellType = CellTypeCurrency
    vaSpread1(0).TypeCurrencyDecPlaces = 0
    vaSpread1(0).TypeCurrencyShowSymbol = False
    vaSpread1(0).Lock = True
    vaSpread1(0).TypeHAlign = TypeHAlignRight
    vaSpread1(0).text = Format(frecest, fg_Pict(6, 0))
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9
    
    vaSpread1(0).Col = 14 '-------> Total
    vaSpread1(0).CellType = CellTypeNumber
    vaSpread1(0).Lock = True
    vaSpread1(0).TypeHAlign = TypeHAlignRight
    vaSpread1(0).Formula = Fg_Sacacremilla("SUM('" & "N" & indini & "':'" & "N" & vaSpread1(0).Row - 1 & "')" & "/" & "SUM('" & "G" & vaSpread1(0).Row & "')")
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9
    
    vecTipoPla(irow - 1, 3) = vaSpread1(0).Row

RS1.Close: Set RS1 = Nothing
Label1(9).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(6, 2))
Label1(11).Caption = Format(totgralreceta, fg_Pict(6, 2))
Set RS1 = vg_db.Execute("SELECT COUNT(b_minuta.min_codigo) AS nreg FROM b_minuta WHERE b_minuta.min_codigo IN (SELECT b_minutadet.mid_codigo FROM b_minutadet WHERE b_minutadet.mid_tipmin = '" & TipMin & "') AND b_minuta.min_subseg = " & subseg & " " & _
                        "AND b_minuta.min_codreg = " & codreg & " AND b_minuta.min_codser = " & codser & " AND substring(convert(char(8),b_minuta.min_fecmin),1,6) = " & anomes & "")
If Not RS1.EOF And RS1!nreg > 0 Then Label1(11).Caption = Format(CCur(totgralreceta / RS1!nreg), fg_Pict(6, 2))
RS1.Close: Set RS1 = Nothing

'-------> mover sub-segmento
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Sub-Segmento"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = fpayuda(0).Caption
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9
'-------> Mover regimen
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Regimen"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = fpayuda(1).Caption
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9
'-------> mover Servicio
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Servicio"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = fpayuda(3).Caption
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

'-------> Mover resumen
vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "RESUMEN COSTO"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 3
vaSpread1(0).text = "Estructura Servicio"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 4
vaSpread1(0).text = "Costo"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 5
vaSpread1(0).text = "Ponderado"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

vaSpread1(0).Col = 6
vaSpread1(0).text = "Total"
vaSpread1(0).Lock = True
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9

indfin = 0
For i = 1 To irow
   vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
   vaSpread1(0).Row = vaSpread1(0).MaxRows
   
'   vaSpread1(0).Col = 3
'   vaSpread1(0).text = vecTipoPla(i, 1)
'   vaSpread1(0).Font.Bold = True
'   vaSpread1(0).Font.Size = 9
   
   vaSpread1(0).Col = 3 '-------> Descripción Estructura
   vaSpread1(0).text = vecTipoPla(i, 2)
   vaSpread1(0).Lock = True
   vaSpread1(0).Font.Bold = True
   vaSpread1(0).Font.Size = 9

   vaSpread1(0).Col = 4 '-------> Total
   vaSpread1(0).CellType = CellTypeNumber
   vaSpread1(0).Lock = True
   vaSpread1(0).TypeHAlign = TypeHAlignRight
   vaSpread1(0).Formula = Fg_Sacacremilla("SUM('" & "N" & vecTipoPla(i, 3) & "':'" & "N" & vecTipoPla(i, 3) & "')")
   vaSpread1(0).Font.Bold = True
   vaSpread1(0).Font.Size = 9

Next i
fg_descarga
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 2
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    Dim x As Boolean
    ' Export Excel file and set result to x
    If Dir(dir_trabajo & "Frecuencia Ingrediente.XLS") <> "" Then Kill dir_trabajo & "Frecuencia Ingrediente.XLS"
    x = vaSpread1(0).ExportToExcel(dir_trabajo & "Frecuencia Ingrediente.XLS", "Test Sheet 1", dir_trabajo & "LOGFILE.TXT")
    ' Display result to user based on T/F value of x
    If x = True Then
'        MsgBox "Export complete.", , "Result"
        Dim XL As Excel.Application
        Set XL = CreateObject("Excel.application")
        XL.Workbooks.Open FileName:=dir_trabajo & "Frecuencia Ingrediente.XLS"
        XL.Cells.Select ''-------> Desactivar proteción
        XL.ActiveSheet.Unprotect
        XL.Rows("1:1").Select '------> Insert Fila
        XL.Selection.Insert 'Shift:=xlDown
        XL.Range("A1").Select
        XL.ActiveCell.FormulaR1C1 = "Estructura Servicio"
        XL.Range("B1").Select
        XL.ActiveCell.FormulaR1C1 = "Código Ingrediente"
        XL.Range("C1").Select
        XL.ActiveCell.FormulaR1C1 = "Descripción"
        XL.Range("D1").Select
        XL.ActiveCell.FormulaR1C1 = "Unidad Ingrediente"
        XL.Range("E1").Select
        XL.ActiveCell.FormulaR1C1 = "Valor Unidad"
        XL.Range("F1").Select
        XL.ActiveCell.FormulaR1C1 = "Tipo Ingrediente"
        XL.Range("G1").Select
        XL.ActiveCell.FormulaR1C1 = "Frecuencia Ingrediente"
        XL.Range("H1").Select
        XL.ActiveCell.FormulaR1C1 = "Código Productos"
        XL.Range("I1").Select
        XL.ActiveCell.FormulaR1C1 = "Descripción"
        XL.Range("J1").Select
        XL.ActiveCell.FormulaR1C1 = "Tipo Productos"
        XL.Range("K1").Select
        XL.ActiveCell.FormulaR1C1 = "Precio"
        XL.Range("L1").Select
        XL.ActiveCell.FormulaR1C1 = "Precio Calculado"
        XL.Range("M1").Select
        XL.ActiveCell.FormulaR1C1 = "Gramaje"
        XL.Range("N1").Select
        XL.ActiveCell.FormulaR1C1 = "Total"
        XL.ActiveWindow.SplitRow = 0.625
        XL.ActiveWindow.SplitRow = 0.6875
        XL.Cells.Select '-------> Activar proteción
        XL.ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
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
If Err = 70 Or Err = 1004 Or Err = 91 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"
End Sub

'Sub ExportarExcel()
'Dim NashXl As Excel.Application
'Dim irow As Long, irow2 As Long
'fg_carga ""
'Set NashXl = CreateObject("excel.application")
'Set NashXl = New Excel.Application
'NashXl.SheetsInNewWorkbook = 1
'NashXl.Workbooks.OpenDatabase ("C:\temp\Frecuencia Ingrediente.XLS")
'
'vaSpread1(0).AllowMultiBlocks = True
'vaSpread1(0).SetSelection 1, -1, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows
'vaSpread1(0).ClipboardCopy
'irow = vaSpread1(0).MaxRows + 1
''------- Pegar vaspread1(1) - Planilla Excel
'NashXl.Range("A1").Select
'NashXl.ActiveSheet.Paste
''------- Asignar color
''NashXl.Range("A1:D" & irow).Select
''With NashXl.Selection.Interior
''     .ColorIndex = 36
''     .Pattern = xlSolid
''End With
''------- Colorear titulo
'NashXl.Range("A1:H1").Select ' samuel 03/0309
'With NashXl.Selection.Interior
'     .ColorIndex = 15
'     .Pattern = xlSolid
'End With
''------- Dibujar marco
'NashXl.Range("A1:H" & irow).Select ' samuel 03/09/09
'NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'With NashXl.Selection.Borders(xlEdgeLeft)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlEdgeTop)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlEdgeBottom)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlEdgeRight)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlInsideVertical)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlInsideHorizontal)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'NashXl.Range("D2" & ":" & "D" & irow).Select
'NashXl.Selection.NumberFormat = "#,##0.00"
''------- Aplicar totales
'
''------- Dibujar marco
'irow = irow + 2
'irow2 = irow + 2
'NashXl.Range("B" & irow).Select
'NashXl.ActiveCell.FormulaR1C1 = Label1(8).Caption
'NashXl.Range("C" & irow).Select
'NashXl.ActiveCell.FormulaR1C1 = Label1(9).Caption
''NashXl.Range("B" & irow2).Select
''NashXl.ActiveCell.FormulaR1C1 = Label1(10).Caption
''NashXl.Range("C" & irow2).Select
''NashXl.ActiveCell.FormulaR1C1 = Label1(11).Caption
'NashXl.Range("B" & irow & ":" & "C" & irow).Select
'NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'With NashXl.Selection.Borders(xlEdgeLeft)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlEdgeTop)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlEdgeBottom)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlEdgeRight)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
'With NashXl.Selection.Borders(xlInsideVertical)
'     .LineStyle = xlContinuous
'     .Weight = xlThin
'     .ColorIndex = xlAutomatic
'End With
''With NashXl.Selection.Borders(xlInsideHorizontal)
''     .LineStyle = xlContinuous
''     .Weight = xlThin
''     .ColorIndex = xlAutomatic
''End With
''NashXl.Selection.Font.Bold = True
''With NashXl.Selection.Interior
''     .ColorIndex = 35
''     .Pattern = xlSolid
''End With
'NashXl.Range("D" & irow & ":" & "D" & irow).Select
'NashXl.Selection.NumberFormat = "#,##0.00"
''------- Ajustar columna
'NashXl.Cells.Select
'NashXl.Cells.EntireColumn.AutoFit
'vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).MaxCols, vaSpread1(0).MaxRows
'fg_descarga
'NashXl.Visible = True
'End Sub

Private Sub vaSpread1_TextTipFetch(Index As Integer, ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If vaSpread1(0).MaxRows < 1 Then Exit Sub
' Set tip to display and set tip's content
vaSpread1(0).Row = Row
TipWidth = 4000
ShowTip = True
MultiLine = 2
Select Case Col
Case 1
    vaSpread1(0).Col = Col
    TipText = "Estructura Servicio : " & vaSpread1(0).text
Case 2
    vaSpread1(0).Col = Col
    TipText = "Código Ingrediente : " & Trim(vaSpread1(0).text)
Case 3
    vaSpread1(0).Col = Col
    TipText = "Descripción Ingrediente : " & Trim(vaSpread1(0).text)
Case 4
    vaSpread1(0).Col = Col
    TipText = "Tipo Ingrediente : " & Trim(vaSpread1(0).text)
Case 5
    vaSpread1(0).Col = Col
    TipText = "Frecuencia Ingrediente : " & Trim(vaSpread1(0).text)
Case 6
    vaSpread1(0).Col = Col
    TipText = "Código Producto : " & Trim(vaSpread1(0).text)
Case 7
    vaSpread1(0).Col = Col
    TipText = "Descripción Producto : " & Trim(vaSpread1(0).text)
Case 8
    vaSpread1(0).Col = Col
    TipText = "Tipo Ingrediente : " & Trim(vaSpread1(0).text)
Case 9
    vaSpread1(0).Col = Col
    TipText = "Precio : " & Trim(vaSpread1(0).text)
Case 10
    vaSpread1(0).Col = Col
    TipText = "Gramaje : " & Trim(vaSpread1(0).text)
End Select
End Sub
