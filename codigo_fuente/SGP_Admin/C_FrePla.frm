VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form C_FrePla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Frecuencia Recetas"
   ClientHeight    =   6840
   ClientLeft      =   1155
   ClientTop       =   2205
   ClientWidth     =   15585
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   15585
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1965
      Left            =   3000
      TabIndex        =   11
      Top             =   0
      Width           =   9375
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   270
         Left            =   7395
         TabIndex        =   12
         Top             =   1320
         Width           =   915
      End
      Begin EditLib.fpDateTime FpFecha 
         Height          =   315
         Left            =   2760
         TabIndex        =   13
         Top             =   1320
         Width           =   1260
         _Version        =   196608
         _ExtentX        =   2222
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
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
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   "03/2025"
         DateCalcMethod  =   4
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   0
         IncHour         =   0
         IncMinute       =   0
         IncSecond       =   0
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   2760
         TabIndex        =   21
         Top             =   960
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2760
         TabIndex        =   20
         Top             =   600
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   19
         Top             =   240
         Width           =   5535
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
         Left            =   1395
         TabIndex        =   18
         Top             =   300
         Width           =   1245
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
         Left            =   1395
         TabIndex        =   17
         Top             =   645
         Width           =   750
      End
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
         Left            =   1395
         TabIndex        =   16
         Top             =   1005
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Periodo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1410
         TabIndex        =   15
         Top             =   1380
         Width           =   675
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   7440
         TabIndex        =   14
         Top             =   1350
         Width           =   915
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2805
         TabIndex        =   22
         Top             =   285
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   2805
         TabIndex        =   23
         Top             =   645
         Width           =   5535
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2805
         TabIndex        =   24
         Top             =   1005
         Width           =   5535
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4695
      Left            =   0
      TabIndex        =   1
      Top             =   2100
      Width           =   14835
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   6240
         TabIndex        =   25
         Top             =   3960
         Width           =   3285
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   26
            Top             =   135
            Width           =   3180
         End
      End
      Begin VB.Frame Frame13 
         Height          =   435
         Left            =   600
         TabIndex        =   4
         Top             =   3960
         Width           =   675
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   570
         End
      End
      Begin VB.Frame Frame12 
         Height          =   435
         Left            =   1410
         TabIndex        =   2
         Top             =   3960
         Width           =   3285
         Begin VB.TextBox TextCai1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   3180
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3675
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   14505
         _Version        =   393216
         _ExtentX        =   25585
         _ExtentY        =   6482
         _StockProps     =   64
         ColsFrozen      =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   49
         MaxRows         =   18
         SpreadDesigner  =   "C_FrePla.frx":0000
         VisibleCols     =   4
         VisibleRows     =   18
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Total Recetas Listadas"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   11325
         TabIndex        =   10
         Top             =   4035
         Width           =   2055
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H00DEFEDE&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Costo Promedio Diario"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   11325
         TabIndex        =   9
         Top             =   4335
         Width           =   2055
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
         Left            =   13380
         TabIndex        =   8
         Top             =   4035
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
         Index           =   11
         Left            =   13380
         TabIndex        =   7
         Top             =   4335
         Width           =   975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6840
      Left            =   14955
      TabIndex        =   0
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12065
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_FrePla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private RS1         As New ADODB.Recordset
Private BtnX        As Variant
Private VarFecHasta As Long
Private VarFecDesde As Long
Private VarCencos   As Variant
Private VarCodReg   As Long
Private VarCodSer   As Long
Private VarTipMin   As String
Private VarTfor     As String
Private FecInicio   As Long
Private FecFin      As Long
Private Est         As Boolean

Private Sub CmdBuscar_Click()
    Let Label1(9).Caption = ""
    Let Label1(11).Caption = ""
    Call LlenarFrecPlan(VarTfor, VarCencos, VarCodReg, VarCodSer, VarFecDesde, VarTipMin, VarFecHasta)
End Sub

Private Sub Form_Activate()
    fg_descarga
End Sub

Private Sub Form_Load()
    Est = True
    Let FpFecha = Mid(VarFecDesde, 5, 2) & "/" & Mid(VarFecDesde, 1, 4)
    Let FecInicio = VarFecDesde
    Let FecFin = VarFecHasta
    Call fg_centra(Me)
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    Est = False
End Sub

Sub LlenarFrecPlan(tfor As String, subseg As Variant, codReg As Long, codser As Long, anomes As Long, TipMin As String, FecFin As Long)
Dim CodReceta           As Long
Dim IRow                As Long
Dim i                   As Long
Dim condia              As Long
Dim auxtip              As Long
Dim codfre              As Long
Dim cosreceta           As Double
Dim canreceta           As Double
Dim totgralreceta       As Double
Dim tippla              As String
Dim dia                 As Long
Dim Confre              As Long
Dim SearchFlagsEqual    As Variant
Dim ind_ini             As Long
Dim CodRec              As String
Dim X                   As Boolean
Dim vecTipoPla()        As Variant
Dim fecfin1              As Long
    Let VarTfor = tfor
    Let VarCencos = subseg
    Let VarCodReg = codReg
    Let VarCodSer = codser
    Let VarFecHasta = FecFin
    Let VarTipMin = TipMin
    Let VarFecDesde = anomes
    fg_carga ""
'-------> Rutina frecuencia de recetas
    Me.Caption = tfor
    MsgTitulo = tfor
    

    If VarSitioRemoto = False Then
        RS1.Open "SELECT sub_codigo, sub_nombre FROM a_subsegmento WHERE sub_codigo = " & subseg & "", vg_db, adOpenForwardOnly
        If Not RS1.EOF Then fpayuda(0).Caption = RS1!sub_nombre
        RS1.Close: Set RS1 = Nothing
        RS1.Open "SELECT reg_nombre FROM a_regimen WHERE reg_codigo = " & codReg & "", vg_db, adOpenForwardOnly
        If Not RS1.EOF Then fpayuda(1).Caption = RS1!reg_nombre
        RS1.Close: Set RS1 = Nothing
        RS1.Open "SELECT ser_nombre FROM a_servicio WHERE ser_codigo = " & codser & "", vg_db, adOpenForwardOnly
        If Not RS1.EOF Then fpayuda(3).Caption = RS1!ser_nombre
        RS1.Close: Set RS1 = Nothing
    Else
    
        Let Label1(0).Caption = "Cliente"
        Set RS1 = vg_db.Execute("SELECT cli_codigo, cli_nombre FROM b_clientes WHERE cli_codigo = '" & subseg & "' and cli_tipo = 0")
        If Not RS1.EOF Then fpayuda(0).Caption = Trim(RS1!Cli_nombre)
        RS1.Close: Set RS1 = Nothing
        Set RS1 = vg_db.Execute("SELECT reg_nombre FROM a_regimen WHERE reg_codigo = " & codReg & "")
        If Not RS1.EOF Then fpayuda(1).Caption = Trim(RS1!reg_nombre)
        RS1.Close: Set RS1 = Nothing
        Set RS1 = vg_db.Execute("SELECT ser_nombre FROM a_servicio WHERE ser_codigo = " & codser & "")
        If Not RS1.EOF Then fpayuda(3).Caption = Trim(RS1!ser_nombre)
        RS1.Close: Set RS1 = Nothing
    
    End If


' Control displays text tips aligned to pointer with focus
    vaSpread1(0).TextTip = 2
    vaSpread1(0).TextTipDelay = 250
    X = vaSpread1(0).SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    vaSpread1(0).Row = -1: vaSpread1(0).Col = 1
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1: vaSpread1(0).Col = 2
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1: vaSpread1(0).Col = 3
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1: vaSpread1(0).Col = 4
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1: vaSpread1(0).Col = 5
    vaSpread1(0).BackColor = &HDEFEDE
    vaSpread1(0).Row = -1: vaSpread1(0).Col = 6
    vaSpread1(0).BackColor = &HDEFEDE
    
    vaSpread1(0).MaxRows = 0
'-------> Buscar Nº días
'RS1.Open "SELECT DISTINCT b_minuta.min_fecmin from b_minuta, b_minutadet WHERE b_minuta.min_codigo=b_minutadet.mid_codigo " & _
'         "AND b_minuta.min_subseg=" & subseg & " AND b_minuta.min_codreg=" & codreg & " AND b_minuta.min_codser=" & codser & " " & _
'         "AND substring(convert(char(8),b_minuta.min_fecmin),1 ,6)=" & anomes & " AND b_minutadet.mid_tipmin='" & tipmin & "' ORDER By b_minuta.min_fecmin", vg_db, adOpenForwardOnly ', adOpenStatic
    If vg_Zona = "" Then
        vg_Zona = 0
    End If

    If VarSitioRemoto = False Then
        Set RS1 = vg_db.Execute("sgpadm_s_frecuenciaplan 3, " & subseg & ", " & codReg & ", " & codser & ", " & anomes & ", 0, 0, '1'," & vg_Zona & "," & vg_codlpr & "")
    Else
        Set RS1 = vg_db.Execute("sgpadm_s_frecuenciaplan 5, " & subseg & ", " & codReg & ", " & codser & ", " & Mid(anomes, 1, 6) & ", 0, " & FecFin & ", '1'," & vg_Zona & "," & vg_codlpr & "")
    End If
    
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: Exit Sub
    Select Case fg_Dia(RS1!min_fecmin)
        Case 1
            dia = 7
        Case 2
            dia = 1
        Case 3
            dia = 2
        Case 4
            dia = 3
        Case 5
            dia = 4
        Case 6
            dia = 5
        Case 7
            dia = 6
    End Select
    dia = dia - 1
    RS1.Close: Set RS1 = Nothing

    'MVA - MVI - ACA ES DONDE SE DEBE CAMBIAR PARA COSTEO DE MINUTA - 2013-01-18
    'trae el codigo de la lista de precios
    Dim Sql As String
    Sql = "select zon_codlpr from a_zona where zon_codigo = (select cli_codzon from b_clientes where cli_codigo = '" & subseg & "')"
    
    Set RS1 = vg_db.Execute(Sql)
    
    If Not RS1.EOF Then
        vg_codlpr = RS1!zon_codlpr
    End If
    'fin trae el codigo de la lista de precios
    
    If VarSitioRemoto = False Then
        Set RS1 = vg_db.Execute("sgpadm_s_frecuenciaplan 4, " & subseg & ", " & codReg & ", " & codser & ", " & anomes & ", 0, 0, '1'," & vg_Zona & "," & vg_codlpr & " ")
    Else
       Dim SeleccionOpt As Integer
       SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
       fecfin1 = Format(dEoM(fg_Ctod1(Format(FpFecha.text, "yyyymmdd"))), "yyyymmdd")
       Set RS1 = vg_db.Execute("sgpadm_Sel_FrecRecetaMinutaBloque '" & subseg & "', " & codReg & ", " & codser & ", " & Mid(anomes, 1, 6) & ", " & fecfin1 & ", '1', " & SeleccionOpt & "")
    End If
    
    DoEvents
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: Exit Sub
    CodReceta = 0: cosreceta = 0: canreceta = 0: totgralreceta = 0: condia = 0
    Dim essnom As String
    essnom = ""
    IRow = 1: Confre = 0
    'definir largo del vector

    ReDim Preserve vecTipoPla(1000, 3)
    Dim AuxCodigoEst As Long
    AuxCodigoEst = 0
    Do While Not RS1.EOF
        DoEvents
        If auxtip <> RS1!rec_tippla Or AuxCodigoEst <> RS1!ess_codigo Then
            If auxtip <> 0 Then
                vecTipoPla(IRow, 1) = tippla
                vecTipoPla(IRow, 2) = Confre
                vecTipoPla(IRow, 3) = essnom
                Confre = 0
                IRow = IRow + 1
            End If
            auxtip = RS1!rec_tippla
            tippla = RS1!nom_tippla
            essnom = RS1!ess_nombre
            AuxCodigoEst = RS1!ess_codigo
        End If
    
        vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
        vaSpread1(0).Row = vaSpread1(0).MaxRows
       
        vaSpread1(0).Col = 1
        vaSpread1(0).CellType = CellTypeStaticText
        vaSpread1(0).TypeHAlign = TypeHAlignLeft
        vaSpread1(0).text = RS1!pas_codrec
             
        vaSpread1(0).Col = 2
        vaSpread1(0).CellType = CellTypeStaticText
        vaSpread1(0).TypeHAlign = TypeHAlignLeft
        vaSpread1(0).text = Trim(RS1!rec_nombre)
       
        vaSpread1(0).Col = 3
        vaSpread1(0).CellType = CellTypeStaticText
        vaSpread1(0).TypeHAlign = TypeHAlignLeft
        vaSpread1(0).text = Trim(RS1!ess_nombre)
       
       ' --------------- Samuel Melendez 03/09/09 ----
        vaSpread1(0).Col = 4
        vaSpread1(0).CellType = CellTypeStaticText
        vaSpread1(0).TypeHAlign = TypeHAlignLeft
        vaSpread1(0).text = Trim(RS1!nom_tippla)
       ' ---------------------------------------------
             
        If VarSitioRemoto = False Then
            vaSpread1(0).Col = 5
            vaSpread1(0).CellType = CellTypeStaticText
            vaSpread1(0).TypeHAlign = TypeHAlignLeft
            vaSpread1(0).text = IIf(RS1!rec_indppr = 1, "Real", "Propuesta")
        End If
             
        vaSpread1(0).Col = 6
        vaSpread1(0).CellType = CellTypeStaticText
        vaSpread1(0).TypeHAlign = TypeHAlignRight
        vaSpread1(0).text = Format(RS1!nrorec, fg_Pict(6, 0))
        vaSpread1(0).ForeColor = &HFF0000
             
        vaSpread1(0).Col = 7
        vaSpread1(0).CellType = CellTypeStaticText
        vaSpread1(0).TypeHAlign = TypeHAlignRight
        vaSpread1(0).text = Format(RS1!rec_prerec, fg_Pict(6, 2))
        vaSpread1(0).ForeColor = &HFF0000
          
        Confre = Confre + RS1!nrorec
        RS1.MoveNext
    Loop

    If auxtip <> 0 Then
        vecTipoPla(IRow, 1) = tippla
        vecTipoPla(IRow, 2) = Confre
        vecTipoPla(IRow, 3) = essnom
    End If
    Confre = 0
    RS1.Close: Set RS1 = Nothing

    If VarSitioRemoto = False Then
       Set RS1 = vg_db.Execute("sgpadm_s_frecuenciaplan 2, " & subseg & ", " & codReg & ", " & codser & ", " & anomes & ", 0, 0, '1'," & vg_Zona & "," & vg_codlpr & " ")
    Else
       Set RS1 = vg_db.Execute("sgpadm_s_frecuenciaplan 7, " & subseg & ", " & codReg & ", " & codser & ", " & Mid(anomes, 1, 6) & ", 0," & FecFin & ", '1'," & vg_Zona & "," & vg_codlpr & " ")
    End If

    Do While Not RS1.EOF
        DoEvents
        CodRec = RS1!mid_codrec
        ind_ini = vaSpread1(0).SearchCol(1, -1, vaSpread1(0).MaxRows, CodRec, SearchFlagsEqual)
        If ind_ini <> -1 Then
            vaSpread1(0).Row = ind_ini
            vaSpread1(0).Col = 7 + (dia + Val(Mid(RS1!min_fecmin, 7, 2)))
            vaSpread1(0).CellType = CellTypeStaticText
            vaSpread1(0).TypeHAlign = TypeHAlignRight
            vaSpread1(0).text = 0
            vaSpread1(0).text = CCur(Val(vaSpread1(0).text) + RS1!mid_numrac) '"X"
            vaSpread1(0).ForeColor = &HFF0000
        End If
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    cosreceta = 0: canreceta = 0: totgralreceta = 0

    For i = 2 To vaSpread1(0).MaxRows
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 6
        canreceta = Val(vaSpread1(0).text)
        vaSpread1(0).Col = 7
        cosreceta = Val(vaSpread1(0).text)
        totgralreceta = CCur(totgralreceta + (cosreceta * canreceta))
    Next i

    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 2
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    vaSpread1(0).Col = 2
    vaSpread1(0).text = "TOTALES X TIPO PLATO"
    vaSpread1(0).Font.Bold = True
    vaSpread1(0).Font.Size = 9
    For i = 1 To IRow
        vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
        vaSpread1(0).Row = vaSpread1(0).MaxRows
        vaSpread1(0).Col = 3
        vaSpread1(0).text = vecTipoPla(i, 3)
        vaSpread1(0).Font.Bold = True
        vaSpread1(0).Font.Size = 9
        
        vaSpread1(0).Col = 4
        vaSpread1(0).text = vecTipoPla(i, 1)
        vaSpread1(0).Font.Bold = True
        vaSpread1(0).Font.Size = 9
        
        vaSpread1(0).Col = 6
        vaSpread1(0).text = vecTipoPla(i, 2)
        vaSpread1(0).Font.Bold = True
        vaSpread1(0).Font.Size = 9
    Next i
    Label1(9).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(6, 2))
    Label1(11).Caption = Format(totgralreceta, fg_Pict(6, 2))

    If VarSitioRemoto = False Then
        RS1.Open "SELECT COUNT(b_minuta.min_codigo) AS nreg " & _
            " FROM b_minuta " & _
            " WHERE b_minuta.min_codigo IN (SELECT b_minutadet.mid_codigo FROM b_minutadet WHERE b_minutadet.mid_tipmin = '" & TipMin & "') " & _
            " AND b_minuta.min_subseg = " & subseg & " " & _
            " AND b_minuta.min_codreg = " & codReg & _
            " AND b_minuta.min_codser = " & codser & _
            " AND substring(convert(char(8),b_minuta.min_fecmin),1,6) = " & anomes & "", vg_db, adOpenForwardOnly ', adOpenStatic
    Else
        Set RS1 = vg_db.Execute("SELECT COUNT(a.min_codigo) AS nreg " & _
            " FROM cas_b_minuta as a  WITH ( NOLOCK ) " & _
            " WHERE a.min_codigo IN (SELECT b.mid_codigo FROM cas_b_minutadet as b  WITH ( NOLOCK ) WHERE b.mid_tipmin = '" & TipMin & "' and  b.mid_cecori = '" & subseg & "') " & _
            " and a.min_cecori = '" & subseg & "' AND a.min_codreg = " & codReg & _
            " AND a.min_codser = " & codser & _
            " AND substring(convert(char(8),a.min_fecmin),1,6) = " & Mid(anomes, 1, 6) & "")
    End If
    
    If Not RS1.EOF And RS1!nReg > 0 Then Label1(11).Caption = Format(CCur(totgralreceta / RS1!nReg), fg_Pict(6, 2))
    RS1.Close: Set RS1 = Nothing
    fg_descarga
End Sub

Private Sub FpFecha_Change()
Dim Fecha As Variant
If Est Then Exit Sub
'    If CDate("01/" & FpFecha) < CDate("01/" & Mid(FecInicio, 5, 2) & "/" & Mid(FecInicio, 1, 4)) Then
'        Let FpFecha = Mid(FecInicio, 5, 2) & "/" & Mid(FecInicio, 1, 4)
'    ElseIf CDate("01/" & FpFecha) > CDate("01/" & Mid(FecInicio, 5, 2) & "/" & Mid(FecInicio, 1, 4)) Then
'        Let FpFecha = Mid(FecInicio, 5, 2) & "/" & Mid(FecInicio, 1, 4)
'    End If
    Let Fecha = Mid(FpFecha, 4, 4) & Mid(FpFecha, 1, 2)
    Let VarFecDesde = CLng(Fecha)
    vaSpread1(0).MaxRows = 0
End Sub

Private Sub TextCai1_Change(Index As Integer)
Dim i As Long
Dim indactivo As Integer
Dim nom As String
Dim icol As Long
icol = IIf(Index = 1, 1, IIf(Index = 2, 2, 4))
Select Case Index
Case 1, 2, 0
    vaSpread1(0).Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       For i = 1 To vaSpread1(0).MaxRows
           If i = 430 Then
              nom = 1
           End If
           vaSpread1(0).Row = i
           vaSpread1(0).Col = icol: nom = UCase(Trim(vaSpread1(0).text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
           vaSpread1(0).Col = icol
           If indactivo = -1 And Trim(vaSpread1(0).text) <> "" Then
              If vaSpread1(0).RowHidden = True Then vaSpread1(0).RowHidden = False
           Else
              If vaSpread1(0).RowHidden = False Then vaSpread1(0).RowHidden = True
           End If
        Next i
        vaSpread1(0).SetActiveCell Index, 1
    End If
    vaSpread1(0).ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1(0).ColUserSortIndicator(IIf(Trim(TextCai1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1(0).SortKey(1) = IIf(Trim(TextCai1(Index).text) = "", 0, 0): vaSpread1(0).SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1(0).Sort -1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows, SortByRow
    If Trim(TextCai1(Index).text) = "" Then
       For i = 1 To vaSpread1(0).MaxRows
           vaSpread1(0).Row = i
           If vaSpread1(0).RowHidden = True Then vaSpread1(0).RowHidden = False
       Next
       vaSpread1(0).SetActiveCell Index, vaSpread1(0).SearchCol(Index, 0, vaSpread1(0).MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(0).SetActiveCell Index, 1
    End If
    vaSpread1(0).Visible = True
End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    ExportarExcel
Case 4
    Me.Hide
    Unload Me
End Select
End Sub

Sub ExportarExcel()
Dim NashXl As excel.Application
Dim IRow As Long, irow2 As Long
fg_carga ""
Set NashXl = CreateObject("excel.application")
Set NashXl = New excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add

vaSpread1(0).AllowMultiBlocks = True
vaSpread1(0).SetSelection 1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
vaSpread1(0).ClipboardCopy
IRow = vaSpread1(0).MaxRows + 1
'------- Pegar vaspread1(1) - Planilla Excel
NashXl.Range("A1").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'NashXl.Range("A1:D" & irow).Select
'With NashXl.Selection.Interior
'     .ColorIndex = 36
'     .Pattern = xlSolid
'End With
'------- Colorear titulo
NashXl.Range("A1:AW1").Select ' samuel 03/0309
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A1:AW" & IRow).Select ' samuel 03/09/09
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
NashXl.Range("D2" & ":" & "D" & IRow).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Aplicar totales

'------- Dibujar marco
IRow = IRow + 2
irow2 = IRow + 2
NashXl.Range("B" & IRow).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(8).Caption
NashXl.Range("C" & IRow).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(9).Caption
NashXl.Range("B" & irow2).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(10).Caption
NashXl.Range("C" & irow2).Select
NashXl.ActiveCell.FormulaR1C1 = Label1(11).Caption
NashXl.Range("B" & IRow & ":" & "C" & irow2).Select
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
NashXl.Selection.Font.Bold = True
'With NashXl.Selection.Interior
'     .ColorIndex = 35
'     .Pattern = xlSolid
'End With
NashXl.Range("D" & IRow & ":" & "D" & irow2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1(0).AllowMultiBlocks = False: vaSpread1(0).SetSelection 1, 0, vaSpread1(0).maxcols, vaSpread1(0).MaxRows
fg_descarga
NashXl.Visible = True
End Sub

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
TipText = "Código : " & vaSpread1(0).text
Case 2
    vaSpread1(0).Col = Col
    TipText = "Nombre Receta : " & Trim(vaSpread1(0).text)
Case 3
    vaSpread1(0).Col = Col
    TipText = "Tipo Plato : " & Trim(vaSpread1(0).text)
Case 4
    vaSpread1(0).Col = Col
    TipText = "Tipo Receta : " & Trim(vaSpread1(0).text)
Case 5
    vaSpread1(0).Col = Col
    TipText = "Frecuencia : " & Trim(vaSpread1(0).text)
Case 6
    vaSpread1(0).Col = Col
    TipText = "Costo : " & Trim(vaSpread1(0).text)
End Select
End Sub




