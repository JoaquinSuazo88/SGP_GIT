VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_GenLpr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación archivos planos lista precio sac"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   6000
      Top             =   8280
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7395
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   12345
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1695
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   240
         Width           =   12165
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   3
            Left            =   6480
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   1200
            Width           =   5415
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   5640
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   1200
            Width           =   855
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   4500
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1200
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   480
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1200
            Width           =   3975
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1560
            TabIndex        =   0
            Top             =   285
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
            Text            =   "05/2010"
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
            Left            =   3240
            TabIndex        =   1
            Top             =   240
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
            Caption         =   "Periodo Negociado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4845
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   12135
         _Version        =   393216
         _ExtentX        =   21405
         _ExtentY        =   8546
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   7
         MaxRows         =   1
         SpreadDesigner  =   "M_GenLpr.frx":0000
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   7140
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enviados"
         Height          =   195
         Index           =   1
         Left            =   9600
         TabIndex        =   13
         Top             =   7080
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   9240
         Top             =   7110
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No Enviados"
         Height          =   195
         Index           =   0
         Left            =   10815
         TabIndex        =   12
         Top             =   7080
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   10455
         Top             =   7110
         Width           =   300
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   12540
      _ExtentX        =   22119
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1680
      Top             =   8040
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
            Picture         =   "M_GenLpr.frx":052D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   0
      OleObjectBlob   =   "M_GenLpr.frx":08C7
      Top             =   8190
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   600
      OleObjectBlob   =   "M_GenLpr.frx":08EB
      Top             =   8160
   End
End
Attribute VB_Name = "M_GenLpr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Msgtitulo  As String
Dim Est As Boolean

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
Dim cdcen As String, cdfil As String, nrosem As Long, periodo As String
Frame1(0).Enabled = False
Select Case Index
Case 0 '-------> Central de compras
    fg_carga ""
    Est = True
    cdcen = Trim(fg_codigocbo(Combo1, 0, 4, ""))
    Combo1(1).Clear
    Combo1(1).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 6) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CICCPA_DTREF FROM VW_SGP_LISTAPRECIO WHERE (TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') AND CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CICCPA_DTREF")
    Do While Not RS.EOF
       DoEvents
       Combo1(1).AddItem Mid(RS!CICCPA_DTREF, 5, 2) & "/" & Mid(RS!CICCPA_DTREF, 1, 4) & Space(150) & "(" & fg_pone_espacio((RS!CICCPA_DTREF), 6) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(1).ListIndex = 0
    
    Combo1(2).Clear
    Combo1(2).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(("0"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CICCPA_NRSEM FROM VW_SGP_LISTAPRECIO WHERE (TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') AND CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CICCPA_NRSEM")
    Do While Not RS.EOF
       DoEvents
       Combo1(2).AddItem RS!CICCPA_NRSEM & Space(150) & "(" & fg_pone_cero(Str(RS!CICCPA_NRSEM), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(2).ListIndex = 0
    
    Combo1(3).Clear
    Combo1(3).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CADFIL_CDFIL, CADFIL_NMFIL FROM VW_SGP_LISTAPRECIO WHERE (TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') AND CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CADFIL_NMFIL")
    Do While Not RS.EOF
       DoEvents
       Combo1(3).AddItem RS!CADFIL_NMFIL & Space(150) & "(" & fg_pone_espacio((RS!CADFIL_CDFIL), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(3).ListIndex = 0
    Est = False
    fg_descarga
    MoverDatoGrilla
Case 1 '-------> Periodo
    fg_carga ""
    Est = True
    cdcen = Trim(fg_codigocbo(Combo1, 0, 4, ""))
    periodo = Trim(fg_codigocbo(Combo1, 1, 6, ""))
    Combo1(2).Clear
    Combo1(2).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(("0"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CICCPA_NRSEM FROM VW_SGP_LISTAPRECIO WHERE (TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') AND CICCPA_DTREF = '" & periodo & "' AND CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CICCPA_NRSEM")
    Do While Not RS.EOF
       DoEvents
       Combo1(2).AddItem RS!CICCPA_NRSEM & Space(150) & "(" & fg_pone_cero(Str(RS!CICCPA_NRSEM), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(2).ListIndex = 0
    
    Combo1(3).Clear
    Combo1(3).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CADFIL_CDFIL, CADFIL_NMFIL FROM VW_SGP_LISTAPRECIO WHERE (TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') AND CICCPA_DTREF = '" & periodo & "' AND CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CADFIL_NMFIL")
    Do While Not RS.EOF
       DoEvents
       Combo1(3).AddItem RS!CADFIL_NMFIL & Space(150) & "(" & fg_pone_espacio((RS!CADFIL_CDFIL), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(3).ListIndex = 0
    Est = False
    fg_descarga
    MoverDatoGrilla
Case 2 '-------> Numero semana
    fg_carga ""
    Est = True
    cdcen = Trim(fg_codigocbo(Combo1, 0, 4, ""))
    periodo = Trim(fg_codigocbo(Combo1, 1, 6, ""))
    nrosem = Trim(fg_codigocbo(Combo1, 2, 10, ""))
    Combo1(3).Clear
    Combo1(3).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CADFIL_CDFIL, CADFIL_NMFIL FROM VW_SGP_LISTAPRECIO WHERE (TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') AND CICCPA_DTREF = '" & periodo & "' AND CICCPA_NRSEM = " & nrosem & " AND CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CADFIL_NMFIL")
    Do While Not RS.EOF
       DoEvents
       Combo1(3).AddItem RS!CADFIL_NMFIL & Space(150) & "(" & fg_pone_espacio((RS!CADFIL_CDFIL), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(3).ListIndex = 0
    Est = False
    fg_descarga
    MoverDatoGrilla
Case 3 '-------> Centro costo
    MoverDatoGrilla
End Select
Frame1(0).Enabled = True
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 9135
Me.Width = 12630
Me.HelpContextID = vg_OpcM
Msgtitulo = "Generación archivos planos lista precio sac"
fg_centra Me
Est = True
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.ToolTipText = "Enviar": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpDateTime1(0).text = Format(Date, "mm/yyyy")
Est = False
SendKeys "+{Tab}"
End Sub

Private Sub MoverDatoGrilla()
Dim cdcen As String, cdfil As String, nrosem As Long, dtref As String
fg_carga ""
cdcen = ""
cdcen = Trim(fg_codigocbo(Combo1, 0, 4, ""))
dtref = Trim(fg_codigocbo(Combo1, 1, 6, ""))
cdfil = Trim(fg_codigocbo(Combo1, 3, 10, ""))
nrosem = fg_codigocbo(Combo1, 2, 10, 0)
'vaSpread1.Visible = False
vaSpread1.MaxRows = 0
Set RS = vg_dbsac.Execute("SELECT DISTINCT a.TABCEN_CDCEN, b.TABCEN_DSCEN, a.CADFIL_CDFIL, a.CADFIL_NMFIL, a.CICCPA_DTREF, a.CICCPA_NRSEM " & _
                          "FROM VW_SGP_LISTAPRECIO a, TABCEN b " & _
                          "WHERE a.TABCEN_CDCEN = b.TABCEN_CDCEN " & _
                          "AND  (a.TABCEN_CDCEN = '" & cdcen & "' OR '" & cdcen & "' = '*') " & _
                          "AND  (a.CADFIL_CDFIL = '" & cdfil & "' OR '" & cdfil & "' = '*') " & _
                          "AND  (a.CICCPA_DTREF = '" & dtref & "' OR '" & dtref & "' = '*') " & _
                          "AND  (a.CICCPA_NRSEM = " & nrosem & "  OR " & nrosem & " = 0) " & _
                          "AND   a.CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' " & _
                          "AND   a.CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' " & _
                          "ORDER BY a.TABCEN_CDCEN, a.CADFIL_CDFIL, a.CADFIL_NMFIL, a.CICCPA_DTREF, a.CICCPA_NRSEM")
DoEvents
Do While Not RS.EOF
   DoEvents
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1: vaSpread1.text = "0"
   vaSpread1.Col = 2: vaSpread1.text = Trim(RS!TABCEN_CDCEN)
   vaSpread1.Col = 3: vaSpread1.text = Trim(RS!TABCEN_DSCEN)
   vaSpread1.Col = 4: vaSpread1.text = Trim(RS!CICCPA_DTREF)
   vaSpread1.Col = 5: vaSpread1.text = Trim(RS!CICCPA_NRSEM)
   vaSpread1.Col = 6: vaSpread1.text = Trim(RS!CADFIL_CDFIL)
   vaSpread1.Col = 7: vaSpread1.text = Trim(RS!CADFIL_NMFIL)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
'vaSpread1.Visible = True
fg_descarga
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).text) = False Then Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 '-------> Envio datos sgp contrato
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Dim i As Long, aAp As String, tlistpre As String, tprod As String, dBo As String, cDBI As String, dbosac As String
    Dim isel As Boolean
    Dim cdcen As String, cdfil As String, dtref As String, nrsem As Long
    Dim cencos As String, nomcencos As String, codpro As String, sourcefile As String, sourcefilezip As String, destinofile As String, destinofilezip As String, mdirserver As String, lognarchsou As String, socsap As String
    Dim fso, codtis As Long, CodSeg As Long
    Dim CHost As String, Cdire As String, Cuser As String, Cpass As String, Cpuer As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    isel = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Toolbar1.Enabled = False
    Frame1(0).Enabled = False
    fg_carga ""
    tlistpre = "": tprod = ""
    '-------> Generar tabla temporal productos
    vg_db.BeginTrans
    aAp = Trim(vg_NUsr) & "_tmp_GenPlanoProdSac"
    fg_CheckTmp aAp: tprod = aAp
    vg_db.Execute "CREATE TABLE " & aAp & " (pro_codigo varchar(20))"
    vg_db.Execute "INSERT INTO " & aAp & " SELECT DISTINCT pro_codigo FROM b_productos"
    vg_db.CommitTrans
       
    vg_db.BeginTrans
    '-------> Generar tabla temporal lista precio sac
    aAp = Trim(vg_NUsr) & "_tmp_GenPlanoListPre"
    fg_CheckTmp aAp: tlistpre = aAp
    vg_db.Execute "CREATE TABLE " & aAp & " (tabcen_cdcen varchar(04), cadfil_cdfil varchar(10), ciccpa_dtref varchar(06), ciccpa_nrsem int)"
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           vaSpread1.Col = 2: cdcen = vaSpread1.text
           vaSpread1.Col = 4: dtref = vaSpread1.text
           vaSpread1.Col = 5: nrsem = vaSpread1.text
           vaSpread1.Col = 6: cdfil = vaSpread1.text
           vg_db.Execute "INSERT INTO " & aAp & " (tabcen_cdcen, cadfil_cdfil, ciccpa_dtref, ciccpa_nrsem) values ('" & cdcen & "', '" & cdfil & "', '" & dtref & "', " & nrsem & ")"
        End If
    Next i
    vg_db.CommitTrans
    '-------> Crear directorio servidor SQLDES
    mdirserver = Dir(dir_trabajo & "\" & "Actualizar", vbDirectory)
    If mdirserver = "" Then MkDir dir_trabajo & "\" & "Actualizar"
    mdirserver = dir_trabajo & "Actualizar" & "\"
    '-------> Fin crear directorio servidor SQLDES
    
    '-------> Generar base padre Access
    sourcefile = "listapreciogeneral" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
    If Dir(mdirpc & sourcefile) <> "" Then Kill mdirpc & sourcefile ' borrar base datos si existe
    
    '-------> base de datos origen
    ' Rutas base de datos access    dBo = dir_trabajo + BaseDeDato
    dBo = "'' [ODBC;Driver={SQL Server};Server=" + vg_SqlNSvr + ";Database=" + vg_SqlBase + ";UID=" + vg_SqlNUsr + ";PWD=" + vg_SqlPass + "]"
    dbosac = "'' [ODBC;Driver={Microsoft ODBC for Oracle};SERVER=" + vgsac_NSvr + ";uid=" + vgsac_NUsr + ";pwd=" + vgsac_Pass + "]"
    GenerarBaseEnviado mdirpc & sourcefile, tprod, tlistpre, dBo, 4, 0, 0, 0
    Bar1(0).Visible = True
    Bar1(0).Value = 0: icopy = False
    For i = 1 To vaSpread1.MaxRows
        DoEvents
        vaSpread1.Row = i
        vaSpread1.Col = 1
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        If vaSpread1.text = "1" Then
           vaSpread1.SetActiveCell 2, vaSpread1.Row
           vaSpread1.Col = 2: cdcen = "": cdcen = vaSpread1.text
           vaSpread1.Col = 4: dtref = "": dtref = vaSpread1.text
           vaSpread1.Col = 5: nrsem = 0: nrsem = vaSpread1.text
           vaSpread1.Col = 6: cdfil = "": cdfil = vaSpread1.text
           vaSpread1.Col = 6: cencos = "": cencos = vaSpread1.text
           '-------> Crear archivos *.MDB y *.ZIP
           destinofile = "sgp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
           destinofilezip = "sgp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           '-------> verificar si existe archivo mdb destino si existe borrar y copiar
           DoEvents
           If Dir(mdirpc & destinofile) <> "" Then Kill mdirpc & destinofile
           FileCopy mdirpc & sourcefile, mdirpc & destinofile
           '---------------------------
           '------- Abrir base contrato
           '---------------------------
           cDBI = mdirpc & destinofile
           Set dbi = New ADODB.Connection
           dbi.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cDBI & "' ;Persist Security Info=False"
           dbi.ConnectionTimeout = 3600
           dbi.CommandTimeout = 3600
           dbi.Open
           '-------> Insert tabla lista precio sac
'           dbi.Execute "INSERT INTO b_sac_listaprecio (lps_cencos, lps_fecini, lps_fecfin, lps_periodo, lps_codsac, lps_precio) SELECT DISTINCT CADFIL_CDFIL, CICCOT_DTPDE, CICCOT_DTPAT, CICCPA_DTREF, CPOPRO_CDPRO, FORPRO_VLPCO FROM vw_sgp_listaprecio IN " & dbosac & " WHERE tabcen_cdcen = '" & cdcen & "' AND cadfil_cdfil = '" & cdfil & "' AND ciccpa_dtref = '" & dtref & "' AND ciccpa_nrsem = " & nrsem & ""
           dbi.Execute "INSERT INTO b_sac_listaprecio (lps_cencos, lps_periodo, lps_codsac, lps_precio) " & _
                       "SELECT DISTINCT CADFIL_CDFIL, '" & Format(fpDateTime1(0).text, "yyyymm") & "', CPOPRO_CDPRO, FORPRO_VLPCO " & _
                       "FROM vw_sgp_listaprecio IN " & dbosac & " " & _
                       "WHERE TABCEN_CDCEN = '" & cdcen & "' " & _
                       "AND   CADFIL_CDFIL = '" & cdfil & "' " & _
                       "AND   CICCPA_DTREF = '" & dtref & "' " & _
                       "AND   CICCPA_NRSEM = " & nrsem & ""
           '-------> Insert tabla lista precio sac
'           dbi.Execute "INSERT INTO b_formatocompras (foc_codsac, foc_codcat, foc_nomsac, foc_unisac, foc_vigini, foc_flexec, foc_vigfin) SELECT DISTINCT a.foc_codsac, a.foc_codcat, a.foc_nomsac, a.foc_unisac, a.foc_vigini, a.foc_flexec, a.foc_vigfin FROM b_formatocompras a IN " & dBo & ", b_sac_listaprecio b WHERE a.foc_codsac = b.cpopro_codpro AND b.tabcen_cdcen = '" & cdcen & "' AND b.cadfil_cdfil = '" & cdfil & "' AND b.ciccpa_dtref = '" & dtref & "' AND b.ciccpa_nrsem = " & nrsem & ""
           dbi.Execute "INSERT INTO b_formatocompras (foc_codsac, foc_codcat, foc_nomsac, foc_unisac, foc_vigini, foc_flexec, foc_vigfin) SELECT DISTINCT foc_codsac, foc_codcat, foc_nomsac, foc_unisac, foc_vigini, foc_flexec, foc_vigfin FROM b_formatocompras IN " & dBo & " WHERE foc_codsac IN (SELECT DISTINCT lps_codsac FROM b_sac_listaprecio WHERE lps_cencos = '" & cdfil & "')" ' AND lps_periodo = '" & dtref & "')"
'           dbi.Execute "INSERT INTO b_formatocomprassgp (fcs_codsac, fcs_codsgp, fcs_sgppre) SELECT DISTINCT a.fcs_codsac, a.fcs_codsgp, a.fcs_sgppre FROM b_formatocomprassgp a, b_formatocompras b, b_saclistaprecio c IN " & dBo & " WHERE a.fcs_sgppre = b.foc_codsac AND b.foc_codsac = c.cpopro_codpro AND c.tabcen_cdcen = '" & cdcen & "' AND c.cadfil_cdfil = '" & cdfil & "' AND c.ciccpa_dtref = '" & dtref & "' AND c.ciccpa_nrsem = " & nrsem & ""
           dbi.Execute "INSERT INTO b_formatocomprassgp (fcs_codsac, fcs_codsgp, fcs_sgppre) SELECT DISTINCT a.fcs_codsac, a.fcs_codsgp, a.fcs_sgppre FROM b_formatocomprassgp a, b_formatocompras b IN " & dBo & " WHERE a.fcs_codsac = b.foc_codsac AND b.foc_codsac IN (SELECT DISTINCT lps_codsac FROM b_sac_listaprecio WHERE lps_cencos = '" & cdfil & "')" ' AND lps_periodo = '" & dtref & "')"
           '-------> Generar parametros ejecutivos contables
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'datcont', mid(cli_nomcontable,1,40), 'C', cli_emailcontable FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT '5etapas', 'Casino 5 Etapas', 'C', iif(cli_subseg=0,'N','S') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT par_codigo, par_nombre, par_tipo, par_valor FROM a_param IN " & dBo & " WHERE par_codigo='porprepro'"
           '-------> Insert concepto grupo vulnerable a tabla a_param.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'opgruvul', 'Opción Grupo Vulnerable', 'C', iif(cli_gruvul='S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
           '-------> Insert concepto modulo paciente.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modpac', 'Modulo Paciente', 'C', iif(cli_modpac='S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
           '-------> Insert concepto parametro proveedor
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modprove', 'Parametro Modificar Proveedor', 'N', '0' FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
           codtis = 0: CodSeg = 0: socsap = ""
           RS.Open "SELECT * FROM b_clientes WHERE cli_codigo='" & cencos & "'", vg_db, adOpenForwardOnly
           If Not RS.EOF Then
              codtis = IIf(IsNull(RS!cli_codtis), 0, RS!cli_codtis)
              CodSeg = IIf(IsNull(RS!cli_codseg), 0, RS!cli_codseg)
              socsap = IIf(IsNull(RS!cli_socsap), "", RS!cli_socsap)
           End If
           RS.Close: Set RS = Nothing
           '-------> Borrar tabla tipo servicio y segmento que no tenga relación con el contrato
           dbi.Execute "DELETE a_tiposervicio FROM a_tiposervicio WHERE tis_codigo NOT IN (" & codtis & ")"
           dbi.Execute "DELETE a_segmento FROM a_segmento WHERE seg_codigo NOT IN (" & CodSeg & ")"

           '-------> Borrar tabla casino envia sap
           dbi.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos NOT IN ('" & cencos & "')"

           '-------> Mover datos a la tabla centro de costo
           dbi.Execute "INSERT INTO a_cencos (cen_codigo, cen_socsap) VALUES ('" & cencos & "', '" & socsap & "')"
           '-------> Cerrar base access
           dbi.Close: Set dbi = Nothing
           DoEvents
           
           '-------> verificar si existe archivo zip destino si existe borrar
           If Dir(mdirpc & destinofilezip) <> "" Then Kill mdirpc & destinofilezip
           AZ1.CreateZip mdirpc & destinofilezip, "": AZ1.AddFile mdirpc & destinofile, "", True, "": AZ1.Close
           '-------> verificar si existe archivo mdb destino si existe borrar
           If Dir(mdirpc & destinofile) <> "" Then Kill mdirpc & destinofile
           '-------> leer casino
           DoEvents
           RS.Open "SELECT * FROM b_clientes WHERE cli_codigo='" & cencos & "'", vg_db, adOpenForwardOnly
           If Not RS.EOF Then
              If RS!cli_openvio = 1 Then
                 '-------> Traer datos FTP
                 Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("ftp")) & "%'")
                 If RS1.EOF Then fg_descarga: RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: Frame1(0).Enabled = True: Frame1(1).Enabled = True: Bar1(0).Visible = False: Bar1(1).Visible = False: MsgBox "No existe Parametrización FTP, Proceso cancelado, contactese con departamento informactica ", vbCritical, Msgtitulo: Exit Sub
                 Do While Not RS1.EOF
                    If RS1!par_codigo = "ftpser" Then CHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftpdir" Then Cdire = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftpusu" Then Cuser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftppas" Then Cpass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftppue" Then Cpuer = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    RS1.MoveNext
                 Loop
                 RS1.Close: Set RS1 = Nothing
'                 Open dir_trabajo & "\sdxftp.ini" For Input As #1
'                 Do While Not EOF(1)
'                    Line Input #1, cpars
'                    If Mid(cpars, 1, InStr(cpars, ",") - 1) = "A" Then
'                       CHost = Mid(cpars, InStr(cpars, ",") + 1)
'                    ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "B" Then
'                       Cuser = Mid(cpars, InStr(cpars, ",") + 1)
'                    ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "C" Then
'                       Cpass = Mid(cpars, InStr(cpars, ",") + 1)
'                    End If
'                 Loop
'                 Close #1
                 a = oFTP.Version
                 oFTP.UseIEProxy = False
                 oFTP.Port = Cpuer '21
                 oFTP.HostName = CHost '"sgp.sodexhochile.cl"
                 oFTP.UserName = Cuser '"userftp"
                 oFTP.password = Cpass '"*sdxo7528*"
                 oFTP.Connect
                 If oFTP.IsConnected Then
                     lDir = oFTP.GetCurrentDirListing("*.*")
                     oFTP.SaveLastError ("aaa.xml")
'                     a = oFTP.ChangeRemoteDir("/casinos/bd")
                     a = oFTP.ChangeRemoteDir(Cdire)
                     oFTP.SaveLastError ("aaa.xml")
                     lDir = oFTP.GetCurrentDirListing("*.*")
                     oFTP.SaveLastError ("aaa.xml")
                     a = oFTP.PutFile(mdirpc & destinofilezip, destinofilezip)
                     oFTP.SaveLastError ("aaa.xml")
                     oFTP.Disconnect
                     If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                        fg_descarga
                        MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no se puede notificar por correo, ya que no tiene asignado el mail", vbInformation + vbOKOnly, Msgtitulo
                        fg_carga ""
                     Else
                        SendMail1 oMail, "Se Informa que el maestro de productos esta disponible. Para que usted pueda actualizar ", "Se Informa que el maestro de productos esta disponible. Para que usted pueda actualizar", mdirpc & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0
                     End If
                 End If
              ElseIf RS!cli_openvio = 2 Then
                 If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                    fg_descarga
                    MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, Msgtitulo
                    fg_carga ""
                 Else
                    SendMail1 oMail, "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), mdirpc & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1
                 End If
              End If
           End If
           RS.Close: Set RS = Nothing
           '-------> Grabar tabla b_minutacasino de enviado la información
           vg_db.BeginTrans
           vg_db.Execute "DELETE b_listapreciocasino FROM b_listapreciocasino WHERE lpc_cencos='" & cencos & "' AND lpc_periodo = '" & dtref & "' AND lpc_nrosem = " & nrsem & ""
           vg_db.Execute "INSERT INTO b_listapreciocasino VALUES ('" & cencos & "', '" & dtref & "', " & nrsem & ", '" & Format(Date, "yyyymmdd") & "')"
           vg_db.CommitTrans
           DoEvents
        End If
    Next i
    Frame1(0).Enabled = True
    Toolbar1.Enabled = True
    Bar1(0).Value = 0: Bar1(0).Visible = False
    If Trim(sourcefile) <> "" Then MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    fg_descarga
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim auxcen As String
Select Case Button.Index
Case 1 '-------> Cargar datos grilla
    fg_carga ""
    Frame1(0).Enabled = False
    AbrirBaseSac
    vaSpread1.MaxRows = 0
    Set RS = vg_dbsac.Execute("SELECT DISTINCT a.TABCEN_CDCEN, b.TABCEN_DSCEN, a.CADFIL_CDFIL, a.CADFIL_NMFIL, a.CICCPA_DTREF, a.CICCPA_NRSEM " & _
             "FROM VW_SGP_LISTAPRECIO a, TABCEN b " & _
             "WHERE a.TABCEN_CDCEN = b.TABCEN_CDCEN " & _
             "AND   a.CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' " & _
             "AND   a.CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' " & _
             "ORDER BY a.TABCEN_CDCEN, a.CADFIL_CDFIL, a.CADFIL_NMFIL, a.CICCPA_DTREF, a.CICCPA_NRSEM")
'    vaSpread1.Visible = False
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
    RS.MoveFirst
    DoEvents
    Do While Not RS.EOF
       DoEvents
'       If RS!CADFIL_CDFIL <> auxcen Then
          Set RS1 = vg_db.Execute("sgpadm_s_enviolistaprecio '" & RS!CADFIL_CDFIL & "', '" & RS!CICCPA_DTREF & "', " & RS!CICCPA_NRSEM & "")
          If Not RS1.EOF Then
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.Row = vaSpread1.MaxRows
             vaSpread1.Col = 1: vaSpread1.text = "0"
             vaSpread1.Col = 2: vaSpread1.text = Trim(RS!TABCEN_CDCEN)
             vaSpread1.Col = 3: vaSpread1.text = Trim(RS!TABCEN_DSCEN)
             vaSpread1.Col = 4: vaSpread1.text = Trim(RS!CICCPA_DTREF)
             vaSpread1.Col = 5: vaSpread1.text = Trim(RS!CICCPA_NRSEM)
             vaSpread1.Col = 6: vaSpread1.text = Trim(RS!CADFIL_CDFIL)
             vaSpread1.Col = 7: vaSpread1.text = Trim(RS!CADFIL_NMFIL)
             If RS1!cli_envio = "1" Then vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
          End If
          RS1.Close: Set RS1 = Nothing
          auxcen = RS!CADFIL_CDFIL
'       End If
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    '-------> Cargar tabla central compras
    Est = True
    Combo1(0).Clear
    Combo1(0).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 4) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT a.TABCEN_CDCEN, b.TABCEN_DSCEN " & _
             "FROM VW_SGP_LISTAPRECIO a, TABCEN b " & _
             "WHERE a.TABCEN_CDCEN = b.TABCEN_CDCEN " & _
             "AND   a.CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' " & _
             "AND   a.CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "'")
    DoEvents
    Do While Not RS.EOF
       DoEvents
       Combo1(0).AddItem RS!TABCEN_DSCEN & Space(150) & "(" & fg_pone_espacio((RS!TABCEN_CDCEN), 4) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(0).ListIndex = 0
    
    '-------> Cargar tabla Periodo
    Combo1(1).Clear
    Combo1(1).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 6) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CICCPA_DTREF FROM VW_SGP_LISTAPRECIO WHERE CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CICCPA_DTREF")
    DoEvents
    Do While Not RS.EOF
       DoEvents
       Combo1(1).AddItem Mid(RS!CICCPA_DTREF, 5, 2) & "/" & Mid(RS!CICCPA_DTREF, 1, 4) & Space(150) & "(" & fg_pone_espacio((RS!CICCPA_DTREF), 6) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(1).ListIndex = 0
    
    '-------> Cargar tabla Semana
    Combo1(2).Clear
    Combo1(2).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(("0"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CICCPA_NRSEM FROM VW_SGP_LISTAPRECIO WHERE CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CICCPA_NRSEM")
    DoEvents
    Do While Not RS.EOF
       DoEvents
       Combo1(2).AddItem RS!CICCPA_NRSEM & Space(150) & "(" & fg_pone_cero(Str(RS!CICCPA_NRSEM), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(2).ListIndex = 0
    
    '-------> Cargar tabla Cencos
    Combo1(3).Clear
    Combo1(3).AddItem "Todos" & Space(150) & "(" & fg_pone_espacio(("*"), 10) & ")"
    Set RS = vg_dbsac.Execute("SELECT DISTINCT CADFIL_CDFIL, CADFIL_NMFIL FROM VW_SGP_LISTAPRECIO WHERE CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY CADFIL_NMFIL")
    DoEvents
    Do While Not RS.EOF
       DoEvents
       Combo1(3).AddItem RS!CADFIL_NMFIL & Space(150) & "(" & fg_pone_espacio((RS!CADFIL_CDFIL), 10) & ")"
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    Combo1(3).ListIndex = 0
    vaSpread1.Visible = True
    Est = False
    Frame1(0).Enabled = True
    fg_descarga
End Select
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = IIf(vaSpread1.Value = "1", "0", "1")
End Sub
