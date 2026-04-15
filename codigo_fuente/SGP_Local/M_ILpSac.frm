VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ILpSac 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Lista de Precio Desde SAC"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   3060
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Top             =   405
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
         Text            =   "12/2009"
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
         Left            =   4320
         TabIndex        =   5
         Top             =   360
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes Actualizar"
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
         Left            =   840
         TabIndex        =   2
         Top             =   450
         Width           =   1260
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   6825
      Left            =   60
      TabIndex        =   3
      Top             =   1200
      Width           =   12375
      _Version        =   393216
      _ExtentX        =   21828
      _ExtentY        =   12039
      _StockProps     =   64
      ButtonDrawMode  =   1
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
      MaxCols         =   12
      MaxRows         =   1000000
      SpreadDesigner  =   "M_ILpSac.frx":0000
      VirtualMode     =   -1  'True
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8580
      Left            =   12495
      TabIndex        =   4
      Top             =   0
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   15134
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
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
            Picture         =   "M_ILpSac.frx":1CE3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   8280
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
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
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   8040
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H80000003&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   11520
      Top             =   8400
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   12000
      Top             =   8400
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_ILpSac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Msgtitulo  As String
Dim cencom() As Variant
Dim tipcal() As Variant
Dim spid As Long
Dim estvec As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim i As Long
On Error GoTo Man_Error
Me.Height = 9060
Me.Width = 13110
estvec = False
fg_centra Me
Msgtitulo = "Importar Lista de Precio Desde SAC"
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fpDateTime1(0).DateTimeFormat = UserDefined
fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
fpDateTime1(0).Text = Format(Date, "mm/yyyy")
vaSpread1.MaxRows = 0
''-------> Llenar vector central de compras
'RS.Open "SELECT COUNT(*) AS nreg FROM b_sac_centralcompras", vg_db, adOpenStatic
'If RS.EOF Or RS!nreg < 1 Then RS.Close: Set RS = Nothing: MsgBox "No existe central de compras, proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'ReDim cencom(RS!nreg, 2)
'RS.Close: Set RS = Nothing
'RS.Open "SELECT * FROM b_sac_centralcompras", vg_db, adOpenStatic
'i = 1
'If Not RS.EOF Then
'   Do While Not RS.EOF
'      cencom(i, 1) = RS!TABCEN_CDCEN
'      cencom(i, 2) = RS!tabcen_dscen
'      i = i + 1
'      RS.MoveNext
'   Loop
'End If
'RS.Close: Set RS = Nothing

'-------> Abrir base sac
AbrirBaseSac
'-------> Llenar vector central de compras
RS.Open "SELECT COUNT(*) AS nreg FROM TABCEN", vg_dbsac, adOpenStatic
If RS.EOF Or RS!nReg < 1 Then RS.Close: Set RS = Nothing: MsgBox "No existe central de compras, proceso cancelado...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
ReDim cencom(RS!nReg, 2)
RS.Close: Set RS = Nothing
RS.Open "SELECT * FROM TABCEN", vg_dbsac, adOpenStatic
i = 1
If Not RS.EOF Then
   Do While Not RS.EOF
      cencom(i, 1) = Trim(RS!TABCEN_CDCEN)
      cencom(i, 2) = Trim(RS!TABCEN_DSCEN)
      i = i + 1
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
'-------> Llenar tipo de cálculo
ReDim tipcal(3, 2)
For i = 1 To 3
    tipcal(i, 1) = i
    tipcal(i, 2) = IIf(i = 1, "Precio Menor", IIf(i = 2, "Precio Promedio", "Precio Mayor"))
Next i
estvec = True
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, titulo
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If IsDate(fpDateTime1(Index).Text) = False Then Exit Sub
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, isel As Integer, Codigo As String, precio As Double, anomes As Long, codlpr As Long, codcal As Long
Dim codcco As String, codcce As String, dtsac As String, nrosem As Long, tipcal As String, codpro As String, nropro As Long
Dim prepro As Double
Select Case Button.Index
Case 1 '-------> Procesar Información
    If vaSpread1.MaxRows < 1 Then MsgBox "Debe seleccionar el periodo a importa precio...", vbCritical, Msgtitulo: Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.Text = "1" Then
           vaSpread1.Col = 5
           codcce = ""
           If vaSpread1.TypeComboBoxCurSel <> -1 Then
              vaSpread1.Col = 6
              codcce = Trim(vaSpread1.Text)
           End If
           vaSpread1.Col = 7
           codcco = Trim(vaSpread1.Text)
           vaSpread1.Col = 9
           dtsac = Mid(vaSpread1.Text, 1, 2) & Mid(vaSpread1.Text, 4, 4)
           vaSpread1.Col = 10
           nrosem = Val(vaSpread1.Text)
'           '-------> Validar central de compras
'           RS.Open "SELECT tabcen_cdcen FROM b_saccentralcompras WHERE tabcen_cdcen = '" & codcce & "'", vg_db, adOpenStatic
'           If RS.EOF Then
'              RS.Close: Set RS = Nothing
'              MsgBox "No existe central de compras", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
'           End If
'           RS.Close: Set RS = Nothing
'           '-------> Validar centro de costo
'           If Trim(codcco) <> "" Then
'              RS.Open "SELECT CADFIL_CDFIL, CADFIL_NMFIL FROM b_saccentrocosto WHERE TABCEN_CDCEN = '" & codcce & "' AND CADFIL_CDFIL = '" & Trim(LimpiaDato(codcco)) & "'", vg_db, adOpenStatic
'              If RS.EOF Then
'                 RS.Close: Set RS = Nothing
'                 MsgBox "No existe centro de costo", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
'              End If
'              RS.Close: Set RS = Nothing
'           End If
'           '-------> Validar periodo
'           RS.Open "SELECT DISTINCT ciccpa_nrsem FROM b_saclistaprecio WHERE tabcen_cdcen = '" & codcce & "' AND (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '') AND ciccpa_dtref = '" & dtsac & "' AND ciccpa_nrsem = " & nrosem & "", vg_db, adOpenStatic
'           If RS.EOF Then
'              RS.Close: Set RS = Nothing
'              MsgBox "No existe periodo", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
'           End If
'           RS.Close: Set RS = Nothing
        End If
    Next i
    isel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.Text = "1" Then
           vaSpread1.Col = 11
           If Trim(vaSpread1.Text) = "" Then MsgBox "Debe seleccionar el tipo de cálculo", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
           isel = 1: Exit For
        End If
    Next i
    If isel = 0 Then MsgBox "Debe seleccionar a lo menor una lista de precio", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = True
    Bar1(0).Visible = True: Bar1(0).Value = 0: Label2(0).Visible = True
    Toolbar1.Enabled = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        DoEvents
        Bar1(0).Value = Val((i / vaSpread1.MaxRows) * 100)
        If vaSpread1.Text = "1" Then
           vaSpread1.SetActiveCell 1, i
           vaSpread1.Col = 3: codlpr = vaSpread1.Text
           vaSpread1.Col = 4: anomes = vaSpread1.Text
           vaSpread1.Col = 6: codcce = Trim(vaSpread1.Text)
           vaSpread1.Col = 7: codcco = IIf(Trim(vaSpread1.Text) = "", "*", Trim(vaSpread1.Text))
           vaSpread1.Col = 9: dtsac = Mid(Trim(vaSpread1.Text), 4, 4) & Mid(Trim(vaSpread1.Text), 1, 2)
           vaSpread1.Col = 10: nrosem = Val(vaSpread1.Text)
           vaSpread1.Col = 12: tipcal = IIf(Val(vaSpread1.Text) = 1, "MIN(forpro_vlpco) AS forpro_vlpco", IIf(Val(vaSpread1.Text) = 2 And codcco = "*", "forpro_vlpco", IIf(Val(vaSpread1.Text) = 2 And codcco <> "*", "AVG(forpro_vlpco) AS forpro_vlpco", "MAX(forpro_vlpco) AS forpro_vlpco")))
           vaSpread1.Col = 12: codcal = Val(vaSpread1.Text)
           If Trim(codcco) = "*" Then
'              RS.Open "SELECT DISTINCT cpopro_cdpro, CPOPRO_DSPRO, MIN(forpro_vlpco) AS forpro_vlpco " & _
'                      "FROM vw_sgp_listaprecio " & _
'                      "WHERE  tabcen_cdcen = '" & codcce & "' " & _
'                      "AND   (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') " & _
'                      "AND    ciccpa_dtref = '" & dtsac & "' " & _
'                      "AND    CICCPA_NRSEM = " & nrosem & " " & _
'                      "GROUP BY cpopro_cdpro, CPOPRO_DSPRO " & _
'                      "", vg_dbsac, adOpenStatic
              If codcal = 2 Then
              RS.Open "SELECT DISTINCT cpopro_cdpro, CPOPRO_DSPRO, " & tipcal & " " & _
                      "FROM vw_sgp_listaprecio " & _
                      "WHERE  tabcen_cdcen = '" & codcce & "' " & _
                      "AND   (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') " & _
                      "AND    ciccpa_dtref = '" & dtsac & "' " & _
                      "AND    CICCPA_NRSEM = " & nrosem & " AND forpro_vlpco > 0 GROUP BY cpopro_cdpro, CPOPRO_DSPRO, " & tipcal & " " & _
                      "", vg_dbsac, adOpenStatic
              Else
                 RS.Open "SELECT DISTINCT cpopro_cdpro, CPOPRO_DSPRO, " & tipcal & " " & _
                         "FROM vw_sgp_listaprecio " & _
                         "WHERE  tabcen_cdcen = '" & codcce & "' " & _
                         "AND   (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') " & _
                         "AND    ciccpa_dtref = '" & dtsac & "' " & _
                         "AND    CICCPA_NRSEM = " & nrosem & " AND forpro_vlpco > 0 GROUP BY cpopro_cdpro, CPOPRO_DSPRO " & _
                         "", vg_dbsac, adOpenStatic
              End If
            Else
'              RS.Open "SELECT DISTINCT cpopro_cdpro, CPOPRO_DSPRO, MIN(forpro_vlpco) AS forpro_vlpco " & _
'                      "FROM vw_sgp_listaprecio " & _
'                      "WHERE  tabcen_cdcen = '" & codcce & "' " & _
'                      "AND   (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') " & _
'                      "AND    ciccpa_dtref = '" & dtsac & "' " & _
'                      "AND    CICCPA_NRSEM = " & nrosem & " " & _
'                      "GROUP BY cpopro_cdpro, CPOPRO_DSPRO " & _
'                      "HAVING COUNT(cpopro_cdpro) = 1", vg_dbsac, adOpenStatic
            
               RS.Open "SELECT DISTINCT cpopro_cdpro, CPOPRO_DSPRO, " & tipcal & " " & _
                       "FROM vw_sgp_listaprecio " & _
                       "WHERE  tabcen_cdcen = '" & codcce & "' " & _
                       "AND   (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') " & _
                       "AND    ciccpa_dtref = '" & dtsac & "' " & _
                       "AND    CICCPA_NRSEM = " & nrosem & " AND forpro_vlpco > 0 " & _
                       "GROUP BY cpopro_cdpro, CPOPRO_DSPRO " & _
                       "HAVING COUNT(cpopro_cdpro) = 1", vg_dbsac, adOpenStatic
            End If
            '-------> Actualizar Encabezado lista de precio
            codpro = "": nropro = 0: prepro = 0
            If codcco = "*" Then codcco = ""
            If Not RS.EOF Then
               vg_db.Execute "UPDATE b_listaprecio " & _
                             "SET lpr_codcec = '" & codcce & "', " & _
                             "    lpr_codcco = '" & codcco & "' " & _
                             "FROM b_listaprecio " & _
                             "WHERE lpr_codigo = " & codlpr & " " & _
                             "AND   lpr_activo = '1'"
               Do While Not RS.EOF
                  DoEvents
                  If Trim(codpro) <> Trim(RS!CPOPRO_CDPRO) Then
                     Label2(0).Caption = Trim(RS!CPOPRO_CDPRO) & " - " & Trim(RS!CPOPRO_DSPRO)
                     If Trim(codpro) <> "" Then
'                                      vg_db.Execute "UPDATE b_detlistaprecio " & _
'                                      "SET b_detlistaprecio.dlp_precio = " & RS!forpro_vlpco & ", " & _
'                                      "    b_detlistaprecio.dlp_fecdig = '" & Format(Date, "mm/dd/yyyy") & "', " & _
'                                      "    b_detlistaprecio.dlp_codcec = '" & codcce & "', " & _
'                                      "    b_detlistaprecio.dlp_codcco = '" & codcco & "', " & _
'                                      "    b_detlistaprecio.dlp_dtsac  = '" & dtsac & "', " & _
'                                      "    b_detlistaprecio.dlp_nrosem = " & nrosem & " " & _
'                                      "FROM b_detlistaprecio a, b_productos c, b_formatocompras d, b_formatocomprassgp e " & _
'                                      "WHERE a.dlp_codpro = c.pro_codigo " & _
'                                      "AND   c.pro_codigo = e.fcs_codsgp " & _
'                                      "AND   d.foc_codsac = e.fcs_codsac " & _
'                                      "AND  (d.foc_flexec = 0 OR (d.foc_flexec = -1 AND d.foc_vigfin > " & Format(Date, "dd/mm/yyyy") & ")) " & _
'                                      "AND   e.fcs_sgppre = 1 " & _
'                                      "AND   d.foc_codsac = '" & RS!cpopro_cdpro & "' " & _
'                                      "AND   a.dlp_codigo = " & codlpr & " " & _
'                                      "AND   a.dlp_anomes = " & anomes & ""
                        vg_db.Execute "UPDATE b_detlistaprecio " & _
                                      "SET b_detlistaprecio.dlp_precio = " & (prepro / nropro) & ", " & _
                                      "    b_detlistaprecio.dlp_fecdig = '" & Format(Date, "mm/dd/yyyy") & "', " & _
                                      "    b_detlistaprecio.dlp_codcec = '" & codcce & "', " & _
                                      "    b_detlistaprecio.dlp_codcco = '" & codcco & "', " & _
                                      "    b_detlistaprecio.dlp_dtsac  = '" & dtsac & "', " & _
                                      "    b_detlistaprecio.dlp_nrosem = " & nrosem & " " & _
                                      "FROM b_detlistaprecio a, b_productos c, b_formatocompras d, b_formatocomprassgp e " & _
                                      "WHERE a.dlp_codpro = c.pro_codigo " & _
                                      "AND   c.pro_codigo = e.fcs_codsgp " & _
                                      "AND   d.foc_codsac = e.fcs_codsac " & _
                                      "AND  (d.foc_flexec = 0 OR (d.foc_flexec = -1 AND d.foc_vigfin > " & Format(Date, "dd/mm/yyyy") & ")) " & _
                                      "AND   e.fcs_sgppre = 1 " & _
                                      "AND   d.foc_codsac = '" & codpro & "' " & _
                                      "AND   a.dlp_codigo = " & codlpr & " " & _
                                      "AND   a.dlp_anomes = " & anomes & ""
                     End If
                     codpro = Trim(RS!CPOPRO_CDPRO)
                     prepro = 0
                     nropro = 0
                  End If
                  prepro = prepro + RS!forpro_vlpco
                  RS.MoveNext: nropro = nropro + 1
               Loop
               If Trim(codpro) <> "" Then
                   vg_db.Execute "UPDATE b_detlistaprecio " & _
                                 "SET b_detlistaprecio.dlp_precio = " & (prepro / nropro) & ", " & _
                                 "    b_detlistaprecio.dlp_fecdig = '" & Format(Date, "mm/dd/yyyy") & "', " & _
                                 "    b_detlistaprecio.dlp_codcec = '" & codcce & "', " & _
                                 "    b_detlistaprecio.dlp_codcco = '" & codcco & "', " & _
                                 "    b_detlistaprecio.dlp_dtsac  = '" & dtsac & "', " & _
                                 "    b_detlistaprecio.dlp_nrosem = " & nrosem & " " & _
                                 "FROM b_detlistaprecio a, b_productos c, b_formatocompras d, b_formatocomprassgp e " & _
                                 "WHERE a.dlp_codpro = c.pro_codigo " & _
                                 "AND   c.pro_codigo = e.fcs_codsgp " & _
                                 "AND   d.foc_codsac = e.fcs_codsac " & _
                                 "AND  (d.foc_flexec = 0 OR (d.foc_flexec = -1 AND d.foc_vigfin > " & Format(Date, "dd/mm/yyyy") & ")) " & _
                                 "AND   e.fcs_sgppre = 1 " & _
                                 "AND   d.foc_codsac = '" & codpro & "' " & _
                                 "AND   a.dlp_codigo = " & codlpr & " " & _
                                 "AND   a.dlp_anomes = " & anomes & ""
               End If
            End If
            RS.Close: Set RS = Nothing
'           vg_db.Execute "sgpadm_u_actuasaclistaprecio " & codlpr & ", " & anomes & ", '" & codcce & "', '" & codcco & "', '" & dtsac & "', " & nrosem & ""
        End If
    Next i
    Bar1(0).Visible = False: Bar1(0).Value = 0: Label2(0).Visible = False
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
    vaSpread1.Col = -1: vaSpread1.Row = -1: vaSpread1.Lock = False
    Toolbar1.Enabled = True
    fg_descarga
Case 3 '-------> Salir
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim lisnom As String, liscod As String, i As Long, codaux As Long, codcce As String
Select Case Button.Index
Case 1 '-------> Traer lista de precio
    fg_carga ""
''    Set RS = vg_db.Execute("sgpadm_s_listaprecio 10, 0, " & Format(fpDateTime1(0).text, "yyyymm") & ", '" & vg_NUsr & "'")
'    '-------> Borrar tabla paso sac lista precio
'    vg_db.Execute "DELETE paso_sac_listaprecio WHERE spid=@@spid and usuario='" & vg_NUsr & "'"
'    '-------> Buscar spid
'    Set RS = vg_db.Execute("SELECT @@spid spid")
'    If Not RS.EOF Then spid = RS!spid
'    RS.Close: Set RS = Nothing
'    Set RS = vg_dbsac.Execute("SELECT DISTINCT " & spid & ", '" & vg_NUsr & "', TABCEN_CDCEN, CADFIL_CDFIL, CADFIL_NMFIL, CICCPA_DTREF, CICCPA_NRSEM FROM vw_sgp_listaprecio WHERE CICCOT_DTPDE >= '" & dBoM("01/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' AND CICCOT_DTPAT <= '" & dEoM("25/" & Format(fpDateTime1(0).text, "mm/yyyy")) & "' ORDER BY TABCEN_CDCEN, CADFIL_CDFIL, CADFIL_NMFIL, CICCPA_DTREF, CICCPA_NRSEM")
'    Do While Not RS.EOF
'       vg_db.Execute "INSERT INTO paso_sac_listaprecio (spid, usuario, TABCEN_CDCEN, CADFIL_CDFIL, CADFIL_NMFIL, CICCPA_DTREF, CICCPA_NRSEM) VALUES (" & spid & ", '" & vg_NUsr & "', '" & RS!TABCEN_CDCEN & "', '" & RS!CADFIL_CDFIL & "', '" & RS!CADFIL_NMFIL & "', '" & RS!CICCPA_DTREF & "', " & RS!CICCPA_NRSEM & ")"
'       RS.MoveNext
'    Loop
'    RS.Close: Set RS = Nothing
    vaSpread1.Visible = 0
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor
    Set RS = vg_db.Execute("sgpadm_s_listaprecio 10, 0, " & Format(fpDateTime1(0).Text, "yyyymm") & ", '" & vg_NUsr & "'")
    If Not RS.EOF Then
       Do While Not RS.EOF
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.Row = vaSpread1.MaxRows
          vaSpread1.Col = 1: vaSpread1.Text = "0"
          vaSpread1.Col = 2: vaSpread1.Text = RS!lpr_codigo & " - " & Trim(RS!lpr_nombre)
          vaSpread1.Col = 3: vaSpread1.Text = RS!lpr_codigo
          vaSpread1.Col = 4: vaSpread1.Text = RS!dlp_anomes
          
          lisnom = "": liscod = "":  encuentra = False
          If estvec Then
          For i = 1 To UBound(cencom)
              vaSpread1.Col = 5: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(cencom(i, 2)) & " (" & Trim(cencom(i, 1)) & ")"
              vaSpread1.Col = 6: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & Trim(cencom(i, 1))
              vaSpread1.Col = 5: vaSpread1.TypeComboBoxList = lisnom
              vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = liscod
          Next i
          End If
          vaSpread1.Col = 6
          codaux = -1
          For i = 0 To vaSpread1.TypeComboBoxCount
              vaSpread1.TypeComboBoxCurSel = i
              If vaSpread1.Text = Trim(RS!lpr_codcec) Then codaux = i: Exit For
              codaux = -1
          Next i
          vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = codaux
          
          lisnom = "": liscod = "":  encuentra = False
          If estvec Then
          For i = 1 To UBound(tipcal)
              vaSpread1.Col = 11: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(tipcal(i, 2)) '& " (" & Trim(tipcal(i, 1)) & ")"
              vaSpread1.Col = 12: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & Trim(tipcal(i, 1))
              vaSpread1.Col = 11: vaSpread1.TypeComboBoxList = lisnom
              vaSpread1.Col = 12: vaSpread1.TypeComboBoxList = liscod
          Next i
          End If
'          vaSpread1.Col = 12
'          codaux = -1
'          For i = 0 To vaSpread1.TypeComboBoxCount
'              vaSpread1.TypeComboBoxCurSel = i
'              If vaSpread1.text = Trim(RS!lpr_codcec) Then codaux = i: Exit For
'              codaux = -1
'          Next i
'          vaSpread1.Col = 11: vaSpread1.TypeComboBoxCurSel = codaux
          
          
          
          vaSpread1.Col = 7: vaSpread1.Text = IIf(IsNull(RS!lpr_codcco), "", RS!lpr_codcco)
          vaSpread1.Col = 8: vaSpread1.Text = ""
          '-------> traer nombre filial desde base sac
          RS1.Open "SELECT * FROM cadfil WHERE cadfil_cdfil = '" & RS!lpr_codcco & "'", vg_dbsac, adOpenStatic
          If Not RS1.EOF Then
             vaSpread1.Text = IIf(IsNull(RS1!CADFIL_NMFIL), "", RS1!CADFIL_NMFIL)
          End If
          RS1.Close: Set RS1 = Nothing
          vaSpread1.Col = 9: vaSpread1.Text = "": vaSpread1.Text = IIf(IsNull(RS!dlp_dtsac), "", Mid(RS!dlp_dtsac, 5, 2) & "/" & Mid(RS!dlp_dtsac, 1, 4))
          vaSpread1.Col = 10: vaSpread1.Text = "": vaSpread1.Text = IIf(IsNull(RS!dlp_nrosem), "", RS!dlp_nrosem)
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    vaSpread1.Visible = True
    fg_descarga
End Select
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, titulo
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
Dim codcce As String, codcco As String, dtsac As String, nrosem As Long
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row
Select Case Col
Case 7 '-------> Traer centro costo
    '-------> Mover código centro de costo
    vaSpread1.Col = 6
    codcce = Trim(vaSpread1.Text)
    vaSpread1.Col = Col
'    RS.Open "SELECT CADFIL_CDFIL, CADFIL_NMFIL FROM b_sac_centrocosto WHERE TABCEN_CDCEN = '" & codcce & "' AND CADFIL_CDFIL = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
    RS.Open "SELECT CADFIL_CDFIL FROM vw_sgp_listaprecio WHERE TABCEN_CDCEN = '" & codcce & "' AND CADFIL_CDFIL = '" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_dbsac, adOpenStatic
    If Not RS.EOF Then
       RS.Close: Set RS = Nothing
       '-------> Validar centro costo oracle
       RS.Open "SELECT * FROM cadfil WHERE cadfil_cdfil = '" & Trim(LimpiaDato(vaSpread1.Text)) & "'", vg_dbsac, adOpenStatic
'       Set RS = vg_db.Execute("sgpadm_s_cliente 13, '" & Trim(LimpiaDato(vaSpread1.text)) & "',''")
       If Not RS.EOF Then
          vaSpread1.Col = 8
          vaSpread1.Text = Trim(RS!CADFIL_NMFIL)
          RS.Close: Set RS = Nothing
        Else
          vaSpread1.Text = ""
          vaSpread1.Col = 8
          vaSpread1.Text = ""
          vaSpread1.SetActiveCell 7, vaSpread1.ActiveRow
          RS.Close: Set RS = Nothing
          MsgBox "No existe centro de costo", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
       End If
    Else
       vaSpread1.Text = ""
       vaSpread1.Col = 8
       vaSpread1.Text = ""
       vaSpread1.SetActiveCell 7, vaSpread1.ActiveRow
       RS.Close: Set RS = Nothing
       MsgBox "No existe centro de costo", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    End If
Case 9 '-------> Traer nro. semana
    vaSpread1.Col = 6: codcce = Trim(vaSpread1.Text)
    vaSpread1.Col = 7: codcco = IIf(Trim(vaSpread1.Text) = "", "*", Trim(vaSpread1.Text))
    vaSpread1.Col = Col: dtsac = Mid(vaSpread1.Text, 4, 4) & Mid(vaSpread1.Text, 1, 2)
'    RS.Open "SELECT DISTINCT ciccpa_nrsem FROM b_sac_listaprecio WHERE tabcen_cdcen = '" & codcce & "' AND (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '') AND ciccpa_dtref = '" & dtsac & "'", vg_db, adOpenStatic
    RS.Open "SELECT DISTINCT ciccpa_nrsem FROM vw_sgp_listaprecio WHERE tabcen_cdcen = '" & codcce & "' AND (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') AND ciccpa_dtref = '" & dtsac & "'", vg_dbsac, adOpenStatic
    If Not RS.EOF Then
       vaSpread1.Col = 10
       vaSpread1.Text = 0
       vaSpread1.Text = CStr(RS!CICCPA_NRSEM)
       RS.Close: Set RS = Nothing
    Else
       vaSpread1.Text = ""
       vaSpread1.Col = 10
       vaSpread1.Text = ""
       vaSpread1.SetActiveCell 9, vaSpread1.ActiveRow
       RS.Close: Set RS = Nothing
       MsgBox "No existe fecha registrada", vbInformation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
Case 10 '-------> Traer nro. semana
    vaSpread1.Col = 6: codcce = Trim(vaSpread1.Text)
    vaSpread1.Col = 7: codcco = IIf(Trim(vaSpread1.Text) = "", "*", Trim(vaSpread1.Text))
    vaSpread1.Col = 9: dtsac = Mid(vaSpread1.Text, 4, 4) & Mid(vaSpread1.Text, 1, 2)
    vaSpread1.Col = Col: nrosem = Val(vaSpread1.Text)
    RS.Open "SELECT DISTINCT ciccpa_nrsem FROM vw_sgp_listaprecio WHERE tabcen_cdcen = '" & codcce & "' AND (cadfil_cdfil = '" & codcco & "' OR '" & codcco & "' = '*') AND ciccpa_dtref = '" & dtsac & "' AND ciccpa_nrsem = " & nrosem & "", vg_dbsac, adOpenStatic
    If Not RS.EOF Then
       RS.Close: Set RS = Nothing
    Else
       vaSpread1.Text = ""
       RS.Close: Set RS = Nothing
       MsgBox "No existe nro semana", vbInformation + vbOKOnly, Msgtitulo
       Exit Sub
       vaSpread1.SetActiveCell 10, vaSpread1.Row
    End If
End Select
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Text = IIf(vaSpread1.Value = "1", "0", "1")
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
Select Case Col
Case 5
    vaSpread1.Row = Row
    vaSpread1.Col = 5: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = indice
Case 11
    vaSpread1.Row = Row
    vaSpread1.Col = 11: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 12: vaSpread1.TypeComboBoxCurSel = indice
End Select
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row
Select Case Col
Case 7 '------->Traer centro costo
    If vaSpread1.MaxRows < 1 Then Exit Sub
    vg_nombre = "": vg_codigo = ""
    vg_left = Me.Left + 7550
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = -1
    If vaSpread1.RowHidden = True Then Exit Sub
    vaSpread1.BackColor = Shape1(2).FillColor
    vaSpread1.Col = 6
    vg_codigo = Trim(vaSpread1.Text)
    B_TabEst.LlenaDatos "b_saccentrocosto", "cco_", "Centro de Costo SAC", "SacCco"
    B_TabEst.Show 1
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor
    If vg_codigo = "" Then Exit Sub
    Row = vaSpread1.ActiveRow
    vaSpread1.Row = Row
    vaSpread1.Col = 7: vaSpread1.Text = Trim(vg_codigo)
    vaSpread1.Col = 8: vaSpread1.Text = Trim(vg_nombre)
    vaSpread1.SetActiveCell 9, vaSpread1.Row
End Select
End Sub
