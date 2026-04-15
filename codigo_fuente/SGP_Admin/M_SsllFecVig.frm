VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form M_SsllFecVig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fecha Vigencia"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3255
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5370
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2775
         _Version        =   393216
         _ExtentX        =   4895
         _ExtentY        =   9472
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
         MaxCols         =   1
         ScrollBars      =   2
         SpreadDesigner  =   "M_SsllFecVig.frx":0000
      End
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1920
      TabIndex        =   3
      Top             =   6000
      Width           =   1455
   End
   Begin VB.Label Label2 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   6000
      Width           =   1455
   End
End
Attribute VB_Name = "M_SsllFecVig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim est As Boolean
Dim strSQL As String

Private Sub Form_Load()

    Label1 = ""
    Label1.Visible = False
    Label2.Visible = False
    
    vg_FecVig = ""
    Me.Height = 6750
    Me.Width = 3600
    Msgtitulo = "Fechas de Vigencia"
    fg_centra Me
    
    CargarFechaVigencia
    vaSpread1.Lock = True

End Sub


Sub CargarFechaVigencia()
    Dim x As Boolean
    
    vaSpread1.TextTip = 2
    vaSpread1.TextTipDelay = 250
    x = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    
    Label2 = vg_AuxCentCost
    
    strSQL = "SELECT DISTINCT dxv_fecvig FROM b_ssll_dxv WHERE dxv_fecvig <> '' AND dxv_codcen = '" & Label2 & "'"

    Set RS = vg_db.Execute(strSQL)
    
    Do While Not RS.EOF
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       
       vaSpread1.Col = 1
       vaSpread1.CellType = CellTypeDate
       vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
       vaSpread1.TypeDateMin = "01011900"
       vaSpread1.TypeDateMax = "31125000"
       vaSpread1.TypeHAlign = TypeHAlignCenter
       vaSpread1.TypeDateCentury = True
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.text = Trim(IIf(IsNull(RS!dxv_fecvig), "", Format(RS!dxv_fecvig, "dd/mm/yyyy")))
       RS.MoveNext
    Loop
    
    RS.Close: Set RS = Nothing

    'vaSpread1.Visible = True
    If vaSpread1.MaxRows > 0 Then
       vaSpread1.Row = 1
       vaSpread1.Col = 1
       vaSpread1.SetActiveCell 1, 1 ': vaSpread1.SetFocus
    End If

End Sub



Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    vaSpread1.Col = Col
    vaSpread1.Row = Row
    
    Label1.Caption = vaSpread1.text
    
    M_SsllDxv.Label3 = vaSpread1.text
    vg_FecVig = vaSpread1.text
    'MsgBox Label1.Caption
    
    Me.Hide
    Call Unload(Me)
    
    'M_sslldvx.Show
End Sub

