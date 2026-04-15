VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_RecCas 
   Caption         =   "Generación archivos planos recetas"
   ClientHeight    =   7800
   ClientLeft      =   1470
   ClientTop       =   2055
   ClientWidth     =   12675
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   12675
   ShowInTaskbar   =   0   'False
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   3600
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Casinos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   945
         Index           =   2
         Left            =   570
         TabIndex        =   9
         Top             =   180
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "M_RecCas.frx":0000
            Left            =   1680
            List            =   "M_RecCas.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   180
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   11
            Top             =   540
            Width           =   2895
            _Version        =   196608
            _ExtentX        =   5106
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   3
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   0
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
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   1
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Texto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   13
            Top             =   600
            Width           =   1470
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Columna"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   11
            Left            =   150
            TabIndex        =   12
            Top             =   285
            Width           =   1485
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3975
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1890
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7011
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   1
         SpreadDesigner  =   "M_RecCas.frx":001E
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   660
         TabIndex        =   15
         Top             =   6060
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
   Begin VB.Frame Frame1 
      Caption         =   "Recetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Index           =   1
      Left            =   5970
      TabIndex        =   0
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   3
         Left            =   570
         TabIndex        =   2
         Top             =   180
         Width           =   4725
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Mostrar solo recetas no enviados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   270
            Index           =   1
            Left            =   150
            TabIndex        =   23
            Top             =   870
            Width           =   3180
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "M_RecCas.frx":0357
            Left            =   1680
            List            =   "M_RecCas.frx":0361
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   180
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   4
            Top             =   540
            Width           =   2895
            _Version        =   196608
            _ExtentX        =   5106
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   3
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   0
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
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483637
            Appearance      =   1
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483637
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Columna"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   285
            Width           =   1485
         End
         Begin VB.Label Label1 
            Caption         =   "Buscar Texto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   600
            Width           =   1470
         End
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   1
         Left            =   660
         TabIndex        =   1
         Top             =   6090
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3975
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1890
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7011
         _StockProps     =   64
         BackColorStyle  =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   1
         SpreadDesigner  =   "M_RecCas.frx":0375
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2805
         Top             =   5910
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No Enviados"
         Height          =   195
         Index           =   0
         Left            =   3165
         TabIndex        =   22
         Top             =   5880
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1590
         Top             =   5910
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enviados"
         Height          =   195
         Index           =   1
         Left            =   1950
         TabIndex        =   21
         Top             =   5880
         Width           =   660
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   9
         Left            =   2295
         TabIndex        =   20
         Top             =   1650
         Width           =   4425
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   8
         Left            =   2295
         TabIndex        =   19
         Top             =   1380
         Width           =   4425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Categoria Dietetica"
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
         Index           =   12
         Left            =   570
         TabIndex        =   18
         Top             =   1380
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Plato"
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
         Index           =   3
         Left            =   570
         TabIndex        =   17
         Top             =   1665
         Width           =   885
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   840
      OleObjectBlob   =   "M_RecCas.frx":06D7
      Top             =   6960
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   1470
      OleObjectBlob   =   "M_RecCas.frx":06FB
      Top             =   6810
   End
End
Attribute VB_Name = "M_RecCas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Est As Boolean

Private Sub Check1_Click(Index As Integer)
If Est Then Exit Sub
MoverDatoGrillaReceta
End Sub

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
vaSpread1(Index).SetFocus
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7245
Me.Width = 12000
Me.HelpContextID = vg_OpcM
Msgtitulo = "Generación archivos planos recetas"
fg_centra Me
Est = True
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.ToolTipText = "Enviar": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Filtro", , tbrDefault, "A_Filtro"): btnX.Visible = True: btnX.ToolTipText = "Filtrar Recetas":: btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(0).ListIndex = 1: Combo1(1).ListIndex = 1
vg_filcatdie = 0: vg_filtippla = 0
RS.Open "select par_valor from a_param where par_codigo='catdefecto'", vg_db, adOpenStatic
If Not RS.EOF Then vg_filcatdie = RS!par_valor: Label2(8).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
RS.Close: Set RS = Nothing
'Label2(8).Caption = "Todos": Label2(9).Caption = "Todos"
Check1(1).Value = 1
MoverDatoGrillaCasino
MoverDatoGrillaReceta
Est = False
SendKeys "+{Tab}"
End Sub

Private Sub fpTnombre_Change(Index As Integer)
Select Case Index
Case 0
    If LimpiaDato(Trim(fptnombre(0).text)) & Chr(KeyAscii) = "" Then Exit Sub
    If Combo1(0).ItemData(Combo1(0).ListIndex) = 0 Then
       RS.Open "select cli_codigo, cli_nombre, cli_tipo from b_clientes Where upper(cli_codigo) like '%" & UCase(LimpiaDato(fptnombre(0).text)) & "%' and cli_tipo=0 order by cli_codigo", vg_db, adOpenStatic
    ElseIf Combo1(0).ItemData(Combo1(0).ListIndex) = 1 Then
       RS.Open "select cli_codigo, cli_nombre, cli_tipo From b_clientes Where upper(cli_nombre) like '%" & UCase(LimpiaDato(fptnombre(0).text)) & "%' and cli_tipo=0 order by cli_nombre", vg_db, adOpenStatic
    End If
    ibusca = RS.RecordCount: vaSpread1(0).MaxRows = RS.RecordCount
    i = 1
    If Not RS.EOF Then
       Do While Not RS.EOF
          vaSpread1(0).Row = i

          vaSpread1(0).Col = 1
          vaSpread1(0).text = "0"
          
          vaSpread1(0).Col = 2
          vaSpread1(0).text = RS!cli_codigo
          
          vaSpread1(0).Col = 3
          vaSpread1(0).TypeHAlign = 0
          vaSpread1(0).text = Trim(RS!cli_nombre)
          RS.MoveNext: i = i + 1
       Loop
    End If
    RS.Close: Set RS = Nothing
Case 1
    If LimpiaDato(Trim(fptnombre(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Then
       If Check1(1).Value = 1 Then
          RS.Open "select distinct b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
                  "FROM b_receta LEFT JOIN b_recetacasino ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
                  "Where IsNull(b_recetacasino.rec_codrec) " & _
                  "and   (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
                  "and (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' " & _
                  "and upper(rec_codigo) like '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "order by b_receta.rec_codigo", vg_db, adOpenStatic
       Else
          RS.Open "select distinct b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
                  "FROM b_receta LEFT JOIN b_recetacasino ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
                  "Where (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
                  "and (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' " & _
                  "and upper(rec_codigo) like '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "order by b_receta.rec_codigo", vg_db, adOpenStatic
       End If
       ibusca = RS.RecordCount: vaSpread1(1).MaxRows = RS.RecordCount
    ElseIf Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
       If Check1(1).Value = 1 Then
          RS.Open "select distinct b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
                  "FROM b_receta LEFT JOIN b_recetacasino ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
                  "Where IsNull(b_recetacasino.rec_codrec) " & _
                  "and   (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
                  "and   (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' " & _
                  "and   upper(rec_nombre) like '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "order by b_receta.rec_nombre", vg_db, adOpenStatic
       Else
          RS.Open "select distinct b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
                  "FROM b_receta LEFT JOIN b_recetacasino ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
                  "Where (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
                  "and   (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' " & _
                  "and   upper(rec_nombre) like '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "order by b_receta.rec_nombre", vg_db, adOpenStatic
       End If
       ibusca = RS.RecordCount: vaSpread1(1).MaxRows = RS.RecordCount
    End If
    i = 1
    If Not RS.EOF Then
       Do While Not RS.EOF
          vaSpread1(1).Row = i
          
          vaSpread1(1).Col = 1
          vaSpread1(1).BackColor = IIf(IsNull(RS!rec_codrec), Shape1(0).FillColor, Shape1(1).FillColor)
          vaSpread1(1).text = "0"
          
          vaSpread1(1).Col = 2
          vaSpread1(1).TypeHAlign = TypeHAlignLeft
          vaSpread1(1).BackColor = IIf(IsNull(RS!rec_codrec), Shape1(0).FillColor, Shape1(1).FillColor)
          vaSpread1(1).text = RS!rec_codigo
        
          vaSpread1(1).Col = 3
          vaSpread1(1).TypeHAlign = TypeHAlignLeft
          vaSpread1(1).BackColor = IIf(IsNull(RS!rec_codrec), Shape1(0).FillColor, Shape1(1).FillColor)
          vaSpread1(1).text = Trim(RS!rec_nombre)
        
          RS.MoveNext: i = i + 1
       Loop
    End If
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fptnombre_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 34 And irow > 0 Then vaSpread1(Index).SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1(0).MaxRows < 1 Or vaSpread1(1).MaxRows < 1 Then Exit Sub
    Dim i As Long, j As Long, codrec As Long
    Dim isel As Boolean, icopy As Boolean
    Dim cencos As String, nomcencos As String, aAp As String, sourcefile As String, sourcefilezip As String, destinofile As String, destinofilezip As String, mdir As String, lognarchsou As String
    Dim CHost As String, Cdire As String, Cuser As String, Cpass As String, Cpuer As Long
    isel = False
    For i = 1 To vaSpread1(0).MaxRows
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un casino", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    isel = False
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor una receta", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
   '------- Creo tabla temporal y chequeo si existe antes
   aAp = Trim(vg_NUsr) & "_tmp_GenPlanoRec"
   fg_CheckTmp aAp
   vg_db.BeginTrans
   vg_db.Execute "create table " & aAp & " (codrec int)"
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then
           vaSpread1(1).Col = 2
           vg_db.Execute "insert into " & aAp & " (codrec) values (" & vaSpread1(1).text & ")"
        End If
    Next i
    '------- Crear directorio si no existe
    mdir = Dir(dir_trabajo & "\" & "Actualizar", vbDirectory)
    If mdir = "" Then MkDir dir_trabajo & "\" & "Actualizar"
    mdir = dir_trabajo & "Actualizar" & "\"
    '------- Fin crear directorio si no existe
    Bar1(0).Visible = True: Bar1(1).Visible = True
    Bar1(0).Value = 0: Bar1(1).Value = 0: icopy = False
    For i = 1 To vaSpread1(0).MaxRows
        Bar1(0).Value = Val((i / vaSpread1(0).MaxRows) * 100)
        vaSpread1(0).Row = i: vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then
           DoEvents
           vaSpread1(0).Col = 3: nomcencos = Trim(vaSpread1(0).text)
           vaSpread1(0).Col = 2: cencos = Trim(vaSpread1(0).text): Bar1(1).Value = 0
           vaSpread1(0).SetActiveCell 2, vaSpread1(0).Row: vaSpread1(0).SetFocus
           If icopy = False Then
              sourcefile = "mr" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
              sourcefilezip = "mr" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           Else
              destinofile = "mr" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
              destinofilezip = "mr" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           End If
           If icopy = True Then
              '------- verificar si existe archivo mdb destino si existe borrar y copiar
              If Dir(mdir & destinofile) <> "" Then Kill mdir & destinofile
              FileCopy mdir & sourcefile, mdir & destinofile
              '------- verificar si existe archivo zip destino si existe borrar
              If Dir(mdir & destinofilezip) <> "" Then Kill mdir & destinofilezip
              AZ1.CreateZip mdir & destinofilezip, "": AZ1.AddFile mdir & destinofile, "", True, "": AZ1.Close
              '------- verificar si existe archivo mdb destino si existe borrar
              If Dir(mdir & destinofile) <> "" Then Kill mdir & destinofile
              '------- leer casino
              RS.Open "select * from b_clientes where cli_codigo='" & cencos & "'", vg_db, adOpenStatic
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
'                    Open dir_trabajo & "\sdxftp.ini" For Input As #1
'                    Do While Not EOF(1)
'                       Line Input #1, cpars
'                       If Mid(cpars, 1, InStr(cpars, ",") - 1) = "A" Then
'                          CHost = Mid(cpars, InStr(cpars, ",") + 1)
'                       ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "B" Then
'                          Cuser = Mid(cpars, InStr(cpars, ",") + 1)
'                       ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "C" Then
'                          Cpass = Mid(cpars, InStr(cpars, ",") + 1)
'                       End If
'                    Loop
'                    Close #1
                    a = oFTP.Version
                    oFTP.UseIEProxy = False
                    oFTP.Port = Cpuer '21
                    oFTP.HostName = CHost '"sgp.sodexhochile.cl" '"64.76.138.76" '"64.76.45.71"  'fg_Desencripta(TipoDato(cHost, ""))
                    oFTP.UserName = Cuser '"userftp" '"sodexho"   'fg_Desencripta(TipoDato(cUser, ""))
                    oFTP.password = Cpass '"*sdxo123*" '"shx873" 'fg_Desencripta(TipoDato(cPass, ""))
                    oFTP.Connect
                    If oFTP.IsConnected Then
                        lDir = oFTP.GetCurrentDirListing("*.*")
                        oFTP.SaveLastError ("aaa.xml")
'                        a = oFTP.ChangeRemoteDir("/casinos/bd")
                        a = oFTP.ChangeRemoteDir(Cdire)
                        oFTP.SaveLastError ("aaa.xml")
                        lDir = oFTP.GetCurrentDirListing("*.*")
                        oFTP.SaveLastError ("aaa.xml")
                        a = oFTP.PutFile(mdir & destinofilezip, destinofilezip)
                        oFTP.SaveLastError ("aaa.xml")
                        oFTP.Disconnect
                        If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                           fg_descarga
                           MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no se puede notificar por correo, ya que no tiene asignado el mail", vbInformation + vbOKOnly, Msgtitulo
                           fg_carga ""
                        Else
                           SendMail oMail, "Actualización maestro de recetas " & Format(Date, "dd/mm/yyyy"), "Se Informa que el maestro de recetas esta disponible para actualizar.", mdir & sourcefilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0
                        End If
                    Else
                       GoTo Man_Error
                    End If
                 ElseIf RS!cli_openvio = 2 Then
                    If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                       fg_descarga
                       MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, Msgtitulo
                       fg_carga ""
                    Else
                       SendMail oMail, "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), mdir & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1
                    End If
                 End If
              End If
              RS.Close: Set RS = Nothing
'           If icopy = True Then
'              '------- verificar si existe archivo mdb destino si existe borrar y copiar
'              If Dir(destinofile) <> "" Then Kill destinofile
'              FileCopy sourcefile, destinofile
           ElseIf icopy = False Then
              '------- verificar si existe archivo mdb y zip si existe borrar
              If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile
              If Dir(mdir & sourcefilezip) <> "" Then Kill mdir & sourcefilezip
              '------- generar archivo mdb
              Set db7 = DBEngine(0).CreateDatabase(mdir & sourcefile, dbLangGeneral)
              db7.Execute "create table a_recetacatdie (car_codigo int, car_nombre char(50), car_previo int)", vg_ModoOpen
              db7.Execute "create table a_recetatippla (tip_codigo int, tip_nombre char(50), tip_previo int)", vg_ModoOpen
              db7.Execute "create table b_receta (rec_codigo int, rec_catdie int, rec_tippla int, rec_nombre char(80), rec_nomfan char(80), rec_metpre longtext, rec_conche longtext, rec_sugere longtext, rec_basrac int, rec_tiprec char(1))", vg_ModoOpen
              db7.Execute "create table b_recetadet (red_codigo int, red_nroite int, red_codpro char(20), red_canpro DOUBLE, red_cospro double, red_pctapr double, red_pctcoc double, red_pctnut double)", vg_ModoOpen
'              Open mdir & "MR" & cencos & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt" For Output As #1
              '------- generar categoria dietetica
              RS.Open "select * from a_recetacatdie", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_recetacatdie Values (" & RS!car_codigo & ", " & IIf(IsNull(RS!car_nombre), "Null", "'" & RS!car_nombre & "'") & ", " & IIf(IsNull(RS!car_previo), "Null", RS!car_previo) & ")", vg_ModoOpen
'                    Print #1, "a_recetacatdie;" & RS!car_codigo & ";insert into a_recetacatdie values (" & RS!car_codigo & "," & IIf(IsNull(RS!car_nombre), "Null", "'" & RS!car_nombre & "'") & "," & IIf(IsNull(RS!car_previo), "Null", RS!car_previo) & ");" & "update a_recetacatdie set car_nombre=" & IIf(IsNull(RS!car_nombre), "Null", "'" & RS!car_nombre & "'") & ", car_previo=" & IIf(IsNull(RS!car_previo), "Null", RS!car_previo) & " where car_codigo=" & RS!car_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar tipo plato
              RS.Open "select * from a_recetatippla", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_recetatippla Values (" & RS!tip_codigo & ", " & IIf(IsNull(RS!tip_nombre), "Null", "'" & RS!tip_nombre & "'") & ", " & IIf(IsNull(RS!tip_previo), "Null", RS!tip_previo) & ")", vg_ModoOpen
'                    Print #1, "a_recetatippla;" & RS!tip_codigo & ";insert into a_recetatippla values (" & RS!tip_codigo & "," & IIf(IsNull(RS!tip_nombre), "Null", "'" & RS!tip_nombre & "'") & "," & IIf(IsNull(RS!tip_previo), "Null", RS!tip_previo) & ");" & "update a_recetatippla set tip_nombre=" & IIf(IsNull(RS!tip_nombre), "Null", "'" & RS!tip_nombre & "'") & ", tip_previo=" & IIf(IsNull(RS!tip_previo), "Null", RS!tip_previo) & " where tip_codigo=" & RS!tip_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
           End If
           For j = 1 To vaSpread1(1).MaxRows
               Bar1(1).Value = Val((j / vaSpread1(1).MaxRows) * 100)
               vaSpread1(1).Row = j: vaSpread1(1).Col = 1
               If vaSpread1(1).text = "1" Then
                  vaSpread1(1).Col = 2
                  codrec = vaSpread1(1).text
                  '------- Leer, insertar y rebrabar recetas casinos
                  RS.Open "select rec_cencos, rec_codrec from b_recetacasino where rec_cencos='" & cencos & "' and rec_codrec=" & codrec & "", vg_db, adOpenStatic
                  If RS.EOF Then
                     vg_db.Execute "insert into b_recetacasino (rec_cencos, rec_codrec, rec_fecenv) values ('" & cencos & "', " & codrec & ", " & Format(Date, "yyyymmdd") & ")"
                  Else
                     vg_db.Execute "update b_recetacasino set rec_fecenv=" & Format(Date, "yyyymmdd") & " where rec_cencos='" & cencos & "' and rec_codrec=" & codrec & ""
                  End If
                  RS.Close: Set RS = Nothing
                  '------- Fin leer, insertar y rebrabar productos casinos
                  If icopy = False Then
                     '------- Generar encabezado receta
                     RS.Open "select * from b_receta where rec_codigo=" & codrec & "", vg_db, adOpenStatic
                     If Not RS.EOF Then
                        Do While Not RS.EOF
                           db7.Execute "insert into b_receta Values (" & RS!rec_codigo & ", " & IIf(IsNull(RS!rec_catdie), "Null", RS!rec_catdie) & ", " & IIf(IsNull(RS!rec_tippla), "Null", RS!rec_tippla) & ", " & IIf(IsNull(RS!rec_nombre), "Null", "'" & RS!rec_nombre & "'") & ", " & _
                                     "" & IIf(IsNull(RS!rec_nomfan), "Null", "'" & RS!rec_nomfan & "'") & ", " & IIf(IsNull(RS!rec_metpre), "Null", "'" & RS!rec_metpre & "'") & ", " & IIf(IsNull(RS!rec_conche), "Null", "'" & RS!rec_conche & "'") & ", " & _
                                     "" & IIf(IsNull(RS!rec_sugere), "Null", "'" & RS!rec_sugere & "'") & ", " & IIf(IsNull(RS!rec_basrac), "Null", RS!rec_basrac) & ", " & IIf(IsNull(RS!rec_nombre), "Null", "'" & RS!rec_tiprec & "'") & ")", vg_ModoOpen

'                           Print #1, "b_receta;" & RS!rec_codigo & ";insert into b_receta values (" & RS!rec_codigo & "," & IIf(IsNull(RS!rec_catdie), "Null", RS!rec_catdie) & "," & IIf(IsNull(RS!rec_tippla), "Null", RS!rec_tippla) & "," & IIf(IsNull(RS!rec_nombre), "Null", "'" & RS!rec_nombre & "'") & "," & _
'                                     "" & IIf(IsNull(RS!rec_nomfan), "Null", "'" & RS!rec_nomfan & "'") & "," & IIf(IsNull(RS!rec_metpre), "Null", "'" & RS!rec_metpre & "'") & "," & IIf(IsNull(RS!rec_conche), "Null", "'" & RS!rec_conche & "'") & "," & IIf(IsNull(RS!rec_sugere), "Null", "'" & RS!rec_sugere & "'") & "," & IIf(IsNull(RS!rec_basrac), "Null", RS!rec_basrac) & "," & _
'                                     "" & IIf(IsNull(RS!rec_tiprec), "Null", "'" & RS!rec_tiprec & "'") & ");" & "update b_receta set rec_catdie=" & IIf(IsNull(RS!rec_catdie), "Null", RS!rec_catdie) & ", rec_tippla=" & IIf(IsNull(RS!rec_tippla), "Null", RS!rec_tippla) & ", rec_nombre=" & IIf(IsNull(RS!rec_nombre), "Null", "'" & RS!rec_nombre & "'") & "," & _
'                                     "rec_nomfan=" & IIf(IsNull(RS!rec_nomfan), "Null", "'" & RS!rec_nomfan & "'") & ", rec_metpre=" & IIf(IsNull(RS!rec_metpre), "Null", "'" & RS!rec_metpre & "'") & ", rec_conche=" & IIf(IsNull(RS!rec_conche), "Null", "'" & RS!rec_conche & "'") & ", rec_sugere=" & IIf(IsNull(RS!rec_sugere), "Null", "'" & RS!rec_sugere & "'") & "," & _
'                                     "rec_basrac=" & IIf(IsNull(RS!rec_basrac), "Null", RS!rec_basrac) & ", rec_tiprec=" & IIf(IsNull(RS!rec_tiprec), "Null", "'" & RS!rec_tiprec & "'") & " where rec_codigo = " & RS!rec_codigo & ""
                           RS.MoveNext
                        Loop
                     End If
                     RS.Close: Set RS = Nothing
                     '------- generar detalle recetas
                     RS.Open "select * from b_recetadet where red_codigo=" & codrec & "", vg_db, adOpenStatic
                     If Not RS.EOF Then
'                        Print #1, "b_recetadet;delete b_recetadet from b_recetadet where red_codigo=" & codrec & ""
                        Do While Not RS.EOF
                           db7.Execute "insert into b_recetadet Values (" & RS!red_codigo & ", " & IIf(IsNull(RS!red_nroite), "Null", RS!red_nroite) & ", " & IIf(IsNull(RS!red_codpro), "Null", "'" & RS!red_codpro & "'") & ", " & IIf(IsNull(RS!red_canpro), "Null", RS!red_canpro) & ", " & IIf(IsNull(RS!red_cospro), "Null", RS!red_cospro) & ", " & IIf(IsNull(RS!red_pctapr), "Null", RS!red_pctapr) & ", " & IIf(IsNull(RS!red_pctcoc), "Null", RS!red_pctcoc) & ", " & IIf(IsNull(RS!red_pctnut), "Null", RS!red_pctnut) & ")", vg_ModoOpen
'                           Print #1, "b_recetadet;insert into b_recetadet values (" & RS!red_codigo & "," & IIf(IsNull(RS!red_nroite), "Null", RS!red_nroite) & "," & IIf(IsNull(RS!red_codpro), "Null", "'" & RS!red_codpro & "'") & "," & IIf(IsNull(RS!red_canpro), "Null", RS!red_canpro) & "," & IIf(IsNull(RS!red_cospro), "Null", RS!red_cospro) & "," & IIf(IsNull(RS!red_pctapr), "Null", RS!red_pctapr) & "," & IIf(IsNull(RS!red_pctcoc), "Null", RS!red_pctcoc) & "," & IIf(IsNull(RS!red_pctnut), "Null", RS!red_pctnut) & ")"
                           RS.MoveNext
                        Loop
                     End If
                     RS.Close: Set RS = Nothing
                  End If
               End If
           Next j
           If icopy = False Then
              '------- cerrar archivo mdb
              db7.Close
              '------- comprimir archivo
              AZ1.CreateZip mdir & sourcefilezip, ""
              AZ1.AddFile mdir & sourcefile, "", True, ""
              AZ1.Close
              '------- leer casino
              RS.Open "select * from b_clientes where cli_codigo='" & cencos & "'", vg_db, adOpenStatic
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
'                    Open dir_trabajo & "\sdxftp.ini" For Input As #1
'                    Do While Not EOF(1)
'                       Line Input #1, cpars
'                       If Mid(cpars, 1, InStr(cpars, ",") - 1) = "A" Then
'                          CHost = Mid(cpars, InStr(cpars, ",") + 1)
'                       ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "B" Then
'                          Cuser = Mid(cpars, InStr(cpars, ",") + 1)
'                       ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "C" Then
'                          Cpass = Mid(cpars, InStr(cpars, ",") + 1)
'                       End If
'                    Loop
'                    Close #1
                    a = oFTP.Version
                    oFTP.UseIEProxy = False
                    oFTP.Port = Cpuer '21
                    oFTP.HostName = CHost '"sgp.sodexhochile.cl" '"64.76.138.76" '"64.76.45.71"  'fg_Desencripta(TipoDato(cHost, ""))
                    oFTP.UserName = Cuser '"userftp" '"sodexho"   'fg_Desencripta(TipoDato(cUser, ""))
                    oFTP.password = Cpass '"*sdxo123*" '"shx873" 'fg_Desencripta(TipoDato(cPass, ""))
                    oFTP.Connect
                    If oFTP.IsConnected Then
                        lDir = oFTP.GetCurrentDirListing("*.*")
                        oFTP.SaveLastError ("aaa.xml")
'                        a = oFTP.ChangeRemoteDir("/casinos/bd")
                        a = oFTP.ChangeRemoteDir(Cdire)
                        oFTP.SaveLastError ("aaa.xml")
                        lDir = oFTP.GetCurrentDirListing("*.*")
                        oFTP.SaveLastError ("aaa.xml")
                        a = oFTP.PutFile(mdir & sourcefilezip, sourcefilezip)
                        oFTP.SaveLastError ("aaa.xml")
                        oFTP.Disconnect
                        If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                           fg_descarga
                           MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no se puede notificar por correo, ya que no tiene asignado el mail", vbInformation + vbOKOnly, Msgtitulo
                           fg_carga ""
                        Else
                           SendMail oMail, "Actualización maestro de recetas " & Format(Date, "dd/mm/yyyy"), "Se Informa que el maestro de recetas esta disponible para actualizar.", mdir & sourcefilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0
                        End If
                    Else
                       GoTo Man_Error
                    End If
                 ElseIf RS!cli_openvio = 2 Then
                    If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                       fg_descarga
                       MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, Msgtitulo
                       fg_carga ""
                    Else
                       SendMail oMail, "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), mdir & sourcefilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1
                    End If
                 End If
              End If
              RS.Close: Set RS = Nothing
           End If
           icopy = True
        End If
    Next i
    '------- verificar si existe archivo mdb destino si existe borrar
    If Dir(mdir & sourcefile) <> "" And Trim(sourcefile) <> "" Then Kill mdir & sourcefile
    '------- fin verificar si existe archivo mdb destino si existe borrar
    vg_db.CommitTrans
    fg_descarga
    Bar1(0).Visible = False: Bar1(1).Visible = False
    If Trim(sourcefile) <> "" Then MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
Case 3
    B_DieTip.Show 1
    Label2(8).Caption = "Todos": Label2(9).Caption = "Todos"
    If vg_filnomtippla <> "" Then Label2(9).Caption = vg_filnomtippla
    If vg_filnomcatdie <> "" Then Label2(8).Caption = vg_filnomcatdie
    If vg_opcion = 2 Then Exit Sub
    MoverDatoGrillaReceta
Case 5
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
fg_descarga
Bar1(0).Visible = False: Bar1(1).Visible = False
Bar1(0).Value = 0: Bar1(1).Value = 0
RS.Close: Set RS = Nothing
Select Case Err
Case 0
    vg_db.RollbackTrans
    MsgBox "Puede que no tenga salida a sitios FTP ó el servicio este sin conexión, conctatese con informatica. Proceso cancelado", vbInformation + vbOKOnly, Msgtitulo
    Exit Sub
Case 35764
    vg_db.RollbackTrans
    DoEvents
    For i = 1 To 1000000
    Next i
    Resume
Case 76
    vg_db.RollbackTrans
    Resume Next
Case -2147467259
    vg_db.RollbackTrans
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub
Case 3034
    vg_db.RollbackTrans: Exit Sub
End Select
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then vaSpread1(Index).Row = -1: vaSpread1(Index).Col = 1: vaSpread1(Index).text = IIf(vaSpread1(Index).Value = "1", "0", "1")
End Sub

Sub MoverDatoGrillaCasino()
On Error GoTo Man_Error
fg_carga ""
'------- Mover casinos
vaSpread1(0).MaxRows = 0
RS.Open "select cli_codigo, cli_nombre from b_clientes where cli_tipo=0 order by cli_nombre", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
      vaSpread1(0).Row = vaSpread1(0).MaxRows
              
      vaSpread1(0).Col = 2
      vaSpread1(0).TypeHAlign = TypeHAlignLeft
      vaSpread1(0).TypeSpin = False
      vaSpread1(0).TypeIntegerSpinInc = 1
      vaSpread1(0).TypeIntegerSpinWrap = False
      vaSpread1(0).text = RS!cli_codigo

      vaSpread1(0).Col = 3
      vaSpread1(0).TypeHAlign = TypeHAlignLeft
      vaSpread1(0).text = Trim(RS!cli_nombre)
      
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Sub MoverDatoGrillaReceta()
On Error GoTo Man_Error
fg_carga ""
'------- Mover productos
vaSpread1(1).MaxRows = 0
'RS.Open "select rec_codigo, rec_nombre From b_receta where (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
'         "and (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' order by rec_nombre", vg_db, adOpenStatic
If Check1(1).Value = 1 Then
   RS.Open "select distinct b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
           "FROM b_receta LEFT JOIN b_recetacasino ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
           "Where IsNull(b_recetacasino.rec_codrec) " & _
           "and   (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
           "and (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' " & _
           "order by b_receta.rec_nombre", vg_db, adOpenStatic
Else
   RS.Open "select distinct b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
           "FROM b_receta LEFT JOIN b_recetacasino ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
           "Where (rec_catdie = " & vg_filcatdie & " or " & vg_filcatdie & "=0) " & _
           "and (rec_tippla= " & vg_filtippla & " or " & vg_filtippla & "=0) and rec_tiprec='0' " & _
           "order by b_receta.rec_nombre", vg_db, adOpenStatic
End If
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
      vaSpread1(1).Row = vaSpread1(1).MaxRows
              
      vaSpread1(1).Col = 1
      vaSpread1(1).BackColor = IIf(IsNull(RS!rec_codrec), Shape1(0).FillColor, Shape1(1).FillColor)
      vaSpread1(1).text = "0"
      
      vaSpread1(1).Col = 2
      vaSpread1(1).TypeHAlign = TypeHAlignLeft
      vaSpread1(1).TypeSpin = False
      vaSpread1(1).TypeIntegerSpinInc = 1
      vaSpread1(1).TypeIntegerSpinWrap = False
      vaSpread1(1).BackColor = IIf(IsNull(RS!rec_codrec), Shape1(0).FillColor, Shape1(1).FillColor)
      vaSpread1(1).Lock = True
      vaSpread1(1).text = RS!rec_codigo

      vaSpread1(1).Col = 3
      vaSpread1(1).TypeHAlign = TypeHAlignLeft
      vaSpread1(1).BackColor = IIf(IsNull(RS!rec_codrec), Shape1(0).FillColor, Shape1(1).FillColor)
      vaSpread1(1).text = Trim(RS!rec_nombre)
      
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub vaSpread1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Or KeyCode = 13 Then Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then fptnombre(Index).text = IIf(KeyCode = 8, fptnombre(Index).text, fptnombre(Index).text & Chr(KeyCode)): fptnombre(Index).SetFocus: fptnombre(Index).SelStart = Len(fptnombre(Index).text)
End Sub
