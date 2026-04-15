VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_GenRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación archivos planos recetas"
   ClientHeight    =   7365
   ClientLeft      =   2775
   ClientTop       =   1845
   ClientWidth     =   15645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   15645
   ShowInTaskbar   =   0   'False
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   9600
      Top             =   6840
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   315
      Left            =   6450
      TabIndex        =   24
      Top             =   6900
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   556
      _Version        =   393217
      TextRTF         =   $"M_GenRec.frx":0000
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
      Height          =   6915
      Index           =   1
      Left            =   9690
      TabIndex        =   8
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame6 
         Height          =   435
         Index           =   1
         Left            =   2280
         TabIndex        =   33
         Top             =   6120
         Width           =   3150
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   34
            Top             =   135
            Width           =   3045
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   1
         Left            =   1320
         TabIndex        =   31
         Top             =   6120
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   32
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1155
         Index           =   3
         Left            =   570
         TabIndex        =   9
         Top             =   180
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "M_GenRec.frx":008B
            Left            =   1680
            List            =   "M_GenRec.frx":0095
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   180
            Width           =   2865
         End
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
            TabIndex        =   10
            Top             =   870
            Width           =   3180
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   12
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
            Index           =   2
            Left            =   150
            TabIndex        =   14
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
            Index           =   1
            Left            =   150
            TabIndex        =   13
            Top             =   285
            Width           =   1485
         End
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   1
         Left            =   660
         TabIndex        =   15
         Top             =   6570
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
         TabIndex        =   16
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
         MaxCols         =   4
         MaxRows         =   1
         SpreadDesigner  =   "M_GenRec.frx":00A9
         TextTip         =   2
         TextTipDelay    =   0
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
         TabIndex        =   22
         Top             =   1665
         Width           =   885
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
         TabIndex        =   21
         Top             =   1380
         Width           =   1650
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   8
         Left            =   2295
         TabIndex        =   20
         Top             =   1380
         Width           =   3465
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   9
         Left            =   2295
         TabIndex        =   19
         Top             =   1650
         Width           =   3465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enviados"
         Height          =   195
         Index           =   1
         Left            =   1950
         TabIndex        =   18
         Top             =   5880
         Width           =   660
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
         Caption         =   "No Enviados"
         Height          =   195
         Index           =   0
         Left            =   3165
         TabIndex        =   17
         Top             =   5880
         Width           =   915
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
      Height          =   6915
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   9465
      Begin VB.Frame Frame6 
         Height          =   435
         Index           =   0
         Left            =   1995
         TabIndex        =   29
         Top             =   5880
         Width           =   7110
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   30
            Top             =   135
            Width           =   7005
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   0
         Left            =   1080
         TabIndex        =   27
         Top             =   5880
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   28
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Envio x Servidor Sodexo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6720
         TabIndex        =   26
         Top             =   1320
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Envio x Outlook"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   4320
         TabIndex        =   25
         Top             =   1320
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   945
         Index           =   2
         Left            =   570
         TabIndex        =   1
         Top             =   180
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "M_GenRec.frx":0463
            Left            =   1680
            List            =   "M_GenRec.frx":046D
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   180
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   3
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
            Index           =   11
            Left            =   150
            TabIndex        =   5
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
            Index           =   0
            Left            =   150
            TabIndex        =   4
            Top             =   600
            Width           =   1470
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3975
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   1890
         Width           =   9255
         _Version        =   393216
         _ExtentX        =   16325
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
         MaxCols         =   5
         MaxRows         =   1
         SpreadDesigner  =   "M_GenRec.frx":0481
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   660
         TabIndex        =   7
         Top             =   6660
         Visible         =   0   'False
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   15645
      _ExtentX        =   27596
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   960
      OleObjectBlob   =   "M_GenRec.frx":0860
      Top             =   6840
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   1470
      OleObjectBlob   =   "M_GenRec.frx":0884
      Top             =   6810
   End
End
Attribute VB_Name = "M_GenRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS  As New ADODB.Recordset
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

Me.Height = 7740
Me.Width = 15735
Me.HelpContextID = vg_OpcM
MsgTitulo = "Generación archivos planos recetas"
fg_centra Me
Est = True
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.ToolTipText = "Enviar": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar Recetas":: BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Combo1(0).ListIndex = 1: Combo1(1).ListIndex = 1
vg_filcatdie = 0: vg_filtippla = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open "SELECT par_valor FROM a_param WITH (NOLOCK ) WHERE par_codigo='catdefecto'", vg_db, adOpenForwardOnly
If Not RS.EOF Then
   
   vg_filcatdie = RS!par_valor
   Label2(8).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
   
End If

RS.Close
Set RS = Nothing

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
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If Combo1(0).ItemData(Combo1(0).ListIndex) = 0 Then
       
       RS.Open "sgpadm_s_cliente_V02 53, '', '%" & UCase(LimpiaDato(fptnombre(0).text)) & "%'", vg_db, adOpenForwardOnly
       If RS.EOF Then vaSpread1(0).MaxRows = 0 Else vaSpread1(0).MaxRows = RS!nReg
    
    ElseIf Combo1(0).ItemData(Combo1(0).ListIndex) = 1 Then
       
       RS.Open "sgpadm_s_cliente_V02 54, '', '%" & UCase(LimpiaDato(fptnombre(0).text)) & "%'", vg_db, adOpenForwardOnly
       If RS.EOF Then vaSpread1(0).MaxRows = 0 Else vaSpread1(0).MaxRows = RS!nReg
    
    End If
    i = 1
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          vaSpread1(0).Row = i

          vaSpread1(0).Col = 1
          vaSpread1(0).text = "0"
          
          vaSpread1(0).Col = 2
          vaSpread1(0).text = RS!Cli_codigo
          
          vaSpread1(0).Col = 3
          vaSpread1(0).TypeHAlign = 0
          vaSpread1(0).text = Trim(RS!Cli_nombre)
          RS.MoveNext: i = i + 1
       
       Loop
    
    End If
    
    RS.Close
    Set RS = Nothing

Case 1
    
    If LimpiaDato(Trim(fptnombre(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    
    If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Then
       
       If Check1(1).Value = 1 Then
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT COUNT(distinct a.rec_codigo) AS nreg " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE b.rec_codrec Is Null " & _
                  "AND (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND Upper(a.rec_codigo) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "", vg_db, adOpenForwardOnly
          If RS.EOF Or RS!nReg = 0 Then
          
             vaSpread1(1).MaxRows = 0
           
          Else
          
             vaSpread1(1).MaxRows = RS!nReg
             
          End If
          
          RS.Close
          Set RS = Nothing
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT DISTINCT a.rec_codigo, a.rec_nombre, b.rec_codrec " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE b.rec_codrec Is Null " & _
                  "AND (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND Upper(a.rec_codigo) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "ORDER BY a.rec_codigo", vg_db, adOpenForwardOnly
'          RS.Open "sgpadm_s_receta_V06 16, 0, '%" & LimpiaDato(UCase(fptnombre(1).Text)) & "%', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'", vg_db, adOpenForwardOnly
'          If RS.EOF Then vaSpread1(1).MaxRows = 0 Else vaSpread1(1).MaxRows = RS!nreg
       Else
       
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT COUNT(distinct a.rec_codigo) AS nreg " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND   (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND Upper(a.rec_codigo) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "", vg_db, adOpenForwardOnly
          If RS.EOF Or RS!nReg = 0 Then
          
             vaSpread1(1).MaxRows = 0
             
          Else
          
             vaSpread1(1).MaxRows = RS!nReg
             
          End If
          RS.Close
          Set RS = Nothing
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT DISTINCT a.rec_codigo, a.rec_nombre, b.rec_codrec " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND   (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND   Upper(a.rec_codigo) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "ORDER BY a.rec_codigo", vg_db, adOpenForwardOnly
'          RS.Open "sgpadm_s_receta_V06 17, 0, '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'", vg_db, adOpenForwardOnly
'          If RS.EOF Then vaSpread1(1).MaxRows = 0 Else vaSpread1(1).MaxRows = RS!nreg
       End If
    ElseIf Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
       
       If Check1(1).Value = 1 Then
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT COUNT(distinct a.rec_codigo) AS nreg " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE b.rec_codrec Is Null " & _
                  "AND   (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND   (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = '1' " & _
                  "AND   Upper(a.rec_nombre) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "", vg_db, adOpenForwardOnly ', adOpenStatic
          
          If RS.EOF Or RS!nReg = 0 Then
          
             vaSpread1(1).MaxRows = 0
          
          Else
          
             vaSpread1(1).MaxRows = RS!nReg
            
          End If
          RS.Close
          Set RS = Nothing
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT DISTINCT a.rec_codigo, a.rec_nombre, b.rec_codrec " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE b.rec_codrec Is Null " & _
                  "AND   (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND   (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND   Upper(a.rec_nombre) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "ORDER BY a.rec_nombre", vg_db, adOpenForwardOnly ', adOpenStatic
'          RS.Open "sgpadm_s_receta_V06 14, 0, '%" & LimpiaDato(UCase(fptnombre(1).Text)) & "%', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'", vg_db, adOpenForwardOnly
'          If RS.EOF Then vaSpread1(1).MaxRows = 0 Else vaSpread1(1).MaxRows = RS!nreg
       Else
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          RS.Open "SELECT COUNT(distinct a.rec_codigo) AS nreg " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec  " & _
                  "WHERE (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND   (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND   Upper(a.rec_nombre) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "", vg_db, adOpenForwardOnly ', adOpenStatic
          If RS.EOF Or RS!nReg = 0 Then vaSpread1(1).MaxRows = 0 Else vaSpread1(1).MaxRows = RS!nReg
          RS.Close
          Set RS = Nothing
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient

          RS.Open "SELECT DISTINCT a.rec_codigo, a.rec_nombre, b.rec_codrec " & _
                  "FROM b_receta a WITH (NOLOCK ) LEFT JOIN b_recetacasino b WITH (NOLOCK ) ON a.rec_codigo = b.rec_codrec " & _
                  "WHERE (a.rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
                  "AND   (a.rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND a.rec_tiprec = '0' AND a.rec_indppr = 1 " & _
                  "AND   Upper(a.rec_nombre) LIKE '%" & LimpiaDato(UCase(fptnombre(1).text)) & "%' " & _
                  "ORDER BY a.rec_nombre", vg_db, adOpenForwardOnly ', adOpenStatic
'          RS.Open "sgpadm_s_receta_V06 15, 0, '%" & LimpiaDato(UCase(fptnombre(1).Text)) & "%', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'", vg_db, adOpenForwardOnly
'          If RS.EOF Then vaSpread1(1).MaxRows = 0 Else vaSpread1(1).MaxRows = RS!nreg
       End If
       
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
    
    RS.Close
    Set RS = Nothing
    
End Select

End Sub

Private Sub fptnombre_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 40 Or KeyCode = 34 And IRow > 0 Then vaSpread1(Index).SetFocus

End Sub

Private Sub TextDet1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet1(Index).text, ",")

If Index = 2 Then
   
   TextDet1(3).text = ""

ElseIf Index = 3 Then
   
   TextDet1(2).text = ""

End If

For i = 1 To vaSpread1(0).MaxRows
           
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 5
    vaSpread1(0).text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread1(0).Visible = False
    
    If Trim(TextDet1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1(0).MaxRows
           
           vaSpread1(0).Row = i
           vaSpread1(0).Col = Index
           indactivo = UCase(Trim(vaSpread1(0).Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1(0).Col = 2
           
           If indactivo = -1 And Trim(vaSpread1(0).text) <> "" Then
              
              vaSpread1(0).Col = 5
              
              If Val(vaSpread1(0).Value) <> 1 Then
                              
                 vaSpread1(0).Col = 1
              
                 If vaSpread1(0).RowHidden = True Then
                 
                    vaSpread1(0).RowHidden = False
                    vaSpread1(0).Col = 5
                    vaSpread1(0).text = 1
                 
                 Else
                 
                    vaSpread1(0).Col = 5
                    vaSpread1(0).text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1(0).Col = 5
              EstBuq = vaSpread1(0).Value
              vaSpread1(0).Col = 2
              
              If vaSpread1(0).RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1(0).RowHidden = True
                 
                 vaSpread1(0).Col = 5
                 vaSpread1(0).text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1(0).SetActiveCell Index + 1, 1
        vaSpread1(0).ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1(0).ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1(0).SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1(0).SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1(0).Sort -1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet1(Index).text) = "" Then
       
       For i = 1 To vaSpread1(0).MaxRows
           
           vaSpread1(0).Row = i
           If vaSpread1(0).RowHidden = True Then vaSpread1(0).RowHidden = False
           
           vaSpread1(0).Col = 5
           vaSpread1(0).text = 0
       
       Next
       
       vaSpread1(0).SetActiveCell Index, vaSpread1(0).SearchCol(Index, 0, vaSpread1(0).MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(0).SetActiveCell Index, 1
    
    End If
    
    vaSpread1(0).Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
fg_carga ""
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 2 Then
   
   TextDet2(3).text = ""

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""

End If

For i = 1 To vaSpread1(1).MaxRows
           
    vaSpread1(1).Row = i
    vaSpread1(1).Col = 4
    vaSpread1(1).text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread1(1).Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1(1).MaxRows
           
           vaSpread1(1).Row = i
           vaSpread1(1).Col = Index
           indactivo = UCase(Trim(vaSpread1(1).Value)) Like IIf(Index = 2, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1(1).Col = 2
           
           If indactivo = -1 And Trim(vaSpread1(1).text) <> "" Then
              
              vaSpread1(1).Col = 4
              
              If Val(vaSpread1(1).Value) <> 1 Then
                              
                 vaSpread1(1).Col = 1
              
                 If vaSpread1(1).RowHidden = True Then
                 
                    vaSpread1(1).RowHidden = False
                    vaSpread1(1).Col = 4
                    vaSpread1(1).text = 1
                 
                 Else
                 
                    vaSpread1(1).Col = 4
                    vaSpread1(1).text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1(1).Col = 4
              EstBuq = vaSpread1(1).Value
              vaSpread1(1).Col = 2
              
              If vaSpread1(1).RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1(1).RowHidden = True
                 
                 vaSpread1(1).Col = 4
                 vaSpread1(1).text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1(1).SetActiveCell Index + 1, 1
        vaSpread1(1).ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1(1).ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1(1).SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1(1).SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1(1).Sort -1, -1, vaSpread1(1).maxcols, vaSpread1(1).MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1(1).MaxRows
           
           vaSpread1(1).Row = i
           If vaSpread1(1).RowHidden = True Then vaSpread1(1).RowHidden = False
           
           vaSpread1(1).Col = 4
           vaSpread1(1).text = 0
       
       Next
       
       vaSpread1(1).SetActiveCell Index, vaSpread1(1).SearchCol(Index, 0, vaSpread1(1).MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(1).SetActiveCell Index, 1
    
    End If
    
    vaSpread1(1).Visible = True

End Select

fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim EstError As Boolean

Select Case Button.Index
Case 1
    
    If vaSpread1(0).MaxRows < 1 Or vaSpread1(1).MaxRows < 1 Then Exit Sub
    
    Dim i              As Long
    Dim j              As Long
    Dim CodRec         As Long
    Dim codzon         As Long
    Dim codtis         As Long
    Dim CodSeg         As Long
    Dim isel           As Boolean
    Dim icopy          As Boolean
    Dim cencos         As String
    Dim nomcencos      As String
    Dim aAprec         As String
    Dim aApprod        As String
    Dim sourcefile     As String
    Dim sourcefilezip  As String
    Dim destinofile    As String
    Dim destinofilezip As String
    Dim mdirserver     As String
    Dim lognarchsou    As String
    Dim codReg         As String
    Dim socsap         As String
    Dim tprod          As String
    Dim treceta        As String
    Dim dBo            As String
    Dim cDBI           As String
    Dim sobrec         As String
    Dim fso
    Dim CHost          As String
    Dim Cdire          As String
    Dim Cuser          As String
    Dim Cpass          As String
    Dim cecsac         As String
    Dim concco         As String
    Dim conreg         As String
    Dim logenv         As String
    Dim Cpuer          As Long
    Dim ccisac         As Long
    Dim codmun         As Long
    Dim codrgi         As Long
    Dim CodOpt         As String
    Dim TipoMinuta     As String
    Dim MyBufferReceta As String
    Dim MyBufferProd   As String
    Dim MyBufferSubSeg As String
    Dim MyBufferReg    As String
    Dim MyBufferCeco   As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    isel = False
    
    For i = 1 To vaSpread1(0).MaxRows
        
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        
        If vaSpread1(0).text = "1" Then
           
           isel = True: Exit For
        
        End If
    
    Next i
    
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un casino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    isel = False
    For i = 1 To vaSpread1(1).MaxRows
        
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then
           
           isel = True: Exit For
        
        End If
    
    Next i
    
    If isel = False Then MsgBox "Debe Seleccionar A lo menor una receta", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fg_carga ""
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    
    '------- Creo tabla temporal y chequeo si existe antes
    Toolbar1.Enabled = False
    Let MyBufferReceta = ""
    Let MyBufferReceta = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferReceta = MyBufferReceta & "<Receta>"
    For i = 1 To vaSpread1(1).MaxRows
        
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then
           
           vaSpread1(1).Col = 2
           MyBufferReceta = MyBufferReceta & " <Recetas"
           MyBufferReceta = MyBufferReceta & " CodReceta = " & Chr(34) & Trim(vaSpread1(1).text) & Chr(34)
           Let MyBufferReceta = MyBufferReceta & "/>"
        
        End If
    
    Next i
    Let MyBufferReceta = MyBufferReceta & "</Receta>"
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbGenerarProducto '" & MyBufferReceta & "'")
    Let MyBufferProd = ""
    Let MyBufferProd = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferProd = MyBufferProd & "<Producto>"
    
    Do While Not RS.EOF
       
       MyBufferProd = MyBufferProd & " <Productos"
       MyBufferProd = MyBufferProd & " CodProducto = " & Chr(34) & RS(0) & Chr(34)
       Let MyBufferProd = MyBufferProd & "/>"
       
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    Let MyBufferProd = MyBufferProd & "</Producto>"
    
    '------- Crear directorio si no existe
    mdirserver = Dir(dir_trabajo & "\" & "Actualizar", vbDirectory)
    If mdirserver = "" Then MkDir dir_trabajo & "\" & "Actualizar"
    mdirserver = dir_trabajo & "Actualizar" & "\"
    'Fin crear directorio si no existe
    
    '------- Generar base padre
    sourcefile = "recetageneral" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
    If Dir(mdirpc & sourcefile) <> "" Then Kill mdirpc & sourcefile ' borrar base datos si existe
    'Base de datos origen
    
    '------- Metodo acceso base de dato access dBo = dir_trabajo + BaseDeDato
    dBo = "'' [ODBC;Driver={SQL Server};Server=" + vg_SqlNSvr + ";Database=" + vg_SqlBase + ";UID=" + vg_SqlNUsr + ";PWD=" + vg_SqlPass + "]"
    DoEvents
    
    '-------> Contatenar centro costo
    concco = ""
    Let MyBufferCeco = ""
    Let MyBufferCeco = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferCeco = MyBufferCeco & "<Ceco>"
    For i = 1 To vaSpread1(0).MaxRows
        
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        
        If vaSpread1(0).text = "1" And vaSpread1(0).RowHidden = False Then
           
           vaSpread1(0).Col = 2
           concco = concco & "'" & vaSpread1(0).text & "',"
           MyBufferCeco = MyBufferCeco & " <Cecos"
           MyBufferCeco = MyBufferCeco & " CodCeco = " & Chr(34) & Trim(vaSpread1(0).text) & Chr(34)
           Let MyBufferCeco = MyBufferCeco & "/>"
        
        End If
    
    Next i
    Let MyBufferCeco = MyBufferCeco & "</Ceco>"
'    concco = Mid(concco, 1, Len(concco) - 1)
    
    '-------> Generar código subseg
    Dim csuse As String, auxseg As Long, auxreg As Long
    csuse = "0,"
    conreg = "0,"
    auxreg = 0
    auxseg = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT DISTINCT a.cli_subseg, b.crs_codreg FROM b_clientes a WITH (NOLOCK ), b_casinoregser b WITH (NOLOCK ) WHERE a.cli_codigo = b.crs_cencos AND a.cli_codigo IN (" & Mid(concco, 1, Len(concco) - 1) & ") AND cli_activo = '1' ORDER BY a.cli_subseg, b.crs_codreg")
    '-------> Xml SubSegmento
    Let MyBufferSubSeg = ""
    Let MyBufferSubSeg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferSubSeg = MyBufferSubSeg & "<SubSegmento>"
    '-------> Xml Regimen
    Let MyBufferReg = ""
    Let MyBufferReg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferReg = MyBufferReg & "<Regimen>"
    Do While Not RS.EOF
       
       If auxseg <> RS!cli_subseg Then
          
          csuse = csuse & "" & RS!cli_subseg & ","
          auxseg = RS!cli_subseg
          MyBufferSubSeg = MyBufferSubSeg & " <SubSegmentos"
          MyBufferSubSeg = MyBufferSubSeg & " CodSubSegmento = " & Chr(34) & RS!cli_subseg & Chr(34)
          Let MyBufferSubSeg = MyBufferSubSeg & "/>"
       
       End If
       
       If auxreg <> RS!crs_codreg Then
          
          conreg = conreg & "" & RS!crs_codreg & ","
          auxreg = RS!crs_codreg
          MyBufferReg = MyBufferReg & " <Regimenes"
          MyBufferReg = MyBufferReg & " CodRegimen = " & Chr(34) & RS!crs_codreg & Chr(34)
          Let MyBufferReg = MyBufferReg & "/>"
       
       End If
       
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    Let MyBufferSubSeg = MyBufferSubSeg & "</SubSegmento>"
    Let MyBufferReg = MyBufferReg & "</Regimen>"
    
    '-------> Inicio mover codigo ceco
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Dim CCeco As String
    CCeco = "'0',"
    Set RS = vg_db.Execute("SELECT DISTINCT isnull(a.cli_codigo,0) as cli_codigo FROM b_clientes a WITH (NOLOCK ) WHERE a.cli_tipo = 0 and a.cli_activo = '1' and a.cli_tipominuta = 3")
    Do While Not RS.EOF
       
       CCeco = CCeco & "'" & RS!Cli_codigo & "',"
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    
    '-------> Fin mover codigo ceco
    GenerarBaseEnviado mdirpc & sourcefile, tprod, treceta, dBo, 1, 0, csuse, conreg, CCeco, MyBufferProd, MyBufferReceta, MyBufferSubSeg, MyBufferReg, MyBufferCeco
    
    '-------> Crear archivo log de envio productos, recetas y planificación
    logenv = "mailSent" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".log"
    If Dir(dir_trabajo & logenv) <> "" Then Kill dir_trabajo & logenv ' borrar base datos si existe
    Open dir_trabajo & logenv For Output As #1 'Crear archivos de errores
    Close #1
        
    Bar1(0).Visible = True
    Bar1(1).Visible = True
    Bar1(0).Value = 0
    Bar1(1).Value = 0
    icopy = False
    EstError = True

    For i = 1 To vaSpread1(0).MaxRows
        
        Bar1(0).Value = Val((i / vaSpread1(0).MaxRows) * 100)
        vaSpread1(0).Row = i
        vaSpread1(0).Col = 1
        
        If vaSpread1(0).text = "1" And vaSpread1(0).RowHidden = False Then
           
           DoEvents
           vaSpread1(0).Col = 3
           nomcencos = Trim(vaSpread1(0).text)
           
           vaSpread1(0).Col = 2
           cencos = Trim(vaSpread1(0).text)
           Bar1(1).Value = 0
           
           vaSpread1(0).SetActiveCell 2, vaSpread1(0).Row
           icopy = True
           
           For j = 1 To vaSpread1(1).MaxRows
               
               DoEvents
               Bar1(1).Value = Val((j / vaSpread1(1).MaxRows) * 100)
               vaSpread1(1).Row = j
               vaSpread1(1).Col = 1
               
               If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then
                  
                  vaSpread1(1).Col = 2
                  icopy = True 'False
                  CodRec = vaSpread1(1).text
               
               End If
           
           Next j
           
           If icopy Then
              
              'Leer, insertar y rebrabar productos y recetas casinos
              vg_db.Execute ("sgpadm_Ins_XmlEnvioProductoCeco '" & MyBufferProd & "', '" & cencos & "' ")
              vg_db.Execute ("sgpadm_Ins_XmlEnvioRecetaCeco '" & MyBufferReceta & "', '" & cencos & "' ")
           
           End If
           
           DoEvents
           destinofile = "sgp" & (cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".kkk"
           destinofilezip = "sgp" & (cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           '------- Verificar si existe archivo mdb destino si existe borrar y copiar
           If Dir(mdirpc & destinofile) <> "" Then Kill mdirpc & destinofile
           FileCopy mdirpc & sourcefile, mdirpc & destinofile
           
           subseg = 0
           codReg = ""
           '---------------------------
           'Abrir base contrato
           '---------------------------
           DoEvents
           cDBI = mdirpc & destinofile
           Set dbi = New ADODB.Connection
           dbi.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cDBI & "' ;Persist Security Info=False"
           dbi.ConnectionTimeout = 3600
           dbi.CommandTimeout = 3600
           dbi.Open
           DoEvents
           codtis = 0
           CodSeg = 0
           socsap = ""
           
           '-------> Xml Regimen
           Let MyBufferReg = ""
           Let MyBufferReg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
           Let MyBufferReg = MyBufferReg & "<Regimen>"
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT DISTINCT a.cli_subseg, b.crs_codreg, a.cli_codzon, a.cli_codtis, a.cli_codseg, a.cli_socsap FROM b_clientes a WITH (NOLOCK ), b_casinoregser b WITH (NOLOCK ) WHERE a.cli_codigo = b.crs_cencos AND a.cli_codigo = '" & cencos & "' AND b.crs_cencos = '" & cencos & "'")
           If Not RS.EOF Then
              
              Do While Not RS.EOF
                 
                 subseg = IIf(IsNull(RS!cli_subseg), 0, RS!cli_subseg)
                 codzon = IIf(IsNull(RS!cli_codzon), 0, RS!cli_codzon)
                 codtis = IIf(IsNull(RS!cli_codtis), 0, RS!cli_codtis)
                 CodSeg = IIf(IsNull(RS!cli_codseg), 0, RS!cli_codseg)
                 socsap = IIf(IsNull(RS!cli_socsap), "", RS!cli_socsap)
                 codReg = codReg & RS!crs_codreg & ","
                 
                 '-------> Xml Regimen
                 MyBufferReg = MyBufferReg & " <Regimenes"
                 MyBufferReg = MyBufferReg & " CodRegimen = " & Chr(34) & RS!crs_codreg & Chr(34)
                 Let MyBufferReg = MyBufferReg & "/>"
                 RS.MoveNext
              
              Loop
           
           End If
           RS.Close
           Set RS = Nothing
           Let MyBufferReg = MyBufferReg & "</Regimen>"
           sobrec = ""
           TipoMinuta = ""
           
           '-------> generar sociedad sap
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT DISTINCT cli_tipominuta, cli_socsap, cli_sobrec, cli_ccisac, cli_cecsac, cli_codmun, cli_codreg FROM b_clientes WITH (NOLOCK ) WHERE cli_codigo = '" & cencos & "' and cli_tipo = 0 and cli_activo = '1'")
           If Not RS.EOF Then
              
              Do While Not RS.EOF
                 
                 socsap = IIf(IsNull(RS!cli_socsap), "", RS!cli_socsap)
                 sobrec = IIf(IsNull(RS!cli_sobrec), "", RS!cli_sobrec)
                 ccisac = IIf(IsNull(RS!cli_ccisac), 0, RS!cli_ccisac)
                 cecsac = IIf(IsNull(RS!cli_cecsac), "", RS!cli_cecsac)
                 codmun = IIf(IsNull(RS!cli_codmun), 0, RS!cli_codmun)
                 codrgi = IIf(IsNull(RS!cli_codreg), 0, RS!cli_codreg)
                 TipoMinuta = IIf(IsNull(RS!cli_tipominuta), "", RS!cli_tipominuta)
                 RS.MoveNext
              
              Loop
           
           End If
           RS.Close
           Set RS = Nothing
           
           '-------> Mover codigo optimun
           CodOpt = ""
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT isnull(Cecos_AX,'') as Cecos_AX FROM Cecos_Sap_AX WHERE Cecos_Sap = '" & cencos & "' and Sociedad_Sap = '" & socsap & "'")
           If Not RS.EOF Then
              
              CodOpt = RS!Cecos_AX
           
           End If
           RS.Close
           Set RS = Nothing
           
           '------- Generar tabla gramaje envio recetas
           If Trim(codReg) <> "" And subseg > 0 Then
              
              codReg = Mid(codReg, 1, Len(codReg) - 1)
              
              '-------> Generar servicio
              If RS.State = 1 Then RS.Close
              RS.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbServicioRegSer '" & cencos & "'")
              Do While Not RS.EOF
                 
                 dbi.Execute "INSERT INTO a_servicio (ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable) " & _
                             "VALUES (" & RS!Ser_codigo & ", '" & RS!ser_nombre & "', " & RS!ser_orden & ", '" & RS!ser_codsap & "', '" & RS!ser_facturable & "')"
                 RS.MoveNext
              
              Loop
              RS.Close
              Set RS = Nothing
              
              '-------> Generar estructura servicio
              If RS.State = 1 Then RS.Close
              RS.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbEstServicioRegSer '" & cencos & "'")
              Do While Not RS.EOF
                 
                 dbi.Execute "INSERT INTO a_estservicio (ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin) " & _
                             "VALUES (" & RS!ess_codser & ", " & RS!ess_codigo & ", '" & RS!ess_nombre & "', " & RS!ess_orden & ", " & RS!ess_codsec & ", " & RS!ess_racmin & ")"
                 RS.MoveNext
              
              Loop
              RS.Close
              Set RS = Nothing
              
              '------- Generar regimen
              If RS.State = 1 Then RS.Close
              RS.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbRegimenRegSer '" & cencos & "'")
              Do While Not RS.EOF
                 
                 dbi.Execute "INSERT INTO a_regimen (reg_codigo, reg_nombre) VALUES (" & RS!Reg_Codigo & ", '" & RS!reg_nombre & "')"
                 RS.MoveNext
              
              Loop
              RS.Close
              Set RS = Nothing
              
              If TipoMinuta = "3" Then

                 '-------> Generar gramaje aux
                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 
                 Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbTablaGramajeCeco '" & MyBufferReceta & "', '" & MyBufferReg & "', '" & cencos & "'")
                 Do While Not RS.EOF
                    
                    DoEvents
                    dbi.Execute "INSERT INTO b_tablagramaje (tgr_codreg, tgr_codrec, tgr_coding, tgr_codzon, tgr_codins, tgr_cantgr)  " & _
                                "VALUES (" & RS!tgc_codreg & ", " & RS!tgc_codrec & ", '" & RS!tgc_coding & "', 1, '" & RS!tgc_codins & "', " & RS!tgc_cantgr & ")"
                    RS.MoveNext
                 
                 Loop
                 RS.Close
                 Set RS = Nothing
              
              Else
                 
                 '-------> Generar gramaje aux
                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 
                 Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbTablaGramajeReceta '" & MyBufferReceta & "', '" & MyBufferReg & "', " & subseg & ", " & codzon & "")
                 Do While Not RS.EOF
                    
                    DoEvents
                    dbi.Execute "INSERT INTO b_tablagramaje (tgr_codreg, tgr_codrec, tgr_coding, tgr_codzon, tgr_codins, tgr_cantgr)  " & _
                                "VALUES (" & RS!tgr_codreg & ", " & RS!tgr_codrec & ", '" & RS!tgr_coding & "', " & RS!tgr_codzon & ", '" & RS!tgr_codins & "', " & RS!tgr_cantgr & ")"
                    RS.MoveNext
                    
                Loop
                RS.Close
                Set RS = Nothing
              
              End If

              dbi.Execute "INSERT INTO gra_receta (rec_codigo) SELECT DISTINCT tgr_codrec FROM b_tablagramaje"
              dbi.Execute "DELETE b_receta.* FROM b_receta INNER JOIN gra_receta ON b_receta.rec_codigo = gra_receta.rec_codigo"
              dbi.Execute "DELETE b_recetadet.* FROM b_recetadet INNER JOIN  gra_receta ON b_recetadet.red_codigo = gra_receta.rec_codigo"
              '------- Insertar receta desde tabla gramaje
              dbi.Execute "INSERT INTO b_receta (rec_codigo, rec_catdie, rec_tippla, rec_nombre, rec_nomfan, rec_metpre, rec_conche, rec_sugere, rec_basrac, rec_tiprec, rec_fecvig, rec_gruvul) SELECT DISTINCT a.rec_codigo, a.rec_catdie, a.rec_tippla, a.rec_nombre, a.rec_nomfan, '', a.rec_conche, a.rec_sugere, a.rec_basrac, a.rec_tiprec, a.rec_fecvig, a.rec_gruvul FROM b_recetaaux a, b_tablagramaje b WHERE a.rec_codigo=b.tgr_codrec"
              dbi.Execute "UPDATE b_receta INNER JOIN b_recetaaux ON b_receta.rec_codigo = b_recetaaux.rec_codigo SET b_receta.rec_metpre=b_recetaaux.rec_metpre"
              dbi.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) select distinct a.red_codigo, a.red_nroite, a.red_codpro, a.red_canpro, a.red_cospro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, 0 FROM b_recetadetaux a, b_tablagramaje b WHERE a.red_codigo=b.tgr_codrec"
              '------- Insertar receta desde tabla gramaje con origen regimen
              dbi.Execute "UPDATE b_receta INNER JOIN b_tablagramaje b ON b_receta.rec_codigo=b.tgr_codrec SET b_receta.rec_tiprec=1"
              dbi.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) SELECT DISTINCT a.red_codigo, a.red_nroite, a.red_codpro, a.red_canpro, a.red_cospro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, b.tgr_codreg FROM b_recetadetaux a, b_tablagramaje b WHERE a.red_codigo=b.tgr_codrec"
              
              If TipoMinuta = "3" Then
                 
                 dbi.Execute "UPDATE b_recetadet INNER JOIN b_tablagramaje ON (b_recetadet.red_tiprec=b_tablagramaje.tgr_codreg) AND (b_recetadet.red_codpro=b_tablagramaje.tgr_coding) AND (b_recetadet.red_codigo=b_tablagramaje.tgr_codrec) SET b_recetadet.red_codpro = [b_tablagramaje].[tgr_codins], b_recetadet.red_canpro = [b_tablagramaje].[tgr_cantgr]"
              
              Else
                 
                 dbi.Execute "UPDATE b_recetadet INNER JOIN b_tablagramaje ON (b_recetadet.red_tiprec=b_tablagramaje.tgr_codreg) AND (b_recetadet.red_codpro=b_tablagramaje.tgr_coding) AND (b_recetadet.red_codigo=b_tablagramaje.tgr_codrec) SET b_recetadet.red_codpro = [b_tablagramaje].[tgr_codins], b_recetadet.red_canpro = [b_tablagramaje].[tgr_cantgr] WHERE b_tablagramaje.tgr_codzon=" & codzon & ""
              
              End If
              
              dbi.Execute "UPDATE b_recetadet INNER JOIN b_ingrediente ON b_recetadet.red_codpro=b_ingrediente.ing_codigo SET b_recetadet.red_pctapr=[b_ingrediente].[ing_pctapr], b_recetadet.red_pctcoc=[b_ingrediente].[ing_pctcoc], b_recetadet.red_pctnut=[b_ingrediente].[ing_pctnut] WHERE b_recetadet.red_tiprec>0"
           
           End If
            
           DoEvents
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
              
           Set RS = vg_db.Execute("SELECT  ISNULL(cli_emailcontable, '') AS cli_emailcontable , isnull(cli_nomcontable, '') as cli_nomcontable, " & _
                    "CASE WHEN ISNULL(cli_subseg, 0) = 0 THEN 'N' " & _
                    "Else 'S' " & _
                    "END AS cli_subseg , " & _
                    "ISNULL(cli_emailenviopedido, '') AS cli_emailenviopedido , " & _
                    "CASE WHEN ISNULL(cli_gruvul, '') = 'S' THEN 'S' " & _
                    "Else 'N' " & _
                    "END AS cli_gruvul , " & _
                    "CASE WHEN ISNULL(cli_modpac, '') = 'S' THEN 'S' " & _
                    "Else 'N' " & _
                    "END AS cli_modpac , " & _
                    "ISNULL(cli_opgped, '') AS cli_opgped , " & _
                    "CASE WHEN ISNULL(cli_hipali, '') = 'S' THEN 'S' " & _
                    "Else 'N' " & _
                    "END AS cli_hipali , " & _
                    "ISNULL(cli_tipope, '') AS cli_tipope , " & _
                    "ISNULL(cli_minsre, '') AS cli_minsre , " & _
                    "ISNULL(cli_blockminteo, '') AS cli_blockminteo , " & _
                    "ISNULL(cli_blockminreal, '') AS cli_blockminreal , " & _
                    "ISNULL(cli_blockmincontrato, '') AS cli_blockmincontrato , " & _
                    "ISNULL(cli_blockmintrabajafinsemana, '') AS cli_blockmintrabajafinsemana FROM b_clientes With(NoLock) WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2) ")

              If Not RS.EOF Then
              
                 '------- Generar parametros ejecutivos contables
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('datcont', '" & Mid(RS!cli_nomcontable, 1, 40) & "', 'C', '" & RS!cli_emailcontable & "')"
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('5etapas', 'Casino 5 Etapas', 'C', '" & RS!cli_subseg & "')"
                 '-------> generar email envio pedido
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('emailenped', 'Email Envio Pedido', 'C', '" & RS!cli_emailenviopedido & "' )"
                 '------- Insert concepto grupo vulnerable a tabla a_param.
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('opgruvul', 'Opción Grupo Vulnerable', 'C', '" & RS!cli_gruvul & "')"
                 '------- Insert concepto modulo paciente.
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('modpac', 'Modulo Paciente', 'C', '" & RS!cli_modpac & "' )"
                 '-------> Insert concepto parametro proveedor
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('modprove', 'Parametro Modificar Proveedor', 'N', '0')"
                 '-------> Insert concepto generación pedido Web o SGP
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('gpedsgpweb', 'Parametro Generación Pedido x SGP o Web', 'C', '" & RS!cli_opgped & "' )"
                 '-------> Insert concepto Hipersensibilidad Alimentaria tabla a_param.
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('hipali', 'Opción Hipersensibilidad Alimentaria', 'C', '" & RS!cli_hipali & "')"
                 '-------> Insert concepto Tipo Operación tabla a_param.
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('tipope', 'Tipo Operación 0=Gravada:1=No Gravada', 'C', '" & RS!cli_tipope & "')"
                 '-------> Insert concepto Minuta Sitio Remoto tabla a_param.
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('minsre', 'Minuta Sitio Remoto 0=No:1=SI', 'C', '" & RS!cli_minsre & "')"
        
        
                 '-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA TEORICA 2013-01-11
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('blockmiteo', 'Bloqueo Minuta Teorica 0=No:1=SI', 'C', '" & RS!cli_blockminteo & "')"
                   
                 '-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA REAL 2013-01-11
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('blockmirea', 'Bloqueo Minuta Real 0=No:1=SI', 'C', '" & RS!cli_blockminreal & "')"
                   
                 '-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA (BLOQUEO MINUTA) 2013-01-11
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('blockmicon', 'Bloqueo Minuta 0=No:1=SI', 'C', '" & RS!cli_blockmincontrato & "')"
        
                 '-------> INSERT - MVA - PARAMETRO DE TRABAJA FIN SEMANA (BLOQUE MINUTA) 2013-03-08
                 dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) values ('trabfinsem', 'Trabaja Fin Semana 0=No:1=SI', 'C', '" & RS!cli_blockmintrabajafinsemana & "')"
              
              End If
              RS.Close
              Set RS = Nothing
              
              
           
'           '------- Generar parametros ejecutivos contables
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'datcont', mid(cli_nomcontable,1,40), 'C', cli_emailcontable FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT '5etapas', 'Casino 5 Etapas', 'C', iif(cli_subseg=0,'N','S') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
'           '-------> generar email envio pedido
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'emailenped', 'Email Envio Pedido', 'C', cli_emailenviopedido FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'           '------- Insert concepto grupo vulnerable a tabla a_param.
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'opgruvul', 'Opción Grupo Vulnerable', 'C', iif(cli_gruvul='S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
'           '------- Insert concepto modulo paciente.
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modpac', 'Modulo Paciente', 'C', iif(cli_modpac='S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
'           '-------> Insert concepto parametro proveedor
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modprove', 'Parametro Modificar Proveedor', 'N', '0' FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & cencos & "' AND (cli_tipo=0 OR cli_tipo=2)"
'           '-------> Insert concepto generación pedido Web o SGP
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT DISTINCT 'gpedsgpweb', 'Parametro Generación Pedido x SGP o Web', 'C', cli_opgped FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'           '-------> Insert concepto Hipersensibilidad Alimentaria tabla a_param.
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'hipali', 'Opción Hipersensibilidad Alimentaria', 'C', iif(cli_hipali = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'           '-------> Insert concepto Tipo Operación tabla a_param.
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'tipope', 'Tipo Operación 0=Gravada:1=No Gravada', 'C', cli_tipope FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'           '-------> Insert concepto Minuta Sitio Remoto tabla a_param.
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'minsre', 'Minuta Sitio Remoto 0=No:1=SI', 'C', cli_minsre FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'
'
'           '-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA TEORICA 2013-01-11
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'blockmiteo', 'Bloqueo Minuta Teorica 0=No:1=SI', 'C', cli_blockminteo FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'
'           '-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA REAL 2013-01-11
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'blockmirea', 'Bloqueo Minuta Real 0=No:1=SI', 'C', cli_blockminreal FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'
'           '-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA (BLOQUEO MINUTA) 2013-01-11
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'blockmicon', 'Bloqueo Minuta 0=No:1=SI', 'C', cli_blockmincontrato FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'
'           '-------> INSERT - MVA - PARAMETRO DE TRABAJA FIN SEMANA (BLOQUE MINUTA) 2013-03-08
'           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'trabfinsem', 'Trabaja Fin Semana 0=No:1=SI', 'C', cli_blockmintrabajafinsemana  FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'

           '------- Borrar tabla tipo servicio y segmento que no tenga relación con el contrato
           dbi.Execute "DELETE a_tiposervicio FROM a_tiposervicio WHERE tis_codigo NOT IN (" & codtis & ")"
           dbi.Execute "DELETE a_segmento FROM a_segmento WHERE seg_codigo NOT IN (" & CodSeg & ")"
           '-------> Borrar tabla casino envia sap
           dbi.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos NOT IN ('" & cencos & "')"
           '-------> Borrar tabla parametro codigo barra <> cencos
           dbi.Execute "DELETE a_par_codigo_barra FROM a_par_codigo_barra WHERE cli_codigo NOT IN ('" & cencos & "')"
          
           '------- Mover datos a la tabla centro de costo
           dbi.Execute "INSERT INTO a_cencos (cen_codigo, cen_socsap, cen_sobrec, cen_codmun, cen_ccisac, cen_cecsac, cen_codreg, cen_codopt) VALUES ('" & cencos & "', '" & socsap & "', '" & sobrec & "', " & codmun & ", " & ccisac & ", '" & cecsac & "', " & codrgi & ", '" & CodOpt & "')"
         
           '-------> Mover datos parametros despachos
           dbi.Execute "INSERT INTO b_paramdesp SELECT DISTINCT pad_cencos, pad_codtip AS pad_codtip, pad_tipo, pad_diaseg, pad_diario FROM b_parametrodespachos IN " & dBo & " WHERE pad_cencos = '" & cencos & "'"
           '-------> Mover datos días inhabiles
           dbi.Execute "INSERT INTO b_Fecha_Inhabiles SELECT DISTINCT CFI_CeCo, CFI_Fecha, CFI_Glosa FROM Cas_b_Fecha_Inhabiles IN " & dBo & " WHERE CFI_CeCo = '" & cencos & "'"
           '-------> Mover datos casino tipo actividades
           dbi.Execute "INSERT INTO b_casinotipoactividades SELECT DISTINCT cta_cencos, cta_tipact FROM b_casinotipoactividades IN " & dBo & " WHERE cta_cencos = '" & cencos & "'"
           '-------> Mover datos casino parametro stock
           dbi.Execute "INSERT INTO b_casinoparametrostock SELECT DISTINCT cps_cencos, cps_invsto, cps_reqmen, cps_porinv, cps_liscri, cps_diario, cps_ajuimp FROM b_casinoparametrostock IN " & dBo & " WHERE cps_cencos = '" & cencos & "'"
           '-------> Mover datos clase documento sap
           dbi.Execute "INSERT INTO a_clasedocsap SELECT DISTINCT cds_coddoc, cds_codreg, cds_cdosap FROM a_clasedocsap IN " & dBo & " WHERE cds_codreg = " & codrgi & ""
           
           dbi.Execute "DROP table b_recetaaux"
           dbi.Execute "DROP table b_recetadetaux"
           dbi.Execute "DROP table b_tablagramajeaux"
           dbi.Execute "DROP table b_tablagramajeauxceco"
           dbi.Execute "DROP table a_subsegmentoaux"
           dbi.Execute "DROP table tmp_receta"
           dbi.Execute "DROP table gra_receta"
           '----------------------------
           'Cerrar base contrato
           '----------------------------
           dbi.Close
           Set dbi = Nothing
'           Dim fso
'           Set fso = CreateObject("Scripting.FileSystemObject")
           
           If Dir(mdirpc & Mid(destinofile, 1, (Len(destinofile) - 3)) & "ldb") = "" And Trim(Environ("OS")) <> "" Then
              
              If Dir(mdirpc & "xxx.mdb") <> "" Then Kill mdirpc & "xxx.mdb"
              DBEngine.CompactDatabase mdirpc & destinofile, mdirpc & "xxx.mdb", dbLangGeneral
              Kill mdirpc & destinofile
              fso.MoveFile mdirpc & "xxx.mdb", mdirpc & destinofile
           
           End If
           '------- verificar si existe archivo zip destino si existe borrar
           If Dir(mdirpc & destinofilezip) <> "" Then Kill mdirpc & destinofilezip
           AZ1.CreateZip mdirpc & destinofilezip, "": AZ1.AddFile mdirpc & destinofile, "", True, "": AZ1.Close
           '------- verificar si existe archivo mdb destino si existe borrar
           If Dir(mdirpc & destinofile) <> "" Then Kill mdirpc & destinofile
           '------- leer casino
           DoEvents
           vg_GlosaEnvioCorreo = ""
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT * FROM b_clientes WITH (NOLOCK ) WHERE cli_codigo='" & cencos & "'")
           If Not RS.EOF Then
              
              If RS!cli_openvio = 1 Then
                 
                 '-------> Traer datos FTP
                 If RS1.State = 1 Then RS1.Close
                 RS1.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 
                 Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param WITH (NOLOCK ) WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("ftp")) & "%'")
                 If RS1.EOF Then fg_descarga: RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: Frame1(0).Enabled = True: Frame1(1).Enabled = True: Bar1(0).Visible = False: Bar1(1).Visible = False: MsgBox "No existe Parametrización FTP, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
                 Do While Not RS1.EOF
                    
                    If RS1!par_codigo = "ftpser" Then CHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftpdir" Then Cdire = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftpusu" Then Cuser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftppas" Then Cpass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    If RS1!par_codigo = "ftppue" Then Cpuer = fg_Desencripta(TipoDato(RS1!par_valor, ""))
                    RS1.MoveNext
                 
                 Loop
                 RS1.Close
                 Set RS1 = Nothing
                 a = oFTP.Version
                 oFTP.UseIEProxy = False
                 oFTP.Port = Cpuer '21
                 oFTP.HostName = CHost '"sgp.sodexhochile.cl" '"64.76.138.76" '"64.76.45.71"  'fg_Desencripta(TipoDato(cHost, ""))
                 oFTP.UserName = Cuser '"userftp" '"sodexho"   'fg_Desencripta(TipoDato(cUser, ""))
                 oFTP.password = Cpass '"*sdxo7528*" '"*sdxo123*" '"shx873" 'fg_Desencripta(TipoDato(cPass, ""))
                 oFTP.Connect
                 
                 If oFTP.IsConnected Then
                    
                    lDir = oFTP.GetCurrentDirListing("*.*")
                    oFTP.SaveLastError ("aaa.xml")
'                    a = oFTP.ChangeRemoteDir("/casinos/bd")
                    a = oFTP.ChangeRemoteDir(Cdire)
                    oFTP.SaveLastError ("aaa.xml")
                    lDir = oFTP.GetCurrentDirListing("*.*")
                    oFTP.SaveLastError ("aaa.xml")
                    a = oFTP.PutFile(mdirpc & destinofilezip, destinofilezip)
                    oFTP.SaveLastError ("aaa.xml")
                    oFTP.Disconnect
                    
                    If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                       
                       fg_descarga
                       MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no se puede notificar por correo, ya que no tiene asignado el mail", vbInformation + vbOKOnly, MsgTitulo
                       fg_carga ""
                    
                    Else
'                       SendMail1 oMail, "Actualización maestro de recetas " & Format(Date, "dd/mm/yyyy"), "Se Informa que el maestro de recetas esta disponible para actualizar.", mdirpc & sourcefilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0, logenv
                        
                        If Option1(0).Value Then
                           
                           SendMailOutlook oMail, "Actualización maestro de recetas " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy") & " Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar", mdirpc & destinofilezip, Trim(RS!Cli_nombre), Trim(RS!cli_email), 1, logenv
                        
                        Else
                           
                           SendMail2 oMail, "Actualización maestro de recetas " & Format(Date, "dd/mm/yyyy"), "Se Informa que el maestro de recetas esta disponible para actualizar.", mdirpc & sourcefilezip, Trim(RS!Cli_nombre), Trim(RS!cli_email), 0, logenv
                        
                        End If
                    
                    End If
                 
                 Else
                    
                    GoTo Man_Error
                 
                 End If
                 
              ElseIf RS!cli_openvio = 2 Then
                 
                 If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                    
                    fg_descarga
                    MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, MsgTitulo
                    fg_carga ""
                 
                 Else
'                    SendMail1 oMail, "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy") & " Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar", mdirpc & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1, logenv
                    
                    If Option1(0).Value Then
                       
                       SendMailOutlook oMail, "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy") & " Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar", mdirpc & destinofilezip, Trim(RS!Cli_nombre), Trim(RS!cli_email), 1, logenv
                    Else
                       
                       SendMail2 oMail, "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo recetas " & Format(Date, "dd/mm/yyyy") & " Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar", mdirpc & destinofilezip, Trim(RS!Cli_nombre), Trim(RS!cli_email), 1, logenv
                    
                    End If
                 
                 End If
                 
              End If
           
              vaSpread1(0).Col = 4
              vaSpread1(0).text = ""
              If Trim(vg_GlosaEnvioCorreo) <> "" Then
                 
                 vaSpread1(0).text = vg_GlosaEnvioCorreo
                 EstError = False
                 
              Else
              
                 vaSpread1(0).text = "Envió exitoso"
                 
              End If
           
           End If
           RS.Close
           Set RS = Nothing
           
        End If
        
    Next i
    '------- verificar si existe archivo mdb destino si existe borrar
    If Dir(mdirpc & sourcefile) <> "" And Trim(sourcefile) <> "" Then Kill mdirpc & sourcefile
    '------- fin verificar si existe archivo mdb destino si existe borrar
    
    '------- Copiar archivos access \\SQLDES\CXCASINO, luego borrar archivos del PC
    fso.CopyFile mdirpc & "sgp*.zip", mdirserver, True
    If Dir(mdirpc & "sgp*.zip") <> "" Then Kill mdirpc & "sgp*.zip"
    '------- Fin copiar archivos access \\SQLDES\CXCASINO, luego borrar archivos del PC
    fg_descarga
    Bar1(0).Visible = False: Bar1(1).Visible = False
'    If Trim(sourcefile) <> "" Then MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    
    If Not EstError Then
        
        MsgBox "Generación finalizado con problema, revise columna de observación de la grilla Nº1 casinos", vbInformation + vbOKOnly, MsgTitulo
    
    Else
        
        MsgBox "Generación finalizado sin problema", vbInformation + vbOKOnly, MsgTitulo
    
    End If
    
    Frame1(0).Enabled = True: Frame1(1).Enabled = True
    Toolbar1.Enabled = True

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
Bar1(0).Visible = False
Bar1(1).Visible = False
Bar1(0).Value = 0
Bar1(1).Value = 0
Frame1(0).Enabled = True: Frame1(1).Enabled = True
Toolbar1.Enabled = True

If Err = 521 Or Err = 3183 Or Err = 3704 Or Err = 424 Or Err = 55 Or Err = 53 Or Err = -2147467259 Then

    Resume Next

End If

If Err <> 76 And Err <> 3704 Then

   If RS.State = 1 Then RS.Close

End If

Select Case Err
Case 0
    
    MsgBox "Puede que no tenga salida a sitios FTP ó el servicio este sin conexión, conctatese con informatica. Proceso cancelado", vbInformation + vbOKOnly, MsgTitulo
    Exit Sub

Case 35764
    
    DoEvents
    For i = 1 To 1000000
    Next i
    Resume

Case 76
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    Exit Sub

Case -2147467259
    
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub

Case 3034
    
    Exit Sub

Case 3704
   
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    Exit Sub

End Select
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_BlockSelected(Index As Integer, ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    Dim i As Long
    vaSpread1(Index).Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1(Index).Row = i
        
        If vaSpread1(Index).RowHidden = False Then
            
           vaSpread1(Index).Value = IIf(vaSpread1(Index).Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1(Index).MaxRows
        
        vaSpread1(Index).Row = i
        
        If vaSpread1(Index).RowHidden = False Then
            
           vaSpread1(Index).Value = IIf(vaSpread1(Index).Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread1(Index).Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1(Index).Row = i
        If vaSpread1(Index).RowHidden = False Then
            
           vaSpread1(Index).Value = IIf(vaSpread1(Index).Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Col = 1 And Row = 0 Then
   
   vaSpread1(Index).Row = -1
   vaSpread1(Index).Col = 1
   vaSpread1(Index).text = IIf(vaSpread1(Index).Value = "1", "0", "1")
   
   If Index = 0 Then
      
      vaSpread1(Index).Col = 4
      vaSpread1(Index).text = ""
   
   End If

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatoGrillaCasino()

On Error GoTo Man_Error

fg_carga ""
'-------> Mover casinos
vaSpread1(0).MaxRows = 0
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "sgpadm_s_cliente_V02 7, '', ''", vg_db, adOpenForwardOnly ', adOpenStatic
If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      If Mid(RS!Cli_codigo, 1, 3) <> "PRO" And Mid(RS!Cli_codigo, 1, 3) <> "DCL" And Mid(RS!Cli_codigo, 1, 3) <> "PPT" And Mid(RS!Cli_codigo, 1, 3) <> "DED" And UCase(Mid(RS!Cli_nombre, 1, 9)) <> "PROPUESTA" And UCase(Mid(RS!Cli_nombre, 1, 6)) <> "DISEÑO" Then
   
          vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
          vaSpread1(0).Row = vaSpread1(0).MaxRows
                  
          vaSpread1(0).Col = 2
          vaSpread1(0).TypeHAlign = TypeHAlignLeft
          vaSpread1(0).TypeSpin = False
          vaSpread1(0).TypeIntegerSpinInc = 1
          vaSpread1(0).TypeIntegerSpinWrap = False
          vaSpread1(0).text = RS!Cli_codigo
    
          vaSpread1(0).Col = 3
          vaSpread1(0).TypeHAlign = TypeHAlignLeft
          vaSpread1(0).text = Trim(RS!Cli_nombre)
          
          vaSpread1(0).Col = 4
          ' Define cell type as edit
          vaSpread1(0).CellType = CellTypeEdit '= SS_CELL_TYPE_EDIT
          ' Display multiple lines of data
          vaSpread1(0).TypeEditMultiLine = True
          vaSpread1(0).TypeMaxEditLen = 10000
          vaSpread1(0).text = ""
      
      End If
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDatoGrillaReceta()

On Error GoTo Man_Error

fg_carga ""
'Mover productos
vaSpread1(1).MaxRows = 0
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Check1(1).Value = 1 Then
   
   RS.Open "SELECT DISTINCT b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
           "FROM b_receta WITH (NOLOCK ) LEFT JOIN b_recetacasino WITH (NOLOCK ) ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
           "WHERE b_recetacasino.rec_codrec Is Null " & _
           "AND  (rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
           "AND  (rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND rec_tiprec = '0' AND b_receta.rec_indppr = 1 " & _
           "ORDER BY b_receta.rec_nombre", vg_db, adOpenForwardOnly ', adOpenStatic
'   RS.Open "sgpadm_s_receta_V06 12, 0, '', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'", vg_db, adOpenForwardOnly

Else
   
   RS.Open "SELECT DISTINCT b_receta.rec_codigo, b_receta.rec_nombre, b_recetacasino.rec_codrec " & _
           "FROM b_receta WITH (NOLOCK ) LEFT JOIN b_recetacasino WITH (NOLOCK ) ON b_receta.rec_codigo = b_recetacasino.rec_codrec " & _
           "WHERE (rec_catdie = " & vg_filcatdie & " OR " & vg_filcatdie & " = 0) " & _
           "AND   (rec_tippla = " & vg_filtippla & " OR " & vg_filtippla & " = 0) AND rec_tiprec = '0'  AND b_receta.rec_indppr = 1 " & _
           "ORDER BY b_receta.rec_nombre", vg_db, adOpenForwardOnly ', adOpenStatic
'   RS.Open "sgpadm_s_receta_V06 13, 0, '', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'", vg_db, adOpenForwardOnly

End If

Dim i As Long

If Not RS.EOF Then
   
   Do While Not RS.EOF

'      If i > 7155 Then
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

      vaSpread1(1).Col = 4
      vaSpread1(1).text = 0

'      End If
'      i = i + 1
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub

Private Sub vaSpread1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

If KeyCode = 27 Or KeyCode = 13 Then Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then fptnombre(Index).text = IIf(KeyCode = 8, fptnombre(Index).text, fptnombre(Index).text & Chr(KeyCode)): fptnombre(Index).SetFocus: fptnombre(Index).SelStart = Len(fptnombre(Index).text)

End Sub


