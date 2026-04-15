VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_Stock 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Stock"
   ClientHeight    =   2565
   ClientLeft      =   3345
   ClientTop       =   3015
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   9090
      Begin VB.Frame Frame3 
         Caption         =   "Familia Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   4575
         TabIndex        =   9
         Top             =   1140
         Width           =   4350
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
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
            ItemData        =   "I_Stock.frx":0000
            Left            =   135
            List            =   "I_Stock.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   4035
         End
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Un Tipo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   11
            Top             =   255
            Width           =   1005
         End
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3330
            TabIndex        =   10
            Top             =   255
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Bodega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   150
         TabIndex        =   5
         Top             =   120
         Width           =   8775
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
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
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   390
            Width           =   4035
         End
         Begin EditLib.fpDateTime Date1 
            Height          =   330
            Index           =   0
            Left            =   7230
            TabIndex        =   7
            Top             =   135
            Visible         =   0   'False
            Width           =   1470
            _Version        =   196608
            _ExtentX        =   2593
            _ExtentY        =   582
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
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "04/09/2004"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Stock"
            Height          =   195
            Index           =   1
            Left            =   7710
            TabIndex        =   8
            Top             =   225
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Cuenta Contable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   150
         TabIndex        =   1
         Top             =   1140
         Width           =   4350
         Begin VB.OptionButton optCUENTA 
            Caption         =   "Todas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3330
            TabIndex        =   4
            Top             =   255
            Width           =   855
         End
         Begin VB.OptionButton optCUENTA 
            Caption         =   "Una Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   3
            Top             =   255
            Width           =   1425
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
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
            ItemData        =   "I_Stock.frx":0004
            Left            =   135
            List            =   "I_Stock.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   540
            Width           =   4035
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_Stock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset

Private Sub Combo1_Click(Index As Integer)
Dim v_codbod  As Long, sqlBO As String
Select Case Index
Case 0
    v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
    sqlBO = v_codbod
    Combo1(2).Clear
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT cta.cta_nombre, cta.cta_codigo FROM a_ctacontable cta, b_productos pro, b_bodegas bod " & _
             "WHERE pro.pro_codigo = bod.bod_codpro AND cta.cta_codigo = pro.pro_ctacon AND bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " " & _
             "GROUP BY cta.cta_nombre, cta.cta_codigo ORDER BY cta_nombre", vg_db, adOpenStatic
    Do While Not RS1.EOF
        Combo1(2).AddItem RS1!cta_nombre & Space(150) & "(" & Space(10 - Len(Trim(RS1!cta_codigo))) & Trim(RS1!cta_codigo) & ")"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Width = 9210
Me.Height = 3030
MsgTitulo = "Imprimir Toma de Inventario"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
'-------> Cargar Combo Familia Productos
CargarDatoCombo Combo1, 1, "a_tipopro", "tip_", "TipPro", "N"
Combo1(0).ListIndex = 0
optTIPPRO(1).Value = True
optCUENTA(1).Value = True
End Sub

Private Sub optTIPPRO_Click(Index As Integer)
    Combo1(1).Enabled = IIf(Index = 0, True, False)
    Combo1(1).ListIndex = IIf(Index = 0 And Combo1(1).listcount > 0, 0, -1)
End Sub

Private Sub optCUENTA_Click(Index As Integer)
    Combo1(2).Enabled = IIf(Index = 0, True, False)
    Combo1(2).ListIndex = IIf(Index = 0 And Combo1(2).listcount > 0, 0, -1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    I_StockxFecha Me  'aApPrin
Case 3
    Me.Hide
    Unload Me
End Select
End Sub
