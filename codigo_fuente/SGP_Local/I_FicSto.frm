VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_FicSto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha Stock"
   ClientHeight    =   5325
   ClientLeft      =   2895
   ClientTop       =   2580
   ClientWidth     =   8070
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   7920
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
         Left            =   120
         TabIndex        =   17
         Top             =   2520
         Width           =   7710
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
            ItemData        =   "I_FicSto.frx":0000
            Left            =   135
            List            =   "I_FicSto.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   540
            Width           =   4035
         End
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Una Familia"
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
            TabIndex        =   19
            Top             =   255
            Width           =   1500
         End
         Begin VB.OptionButton optTIPPRO 
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
            TabIndex        =   18
            Top             =   255
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   150
         TabIndex        =   3
         Top             =   1230
         Width           =   7665
         Begin VB.OptionButton Option1 
            Caption         =   "Uno"
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
            Index           =   2
            Left            =   120
            TabIndex        =   5
            Top             =   300
            Width           =   1395
         End
         Begin VB.OptionButton Option1 
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
            Index           =   3
            Left            =   3285
            TabIndex        =   4
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   1005
            TabIndex        =   6
            Top             =   690
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
            _ExtentY        =   556
            Enabled         =   0   'False
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
            NoSpecialKeys   =   0
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
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   2400
            Picture         =   "I_FicSto.frx":0004
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Producto"
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
            Left            =   165
            TabIndex        =   8
            Top             =   750
            Width           =   780
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2895
            TabIndex        =   7
            Top             =   690
            Width           =   4545
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   2940
            TabIndex        =   9
            Top             =   720
            Width           =   4530
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
         Height          =   1020
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   7725
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
            TabIndex        =   2
            Top             =   450
            Width           =   4035
         End
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1950
         TabIndex        =   10
         Top             =   3780
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ButtonStyle     =   3
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
         Text            =   "04/01/2005"
         DateCalcMethod  =   3
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   5310
         TabIndex        =   11
         Top             =   3795
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ButtonStyle     =   3
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
         Text            =   "04/01/2005"
         DateCalcMethod  =   3
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   2640
         TabIndex        =   16
         Top             =   2340
         Width           =   435
         _Version        =   196608
         _ExtentX        =   767
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "30"
         MaxValue        =   "60"
         MinValue        =   "1"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   4170
         Visible         =   0   'False
         Width           =   7650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Termino"
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
         Left            =   3705
         TabIndex        =   13
         Top             =   3870
         Width           =   1275
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio"
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
         Left            =   540
         TabIndex        =   12
         Top             =   3870
         Width           =   1065
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_FicSto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Public lc_Aux As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
    Me.Width = 8085
    If lc_Aux = "InfInt" Or lc_Aux = "ProMov" Then
        Me.Height = 4605
    End If
    If lc_Aux = "FicSto" Or lc_Aux = "DetCarInv" Then
        Me.Height = 5805
    End If
    'Me.Height = IIf(lc_Aux <> "FicSto", 4605, 5805)
    fg_centra Me
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    '-------> Cargar Combo Bodega
    CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
    '-------> Cargar Combo Familia Productos
    CargarDatoCombo Combo1, 1, "a_tipopro", "tip_", "TipPro", "N"
    If lc_Aux = "InfInt" Or lc_Aux = "ProMov" Then
        Frame1.Height = 3375
    End If
    If lc_Aux = "FicSto" Or lc_Aux = "DetCarInv" Then
        Frame1.Height = 4575
    End If
    'Frame1.Height = IIf(lc_Aux <> "FicSto", 3375, 4575)
    If lc_Aux = "FicSto" Then
       MsgTitulo = "Imprimir Ficha Stock"
       Me.Caption = "Ficha Stock"
       fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
       fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
       fpLongInteger1(0).Visible = False
    ElseIf lc_Aux = "ProMov" Then
       MsgTitulo = "Imprimir Producto Sin Movimiento"
       Me.Caption = "Producto Sin Movimiento"
       Label2(1).Caption = "Sin movimientos los últimos "
       Label2(2).Caption = "días "
       fpDateTime1(0).DateTimeFormat = UserDefined
       fpDateTime1(0).UserDefinedFormat = "dd/mm/yyyy"
       fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
       fpDateTime1(0).Visible = False
       fpDateTime1(1).Visible = False
       fpLongInteger1(0).Visible = True
       fpLongInteger1(0).Left = 3080
       fpLongInteger1(0).Top = 2580
       fpLongInteger1(0).Value = 30
       Frame3.Visible = False
       Label2(1).Top = 2670
       Label2(2).Top = 2670
    ElseIf lc_Aux = "DetCarInv" Then
       MsgTitulo = "Imprimir Detalle Cartola de Inventario"
       Me.Caption = "Detalle Cartola de Inventario"
       Label2(1).Caption = "Fecha Invent."
       fpDateTime1(0).DateTimeFormat = UserDefined
       fpDateTime1(0).UserDefinedFormat = "dd/mm/yyyy"
       fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
       Label2(2).Visible = False
       fpDateTime1(1).Visible = False
       fpLongInteger1(0).Visible = False
    ElseIf lc_Aux = "InfInt" Then
       MsgTitulo = "Imprimir Inflación Interna"
       Me.Caption = "Inflación Interna"
       Label2(1).Top = 2670
       Label2(2).Top = 2670
       fpDateTime1(0).Top = 2580
       fpDateTime1(1).Top = 2580
       Label2(1).Caption = "Periodo Desde"
       fpDateTime1(0).DateTimeFormat = UserDefined
       fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
       fpDateTime1(0).text = Format(Date, "mm/yyyy")
       Label2(2).Caption = "Periodo Hasta"
       fpDateTime1(1).UserDefinedFormat = "mm/yyyy"
       fpDateTime1(1).text = Format(Date, "mm/yyyy")
       fpLongInteger1(0).Visible = False
       Frame3.Visible = False
    ElseIf lc_Aux = "AnaCpf" Then
       MsgTitulo = "Imprimir Analisis de Consumo Precio Fijo"
       Me.Caption = "Analisis de Consumo Precio Fijo"
       Label2(1).Top = 2670
       Label2(2).Top = 2670
       fpDateTime1(0).Top = 2580
       fpDateTime1(1).Top = 2580
       Label2(1).Caption = "Periodo Desde"
       fpDateTime1(0).DateTimeFormat = UserDefined
       fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
       fpDateTime1(0).text = Format(Date, "mm/yyyy")
       Label2(2).Caption = "Periodo Hasta"
       fpDateTime1(1).UserDefinedFormat = "mm/yyyy"
       fpDateTime1(1).text = Format(Date, "mm/yyyy")
       fpLongInteger1(0).Visible = False
       Frame3.Visible = False
    End If
    Combo1(0).ListIndex = 0
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If Trim(fpDateTime1(0).text) = "" Or Trim(fpDateTime1(1).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If Trim(fpText1(0).text) = "" Then fpayuda(0).Caption = "": Exit Sub
RS.Open "SELECT DISTINCT a.pro_nombre FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_codigo = '" & fpText1(0).text & "'", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      fpayuda(0).Caption = RS!pro_nombre
      RS.MoveNext
   Loop
Else
   fpText1(0).text = "": fpayuda(0).Caption = ""
   MsgBox "Producto no existe...", vbExclamation + vbOKOnly, MsgTitulo
End If
RS.Close: Set RS = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = "": vg_nombre = ""
vg_left = fpayuda(0).Left + 4800
B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gpr"
B_TabEst.Show 1
Me.Refresh
If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
fpText1(Index) = Trim(vg_codigo)
fpayuda(Index).Caption = vg_nombre
fpText1_LostFocus 0
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 2
    Image1(0).Enabled = True
    fpText1(0).Enabled = True
Case 3
    Image1(0).Enabled = False
    fpText1(0).text = "": fpText1(0).Enabled = False
    fpayuda(0).Caption = ""
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim RS As New ADODB.Recordset
Select Case Button.Index
Case 1
    Dim codpro As String, codtip As Long
    codpro = "": codtip = 0
    If optTIPPRO(0).Value = True Then codtip = fg_codigocbo(Combo1, 1, 10, "")
    If Option1(2).Value = True Then
    
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

       RS.Open "SELECT * FROM b_productospmpdia WHERE ppd_codpro = '" & LimpiaDato(Trim(fpText1(0).text)) & "' AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(dEoM(CDate(vg_ciedia)), "yyyymmdd") & "", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: fpText1(0).text = "": fpayuda(0).Caption = "": MsgBox "No existe producto", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       codpro = RS!ppd_codpro
       RS.Close: Set RS = Nothing
    
    End If
    
    If lc_Aux = "FicSto" Then
       
       fpayuda(1).Visible = True
'       I_FichaStock Me, vg_codbod, codpro, Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd")
       I_FichaStockMichel Me, vg_codbod, codpro, codtip, Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd")
       fpayuda(1).Visible = False
    
    ElseIf lc_Aux = "ProMov" Then
       
       I_ProductoSinMovimiento vg_codbod, codpro, Format(fpDateTime1(0).text, "yyyymmdd"), Val(fpLongInteger1(0).Value)
    
    ElseIf lc_Aux = "InfInt" Or lc_Aux = "AnaCpf" Then
       
       If Format(fpDateTime1(0).text, "yyyymm") > Format(fpDateTime1(1).text, "yyyymm") Or Format(fpDateTime1(0).text, "mmyyyy") = Format(fpDateTime1(1).text, "mmyyyy") Then MsgBox "Periodo Desde debe ser menor al periodo hasta...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       
       If lc_Aux = "InfInt" Then
          
          I_InflacionInterna vg_codbod, codpro, Format(fpDateTime1(0).text, "yyyymm"), Format(fpDateTime1(1).text, "yyyymm")
       
       ElseIf lc_Aux = "AnaCpf" Then
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient

          '-------> Traer fecha del periodo desde
          RS.Open "SELECT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_periodo = " & Format(fpDateTime1(0).text, "yyyymm") & " and cie_estado = 0", vg_db, adOpenStatic
          If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Periodo inicial debe estar cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          
          If RS.State = 1 Then RS.Close
          RS.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient

          '-------> Traer fecha del periodo hasta
          RS.Open "SELECT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_periodo = " & Format(fpDateTime1(1).text, "yyyymm") & " and cie_estado = 0", vg_db, adOpenStatic
          If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Periodo final debe estar cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          RS.Close: Set RS = Nothing
          '-------> Validar que sea una diferencia de dos meses
          If Format(fpDateTime1(0).text, "yyyymm") > Format(fpDateTime1(1).text, "yyyymm") Then MsgBox "Mes inicial debe ser menor mes destino...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          If Format(fpDateTime1(1).text, "yyyymm") - Format(fpDateTime1(0).text, "yyyymm") <> 1 Then MsgBox "Debe ser una diferencia de un mes...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
          I_AnalisisConsumoPrecioFijo vg_codbod, codpro, Format(fpDateTime1(0).text, "yyyymm"), Format(fpDateTime1(1).text, "yyyymm")
       
       End If
       
    ElseIf lc_Aux = "DetCarInv" Then
       
       fpayuda(1).Visible = True
       I_DetCarInvMichel Me, vg_codbod, codpro, codtip, Format(fpDateTime1(0).text, "yyyymmdd"), Format(fpDateTime1(1).text, "yyyymmdd")
       fpayuda(1).Visible = False
    
    End If

Case 3
    
    Me.Hide
    Unload Me

End Select
End Sub
