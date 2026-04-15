VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form I_Traspa 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Traspasos"
   ClientHeight    =   7245
   ClientLeft      =   3105
   ClientTop       =   2010
   ClientWidth     =   7845
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7650
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Informe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   7350
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   140
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   285
            Width           =   7065
         End
      End
      Begin VB.Frame Frame6 
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
         Height          =   1215
         Left            =   140
         TabIndex        =   15
         Top             =   5160
         Width           =   7290
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
            Left            =   6285
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
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
            TabIndex        =   16
            Top             =   360
            Width           =   1395
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   135
            TabIndex        =   18
            Top             =   630
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2055
            TabIndex        =   19
            Top             =   630
            Width           =   5025
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   1530
            Picture         =   "I_Traspa.frx":0000
            Top             =   525
            Width           =   480
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   2100
            TabIndex        =   20
            Top             =   660
            Width           =   5020
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   140
         TabIndex        =   9
         Top             =   1095
         Width           =   7365
         Begin EditLib.fpDateTime Fecha 
            Height          =   315
            Index           =   0
            Left            =   1635
            TabIndex        =   10
            Top             =   255
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
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
            AllowNull       =   -1  'True
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
            Text            =   ""
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
            Appearance      =   2
            BorderDropShadow=   0
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
         Begin EditLib.fpDateTime Fecha 
            Height          =   315
            Index           =   1
            Left            =   5835
            TabIndex        =   11
            Top             =   225
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
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
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   2
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
            Appearance      =   2
            BorderDropShadow=   0
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
            Caption         =   "Fecha Desde"
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
            Left            =   140
            TabIndex        =   13
            Top             =   300
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta"
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
            Left            =   4320
            TabIndex        =   12
            Top             =   300
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
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
         Height          =   795
         Index           =   0
         Left            =   140
         TabIndex        =   7
         Top             =   1935
         Width           =   7350
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   140
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   285
            Width           =   7065
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1260
         Left            =   140
         TabIndex        =   3
         Top             =   2790
         Width           =   7305
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
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   360
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
            Index           =   1
            Left            =   6285
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   135
            TabIndex        =   4
            Top             =   705
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   0
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
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   20
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
         Begin VB.Label label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2115
            TabIndex        =   5
            Top             =   705
            Width           =   5010
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   1620
            Picture         =   "I_Traspa.frx":030A
            Top             =   600
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   2160
            TabIndex        =   6
            Top             =   750
            Width           =   5010
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Traspasos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   140
         TabIndex        =   1
         Top             =   4200
         Width           =   7290
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   140
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   330
            Width           =   6975
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   7845
      _ExtentX        =   13838
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_Traspa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim Msgtitulo As String, est As Boolean

Private Sub Combo1_Click(Index As Integer)
If est Then Exit Sub
Frame5.Enabled = True
Select Case Index
Case 0
    Frame6.Enabled = IIf(fg_codigocbo(Combo1, 0, 2, 0) = "01", False, True)
    If fg_codigocbo(Combo1, 0, 2, 0) = "01" Then fpText1(1).text = "": fpayuda(0).Caption = "": Option1(2).Value = False: Option1(3).Value = True
    If fg_codigocbo(Combo1, 0, 2, 0) = "03" Then Combo1(2).ListIndex = 1:        Frame5.Enabled = False
Case 2
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Local Error GoTo Error_Partida
'------- Crear botones para botonera
'Dim btnX As Button
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
EspFecha Fecha(0)
EspFecha Fecha(1)
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Me.Height = 7755
Me.Width = 7965
Msgtitulo = "Traspasos"
fg_centra Me
est = True
'------- Carga Combos con Información Necesaria
With Combo1(0)
    .Clear
    .AddItem "Resumen Traspasos por Periodo" & Space(150) & "(01)"
    .AddItem "Detalle Traspasos por Periodo" & Space(150) & "(02)"
    .AddItem "Diferencia entre Contrato" & Space(150) & "(03)"
    If .listcount > 0 Then .ListIndex = 0
End With
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"
If Combo1(1).listcount > 0 Then Combo1(1).ListIndex = 0
With Combo1(2)
    .Clear
    .AddItem "TODOS" & Space(150) & "(01)"
    .AddItem "ENTRADAS" & Space(150) & "(02)"
    .AddItem "SALIDAS" & Space(150) & "(03)"
    If .listcount > 0 Then .ListIndex = 0
End With
Frame6.Enabled = False
est = False
Exit Sub
Error_Partida:
    MsgBox "Error: " & Err.Number & "-" & Err.Description, vbExclamation, Msgtitulo
End Sub

Private Sub fpText1_Change(Index As Integer)
Select Case Index
Case 0
    RS.Open RutinaLectura.Cliente(5, LimpiaDato(Trim(fpText1(0).text)), ""), vg_db, adOpenStatic
    If RS.EOF Then Label2(0).Caption = "": RS.Close: Set RS = Nothing: Exit Sub
    Label2(0).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
Case 1
    RS.Open "SELECT DISTINCT a.pro_nombre FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro<1) AND c.cli_codigo='" & MuestraCasino(1) & "' AND (b.tis_codigo=a.pro_maepro OR a.pro_maepro<1) AND a.pro_codigo='" & LimpiaDato(Trim(fpText1(1).text)) & "'", vg_db, adOpenStatic
    If RS.EOF Then fpayuda(0).Caption = "": RS.Close: Set RS = Nothing: Exit Sub
    fpayuda(0).Caption = RS!pro_nombre
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_codigo = 0
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Traspaso"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(0).text = Trim(vg_codigo)
    Label2(0).Caption = vg_nombre
Case 1
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 4800
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gpr"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(1).text = Trim(vg_codigo)
    fpayuda(0).Caption = vg_nombre
End Select
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    fpText1(0).Enabled = True: Image1(0).Enabled = True
Case 1
    fpText1(0).Enabled = False: Image1(0).Enabled = False
    fpText1(0).text = "": Label2(0).Caption = ""
Case 2
    fpText1(1).Enabled = True: Image1(1).Enabled = True
Case 3
    fpText1(1).Enabled = False: Image1(1).Enabled = False
    fpText1(1).text = "": fpayuda(0).Caption = ""
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim paso As Long
paso = 0
Select Case Button.Index
Case 1 '------- Previsualizar
    If IsDate(Fecha(0)) = False Then MsgBox "Rango de fechas no valido...", vbExclamation, Msgtitulo: Exit Sub
    If IsDate(Fecha(1)) = False Then MsgBox "Rango de fechas no valido...", vbExclamation, Msgtitulo: Exit Sub
    If CDate(Fecha(0).text) > CDate(Fecha(1).text) Then MsgBox "Rango de fechas no valido...", vbExclamation, Msgtitulo: Exit Sub
    If Trim(Label2(0).Caption) = "" And Option1(0).Value = True Then MsgBox "Contrato no valido...", vbExclamation, Msgtitulo: Exit Sub
    If Trim(fpayuda(0).Caption) = "" And Option1(2).Value = True Then MsgBox "Producto no valido...", vbExclamation, Msgtitulo: Exit Sub
    If Combo1(0).ListIndex = -1 Then MsgBox "Bodega no valida...", vbExclamation, Msgtitulo: Exit Sub
    If fg_codigocbo(Combo1, 0, 2, 0) = "01" Then
       I_ResumenTraspasos Format(Fecha(0).text, "dd/mm/yyyy"), Format(Fecha(1).text, "dd/mm/yyyy"), fg_codigocbo(Combo1, 1, 2, 0), Trim(fpText1(0).text), IIf(fg_codigocbo(Combo1, 2, 2, 0) = "01", -1, IIf(fg_codigocbo(Combo1, 2, 2, 0) = "02", 1, 0))
    ElseIf fg_codigocbo(Combo1, 0, 2, 0) = "02" Then
       I_DetalleTraspasos Format(Fecha(0).text, "dd/mm/yyyy"), Format(Fecha(1).text, "dd/mm/yyyy"), fg_codigocbo(Combo1, 1, 2, 0), Trim(fpText1(0).text), IIf(fg_codigocbo(Combo1, 2, 2, 0) = "01", -1, IIf(fg_codigocbo(Combo1, 2, 2, 0) = "02", 1, 0)), LimpiaDato(Trim(fpText1(1).text))
    ElseIf fg_codigocbo(Combo1, 0, 2, 0) = "03" Then
       I_DiferenciaTraspasos Format(Fecha(0).text, "dd/mm/yyyy"), Format(Fecha(1).text, "dd/mm/yyyy"), fg_codigocbo(Combo1, 1, 2, 0), Trim(fpText1(0).text), 1, LimpiaDato(Trim(fpText1(1).text))
    End If
Case 3 '------- Salir
    Me.Hide
    Unload Me
End Select
End Sub
