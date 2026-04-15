VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_GenPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Pedidos"
   ClientHeight    =   8580
   ClientLeft      =   1455
   ClientTop       =   1755
   ClientWidth     =   18300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   18300
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   945
      Index           =   0
      Left            =   3360
      TabIndex        =   16
      Top             =   240
      Width           =   8535
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1500
         TabIndex        =   17
         Top             =   405
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
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
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   3165
         TabIndex        =   19
         Top             =   405
         Width           =   4935
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   2730
         Picture         =   "M_GenPed.frx":0000
         Top             =   315
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   18
         Top             =   480
         Width           =   735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3210
         TabIndex        =   20
         Top             =   450
         Width           =   4935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   8580
      Left            =   17790
      TabIndex        =   0
      Top             =   0
      Width           =   510
      _ExtentX        =   900
      _ExtentY        =   15134
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   615
      Left            =   0
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   1215
      _Version        =   393216
      _ExtentX        =   2143
      _ExtentY        =   1085
      _StockProps     =   64
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
      SpreadDesigner  =   "M_GenPed.frx":030A
   End
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
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
            Picture         =   "M_GenPed.frx":04EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6855
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   12091
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Resumen Pedido"
      TabPicture(0)   =   "M_GenPed.frx":0886
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ingreso Pedido"
      TabPicture(1)   =   "M_GenPed.frx":08A2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "vaSpread1"
      Tab(1).Control(1)=   "Frame1(1)"
      Tab(1).Control(2)=   "Frame2"
      Tab(1).ControlCount=   3
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   960
         TabIndex        =   12
         Top             =   600
         Width           =   15495
         Begin VB.CommandButton Command3 
            Caption         =   "Borrar Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   12480
            TabIndex        =   14
            Top             =   5040
            Width           =   1455
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Enviar Pel"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   14040
            TabIndex        =   13
            Top             =   5040
            Width           =   1095
         End
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   4395
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Width           =   14970
            _Version        =   393216
            _ExtentX        =   26405
            _ExtentY        =   7752
            _StockProps     =   64
            AutoClipboard   =   0   'False
            ButtonDrawMode  =   1
            EditEnterAction =   2
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
            MaxCols         =   5
            MaxRows         =   1
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "M_GenPed.frx":08BE
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   -69120
         TabIndex        =   9
         Top             =   3840
         Width           =   5055
         Begin VB.Label Label3 
            Caption         =   "Un Momento Generando Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   10
            Top             =   480
            Width           =   3735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   705
         Index           =   1
         Left            =   -70920
         TabIndex        =   3
         Top             =   720
         Width           =   8895
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1980
            TabIndex        =   4
            Top             =   180
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Text            =   "13/07/2004"
            DateCalcMethod  =   4
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
            Left            =   5655
            TabIndex        =   5
            Top             =   180
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
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
            Text            =   "13/07/2004"
            DateCalcMethod  =   4
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
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   7320
            TabIndex        =   6
            Top             =   120
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
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicial"
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
            Left            =   690
            TabIndex        =   8
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Final"
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
            Left            =   4440
            TabIndex        =   7
            Top             =   240
            Width           =   1005
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4515
         Left            =   -74760
         TabIndex        =   11
         Top             =   1800
         Width           =   16890
         _Version        =   393216
         _ExtentX        =   29792
         _ExtentY        =   7964
         _StockProps     =   64
         AutoClipboard   =   0   'False
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         MaxCols         =   11
         MaxRows         =   1
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_GenPed.frx":0CBA
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFC0&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   2
      Left            =   180
      Top             =   1000
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_GenPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub LimpiarControles()
    vaSpread1.MaxRows = 0
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_carga ""
Me.Height = 9090
Me.Width = 18420
fg_centra Me
Msgtitulo = "Generación Pedidos"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")
fpText.text = ""
Label3.Visible = False
Frame2.Visible = False
fg_descarga
Call LimpiarControles
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub
End Sub

Private Sub fpText_Change()
Dim RS As New ADODB.Recordset
Dim sql As String
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
sql = Trim(LimpiaDato(fpText.text))
Set RS = vg_db.Execute("sgpadm_s_cliente 29, '" & sql & "', ''")
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda.Caption = "": Exit Sub
fpayuda.Caption = Trim(RS!cli_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 120
    Image1_Click
End Select
End Sub

Private Sub Image1_Click()
vg_left = fpayuda.Left + 2300
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Gen"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
If Me.Visible Then fpDateTime1(0).SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 5 '-------> Salir
    Me.Hide
    Unload Me
 
Case 2 'Generar el Pedido
    
     Valida_Generacion_pedido M_GenPed.vaSpread1
     
     If valida = True Then
        If MsgBox("Esta Seguro Generar el Pedido ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        Call generar_pedido
     End If
    
End Select
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub generar_pedido()
    Dim MyBuffer As Variant
    Dim IdRuta As Long
    Dim CodIngrediente As String
    Dim CodProveedor As String
    Dim FamProducto As String
    Dim CenCosto As String
    Dim codproducto As String
    Dim FechaDespacho As String
    Dim total As Double
    Dim CodProductoSGP As String
    Dim CantidadIngrediente As Double
    Dim CantidadProducto As Double
    Dim pedido   As Integer
    Dim Linea    As Integer
    Dim activo   As Integer
    Dim observacion   As String
    
    
    
    
    '-------> General Pedido & Minuta Real
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaDetallePedido>"

    For i = 1 To M_GenPed.vaSpread1.MaxRows
        Let MyBuffer = MyBuffer & " <DetallePedido"
        MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)
        desc = Replace(Trim(desc), Chr(34), "&quot;")
        desc = Replace(Trim(desc), Chr(38), "&amp;")
        desc = Replace(Trim(desc), Chr(39), "&apos;")
        desc = Replace(Trim(desc), Chr(60), "&lt;")
        desc = Replace(Trim(desc), Chr(62), "&gt;")

        M_GenPed.vaSpread1.Row = i
        
        M_GenPed.vaSpread1.Col = 1 'Id Ruta de Compras
        IdRuta = IIf(M_GenPed.vaSpread1.text = "", 0, M_GenPed.vaSpread1.text)

        M_GenPed.vaSpread1.Col = 2 'Código Ingrediente
        CodIngrediente = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 3 'Descripción Ingrediente

        M_GenPed.vaSpread1.Col = 4 'Código Proveedor SAP
        CodProveedor = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 5 'Código Familia Producto

        M_GenPed.vaSpread1.Col = 6 'Centro costo

        M_GenPed.vaSpread1.Col = 7 'Código Producto SAP
        codproducto = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 8 'Descripción Producto

        M_GenPed.vaSpread1.Col = 9 'Unidad

        M_GenPed.vaSpread1.Col = 10 'Fecha Despacho
        FechaDespacho = IIf(M_GenPed.vaSpread1.text = "", 0, Format(M_GenPed.vaSpread1.text, "yyyymmdd"))

        M_GenPed.vaSpread1.Col = 11 'Total
        total = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 12 'Còdigo Producto SGP
        CodProductoSGP = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 13 'Cantidad Ingrediente SGP
        CantidadIngrediente = M_GenPed.vaSpread1.text

        M_GenPed.vaSpread1.Col = 14 'Cantidad Producto SGP
        CantidadProducto = M_GenPed.vaSpread1.text

        MyBuffer = MyBuffer & " pedido  = " & Chr(34) & 99999 & Chr(34)
        MyBuffer = MyBuffer & " linea  = " & Chr(34) & i & Chr(34)
        MyBuffer = MyBuffer & " CodIngrediente = " & Chr(34) & CodIngrediente & Chr(34)
        MyBuffer = MyBuffer & " CodProductoSGP  = " & Chr(34) & CodProductoSGP & Chr(34)
        MyBuffer = MyBuffer & " CodProducto  = " & Chr(34) & codproducto & Chr(34)
       '- MyBuffer = MyBuffer & " FechaDespacho  = " & Chr(34) & FechaDespacho & Chr(34)
        
        MyBuffer = MyBuffer & " FechaDespacho  = " & Chr(34) & FechaDespacho & Chr(34)
       
        MyBuffer = MyBuffer & " CantidadIngrediente  = " & Chr(34) & CantidadIngrediente & Chr(34)
        MyBuffer = MyBuffer & " CantidadProducto  = " & Chr(34) & CantidadProducto & Chr(34)
        MyBuffer = MyBuffer & " Total  = " & Chr(34) & total & Chr(34)
        MyBuffer = MyBuffer & " IdRuta  = " & Chr(34) & IdRuta & Chr(34)
        MyBuffer = MyBuffer & " activo  = " & Chr(34) & 1 & Chr(34)
        MyBuffer = MyBuffer & " observacion  = " & Chr(34) & Null & Chr(34)
        
  
        
        Let MyBuffer = MyBuffer & "/>"

    Next i

    Let MyBuffer = MyBuffer & "</GrabaDetallePedido>"
    vg_db.Execute ("sgpadm_Ins_GrabaDetallePedidoNuevo '" & MyBuffer & "', '" & LimpiaDato(fpText.text) & "', " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & Format(fpDateTime1(1).text, "yyyymmdd") & "")

    MsgBox "Generación pedido finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
    Toolbar1.Enabled = True
    fg_descarga


Exit Sub

Man_Error:
Toolbar1.Enabled = True
Label3.Visible = False
Frame2.Visible = False
Label3.Caption = ""
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Dim RS As New ADODB.Recordset
Dim sql As String
Dim NomExcelZip As String
Dim i As Long
Dim NameTemp As String

Select Case Button.Index
Case 1
    '-------> Validar centro de costo
    sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_s_cliente 29, '" & sql & "', ''")
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       MsgBox "Ceco no corresponde, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
       fpayuda.Caption = ""
       Exit Sub
    End If
    '-------> Validar fecha nulas
    If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then
       MsgBox "Fecha no corresponde, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    '-------> Validar si fecha final es menor inicial
    If fpDateTime1(1).text < fpDateTime1(0).text Then
       MsgBox "Fecha Inicial es mayor Final, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    '-------> Validar pedido bloqueado
    
    '-------> Validar si existe datos ruta carga
    sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarRutaDespacho '" & sql & "'")
    If RS.EOF Then
        RS.Close: Set RS = Nothing
        MsgBox "No existe datos cargados rutas compras, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
        Exit Sub
    End If
    RS.Close: Set RS = Nothing
    
    '-------> Validar si existe datos convenios
    sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarConvenios '" & sql & "'")
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       MsgBox "No existe datos cargados convenios, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing

    '-------> Validar si existe minuta bloque
    
  ' FIN ARI
    
    sql = Trim(LimpiaDato(fpText.text))
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarMinutaBloqueACT '" & sql & "', " & Format(fpDateTime1(0).text, "yyyymmdd") & ", " & Format(fpDateTime1(1).text, "yyyymmdd") & "")
   
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       MsgBox "No existe minuta bloque, pedido cancelado", vbExclamation + vbOKOnly, Msgtitulo
       Exit Sub
    End If
    RS.Close: Set RS = Nothing
   
   'FIN ARI1
    
    Label3.Visible = True
    Frame2.Visible = True
    Label3.Caption = "Un momento generando pedido ..."
    
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = &HC0FFFF
    estexi = True
   
   'dispara sp ppal.
   fg_carga ""
   Toolbar1.Enabled = False
   sql = " sgpadm_Sel_GeneracionPedidoNuevo"
   sql = sql & " '" & fpText & "'"
   sql = sql & " , " & Format(fpDateTime1(0).text, "yyyymmdd") & " "
   sql = sql & " , " & Format(fpDateTime1(1).text, "yyyymmdd") & ""

   Set RS = vg_db.Execute(sql)
    
   '-------> Inicio LLenar grilla
   Dim AuxCodIngrediente As String
   AuxIngrediente = ""
   vaSpread1.MaxRows = 0
    If Not RS.EOF Then
       Toolbar1.Buttons(3).Enabled = True
    Else
       Toolbar1.Buttons(3).Enabled = False
    End If
    Do While Not RS.EOF
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
      '  If AuxCodIngrediente <> RS(1) Then
      '     For i = 1 To 12
      '         vaSpread1.Col = i
      '         vaSpread1.BackColor = Shape1(2).FillColor
      '     Next i
      '     AuxCodIngrediente = RS(1)
      '  End If
        vaSpread1.Col = 1 ' IdCompra
        vaSpread1.text = IIf(RS(0) = 0, "", RS(0))
        
        vaSpread1.Col = 2 ' Cod. Ingrediente
        vaSpread1.text = RS(1)
        vaSpread1.Col = 3 ' Des. Ingrediente
        vaSpread1.text = RS(2)
        vaSpread1.Col = 4 ' Proveedor
        vaSpread1.text = RS(3)
        vaSpread1.Col = 5 ' Familia Producto
        vaSpread1.text = RS(4)
        vaSpread1.Col = 6 ' Centro Costo
        vaSpread1.text = RS(5)
        vaSpread1.Col = 7 ' Codigo Producto SAP
        vaSpread1.text = RS(6)
        vaSpread1.Col = 8 ' Des. Producto SAp
        vaSpread1.text = RS(7)
        vaSpread1.Col = 9 ' Unidad
        vaSpread1.text = RS(8)
        vaSpread1.Col = 10 ' Fecha Despacho
        vaSpread1.text = RS(9)
        vaSpread1.Col = 11 ' Cantidad Solicitar
        vaSpread1.text = RS(10)
        vaSpread1.Col = 12 'Código Productos
        vaSpread1.text = RS(11)
        vaSpread1.Col = 13 ' Cantidad Ingrediente
        vaSpread1.text = RS(12)
        vaSpread1.Col = 14 ' Cantidad Producto
        vaSpread1.text = RS(13)
        
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    
    Label3.Visible = False
    Frame2.Visible = False
    Label3.Caption = ""
    
    '-------> validar si existe pedido
    
    If vaSpread1.MaxRows < 1 Then
       fg_descarga
       Label3.Visible = False
       Frame2.Visible = False
       Label3.Caption = ""
'       DropTeblaTmp (NameTemp)
       MsgBox "Por favor verificar si existen " & VgLinea & VgLinea & "- Rutas para la fecha consultada " & VgLinea & "- Convenios vigentes para la fecha consultada " & VgLinea, vbInformation + vbOKOnly, Msgtitulo
       Toolbar1.Enabled = True
       Exit Sub
    End If
       Toolbar1.Enabled = True
       Toolbar1.Buttons(2).Visible = True
       Toolbar1.Buttons(3).Visible = False
                    
       fg_descarga
End Select
Exit Sub
End Sub

