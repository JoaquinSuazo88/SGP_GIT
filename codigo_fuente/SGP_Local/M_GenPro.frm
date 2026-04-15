VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_GenPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación archivos planos productos"
   ClientHeight    =   7575
   ClientLeft      =   1785
   ClientTop       =   1755
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   5880
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
      TabIndex        =   12
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   2
         Left            =   570
         TabIndex        =   13
         Top             =   240
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "M_GenPro.frx":0000
            Left            =   1680
            List            =   "M_GenPro.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   15
            Top             =   600
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
            TabIndex        =   17
            Top             =   660
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
            TabIndex        =   16
            Top             =   345
            Width           =   1485
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4125
         Index           =   0
         Left            =   120
         TabIndex        =   18
         Top             =   1680
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7276
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
         MaxCols         =   3
         MaxRows         =   1
         SpreadDesigner  =   "M_GenPro.frx":001E
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   660
         TabIndex        =   19
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
      Height          =   6315
      Index           =   1
      Left            =   5970
      TabIndex        =   0
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1245
         Index           =   3
         Left            =   570
         TabIndex        =   2
         Top             =   240
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "M_GenPro.frx":03DE
            Left            =   1680
            List            =   "M_GenPro.frx":03EB
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2865
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Mostrar solo productos no enviados"
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
            TabIndex        =   3
            Top             =   930
            Width           =   3390
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   5
            Top             =   600
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
            TabIndex        =   7
            Top             =   345
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
            TabIndex        =   6
            Top             =   660
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
         Height          =   4125
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7276
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
         MaxCols         =   4
         MaxRows         =   2
         SpreadDesigner  =   "M_GenPro.frx":0408
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   435
         Index           =   2
         Left            =   4500
         TabIndex        =   9
         Top             =   5850
         Visible         =   0   'False
         Width           =   1095
         _Version        =   393216
         _ExtentX        =   1931
         _ExtentY        =   767
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
         MaxRows         =   2
         SpreadDesigner  =   "M_GenPro.frx":0885
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2775
         Top             =   5880
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "No Enviados"
         Height          =   195
         Index           =   0
         Left            =   3135
         TabIndex        =   11
         Top             =   5850
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1560
         Top             =   5880
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enviados"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   10
         Top             =   5850
         Width           =   660
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   12285
      _ExtentX        =   21669
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   4920
      OleObjectBlob   =   "M_GenPro.frx":0C9C
      Top             =   6720
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   4320
      OleObjectBlob   =   "M_GenPro.frx":0CC0
      Top             =   6750
   End
End
Attribute VB_Name = "M_GenPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Est As Boolean, estado As Boolean
Dim codtip As Long, ibusca As Long, i As Long, j As Long
Dim aAp As String

Private Sub Check1_Click(Index As Integer)
If Est Then Exit Sub
fg_carga ""
If Check1(1).Value = 0 Then
   If Combo1(1).ListIndex = 2 Then
      a = fptnombre(1).text
      fptnombre(1).text = " "
      fptnombre(1).text = a
   Else
      vaSpread1(1).Visible = False
      For i = 1 To vaSpread1(1).MaxRows
          vaSpread1(1).Row = i
          vaSpread1(1).Col = 1
          vaSpread1(1).text = "0"
          vaSpread1(1).RowHidden = False
      Next i
      vaSpread1(1).SetActiveCell 1, 1
      vaSpread1(1).Visible = True
  End If
Else
   vaSpread1(1).Visible = False
   For i = 1 To vaSpread1(1).MaxRows
       vaSpread1(1).Row = i
       vaSpread1(1).Col = 1
       vaSpread1(1).text = "0"
       If vaSpread1(1).BackColor = Shape1(1).FillColor Then vaSpread1(1).RowHidden = True
   Next i
   vaSpread1(1).SetActiveCell 1, 1
   vaSpread1(1).Visible = True
End If
fg_descarga
'MoverDatoGrilla
End Sub

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
Select Case Index
Case 0
    vaSpread1(0).SetFocus
Case 1
    If Combo1(1).ListIndex = 2 Then
       vg_left = Frame1(0).Left + Combo1(1).Left + 1920
       B_TabEst.LlenaDatos "a_tipopro", "tip_", "Familia del Producto", "Gen"
       B_TabEst.Show 1
       Me.Refresh
       If Val(vg_codigo) = 0 Then Combo1(1).ListIndex = 1: fptnombre(1).Enabled = True: fptnombre(1).text = "": Exit Sub
       codtip = Val(vg_codigo)
       fptnombre(1).text = codtip & " " & vg_nombre
       fptnombre(1).Enabled = False
   Else
      fptnombre(1).Enabled = True
      fptnombre(1).text = ""
   End If
   If vaSpread1(1).MaxRows > 0 Then vaSpread1(1).SetFocus
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7245
Me.Width = 11985
Me.HelpContextID = vg_OpcM
Msgtitulo = "Generación archivos planos productos"
fg_centra Me
Est = True: ibusca = 0
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.ToolTipText = "Enviar": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Combo1(0).ListIndex = 1: Combo1(1).ListIndex = 1
Check1(1).Value = 1
MoverDatoGrilla
SendKeys "+{Tab}"
Est = False
End Sub

Private Sub fpTnombre_Change(Index As Integer)
Select Case Index
Case 0
    Est = True
    If LimpiaDato(Trim(fptnombre(0).text)) & Chr(KeyAscii) = "" Then Exit Sub
    If Combo1(0).ItemData(Combo1(0).ListIndex) = 0 Then
       RS.Open "sgpadm_s_cliente 10, '', '%" & UCase(LimpiaDato(fptnombre(0).text)) & "%'", vg_db, adOpenForwardOnly
       If RS.EOF Then vaSpread1(0).MaxRows = 0 Else vaSpread1(0).MaxRows = RS!nreg
    ElseIf Combo1(0).ItemData(Combo1(0).ListIndex) = 1 Then
       RS.Open "sgpadm_s_cliente 11, '', '%" & UCase(LimpiaDato(fptnombre(0).text)) & "%'", vg_db, adOpenForwardOnly
       If RS.EOF Then vaSpread1(0).MaxRows = 0 Else vaSpread1(0).MaxRows = RS!nreg
    End If
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
    Est = False
Case 1
    If LimpiaDato(Trim(fptnombre(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    FindString = Trim(fptnombre(1).text)
    If fptnombre(1).text = "" Then
       vaSpread1(1).Visible = False
       SwActiva = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If Check1(1).Value = 1 And vaSpread1(1).BackColor = Shape1(0).FillColor Then
              vaSpread1(1).RowHidden = False
           ElseIf Check1(1).Value = 0 Then
              vaSpread1(1).RowHidden = False
           End If
           If SwActiva = 0 Then SwActiva = 1
       Next i
       vaSpread1(1).Visible = True
    Else
       SwActiva = 0
       vaSpread1(1).Visible = False
       irow = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Or Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
              vaSpread1(1).Col = IIf(Combo1(1).ItemData(Combo1(1).ListIndex) = 0, 2, 3)
           Else
              FindString = Trim(Str(codtip))
              vaSpread1(1).Col = 4
           End If
           SourceString = Trim(vaSpread1(1).text)
           indactivo = UCase(Trim(SourceString)) Like "*" & UCase(FindString) & "*"
           If indactivo = -1 Then
              If SwActiva = 0 Then SwActiva = 1
              If vaSpread1(1).RowHidden = True And Check1(1).Value = 1 And vaSpread1(1).BackColor = Shape1(0).FillColor Then
                 vaSpread1(1).RowHidden = False
              ElseIf vaSpread1(1).RowHidden = True And Check1(1).Value = 0 Then
                 vaSpread1(1).RowHidden = False
              End If
              irow = irow + 1
           Else
              If vaSpread1(1).RowHidden = False Then vaSpread1(1).RowHidden = True
           End If
       Next i
       vaSpread1(1).Visible = True
       End If
End Select
End Sub

Private Sub fptnombre_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 34 And irow > 0 Then vaSpread1(Index).SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim CHost As String, Cdire As String, Cuser As String, Cpass As String, Cpuer As String
Dim codmun As Long, codreg As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1(0).MaxRows < 1 Or vaSpread1(1).MaxRows < 1 Then Exit Sub
    Dim i As Long, j As Long
    Dim isel As Boolean, icopy As Boolean
    Dim cencos As String, nomcencos As String, codpro As String, sourcefile As String, sourcefilezip As String, destinofile As String, destinofilezip As String, mdirserver As String, lognarchsou As String, socsap As String
    Dim tprod As String, treceta As String, dBo As String, cDBI As String
    Dim fso, codtis As Long, CodSeg As Long, sobrec As String
    Set fso = CreateObject("Scripting.FileSystemObject")
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
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un producto", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    Frame1(0).Enabled = False: Frame1(1).Enabled = False
   '------- Creo tabla temporal y chequeo si existe antes
   treceta = "": tprod = ""
   vg_db.BeginTrans
   aAp = Trim(vg_NUsr) & "_tmp_GenPlanoProd"
   fg_CheckTmp aAp: tprod = aAp
   vg_db.Execute "create table " & aAp & " (pro_codigo varchar(20))"
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then
           vaSpread1(1).Col = 2
           vg_db.Execute "INSERT INTO " & aAp & " (pro_codigo) values ('" & Trim(vaSpread1(1).text) & "')"
        End If
    Next i
    vg_db.CommitTrans
    
    '------- Crear directorio servidor SQLDES
    mdirserver = Dir(dir_trabajo & "\" & "Actualizar", vbDirectory)
    If mdirserver = "" Then MkDir dir_trabajo & "\" & "Actualizar"
    mdirserver = dir_trabajo & "Actualizar" & "\"
    '------- Fin crear directorio servidor SQLDES
    
    '------- Generar base padre Access
    sourcefile = "productogeneral" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
    If Dir(mdirpc & sourcefile) <> "" Then Kill mdirpc & sourcefile ' borrar base datos si existe
    
    '------- base de datos origen
    ' Rutas base de datos access    dBo = dir_trabajo + BaseDeDato
    dBo = "'' [ODBC;Driver={SQL Server};Server=" + vg_SqlNSvr + ";Database=" + vg_SqlBase + ";UID=" + vg_SqlNUsr + ";PWD=" + vg_SqlPass + "]"
    GenerarBaseEnviado mdirpc & sourcefile, tprod, treceta, dBo, 0, 0, 0, 0
    Bar1(0).Visible = True: Bar1(1).Visible = True
    Bar1(0).Value = 0: Bar1(1).Value = 0: icopy = False
    For i = 1 To vaSpread1(0).MaxRows
        Bar1(0).Value = Val((i / vaSpread1(0).MaxRows) * 100)
        vaSpread1(0).Row = i: vaSpread1(0).Col = 1
        If vaSpread1(0).text = "1" Then
           DoEvents
           vaSpread1(0).Col = 3: nomcencos = Trim(vaSpread1(0).text)
           vaSpread1(0).Col = 2: cencos = Trim(vaSpread1(0).text): Bar1(1).Value = 0
           vaSpread1(0).SetActiveCell 2, vaSpread1(0).Row
           codpro = ""
           vg_db.BeginTrans
           For j = 1 To vaSpread1(1).MaxRows
                DoEvents
               Bar1(1).Value = Val((j / vaSpread1(1).MaxRows) * 100)
               vaSpread1(1).Row = j: vaSpread1(1).Col = 1
               '------- Leer, insertar y regrabar productos casinos
               If Trim(codpro) = "" Then
                  vg_db.Execute "DELETE b_productocasino FROM b_productocasino a , " & aAp & " b WHERE a.prc_cencos = '" & cencos & "' AND a.prc_codpro = b.pro_codigo"
                  vg_db.Execute "INSERT INTO b_productocasino (prc_cencos, prc_codpro, prc_fecenv) SELECT '" & cencos & "', pro_codigo, " & Format(Date, "yyyymmdd") & " FROM " & aAp & ""
               End If
               '------- Fin leer, insertar y regrabar productos casinos
               If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then vaSpread1(1).Col = 2: codpro = Trim(Str(vaSpread1(1).text))
           Next j
           vg_db.CommitTrans
           DoEvents
           '------- Crear archivos *.MDB y *.ZIP
           '------- Modificación de archivo *.mdb x *.kkk, ya que el correo esta eliminado archivo
'           destinofile = "sgp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
'           destinofile = "sgp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".kkk"
'           destinofilezip = "sgp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           destinofile = "sgp" & (cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".kkk"
           destinofilezip = "sgp" & (cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           '------- verificar si existe archivo mdb destino si existe borrar y copiar
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
           '------- Generar parametros ejecutivos contables
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'datcont', mid(cli_nomcontable,1,40), 'C', cli_emailcontable FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT '5etapas', 'Casino 5 Etapas', 'C', iif(cli_subseg = 0,'N','S') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT par_codigo, par_nombre, par_tipo, par_valor FROM a_param IN " & dBo & " WHERE par_codigo = 'porprepro'"
           '-------> Insert concepto grupo vulnerable tabla a_param.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'opgruvul', 'Opción Grupo Vulnerable', 'C', iif(cli_gruvul = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           '-------> Insert concepto modulo paciente.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modpac', 'Modulo Paciente', 'C', iif(cli_modpac = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           '-------> Insert concepto parametro proveedor
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modprove', 'Parametro Modificar Proveedor', 'N', '0' FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           '-------> Insert concepto generación pedido Web o SGP
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT DISTINCT 'gpedsgpweb', 'Parametro Generación Pedido x SGP o Web', 'C', cli_opgped FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           '-------> Insert concepto Hipersensibilidad Alimentaria tabla a_param.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'hipali', 'Opción Hipersensibilidad Alimentaria', 'C', iif(cli_hipali = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           '-------> Insert concepto Tipo Operación tabla a_param.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'tipope', 'Tipo Operación 0=Gravada:1=No Gravada', 'C', cli_tipope FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           '-------> Insert concepto Minuta Sitio Remoto tabla a_param.
           dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'minsre', 'Minuta Sitio Remoto 0=No:1=SI', 'C', cli_minsre FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
           codtis = 0: CodSeg = 0: socsap = "": sobrec = "": codmun = 0: codreg = 0
           Dim ccisac As Long
           Dim cecsac As String
           RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & cencos & "'", vg_db, adOpenStatic
           If Not RS.EOF Then
              codtis = IIf(IsNull(RS!cli_codtis), 0, RS!cli_codtis)
              CodSeg = IIf(IsNull(RS!cli_codseg), 0, RS!cli_codseg)
              socsap = IIf(IsNull(RS!cli_socsap), "", RS!cli_socsap)
              sobrec = IIf(IsNull(RS!cli_sobrec), "", RS!cli_sobrec)
              codmun = IIf(IsNull(RS!cli_codmun), 0, RS!cli_codmun)
              ccisac = IIf(IsNull(RS!cli_ccisac), 0, RS!cli_ccisac)
              cecsac = IIf(IsNull(RS!cli_cecsac), "", RS!cli_cecsac)
              codreg = IIf(IsNull(RS!cli_codreg), 0, RS!cli_codreg)
           End If
           RS.Close: Set RS = Nothing
           '-------> Borrar tabla tipo servicio y segmento que no tenga relación con el contrato
           dbi.Execute "DELETE a_tiposervicio FROM a_tiposervicio WHERE tis_codigo NOT IN (" & codtis & ")"
           dbi.Execute "DELETE a_segmento FROM a_segmento WHERE seg_codigo NOT IN (" & CodSeg & ")"

           '-------> Borrar tabla casino envia sap
           dbi.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos NOT IN ('" & cencos & "')"

           '-------> Mover datos a la tabla centro de costo
           dbi.Execute "INSERT INTO a_cencos (cen_codigo, cen_socsap, cen_sobrec, cen_codmun, cen_ccisac, cen_cecsac, cen_codreg) VALUES ('" & cencos & "', '" & socsap & "', '" & sobrec & "', " & codmun & ", " & ccisac & ", '" & cecsac & "', " & codreg & ")"
           '-------> Mover datos parametros despachos
           dbi.Execute "INSERT INTO b_paramdesp SELECT DISTINCT pad_cencos, pad_codtip AS pad_codtip, pad_tipo, pad_diaseg, pad_diario FROM b_parametrodespachos IN " & dBo & " WHERE pad_cencos = '" & cencos & "'"
           '-------> Mover datos días inhabiles
           dbi.Execute "INSERT INTO b_Fecha_Inhabiles SELECT DISTINCT CFI_CeCo, CFI_Fecha, CFI_Glosa FROM Cas_b_Fecha_Inhabiles IN " & dBo & " WHERE CFI_CeCo = '" & cencos & "'"
           '-------> Mover datos casino tipo actividades
           dbi.Execute "INSERT INTO b_casinotipoactividades SELECT DISTINCT cta_cencos, cta_tipact FROM b_casinotipoactividades IN " & dBo & " WHERE cta_cencos = '" & cencos & "'"
           '-------> Mover datos casino parametro stock
           dbi.Execute "INSERT INTO b_casinoparametrostock SELECT DISTINCT cps_cencos, cps_invsto, cps_reqmen, cps_porinv, cps_liscri, cps_diario, cps_ajuimp FROM b_casinoparametrostock IN " & dBo & " WHERE cps_cencos = '" & cencos & "'"
           '-------> Mover datos clase documento sap
           dbi.Execute "INSERT INTO a_clasedocsap SELECT DISTINCT cds_coddoc, cds_codreg, cds_cdosap FROM a_clasedocsap IN " & dBo & " WHERE cds_codreg = " & codreg & ""
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
           RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & cencos & "'", vg_db, adOpenStatic
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
'                       cHost = Mid(cpars, InStr(cpars, ",") + 1)
'                    ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "B" Then
'                       cUser = Mid(cpars, InStr(cpars, ",") + 1)
'                    ElseIf Mid(cpars, 1, InStr(cpars, ",") - 1) = "C" Then
'                       cPass = Mid(cpars, InStr(cpars, ",") + 1)
'                    End If
'                 Loop
'                 Close #1
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
                    SendMail1 oMail, "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy") & " Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar", mdirpc & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1
                 End If
              End If
           End If
           RS.Close: Set RS = Nothing
           DoEvents
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
    If Trim(sourcefile) <> "" Then MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    Frame1(0).Enabled = True: Frame1(1).Enabled = True
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
fg_descarga
Frame1(0).Enabled = True: Frame1(1).Enabled = True
Bar1(0).Visible = False: Bar1(1).Visible = False
If Err = -2147168242 Then
    vg_db.RollbackTrans
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    DoEvents
    Exit Sub
End If
RS.Close: Set RS = Nothing
Man_Error:
Select Case Err
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

Private Sub vaSpread1_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo Man_Error
If Index = 1 Or Est Then Exit Sub
vaSpread1(0).Row = Row
Select Case Col
Case 1
    If Row = 0 Or Row = -1 Then x = vaSpread1(0).MaxRows: j = 1 Else x = vaSpread1(0).Row: j = vaSpread1(0).Row
    For j = j To x
        fg_carga ""
        vaSpread1(0).Row = j
        vaSpread1(0).Col = 1
        If Trim(vaSpread1(0).text) = "1" Then
           vaSpread1(0).Col = 2
'           Set RS1 = vg_db.Execute("sp_s_productonoenviado '" & Trim(vaSpread1(0).Text) & "'")
           aAp = Trim(vg_NUsr) & "_tmp_CasinoProductos"
           vg_db.Execute "DELETE " & aAp & " FROM " & aAp & ""
Paso:
''           RS1.Open "select * from " & aAp & "", vg_db, adOpenStatic
''           RS1.Close: Set RS1 = Nothing
''           fg_CheckTmp aAp
''           RS1.Open "select distinct pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
''                    "into " & aAp & " " & _
''                    "from b_productos pro inner join b_productocasino pri " & _
''                    "on pro.pro_codigo = pri.prc_codpro " & _
''                    "where pri.prc_cencos='" & Trim(vaSpread1(0).Text) & "'", vg_db, adOpenStatic

           vg_db.Execute "INSERT INTO " & aAp & " SELECT DISTINCT pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
                         "FROM b_productos pro inner join b_productocasino pri " & _
                         "on pro.pro_codigo    = pri.prc_codpro " & _
                         "WHERE pri.prc_cencos = '" & Trim(vaSpread1(0).text) & "'"
           Set RS1 = Nothing

           RS1.Open "SELECT pro.pro_codigo, pro.pro_nombre, tmp.prc_codpro " & _
                    "FROM b_productos pro left join " & aAp & " tmp on pro.pro_codigo = tmp.pro_codigo " & _
                    "WHERE (tmp.pro_codigo) IS NULL ORDER BY tmp.pro_codigo", vg_db, adOpenForwardOnly ', adOpenStatic
           If Not RS1.EOF Then
              vaSpread1(1).Visible = False
              Do While Not RS1.EOF
                 If IsNull(RS1!prc_codpro) Then
                    i = vaSpread1(1).SearchCol(2, 1, vaSpread1(1).MaxRows, Trim(RS1!pro_codigo), SearchFlagsEqual) 'SearchFlagsGreaterOrEqual)
                    vaSpread1(1).Row = i
                    If vaSpread1(1).BackColor = Shape1(1).FillColor Then
                        vaSpread1(1).Col = 1
                        vaSpread1(1).BackColor = Shape1(0).FillColor
                        vaSpread1(1).Col = 2
                        vaSpread1(1).BackColor = Shape1(0).FillColor
                        vaSpread1(1).Col = 3
                        vaSpread1(1).BackColor = Shape1(0).FillColor
                        vaSpread1(1).Col = 4
                        vaSpread1(1).BackColor = Shape1(0).FillColor
                        vaSpread1(1).RowHidden = False
                    End If
                 Else
                    Exit Do
                 End If
                 RS1.MoveNext
              Loop
              vaSpread1(1).Visible = True
           End If
           RS1.Close: Set RS1 = Nothing
        End If
    Next j
    fg_descarga
End Select
Exit Sub
Man_Error:
Select Case Err
Case -2147217865
   vg_db.Execute "CREATE TABLE " & aAp & " (pro_codigo varchar(20), pro_nombre varchar(50), prc_codpro varchar(20))"
'     RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
'              "INTO " & aAp & " " & _
'              "FROM b_productos pro INNER JOIN b_productocasino pri " & _
'              "ON pro.pro_codigo = pri.prc_codpro " & _
'              "WHERE pri.prc_cencos='" & Trim(vaSpread1(0).text) & "'", vg_db, adOpenStatic
    DoEvents
    GoTo Paso
End Select
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub MoverDatoGrilla()
On Error GoTo Man_Error
fg_carga "": estado = True: i = 1
'-------> Mover casinos
If Est Then
   vaSpread1(0).MaxRows = 0
   RS.Open "sgpadm_s_cliente 12, '', ''", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         DoEvents
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
         If estado = True Then
            estado = False
            Set RS1 = vg_db.Execute("sp_s_productonoenviado '" & RS!cli_codigo & "'")
'            aAp = Trim(vg_NUsr) & "_tmp_CasinoProductos"
'            fg_CheckTmp aAp
'            RS1.Open "SELECT DISTINCT pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
'                     "INTO " & aAp & " " & _
'                     "FROM b_productos pro INNER JOIN b_productocasino pri " & _
'                     "ON pro.pro_codigo = pri.prc_codpro " & _
'                     "WHERE pri.prc_cencos='" & RS!cli_codigo & "'", vg_db, adOpenForwardOnly ', adOpenStatic
'            Set RS1 = Nothing

'            RS1.Open "SELECT pro.pro_codigo, pro.pro_nombre, tmp.prc_codpro " & _
'                     "FROM b_productos pro LEFT JOIN " & aAp & " tmp ON pro.pro_codigo = tmp.pro_codigo " & _
'                     "WHERE (tmp.pro_codigo) IS NULL", vg_db, adOpenForwardOnly ', adOpenStatic
''            RS1.Open "select distinct b_productos.pro_codigo, b_productos.pro_nombre, b_productocasino.prc_codpro " & _
''                     "from b_productos left join b_productocasino ON b_productos.pro_codigo = b_productocasino.prc_codpro " & _
''                     "where isnull(b_productocasino.prc_codpro) or b_productocasino.prc_cencos='" & RS!cli_codigo & "' order by b_productocasino.prc_codpro", vg_db, adOpenStatic
            If Not RS1.EOF Then
               vaSpread1(2).MaxRows = 0
               Do While Not RS1.EOF
                  If IsNull(RS1!prc_codpro) Then
                     vaSpread1(2).MaxRows = vaSpread1(2).MaxRows + 1
                     vaSpread1(2).Row = vaSpread1(2).MaxRows
                  
                     vaSpread1(2).Col = 2
                     vaSpread1(2).text = RS1!pro_codigo
                  
                     vaSpread1(2).Col = 3
                     vaSpread1(2).text = Trim(RS1!pro_nombre)

                     i = vaSpread1(0).MaxRows
                  Else
                     Exit Do
                  End If
                  RS1.MoveNext
               Loop
            End If
            RS1.Close: Set RS1 = Nothing
         End If
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
End If
vaSpread1(0).SetActiveCell 1, i

'-------> Mover productos
vaSpread1(1).MaxRows = 0
'RS.Open "SELECT pro_codigo, pro_nombre, pro_codtip FROM b_productos ORDER BY pro_nombre", vg_db, adOpenForwardOnly ', adOpenStatic
RS.Open "sgpadm_s_productos 1, '1', '', '" & vg_NUsr & "'", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      DoEvents
      vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
      vaSpread1(1).Row = vaSpread1(1).MaxRows
      '-------> Validar si productos existe
      If -1 <> vaSpread1(2).SearchCol(2, 1, vaSpread1(2).MaxRows, Trim(RS!pro_codigo), SearchFlagsEqual) Then
         vaSpread1(1).RowHidden = False
         estado = True
      Else
         vaSpread1(1).RowHidden = True
         estado = False
      End If
'      For i = 1 To vaSpread1(2).MaxRows
'          vaSpread1(2).Row = i
'          vaSpread1(2).Col = 2
'          If RS!pro_codigo = Trim(vaSpread1(2).text) Then estado = True: Exit For
'      Next i
      vaSpread1(1).Col = 1
      vaSpread1(1).BackColor = IIf(estado = True, Shape1(0).FillColor, Shape1(1).FillColor)
      vaSpread1(1).text = "0"
      
      vaSpread1(1).Col = 2
      vaSpread1(1).TypeHAlign = TypeHAlignLeft
      vaSpread1(1).TypeSpin = False
      vaSpread1(1).TypeIntegerSpinInc = 1
      vaSpread1(1).TypeIntegerSpinWrap = False
      vaSpread1(1).BackColor = IIf(estado = True, Shape1(0).FillColor, Shape1(1).FillColor)
'      vaSpread1(1).CellType = CellTypeStaticText
      vaSpread1(1).Lock = True
      vaSpread1(1).text = Trim(RS!pro_codigo)

      vaSpread1(1).Col = 3
      vaSpread1(1).TypeHAlign = TypeHAlignLeft
      vaSpread1(1).BackColor = IIf(estado = True, Shape1(0).FillColor, Shape1(1).FillColor)
      vaSpread1(1).text = Trim(RS!pro_nombre)
      
      vaSpread1(1).Col = 4
      vaSpread1(1).BackColor = IIf(estado = True, Shape1(0).FillColor, Shape1(1).FillColor)
      vaSpread1(1).text = RS!pro_codtip
            
       RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then vaSpread1(Index).Row = -1: vaSpread1(Index).Col = 1: vaSpread1(Index).text = IIf(vaSpread1(Index).Value = "1", "0", "1")
End Sub

Private Sub vaSpread1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Or KeyCode = 13 Then Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then fptnombre(Index).text = IIf(KeyCode = 8, fptnombre(Index).text, fptnombre(Index).text & Chr(KeyCode)): fptnombre(Index).SetFocus: fptnombre(Index).SelStart = Len(fptnombre(Index).text)
End Sub
