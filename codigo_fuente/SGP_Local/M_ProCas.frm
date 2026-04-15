VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_ProCas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación archivos planos productos"
   ClientHeight    =   7590
   ClientLeft      =   1620
   ClientTop       =   1665
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12555
   ShowInTaskbar   =   0   'False
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   6360
      Top             =   6960
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
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
      TabIndex        =   7
      Top             =   390
      Width           =   5865
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   1
         Left            =   660
         TabIndex        =   16
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
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1245
         Index           =   3
         Left            =   570
         TabIndex        =   11
         Top             =   240
         Width           =   4725
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
            TabIndex        =   19
            Top             =   930
            Width           =   3390
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "M_ProCas.frx":0000
            Left            =   1680
            List            =   "M_ProCas.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   4
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
            Index           =   2
            Left            =   150
            TabIndex        =   13
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
            Index           =   1
            Left            =   150
            TabIndex        =   12
            Top             =   345
            Width           =   1485
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4125
         Index           =   1
         Left            =   120
         TabIndex        =   5
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
         SpreadDesigner  =   "M_ProCas.frx":002A
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   435
         Index           =   2
         Left            =   4500
         TabIndex        =   20
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
         SpreadDesigner  =   "M_ProCas.frx":0430
         TextTip         =   2
         TextTipDelay    =   0
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Enviados"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   18
         Top             =   5850
         Width           =   660
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
         Caption         =   "No Enviados"
         Height          =   195
         Index           =   0
         Left            =   3135
         TabIndex        =   17
         Top             =   5850
         Width           =   915
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
      TabIndex        =   6
      Top             =   390
      Width           =   5865
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   2
         Left            =   570
         TabIndex        =   8
         Top             =   240
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            ItemData        =   "M_ProCas.frx":07D0
            Left            =   1680
            List            =   "M_ProCas.frx":07DA
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   1
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
            Index           =   11
            Left            =   150
            TabIndex        =   10
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
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   660
            Width           =   1470
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4125
         Index           =   0
         Left            =   120
         TabIndex        =   2
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
         SpreadDesigner  =   "M_ProCas.frx":07EE
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   5040
      OleObjectBlob   =   "M_ProCas.frx":0B37
      Top             =   6840
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   4320
      OleObjectBlob   =   "M_ProCas.frx":0B5B
      Top             =   6750
   End
End
Attribute VB_Name = "M_ProCas"
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
Me.Width = 12000
Me.HelpContextID = vg_OpcM
Msgtitulo = "Generación archivos planos productos"
fg_centra Me
Est = True: ibusca = 0
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.ToolTipText = "Enviar": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(0).ListIndex = 1: Combo1(1).ListIndex = 1
Check1(1).Value = 1
MoverDatoGrilla
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
    findstring = Trim(fptnombre(1).text)
    If fptnombre(1).text = "" Then
       vaSpread1(1).Visible = False
       swactiva = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If Check1(1).Value = 1 And vaSpread1(1).BackColor = Shape1(0).FillColor Then
              vaSpread1(1).RowHidden = False
           ElseIf Check1(1).Value = 0 Then
              vaSpread1(1).RowHidden = False
           End If
           If swactiva = 0 Then swactiva = 1
       Next i
       vaSpread1(1).Visible = True
    Else
       swactiva = 0
       vaSpread1(1).Visible = False
       irow = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Or Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
              vaSpread1(1).Col = IIf(Combo1(1).ItemData(Combo1(1).ListIndex) = 0, 2, 3)
           Else
              findstring = Trim(Str(codtip))
              vaSpread1(1).Col = 4
           End If
           sourcestring = Trim(vaSpread1(1).text)
           indactivo = UCase(Trim(sourcestring)) Like "*" & UCase(findstring) & "*"
           If indactivo = -1 Then
              If swactiva = 0 Then swactiva = 1
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
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1(0).MaxRows < 1 Or vaSpread1(1).MaxRows < 1 Then Exit Sub
    Dim i As Long, j As Long
    Dim isel As Boolean, icopy As Boolean
    Dim cencos As String, nomcencos As String, codpro As String, sourcefile As String, sourcefilezip As String, destinofile As String, destinofilezip As String, mdir As String, lognarchsou As String
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
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un producto", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
   '------- Creo tabla temporal y chequeo si existe antes
   aAp = Trim(vg_NUsr) & "_tmp_GenPlano"
   fg_CheckTmp aAp
   vg_db.BeginTrans
   vg_db.Execute "create table " & aAp & " (codpro varchar(20))"
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" Then
           vaSpread1(1).Col = 2
           vg_db.Execute "insert into " & aAp & " (codpro) values ('" & Trim(vaSpread1(1).text) & "')"
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
           vaSpread1(0).Col = 3: nomcencos = Trim(vaSpread1(0).text)
           vaSpread1(0).Col = 2: cencos = Trim(vaSpread1(0).text): Bar1(1).Value = 0
           vaSpread1(0).SetActiveCell 2, vaSpread1(0).Row: vaSpread1(0).SetFocus
''           If icopy = False Then sourcefile = dir_trabajo & "mp" & lcase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt" Else destinofile = dir_trabajo & "mp" & lcase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt"
           If icopy = False Then
              sourcefile = "mp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
              sourcefilezip = "mp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
           Else
              destinofile = "mp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
              destinofilezip = "mp" & LCase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
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
                           SendMail oMail, "Se Informa que el maestro de productos esta disponible. Para que usted pueda actualizar ", "Se Informa que el maestro de productos esta disponible. Para que usted pueda actualizar", mdir & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0
                        End If
                    End If
                 ElseIf RS!cli_openvio = 2 Then
                    If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                       fg_descarga
                       MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, Msgtitulo
                       fg_carga ""
                    Else
                       SendMail oMail, "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), mdir & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1
                    End If
                 End If
              End If
              RS.Close: Set RS = Nothing
           ElseIf icopy = False Then
              '------- verificar si existe archivo mdb y zip si existe borrar
              If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile
              If Dir(mdir & sourcefilezip) <> "" Then Kill mdir & sourcefilezip
              '------- generar archivo mdb
''              Open dir_trabajo & "mp" & lcase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt" For Output As #1
'              Open mdir & "mp" & lcase(cencos) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt" For Output As #1
              Set db7 = DBEngine(0).CreateDatabase(mdir & sourcefile, dbLangGeneral)
              db7.Execute "create table a_tipopro (tip_codigo int, tip_nombre char(35), tip_previo int)", vg_ModoOpen
              db7.Execute "create table a_unidad (uni_codigo int, uni_nombre char(10), uni_nomcor char(5))", vg_ModoOpen
              db7.Execute "create table a_embalaje (emb_codigo int, emb_nombre char(20), emb_nomcor char(5))", vg_ModoOpen
              db7.Execute "create table a_ctacontable (cta_codigo char(10), cta_nombre char(40))", vg_ModoOpen
              db7.Execute "create table a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255))", vg_ModoOpen
              db7.Execute "create table a_impuesto (imp_codigo int, imp_nombre char(15), imp_pctimp double, imp_inccos int)", vg_ModoOpen
              db7.Execute "create table a_unidadmed (unm_codigo int, unm_nombre char(10), unm_nomcor char(5))", vg_ModoOpen
              db7.Execute "create table a_nutriente (nut_codigo int, nut_nombre char(30), nut_nomuni char(5), nut_indpri int, nut_secnro int)", vg_ModoOpen
              db7.Execute "create table b_productos (pro_codigo char(20), pro_codbar char(20), pro_codcom char(20), pro_codtip int, pro_nombre char(50), pro_coduni int, pro_facing double, pro_facsto double, pro_codemb int, pro_uniemb double, pro_upreco double, pro_fecuco datetime, pro_propon double, pro_ctacon char(10), pro_fecven int, pro_ctrsto int)", vg_ModoOpen
              db7.Execute "create table b_productosimp (ipr_codpro char(20), ipr_codimp int)", vg_ModoOpen
              db7.Execute "create table b_productosing (pri_codpro char(20), pri_coding char(20))", vg_ModoOpen
              db7.Execute "create table b_ingrediente (ing_codigo char(20), ing_nombre char(50), ing_nomfan char(50), ing_unimed int, ing_pctapr double, ing_pctcoc double, ing_pctnut double, ing_facnut double, ing_indpav int, ing_indgrv int, ing_precos double, ing_feccos int, ing_codcom char(20), ing_codped char(20))", vg_ModoOpen
              db7.Execute "create table b_productonut (pnu_codpro char(20), pnu_codapo int, pnu_canapo double)", vg_ModoOpen
              db7.Execute "create table b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50))", vg_ModoOpen
              '------- generar familia productos
''              RS.Open "select distinct tip.* from a_tipopro tip, " & aAp & " tpro, b_productos pro where tip.tip_codigo=pro.pro_codtip and pro.pro_codigo=tpro.codpro", vg_db, adOpenStatic
              RS.Open "select * from a_tipopro", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_tipopro Values (" & RS!tip_codigo & ", " & IIf(IsNull(RS!tip_nombre), "Null", "'" & RS!tip_nombre & "'") & ", " & IIf(IsNull(RS!tip_previo), "Null", RS!tip_previo) & ")", vg_ModoOpen
'                    Print #1, "a_tipopro;" & RS!tip_codigo & ";insert into a_tipopro values (" & RS!tip_codigo & "," & IIf(IsNull(RS!tip_nombre), "Null", "'" & RS!tip_nombre & "'") & "," & IIf(IsNull(RS!tip_previo), "Null", RS!tip_previo) & ");" & "update a_tipopro set tip_nombre=" & IIf(IsNull(RS!tip_nombre), "Null", "'" & RS!tip_nombre & "'") & ", tip_previo=" & IIf(IsNull(RS!tip_previo), "Null", RS!tip_previo) & " where tip_codigo=" & RS!tip_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar unidad medida productos
              RS.Open "select distinct uni.* from a_unidad uni, " & aAp & " tpro, b_productos pro where uni.uni_codigo=pro.pro_coduni and pro.pro_codigo=tpro.codpro", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_unidad values (" & RS!uni_codigo & ", " & IIf(IsNull(RS!uni_nombre), "Null", "'" & RS!uni_nombre & "'") & ", " & IIf(IsNull(RS!uni_nomcor), "Null", "'" & RS!uni_nomcor & "'") & ")", vg_ModoOpen
'                    Print #1, "a_unidad;" & RS!uni_codigo & ";insert into a_unidad values (" & RS!uni_codigo & "," & IIf(IsNull(RS!uni_nombre), "Null", "'" & RS!uni_nombre & "'") & "," & IIf(IsNull(RS!uni_nomcor), "Null", "'" & RS!uni_nomcor & "'") & ");" & "update a_unidad set uni_nombre=" & IIf(IsNull(RS!uni_nombre), "Null", "'" & RS!uni_nombre & "'") & ", uni_nomcor=" & IIf(IsNull(RS!uni_nomcor), "Null", "'" & RS!uni_nomcor & "'") & " where uni_codigo=" & RS!uni_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar embalaje productos
              RS.Open "select distinct emb.* from a_embalaje emb, " & aAp & " tpro, b_productos pro where emb.emb_codigo=pro.pro_codemb and pro.pro_codigo=tpro.codpro", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_embalaje values (" & RS!emb_codigo & ", " & IIf(IsNull(RS!emb_nombre), "Null", "'" & RS!emb_nombre & "'") & ", " & IIf(IsNull(RS!emb_nomcor), "Null", "'" & RS!emb_nomcor & "'") & ")", vg_ModoOpen
'                    Print #1, "a_embalaje;" & RS!emb_codigo & ";insert into a_embalaje values (" & RS!emb_codigo & "," & IIf(IsNull(RS!emb_nombre), "Null", "'" & RS!emb_nombre & "'") & "," & IIf(IsNull(RS!emb_nomcor), "Null", "'" & RS!emb_nomcor & "'") & ");" & "update a_embalaje set emb_nombre=" & IIf(IsNull(RS!emb_nombre), "Null", "'" & RS!emb_nombre & "'") & ", emb_nomcor=" & IIf(IsNull(RS!emb_nomcor), "Null", "'" & RS!emb_nomcor & "'") & " where emb_codigo=" & RS!emb_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar cuentas contables productos
              RS.Open "select distinct cta.* from a_ctacontable cta, " & aAp & " tpro, b_productos pro where cta.cta_codigo=pro.pro_ctacon and pro.pro_codigo=tpro.codpro", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_ctacontable values (" & RS!cta_codigo & ", " & IIf(IsNull(RS!cta_nombre), "Null", "'" & RS!cta_nombre & "'") & ")", vg_ModoOpen
'                    Print #1, "a_ctacontable;" & RS!cta_codigo & ";insert into a_ctacontable values (" & RS!cta_codigo & "," & IIf(IsNull(RS!cta_nombre), "Null", "'" & RS!cta_nombre & "'") & ");" & "update a_ctacontable set cta_nombre=" & IIf(IsNull(RS!cta_nombre), "Null", "'" & RS!cta_nombre & "'") & " where cta_codigo=" & RS!cta_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar parametros cuentas contables
              RS.Open "select * from a_param where par_codigo in ('ctagastos','ctainsumo','ctalimdes')", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_param values ('" & RS!par_codigo & "', " & IIf(IsNull(RS!par_nombre), "Null", "'" & RS!par_nombre & "'") & ", " & IIf(IsNull(RS!par_tipo), "Null", "'" & RS!par_tipo & "'") & ", " & IIf(IsNull(RS!par_valor), "Null", "'" & RS!par_valor & "'") & ")", vg_ModoOpen
'                    Print #1, "a_param;" & RS!par_codigo & ";insert into a_param values (" & RS!par_codigo & "," & IIf(IsNull(RS!par_nombre), "Null", "'" & RS!par_nombre & "'") & "," & IIf(IsNull(RS!par_tipo), "Null", "'" & RS!par_tipo & "'") & "," & IIf(IsNull(RS!par_valor), "Null", "'" & RS!par_valor & "'") & ");" & "update a_param set par_nombre=" & IIf(IsNull(RS!par_nombre), "Null", "'" & RS!par_nombre & "'") & ", par_tipo=" & IIf(IsNull(RS!par_tipo), "Null", "'" & RS!par_tipo & "'") & ", par_valor=" & IIf(IsNull(RS!par_valor), "Null", "'" & RS!par_valor & "'") & " where cta_codigo=" & RS!par_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar impuesto productos
              RS.Open "select distinct impu.* from a_impuesto impu, " & aAp & " tpro, b_productosimp ipr where impu.imp_codigo=ipr.ipr_codimp and ipr.ipr_codpro=tpro.codpro", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_impuesto values (" & RS!imp_codigo & ", " & IIf(IsNull(RS!imp_nombre), "Null", "'" & RS!imp_nombre & "'") & ", " & IIf(IsNull(RS!imp_pctimp), "Null", RS!imp_pctimp) & ", " & IIf(IsNull(RS!imp_inccos), "Null", RS!imp_inccos) & ")", vg_ModoOpen
'                    Print #1, "a_impuesto;" & RS!imp_codigo & ";insert into a_impuesto values (" & RS!imp_codigo & "," & IIf(IsNull(RS!imp_nombre), "Null", "'" & RS!imp_nombre & "'") & "," & IIf(IsNull(RS!imp_pctimp), "Null", RS!imp_pctimp) & "," & IIf(IsNull(RS!imp_inccos), "Null", RS!imp_inccos) & ");" & "update a_impuesto set imp_nombre=" & IIf(IsNull(RS!imp_nombre), "Null", "'" & RS!imp_nombre & "'") & ", imp_pctimp=" & IIf(IsNull(RS!imp_pctimp), "Null", RS!imp_pctimp) & ", imp_inccos=" & IIf(IsNull(RS!imp_inccos), "Null", RS!imp_inccos) & " where imp_codigo=" & RS!imp_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar unidad medida ingrediente
              RS.Open "select distinct unm.* from a_unidadmed unm, " & aAp & " tpro, b_productosing pri, b_ingrediente ing where unm.unm_codigo=ing.ing_unimed and ing.ing_codigo=pri.pri_coding and pri.pri_codpro=tpro.codpro", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                   db7.Execute "insert into a_unidadmed values (" & RS!unm_codigo & ", " & IIf(IsNull(RS!unm_nombre), "Null", "'" & RS!unm_nombre & "'") & ", " & IIf(IsNull(RS!unm_nomcor), "Null", "'" & RS!unm_nomcor & "'") & ")", vg_ModoOpen
'                    Print #1, "a_unidadmed;" & RS!unm_codigo & ";insert into a_unidadmed values (" & RS!unm_codigo & "," & IIf(IsNull(RS!unm_nombre), "Null", "'" & RS!unm_nombre & "'") & "," & IIf(IsNull(RS!unm_nomcor), "Null", "'" & RS!unm_nomcor & "'") & ");" & "update a_unidadmed set unm_nombre=" & IIf(IsNull(RS!unm_nombre), "Null", "'" & RS!unm_nombre & "'") & ", unm_nomcor=" & IIf(IsNull(RS!unm_nomcor), "Null", "'" & RS!unm_nomcor & "'") & " where unm_codigo=" & RS!unm_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar nutriente aporte
              RS.Open "select distinct nut.* from a_nutriente nut, " & aAp & " tpro, b_productonut pnu, b_productosing pri where nut.nut_codigo=pnu.pnu_codapo and pnu.pnu_codpro=pri.pri_coding and pri.pri_codpro=tpro.codpro and pnu.pnu_canapo>0", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into a_nutriente values (" & RS!nut_codigo & ", " & IIf(IsNull(RS!nut_nombre), "Null", "'" & RS!nut_nombre & "'") & ", " & IIf(IsNull(RS!nut_nomuni), "Null", "'" & RS!nut_nomuni & "'") & ", " & IIf(IsNull(RS!nut_indpri), "Null", RS!nut_indpri) & ", " & IIf(IsNull(RS!nut_secnro), "Null", RS!nut_secnro) & ")", vg_ModoOpen
'                    Print #1, "a_nutriente;" & RS!nut_codigo & ";insert into a_nutriente values (" & RS!nut_codigo & "," & IIf(IsNull(RS!nut_nombre), "Null", "'" & RS!nut_nombre & "'") & "," & IIf(IsNull(RS!nut_nomuni), "Null", "'" & RS!nut_nomuni & "'") & "," & IIf(IsNull(RS!nut_indpri), "Null", RS!nut_indpri) & "," & IIf(IsNull(RS!nut_secnro), "Null", RS!nut_secnro) & ");" & "update a_nutriente set nut_nombre=" & IIf(IsNull(RS!nut_nombre), "Null", "'" & RS!nut_nombre & "'") & ", nut_nomuni=" & IIf(IsNull(RS!nut_nomuni), "Null", "'" & RS!nut_nomuni & "'") & ", nut_indpri=" & IIf(IsNull(RS!nut_indpri), "Null", RS!nut_indpri) & ", rs!nut_secnro=" & IIf(IsNull(RS!nut_secnro), "Null", RS!nut_secnro) & " where nut_codigo=" & RS!nut_codigo & ""
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
              '------- generar proveedores
              RS.Open "select * from b_proveedor", vg_db, adOpenStatic
              If Not RS.EOF Then
                 Do While Not RS.EOF
                    db7.Execute "insert into b_proveedor values ('" & RS!prv_codigo & "', '" & LimpiaDato(TipoDato(RS!prv_nombre, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_direccion, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_comuna, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_ciudad, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_fono1, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_fono2, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_fax, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_percon, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_giro, "")) & "', '" & LimpiaDato(TipoDato(RS!prv_emapro, "")) & "')", vg_ModoOpen
                    RS.MoveNext
                 Loop
              End If
              RS.Close: Set RS = Nothing
           End If
           For j = 1 To vaSpread1(1).MaxRows
               Bar1(1).Value = Val((j / vaSpread1(1).MaxRows) * 100)
               vaSpread1(1).Row = j: vaSpread1(1).Col = 1
               If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then
                  vaSpread1(1).Col = 2
                  codpro = Trim(Str(vaSpread1(1).text))
                  '------- Leer, insertar y regrabar productos casinos
                  RS.Open "select prc_cencos, prc_codpro from b_productocasino where prc_cencos='" & cencos & "' and prc_codpro='" & codpro & "'", vg_db, adOpenStatic
                  If RS.EOF Then
                     vg_db.Execute "insert into b_productocasino (prc_cencos, prc_codpro, prc_fecenv) values ('" & cencos & "', '" & codpro & "', " & Format(Date, "yyyymmdd") & ")"
                  Else
                     vg_db.Execute "update b_productocasino set prc_fecenv=" & Format(Date, "yyyymmdd") & " where prc_cencos='" & cencos & "' and prc_codpro='" & codpro & "'"
                  End If
                  RS.Close: Set RS = Nothing
                  '------- Fin leer, insertar y regrabar productos casinos
                  If icopy = False Then
                     '------- Generar productos
                     RS.Open "select distinct pro.* from b_productos pro where pro.pro_codigo='" & codpro & "'", vg_db, adOpenStatic
                     If Not RS.EOF Then
                        Do While Not RS.EOF
                           db7.Execute "insert into b_productos values ('" & RS!pro_codigo & "' ," & IIf(IsNull(RS!pro_codbar), "Null", "'" & RS!pro_codbar & "'") & ", " & IIf(IsNull(RS!pro_codcom), "Null", "'" & RS!pro_codcom & "'") & ", " & IIf(IsNull(RS!pro_codtip), "Null", RS!pro_codtip) & ", " & _
                                     "" & IIf(IsNull(RS!pro_nombre), "Null", "'" & RS!pro_nombre & "'") & ", " & IIf(IsNull(RS!pro_coduni), "Null", RS!pro_coduni) & ", " & IIf(IsNull(RS!pro_facing), "Null", RS!pro_facing) & ", " & IIf(IsNull(RS!pro_facsto), "Null", RS!pro_facsto) & ", " & IIf(IsNull(RS!pro_codemb), "Null", RS!pro_codemb) & ", " & _
                                     "" & IIf(IsNull(RS!pro_uniemb), "Null", RS!pro_uniemb) & ", " & IIf(IsNull(RS!pro_upreco), "Null", RS!pro_upreco) & ", " & IIf(IsNull(RS!pro_fecuco), "Null", "'" & RS!pro_fecuco & "'") & ", " & IIf(IsNull(RS!pro_propon), "Null", RS!pro_propon) & ", " & IIf(IsNull(RS!pro_ctacon), "Null", "'" & RS!pro_ctacon & "'") & ", " & _
                                     "" & IIf(IsNull(RS!pro_fecven), "Null", RS!pro_fecven) & ", " & IIf(IsNull(RS!pro_ctrsto), "Null", RS!pro_ctrsto) & ")", vg_ModoOpen
'                           Print #1, "b_productos;" & RS!pro_codigo & ";insert into b_productos values ('" & RS!pro_codigo & "'," & IIf(IsNull(RS!pro_codbar), "Null", "'" & RS!pro_codbar & "'") & "," & IIf(IsNull(RS!pro_codcom), "Null", "'" & RS!pro_codcom & "'") & "," & IIf(IsNull(RS!pro_codtip), "Null", RS!pro_codtip) & "," & _
'                                     "" & IIf(IsNull(RS!pro_nombre), "Null", "'" & RS!pro_nombre & "'") & "," & IIf(IsNull(RS!pro_coduni), "Null", RS!pro_coduni) & "," & IIf(IsNull(RS!pro_facing), "Null", RS!pro_facing) & "," & IIf(IsNull(RS!pro_facsto), "Null", RS!pro_facsto) & "," & IIf(IsNull(RS!pro_codemb), "Null", RS!pro_codemb) & "," & _
'                                     "" & IIf(IsNull(RS!pro_uniemb), "Null", RS!pro_uniemb) & "," & IIf(IsNull(RS!pro_upreco), "Null", RS!pro_upreco) & "," & IIf(IsNull(RS!pro_fecuco), "Null", RS!pro_fecuco) & "," & IIf(IsNull(RS!pro_propon), "Null", RS!pro_propon) & "," & IIf(IsNull(RS!pro_ctacon), "Null", "'" & RS!pro_ctacon & "'") & "," & _
'                                     "" & IIf(IsNull(RS!pro_fecven), "Null", RS!pro_fecven) & ");" & "update b_productos set pro_codbar=" & IIf(IsNull(RS!pro_codbar), "Null", "'" & RS!pro_codbar & "'") & ", pro_codcom=" & IIf(IsNull(RS!pro_codcom), "Null", "'" & RS!pro_codcom & "'") & ", pro_codtip=" & IIf(IsNull(RS!pro_codtip), "Null", RS!pro_codtip) & "," & _
'                                     "pro_nombre=" & IIf(IsNull(RS!pro_nombre), "Null", "'" & RS!pro_nombre & "'") & ", pro_coduni=" & IIf(IsNull(RS!pro_coduni), "Null", RS!pro_coduni) & ", pro_facing=" & IIf(IsNull(RS!pro_facing), "Null", RS!pro_facing) & ", pro_facsto=" & IIf(IsNull(RS!pro_facsto), "Null", RS!pro_facsto) & "," & _
'                                     "pro_codemb=" & IIf(IsNull(RS!pro_codemb), "Null", RS!pro_codemb) & ", pro_uniemb=" & IIf(IsNull(RS!pro_uniemb), "Null", RS!pro_uniemb) & ", pro_ctacon=" & IIf(IsNull(RS!pro_ctacon), "Null", "'" & RS!pro_ctacon & "'") & ", pro_fecven=" & IIf(IsNull(RS!pro_fecven), "Null", RS!pro_fecven) & " where cta_codigo = " & RS!pro_codigo & ""
                           RS.MoveNext
                        Loop
                     End If
                     RS.Close: Set RS = Nothing
                     '------- generar productos impuestos
'                     Print #1, "b_productosimp;delete b_productosimp from b_productosimp where ipr_codpro=" & CodPro & ""
                     RS.Open "select distinct ipr.* from b_productosimp ipr where ipr.ipr_codpro='" & codpro & "'", vg_db, adOpenStatic
                     If Not RS.EOF Then
                        Do While Not RS.EOF
                           db7.Execute "insert into b_productosimp values ('" & RS!ipr_codpro & "', " & IIf(IsNull(RS!ipr_codimp), "Null", RS!ipr_codimp) & ")", vg_ModoOpen
'                           Print #1, "b_productosimp;insert into b_productosimp values (" & RS!ipr_codpro & "," & IIf(IsNull(RS!ipr_codimp), "Null", RS!ipr_codimp) & ")"
                           RS.MoveNext
                        Loop
                     End If
                     RS.Close: Set RS = Nothing
                     '------- generar productos ingredientes & ingredientes
'                     Print #1, "b_productosing;delete b_productosing from b_productosing where pri_codpro=" & CodPro & ""
                     RS.Open "select distinct pri.*, ing.* from b_productosing pri, b_ingrediente ing where ing.ing_codigo=pri.pri_coding and pri.pri_codpro='" & codpro & "'", vg_db, adOpenStatic
                     If Not RS.EOF Then
                        Do While Not RS.EOF
                           db7.Execute "insert into b_productosing values ('" & RS!pri_codpro & "', " & IIf(IsNull(RS!pri_coding), "Null", "'" & RS!pri_coding & "'") & ")", vg_ModoOpen
                           db7.Execute "insert into b_ingrediente values ('" & RS!ing_codigo & "', " & IIf(IsNull(RS!ing_nombre), "Null", "'" & RS!ing_nombre & "'") & ", " & IIf(IsNull(RS!ing_nomfan), "Null", "'" & RS!ing_nomfan & "'") & ", " & _
                                       "" & IIf(IsNull(RS!ing_unimed), "Null", RS!ing_unimed) & ", " & IIf(IsNull(RS!ing_pctapr), "Null", RS!ing_pctapr) & ", " & IIf(IsNull(RS!ing_pctcoc), "Null", RS!ing_pctcoc) & ", " & IIf(IsNull(RS!ing_pctnut), "Null", RS!ing_pctnut) & ", " & IIf(IsNull(RS!ing_facnut), "Null", RS!ing_facnut) & ", " & IIf(IsNull(RS!ing_indpav), "Null", RS!ing_indpav) & ", " & _
                                       "" & IIf(IsNull(RS!ing_indgrv), "Null", RS!ing_indgrv) & ", " & IIf(IsNull(RS!ing_precos), "Null", RS!ing_precos) & ", " & IIf(IsNull(RS!ing_feccos), "Null", RS!ing_feccos) & ", " & IIf(IsNull(RS!ing_codcom), "Null", "'" & RS!ing_codcom & "'") & ", " & IIf(IsNull(RS!ing_codped), "Null", "'" & RS!ing_codped & "'") & ")", vg_ModoOpen

'                           Print #1, "b_productosing;insert into b_productosing values (" & RS!pri_codpro & "," & IIf(IsNull(RS!pri_coding), "Null", RS!pri_coding) & ");" & RS!ing_codigo & ";insert into b_ingrediente values (" & RS!ing_codigo & "," & IIf(IsNull(RS!ing_nombre), "Null", "'" & RS!ing_nombre & "'") & "," & IIf(IsNull(RS!ing_nomfan), "Null", "'" & RS!ing_nomfan & "'") & "," & _
'                                     "" & IIf(IsNull(RS!ing_unimed), "Null", RS!ing_unimed) & "," & IIf(IsNull(RS!ing_pctapr), "Null", RS!ing_pctapr) & "," & IIf(IsNull(RS!ing_pctcoc), "Null", RS!ing_pctcoc) & "," & IIf(IsNull(RS!ing_pctnut), "Null", RS!ing_pctnut) & "," & IIf(IsNull(RS!ing_facnut), "Null", RS!ing_facnut) & "," & IIf(IsNull(RS!ing_indpav), "Null", RS!ing_indpav) & "," & _
'                                     "" & IIf(IsNull(RS!ing_indgrv), "Null", RS!ing_indgrv) & "," & IIf(IsNull(RS!ing_precos), "Null", RS!ing_precos) & "," & IIf(IsNull(RS!ing_feccos), "Null", RS!ing_feccos) & "," & IIf(IsNull(RS!ing_codcom), "Null", "'" & RS!ing_codcom & "'") & "," & IIf(IsNull(RS!ing_codped), "Null", "'" & RS!ing_codped & "'") & ");" & "update b_ingrediente set " & _
'                                     "ing_nombre=" & IIf(IsNull(RS!ing_nombre), "Null", "'" & RS!ing_nombre & "'") & ", ing_nomfan=" & IIf(IsNull(RS!ing_nomfan), "Null", "'" & RS!ing_nomfan & "'") & ", ing_unimed=" & IIf(IsNull(RS!ing_unimed), "Null", RS!ing_unimed) & ", ing_pctapr=" & IIf(IsNull(RS!ing_pctapr), "Null", RS!ing_pctapr) & ", ing_pctcoc=" & IIf(IsNull(RS!ing_pctcoc), "Null", RS!ing_pctcoc) & "," & _
'                                     "ing_pctnut=" & IIf(IsNull(RS!ing_pctnut), "Null", RS!ing_pctnut) & ", ing_facnut=" & IIf(IsNull(RS!ing_facnut), "Null", RS!ing_pctnut) & ", ing_indpav=" & IIf(IsNull(RS!ing_indpav), "Null", RS!ing_indpav) & ", ing_indgrv=" & IIf(IsNull(RS!ing_indgrv), "Null", RS!ing_indgrv) & " where ing_codigo=" & RS!ing_codigo & ""
                           RS.MoveNext
                        Loop
                     End If
                     RS.Close: Set RS = Nothing
                     '------- generar nutriente del ingrediente
'                     Print #1, "b_productonut;delete b_productonut from b_productonut where pnu_codpro=" & CodPro & ""
                     RS.Open "select distinct pnu.* from b_productosing pri, b_productonut pnu where pnu.pnu_codpro=pri.pri_coding and pri.pri_codpro='" & codpro & "' and pnu.pnu_canapo>0", vg_db, adOpenStatic
                     If Not RS.EOF Then
                        Do While Not RS.EOF
                           db7.Execute "insert into b_productonut values ('" & RS!pnu_codpro & "', " & IIf(IsNull(RS!pnu_codapo), "Null", RS!pnu_codapo) & ", " & IIf(IsNull(RS!pnu_canapo), "Null", RS!pnu_canapo) & ")", vg_ModoOpen
'                           Print #1, "b_productonut;insert into b_productonut values (" & RS!pnu_codpro & "," & IIf(IsNull(RS!pnu_codapo), "Null", RS!pnu_codapo) & "," & IIf(IsNull(RS!pnu_canapo), "Null", RS!pnu_canapo) & ")"
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
'                Dim cHost As String, cUser As String, cPass As String
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
                    oFTP.password = Cpass '"*sdxo123*" ' "shx873" 'fg_Desencripta(TipoDato(cPass, ""))
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
                           SendMail oMail, "Se Informa que el maestro de productos esta disponible. Para que usted pueda actualizar", "Se Informa que el maestro de productos esta disponible. Para que usted pueda actualizar", mdir & sourcefilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0
                        End If
                    End If
                 ElseIf RS!cli_openvio = 2 Then
                    If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
                       fg_descarga
                       MsgBox "Casino : (" & Trim(cencos) & ") " & nomcencos & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, Msgtitulo
                       fg_carga ""
                    Else
                       SendMail oMail, "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), "Adjunto archivo productos " & Format(Date, "dd/mm/yyyy"), mdir & sourcefilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1
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
    Me.Hide
    Unload Me
End Select

Exit Sub
fg_descarga
Bar1(0).Visible = False: Bar1(1).Visible = False
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
If Index = 1 Then Exit Sub
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
           aAp = Trim(vg_NUsr) & "_tmp_CasinoProductos"
           vg_db.Execute "delete " & aAp & " from " & aAp & ""
'           RS1.Open "select * from " & aAp & "", vg_db, adOpenStatic
'           RS1.Close: Set RS1 = Nothing
'           fg_CheckTmp aAp
'           RS1.Open "select distinct pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
'                    "into " & aAp & " " & _
'                    "from b_productos pro inner join b_productocasino pri " & _
'                    "on pro.pro_codigo = pri.prc_codpro " & _
'                    "where pri.prc_cencos='" & Trim(vaSpread1(0).Text) & "'", vg_db, adOpenStatic
           
           vg_db.Execute "insert into " & aAp & " select distinct pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
                         "from b_productos pro inner join b_productocasino pri " & _
                         "on pro.pro_codigo = pri.prc_codpro " & _
                         "where pri.prc_cencos='" & Trim(vaSpread1(0).text) & "'"
           Set RS1 = Nothing
            
           RS1.Open "select pro.pro_codigo, pro.pro_nombre, tmp.prc_codpro " & _
                    "from b_productos pro left join " & aAp & " tmp on pro.pro_codigo = tmp.pro_codigo " & _
                    "where IsNull(tmp.pro_codigo) order by tmp.pro_codigo", vg_db, adOpenStatic
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
End Sub

Sub MoverDatoGrilla()
On Error GoTo Man_Error
fg_carga "": estado = True: i = 1
'------- Mover casinos
If Est Then
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
         If estado = True Then
                     estado = False
            aAp = Trim(vg_NUsr) & "_tmp_CasinoProductos"
            fg_CheckTmp aAp
            RS1.Open "select distinct pro.pro_codigo, pro.pro_nombre, pri.prc_codpro " & _
                     "into " & aAp & " " & _
                     "from b_productos pro inner join b_productocasino pri " & _
                     "on pro.pro_codigo = pri.prc_codpro " & _
                     "where pri.prc_cencos='" & RS!cli_codigo & "'", vg_db, adOpenStatic
            Set RS1 = Nothing

            RS1.Open "select pro.pro_codigo, pro.pro_nombre, tmp.prc_codpro " & _
                     "from b_productos pro left join " & aAp & " tmp on pro.pro_codigo = tmp.pro_codigo " & _
                     "where IsNull(tmp.pro_codigo)", vg_db, adOpenStatic
'            RS1.Open "select distinct b_productos.pro_codigo, b_productos.pro_nombre, b_productocasino.prc_codpro " & _
'                     "from b_productos left join b_productocasino ON b_productos.pro_codigo = b_productocasino.prc_codpro " & _
'                     "where isnull(b_productocasino.prc_codpro) or b_productocasino.prc_cencos='" & RS!cli_codigo & "' order by b_productocasino.prc_codpro", vg_db, adOpenStatic
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

'------- Mover productos
vaSpread1(1).MaxRows = 0
RS.Open "select pro_codigo, pro_nombre, pro_codtip from b_productos order by pro_nombre", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
      vaSpread1(1).Row = vaSpread1(1).MaxRows
      estado = False
      For i = 1 To vaSpread1(2).MaxRows
          vaSpread1(2).Row = i
          vaSpread1(2).Col = 2
          If RS!pro_codigo = Trim(vaSpread1(2).text) Then estado = True: Exit For
      Next i
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
            
      If estado = False Then vaSpread1(1).RowHidden = True
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
