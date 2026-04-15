VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_Produc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos"
   ClientHeight    =   7230
   ClientLeft      =   3675
   ClientTop       =   2475
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   6705
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   6420
      Begin VB.Frame Frame1 
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
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   0
         Left            =   165
         TabIndex        =   8
         Top             =   1200
         Width           =   6150
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
            ItemData        =   "B_Produc.frx":0000
            Left            =   1875
            List            =   "B_Produc.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   240
            Width           =   2895
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1875
            TabIndex        =   10
            Top             =   555
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
               Weight          =   700
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
            AutoAdvance     =   -1  'True
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
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
            Left            =   4860
            TabIndex        =   13
            Top             =   645
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   " Buscar Texto"
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
            Left            =   360
            TabIndex        =   12
            Top             =   645
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   " Buscar Columna"
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
            Left            =   360
            TabIndex        =   11
            Top             =   345
            Width           =   1440
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   165
         TabIndex        =   4
         Top             =   210
         Width           =   6150
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
            Index           =   0
            ItemData        =   "B_Produc.frx":001E
            Left            =   855
            List            =   "B_Produc.frx":0020
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   540
            Width           =   4110
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
            Left            =   855
            TabIndex        =   6
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
            Left            =   4110
            TabIndex        =   5
            Top             =   255
            Width           =   855
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Selección Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   165
         TabIndex        =   1
         Top             =   2340
         Width           =   6150
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3735
            Left            =   135
            TabIndex        =   2
            Top             =   330
            Width           =   5940
            _Version        =   393216
            _ExtentX        =   10478
            _ExtentY        =   6588
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            AutoClipboard   =   0   'False
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
            MaxRows         =   20
            SpreadDesigner  =   "B_Produc.frx":0022
            ScrollBarTrack  =   3
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "B_Produc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim sqlTIP As String, sqlALE As String

Private Sub Combo1_Click(Index As Integer)
Select Case Index
Case 0 'Tipo Producto
    Dim codtip As Long
    sqlTIP = ""
    codtip = fg_codigocbo(Combo1, 0, 10, "")
    If Combo1(0).ListIndex > -1 Then sqlTIP = "pro_codtip = " & codtip
    MuestraGrilla
Case 1 'Aleatorio
    fpText.Text = ""
End Select
End Sub

Private Sub Form_Load()
Me.Width = 6600
Me.Height = 7635
MsgTitulo = "Productos"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirma"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(0).Clear
RS1.Open "select * from a_tipopro order by tip_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo1(0).AddItem RS1!tip_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!tip_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
optTIPPRO(1).Value = True
sqlALE = "": sqlTIP = ""
Combo1(1).ListIndex = 1
MuestraGrilla
End Sub

Private Sub fpText_Change()
sqlALE = ""
If LimpiaDato(Trim(fpText.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Then
    sqlALE = "pro_codigo like '%" + UCase(LimpiaDato(fpText.Text)) + "%'"
ElseIf Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
    sqlALE = "Ucase(pro_nombre) like '%" + UCase(LimpiaDato(fpText.Text)) + "%'"
End If
MuestraGrilla
End Sub

Private Sub optTIPPRO_Click(Index As Integer)
    Combo1(0).Enabled = IIf(Index = 0, True, False)
    Combo1(0).ListIndex = IIf(Index = 0, 0, -1)
End Sub

Private Sub MuestraGrilla()
Dim sqlFIN As String
sqlFIN = "select pro_codigo, pro_nombre from b_productos"
If sqlALE <> "" And sqlTIP = "" Then sqlFIN = sqlFIN & " where " & sqlALE
If sqlALE = "" And sqlTIP <> "" Then sqlFIN = sqlFIN & " where " & sqlTIP
If sqlALE <> "" And sqlTIP <> "" Then sqlFIN = sqlFIN & " where " & sqlALE & " and " & sqlTIP
sqlFIN = sqlFIN & " order by pro_nombre"
RS1.Open sqlFIN, vg_db, adOpenStatic
vaSpread1.MaxRows = 0
Do While Not RS1.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    'vaSpread1.Col = 1: vaSpread1.Value = "1"
    vaSpread1.Col = 2: vaSpread1.TypeHAlign = 1: vaSpread1.CellType = 5
    vaSpread1.Value = RS1!pro_codigo
    vaSpread1.Col = 3: vaSpread1.TypeHAlign = 0: vaSpread1.CellType = 5
    vaSpread1.Value = RS1!pro_nombre
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
vaSpread1.Col = 1: vaSpread1.Row = -1: vaSpread1.Value = "1"
If fpText.Text = "" Then
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim aAp As String, i As Long
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows = 0 Then GoTo Salir
    vg_codigo = "|Ok|"
    aAp = Trim(vg_NUsr) & "_tmp_filtomainv"
    fg_CheckTmp aAp
    vg_db.BeginTrans
    vg_db.Execute "create table " & aAp & " (tem_codigo varchar(20))"
    vg_db.CommitTrans
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.Value = "1" Then
            vaSpread1.Col = 2
            vg_db.BeginTrans
            vg_db.Execute "insert into " & aAp & " (tem_codigo) values ('" & vaSpread1.Text & "')"
            vg_db.CommitTrans
        End If
    Next i
End Select
Salir:
Me.Hide
Unload Me
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
'If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
vaSpread1.Col = 1
For i = BlockRow To BlockRow2
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
Next
End Sub
