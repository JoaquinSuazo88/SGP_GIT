VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form I_SsLlGen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe Servicios Logístico"
   ClientHeight    =   2970
   ClientLeft      =   5715
   ClientTop       =   4710
   ClientWidth     =   9720
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   9720
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Width           =   9255
      Begin VB.Frame Frame5 
         ForeColor       =   &H80000000&
         Height          =   1695
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   8775
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   6135
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Left            =   1680
            TabIndex        =   3
            Top             =   1120
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
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
            Text            =   "09/2019"
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/yyyy"
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   1
            Top             =   720
            Width           =   1275
            _Version        =   196608
            _ExtentX        =   2249
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   1700
            TabIndex        =   14
            Top             =   360
            Width           =   6160
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo Informes"
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
            Left            =   120
            TabIndex        =   13
            Top             =   340
            Width           =   1335
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3480
            TabIndex        =   2
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label Label0 
            Caption         =   "Periodo"
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
            Left            =   120
            TabIndex        =   11
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3510
            TabIndex        =   10
            Top             =   765
            Width           =   4335
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   3000
            Picture         =   "I_SsLlGen.frx":0000
            Top             =   600
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Centro De Costo"
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
            Index           =   5
            Left            =   120
            TabIndex        =   9
            Top             =   795
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Left            =   3480
            TabIndex        =   8
            Top             =   720
            Width           =   2835
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
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
            Height          =   255
            Left            =   5280
            TabIndex        =   7
            Top             =   720
            Width           =   3015
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Left            =   3480
            TabIndex        =   6
            Top             =   720
            Visible         =   0   'False
            Width           =   2835
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_SsLlGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim RS As ADODB.Recordset
Dim strSQL As String
Dim MsgTitulo As Variant

Private Sub Combo1_Click(Index As Integer)
fpDateTime1.Enabled = True
Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
Case 0
    Me.Caption = "Consolidado Facturación Cliente"
    MsgTitulo = "Consolidado Facturación Cliente"
Case 1
    Me.Caption = "Top 10 de Productos"
    MsgTitulo = "Top 10 de Productos"
Case 2
    Me.Caption = "Canasta de Medición"
    MsgTitulo = "Canasta de Medición"
Case 3
    Me.Caption = "80 - 20 Consumo"
    MsgTitulo = "80 - 20 Consumo"
Case 4
    fpDateTime1.Enabled = False
    Me.Caption = "Nivel de Servicio"
    MsgTitulo = "Nivel de Servicio"
Case 5
    fpDateTime1.Enabled = False
    Me.Caption = "Evolución Compras Familia"
    MsgTitulo = "Evolución Compras Familia"
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
    fpayuda(0).Caption = ""
End Sub

Private Sub fpText_LostFocus(Index As Integer)
    Dim codi As Long, Bd As String, Ul As String
    On Error GoTo Man_Error
    If fpText(Index).text = "" Then fpayuda(0).Caption = "": codi = 0: Exit Sub
    
    codi = fpText(Index).text
    Bd = IIf(Index = 0, "b_clientes", "")
    Ul = IIf(Bd = "b_clientes", "cli", "")
    
    Set RS1 = Nothing
    
    strSQL = "SELECT " & Ul & "_codigo, " & Ul & "_nombre FROM " & Bd & " WHERE " & Ul & "_codigo = " & IIf(Ul = "cli", "'" & codi & "'", codi) & ""
    Set RS1 = vg_db.Execute(strSQL)
    
    If Not RS1.EOF Then
        fpayuda(0).Caption = IIf(IsNull(Trim(RS1!cli_nombre) = ""), "", RS1!cli_nombre)
        vg_codigo = RS1!cli_codigo
        codi = 0
    Else
        MsgBox "No existe codigo en la tabla..."
        fpayuda(0).Caption = ""
        fpText(Index).text = ""
        codi = 0
        On Error Resume Next: fpText(Index).SetFocus
    End If
    
    RS1.Close: Set RS1 = Nothing
    Exit Sub
    
Man_Error:
    If Err = 3034 Then Exit Sub
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Image1_Click(Index As Integer)
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Mantenedor Centro Costo", "CentCost"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(0).Caption = vg_nombre
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1 'IMPRIMIR
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then MsgBox "No Existen seleccionado tipo informe.", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    If Trim(fpText(0).text) = "" Or Trim(fpayuda(0).Caption) = "" Then MsgBox "No Existen centro de costo.", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    If IsNull(fpDateTime1.text) Then MsgBox "No esta definida la fecha.", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    'I_SsLlGen.Label6.Caption = fpDateTime1.text
    
    Select Case Val(fg_codigocbo(Combo1, 0, 1, ""))
    Case 0
        I_ConsolidadoFacturacion fpText(0).text, Format(fpDateTime1.text, "yyyymm") 'Informe Consolidado de Facturación al Cliente
    Case 1
        I_SSLL_Top10 fpText(0).text, Format(fpDateTime1.text, "yyyymm") 'Top10
    Case 2
        I_CanastaMedicion fpText(0).text, Format(fpDateTime1.text, "yyyymm") 'Canasta de Medición
    Case 3
        I_OchentaVeinte fpText(0).text, Format(fpDateTime1.text, "yyyymm") '80/20
    Case 4
        I_SSLL_NivelServicio fpText(0).text 'Nivel de Servicio
    Case 5
        I_SSLL_ComprasEvolucion fpText(0).text  'Evolución Compras Familias
    End Select
Case 3 'SALIR
    Me.Hide
    Unload Me
End Select

Man_Error:
    If Err = 3034 Then Exit Sub
    If Err = 13 Then Exit Sub
End Sub

Private Sub Form_Load()

Combo1(0).Clear
Combo1(0).AddItem "Consolidado Facturación Cliente" & Space(150) & "(0)"
Combo1(0).AddItem "Top 10 de Productos" & Space(150) & "(1)"
Combo1(0).AddItem "Canasta de Medición" & Space(150) & "(2)"
Combo1(0).AddItem "80 - 20 Consumo" & Space(150) & "(3)"
Combo1(0).AddItem "Nivel de Servicio" & Space(150) & "(4)"
Combo1(0).AddItem "Evolución Compras Familia" & Space(150) & "(5)"
Combo1(0).ListIndex = -1

Me.Height = 3480
Me.Width = 9840
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub
