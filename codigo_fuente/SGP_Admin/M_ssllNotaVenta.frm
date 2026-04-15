VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ssllNotaVenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Venta"
   ClientHeight    =   9825
   ClientLeft      =   5085
   ClientTop       =   855
   ClientWidth     =   9270
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9825
   ScaleWidth      =   9270
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   6615
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   8775
      Begin VB.CommandButton Command1 
         Caption         =   "Emitir Nota de Venta"
         Height          =   375
         Left            =   6360
         TabIndex        =   5
         Top             =   6120
         Width           =   1815
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5775
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   7575
         _Version        =   393216
         _ExtentX        =   13361
         _ExtentY        =   10186
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
         MaxCols         =   2
         SpreadDesigner  =   "M_ssllNotaVenta.frx":0000
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   855
         Left            =   360
         TabIndex        =   15
         Top             =   4560
         Visible         =   0   'False
         Width           =   5415
         _Version        =   393216
         _ExtentX        =   9551
         _ExtentY        =   1508
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
         MaxCols         =   6
         SpreadDesigner  =   "M_ssllNotaVenta.frx":18E0
      End
   End
   Begin VB.Frame Frame5 
      ForeColor       =   &H80000000&
      Height          =   2175
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   8775
      Begin VB.Frame Frame1 
         Caption         =   "Porcentaje de Costo Servicio"
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
         Height          =   975
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   8415
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   375
            Left            =   2280
            TabIndex        =   18
            Top             =   360
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   661
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
            ButtonStyle     =   0
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            Text            =   "0"
            MaxValue        =   "100"
            MinValue        =   "0"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   360
            Left            =   3840
            TabIndex        =   17
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   8040
         Top             =   120
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
               Picture         =   "M_ssllNotaVenta.frx":317E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3480
         TabIndex        =   3
         Top             =   600
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   600
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
         Text            =   "05/2016"
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
         TabIndex        =   0
         Top             =   240
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3480
         TabIndex        =   1
         Top             =   240
         Width           =   4335
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
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
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
         TabIndex        =   11
         Top             =   720
         Width           =   3015
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
         TabIndex        =   10
         Top             =   720
         Width           =   2835
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
         Top             =   315
         Width           =   1410
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3000
         Picture         =   "M_ssllNotaVenta.frx":3518
         Top             =   120
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3495
         TabIndex        =   8
         Top             =   290
         Width           =   4335
      End
      Begin VB.Label Label0 
         Caption         =   "Facturado Al"
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
         TabIndex        =   7
         Top             =   675
         Width           =   1335
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ssllNotaVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS, RS1, RS2, RS3 As ADODB.Recordset
Dim strSQL As String
Dim modo As String
Dim codigo, Msgtitulo As String
Dim est As Boolean

Private Sub Command1_Click()
    If Not IsDate(fpDateTime1.text) Or Trim(fpDateTime1.text) = "" Then MsgBox "Periodo no valido", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Trim(fpText(0).text) = "" Or Trim(fpayuda(0).Caption) = "" Then MsgBox "No existe centro costo", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_EmitirNotaVenta Format(fpDateTime1.text, "yyyymm"), LimpiaDato(fpText(0).text)
End Sub

Private Sub fpDateTime1_Change()
If Not IsDate(fpDateTime1.text) Then Exit Sub
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
fpLongInteger1.Value = Format(0, fg_Pict(3, 0))
Toolbar3.Buttons(1).Visible = False
Toolbar3.Buttons(2).Visible = True
Toolbar3.Buttons(3).Visible = False
Toolbar3.Buttons(4).Visible = True
End Sub

Private Sub fpLongInteger1_Change()
If est Then Exit Sub
If Toolbar3.Buttons(1).Visible = True Then Exit Sub
Toolbar3.Buttons(1).Visible = True
Toolbar3.Buttons(2).Visible = False
Toolbar3.Buttons(3).Visible = True
Toolbar3.Buttons(4).Visible = False
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
    If est Then Exit Sub
    fpayuda(0).Caption = ""
    If modo = "" Then modo = "M"
    vaSpread1.MaxRows = 0
    vaSpread2.MaxRows = 0
    fpLongInteger1.Value = Format(0, fg_Pict(3, 0))
    Toolbar3.Buttons(1).Visible = False
    Toolbar3.Buttons(2).Visible = True
    Toolbar3.Buttons(3).Visible = False
    Toolbar3.Buttons(4).Visible = True
    Command1.Enabled = False
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
    'RS1.Open strSQL, vg_db, adOpenStatic
    Set RS1 = vg_db.Execute(strSQL)
    
    If Not RS1.EOF Then
        fpayuda(0).Caption = IIf(IsNull(Trim(RS1!cli_nombre) = ""), "", RS1!cli_nombre)
        vg_codigo = RS1!cli_codigo
        codi = 0
        'Command1.Enabled = True
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
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Image1_Click(Index As Integer)
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Mantenedor Centro Costo", "CentCost"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If Not (fpText(0).text = "") Then
            est = True
            MoverDatosGrilla
            Toolbar1.Buttons(15).Enabled = True
            
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = 2
            
            If vaSpread1.text <> 0 Then
                Command1.Enabled = True
            End If
            est = False
        End If
    End Select
End Sub

Private Sub Form_Load()
    Dim i As Long
    
    Me.HelpContextID = vg_OpcM
    Me.Height = 10335
    Me.Width = 9450
    fg_centra Me
    Msgtitulo = "Nota de Venta"
    modo = ""
    est = True
    Gl_Mo_Botones Me, 1
    Gl_Ac_Botones Me, 1, 1, modo
    
    Gl_Mo_Botones Me, 18
    
    For i = 1 To 14
        Toolbar1.Buttons(i).Visible = False
    Next i
    fpLongInteger1.MinValue = 0
    fpLongInteger1.MaxValue = 100
    fpLongInteger1.Value = Format(0, fg_Pict(3, 0))
    Toolbar1.Buttons(15).Enabled = False
    Command1.Enabled = False
    est = False
End Sub

Sub MoverDatosGrilla()
    fg_carga ""
    Dim RS, RS1, RS2, RS3, RS4 As ADODB.Recordset
    Dim x As Boolean
    Dim i As Long
    Dim periodo As String
    Dim SumaTotal As Double
    Dim Arreglo As String
    Dim SumaFamilia As Double
    Dim varStrFamilia As String
    
    vaSpread1.TextTip = 2
    vaSpread1.TextTipDelay = 250
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    vaSpread1.Lock = True
    
    periodo = Trim(Format(fpDateTime1, "yyyymm"))
    
    strSQL = "SELECT * FROM a_tipopro a WHERE a.tip_previo = 0 AND a.tip_activo = 1"
    Set RS1 = vg_db.Execute(strSQL)
    
    codtip = 0
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    
'---------------------------------------------------------------------------------------------------------
'    strSQL = "ssll_s_NotaVentaxFamilia02 " & periodo & ", '" & fpText(0).text & "'"
'    Set RS3 = vg_db.Execute(strSQL)
'
'    Dim personas() As String
'    Dim Contador As Integer
'
'    Contador = 1000 'RS3.RecordCount
'
'    ReDim Preserve personas(Contador, 5) As String
'
'    For i = 1 To Contador
'        personas(i, 1) = RS3("codpro")
'        personas(i, 2) = RS3("codtip")
'        personas(i, 3) = RS3("canmer")
'        personas(i, 4) = RS3("precos")
'        personas(i, 5) = RS3("unimed")
'    Next i
'---------------------------------------------------------------------------------------------------------

    strSQL = "ssll_s_NotaVentaxFamilia02 '" & periodo & "', '" & fpText(0).text & "'"
    Set RS3 = vg_db.Execute(strSQL)
    
    vaSpread2.MaxRows = 0
    k = 0
    
    Do While Not RS3.EOF
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 1: vaSpread2.text = RS3!codpro
        vaSpread2.Col = 2: vaSpread2.text = RS3!codtip
        vaSpread2.Col = 3: vaSpread2.text = RS3!canmer
        vaSpread2.Col = 4: vaSpread2.text = RS3!precos
        vaSpread2.Col = 5: vaSpread2.text = RS3!unimed
        vaSpread2.Col = 6: vaSpread2.text = IIf(IsNull(RS3!fampro), "", RS3!fampro)
        
        RS3.MoveNext
    Loop
'---------------------------------------------------------------------------------------------------------
    
    
    Do While Not RS1.EOF
    
        strSQL = "SELECT * " & _
                 "FROM a_tipopro " & _
                 "WHERE tip_previo = " & RS1!tip_codigo & " ORDER BY tip_nombre"
                 
        Set RS2 = vg_db.Execute(strSQL)
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1: vaSpread1.Lock = True: vaSpread1.Font.Bold = True: vaSpread1.text = RS1!tip_nombre 'fg_BuscaenArbolDosNiveles(RS1!tip_codigo, "a_tipopro", "tip_codigo")
        vaSpread1.BackColorStyle = BackColorStyleUnderGrid: vaSpread1.BackColor = &HC0C0FF '&HE0FEFE
        vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = ""
        vaSpread1.BackColorStyle = BackColorStyleUnderGrid: vaSpread1.BackColor = &HC0C0FF '&HE0FEFE
                
        Dim wvarFamPro, wvarCanMer, wvarPreCos As String
        
        Do While Not RS2.EOF
            varStrFamilia = Trim(RS1!tip_nombre & "\" & RS2!tip_nombre)
        
            If RS2!tip_previo <> 0 Then
                SumaFamilia = 0
                
               For i = vaSpread2.SearchCol(6, -1, vaSpread2.MaxRows, Trim(CStr(varStrFamilia)), SearchFlagsEqual) To vaSpread2.MaxRows
                   vaSpread2.Row = i 'vaSpread2.SearchCol(6, -1, vaSpread2.MaxRows, Trim(CStr(varStrFamilia)), SearchFlagsEqual)
                   vaSpread2.Col = 6: wvarFamPro = vaSpread2.text
                   If (varStrFamilia = wvarFamPro) Then
                      vaSpread2.Col = 3: wvarCanMer = vaSpread2.text
                      vaSpread2.Col = 4: wvarPreCos = vaSpread2.text
                      SumaFamilia = SumaFamilia + ((Trim(wvarCanMer)) * Trim(wvarPreCos))
                   Else
                      Exit For
                   End If
               Next i
'                For i = 1 To vaSpread2.MaxRows
'                    vaSpread2.Row = i
'                    vaSpread2.Col = 3: wvarCanMer = vaSpread2.text
'                    vaSpread2.Col = 4: wvarPreCos = vaSpread2.text
'                    vaSpread2.Col = 6: wvarFamPro = vaSpread2.text
'                    If (varStrFamilia = wvarFamPro) Then
'                        SumaFamilia = SumaFamilia + ((Trim(wvarCanMer)) * Trim(wvarPreCos))
'                    End If
'                Next i
                
                vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                vaSpread1.Row = vaSpread1.MaxRows
                vaSpread1.Col = 1: vaSpread1.Lock = True: vaSpread1.text = Trim(RS2!tip_nombre)
                vaSpread1.Col = 2: vaSpread1.Lock = True: vaSpread1.text = SumaFamilia
            End If
        RS2.MoveNext
        Loop
    RS1.MoveNext
    Loop
        
        
    vaSpread1.Col = 2
    
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        If vaSpread1.text <> "" Then
            SumaTotal = SumaTotal + Val(CDbl(vaSpread1.text))
        End If
    Next i
    fpLongInteger1.Value = Format(0, fg_Pict(3, 0))
    If vaSpread2.MaxRows > 0 Then '------- Traer porcentaje costo servicio
'        a = BoM(Format(fpDateTime1.Value, "dd/mm/yyyy"))
        strSQL = "SELECT * " & _
                 "FROM b_ssll_porctoser " & _
                 "WHERE pcs_codcen = '" & fpText(0).text & "' and pcs_period = '" & periodo & "'"
        Set RS4 = vg_db.Execute(strSQL)
        If RS4.EOF Then
           RS4.Close: Set RS4 = Nothing
           strSQL = "SELECT * " & _
                    "FROM b_ssll_porctoser " & _
                    "WHERE pcs_codcen = '" & fpText(0).text & "' and pcs_period = '" & Format(BoM(Format(fpDateTime1.Value, "dd/mm/yyyy")), "yyyymm") & "'"
           Set RS4 = vg_db.Execute(strSQL)
           If Not RS4.EOF Then
              vg_db.Execute "INSERT INTO b_ssll_porctoser (pcs_codcen, pcs_period, pcs_porcen) values ('" & fpText(0).text & "', '" & periodo & "', " & RS4!pcs_porcen & ")"
              fpLongInteger1.Value = Format(RS4!pcs_porcen, fg_Pict(3, 0))
           End If
        Else
           fpLongInteger1.Value = Format(RS4!pcs_porcen, fg_Pict(3, 0))
        End If
        RS4.Close: Set RS4 = Nothing
        Frame1.Enabled = True
    Else
        Frame1.Enabled = False
    End If
    vaSpread1.MaxRows = vaSpread1.MaxRows + 2
    vaSpread1.Row = vaSpread1.MaxRows

    vaSpread1.Col = 1: vaSpread1.Font.Bold = True: vaSpread1.text = "TOTAL (sin IVA)"
    vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.text = SumaTotal

    vaSpread1.SetActiveCell 1, 1
        
    'RS.Close: Set RS = Nothing
    RS1.Close: Set RS1 = Nothing
    RS2.Close: Set RS2 = Nothing
    RS3.Close: Set RS3 = Nothing
    
    vaSpread1.Visible = True
    
    If (SumaTotal = 0) Then
        Command1.Enabled = False
    Else
        Command1.Enabled = True
    End If
    
    fg_descarga
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error

Select Case Button.Index

Case 15 'IMPRIMIR
    vaSpread1.Row = 1
    vaSpread1.Col = 1

    If Trim(vaSpread1.text) = "" Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_SsllNotaVenta

Case 18 'SALIR
    Me.Hide
    Unload Me
End Select

Man_Error:
    If Err = 3034 Then Exit Sub
    If Err = 13 Then Exit Sub
End Sub

Public Function fg_BuscaenArbolDosNiveles(codigo As Long, Tabla As String, CampoBus As String) As String
Dim RS5 As New ADODB.Recordset
Dim Nombre As String
Dim i As Long
    
    Nombre = ""
    
    For i = 1 To 4
        If codigo = 0 Then Exit For
        Set RS5 = vg_db.Execute("SELECT * FROM " & Tabla & " WHERE " & CampoBus & " = " & codigo & "")
        
        If RS5.EOF Then RS5.Close: Set RS5 = Nothing: Exit For
        
        Nombre = Trim(RS5(1)) & "\" & Nombre
        codigo = RS5(2)
        
        If (i = 1) Then
            wvartexto1 = Trim(RS5(1))
        ElseIf (i = 2) Then
            wvartexto2 = Trim(RS5(1))
        ElseIf (i = 3) Then
            wvartexto3 = Trim(RS5(1))
        ElseIf (i = 4) Then
            wvartexto4 = Trim(RS5(1))
        End If
        
        If RS5(0) = 0 Then RS5.Close: Set RS5 = Nothing: Exit For
        RS5.Close: Set RS5 = Nothing
    Next
    
    'If Right(Trim(Nombre), 1) = "\" Then
        
        'Nombre2Niveles = Left(Trim(Nombre), Len(Nombre) - 1)
        Nombre2Niveles = Trim(Nombre)
        
    'End If
    
    
    
    If Trim(Nombre2Niveles) <> "" Then fg_BuscaenArbolDosNiveles = Mid(Nombre2Niveles, 1, Len(Nombre2Niveles) - 1) Else fg_BuscaenArbolDosNiveles = ""
    
End Function

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 'cancelar
    fpLongInteger1.Value = Format(0, fg_Pict(3, 0))
    strSQL = "select * " & _
             "from b_ssll_porctoser " & _
             "where pcs_codcen = '" & fpText(0).text & "' and pcs_period = '" & Format(fpDateTime1.Value, "yyyymm") & "'"
    Set RS4 = vg_db.Execute(strSQL)
    If Not RS4.EOF Then
       fpLongInteger1.Value = Format(RS4!pcs_porcen, fg_Pict(3, 0))
    End If
    RS4.Close: Set RS4 = Nothing
Case 3 'grabar
    If Not IsDate(fpDateTime1.text) Or Trim(fpDateTime1.text) = "" Then MsgBox "Periodo no valido", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Trim(fpText(0).text) = "" Or Trim(fpayuda(0).Caption) = "" Then MsgBox "No existe centro costo", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    strSQL = "select * " & _
             "from b_ssll_porctoser " & _
             "where pcs_codcen = '" & fpText(0).text & "' and pcs_period = '" & Format(fpDateTime1.Value, "yyyymm") & "'"
    Set RS4 = vg_db.Execute(strSQL)
    If RS4.EOF Then
       vg_db.Execute "insert into b_ssll_porctoser (pcs_codcen, pcs_period, pcs_porcen) values ('" & fpText(0).text & "', '" & Format(fpDateTime1.Value, "yyyymm") & "', " & fpLongInteger1.Value & ")"
    Else
       vg_db.Execute "update b_ssll_porctoser set pcs_porcen = " & fpLongInteger1.Value & " where pcs_codcen = '" & fpText(0).text & "' and pcs_period = '" & Format(fpDateTime1.Value, "yyyymm") & "'"
    End If
    RS4.Close: Set RS4 = Nothing
End Select
Toolbar3.Buttons(1).Visible = False
Toolbar3.Buttons(2).Visible = True
Toolbar3.Buttons(3).Visible = False
Toolbar3.Buttons(4).Visible = True
End Sub
