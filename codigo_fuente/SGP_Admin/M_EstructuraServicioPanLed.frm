VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_EstructuraServicioPanLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parametrizar Estructura Servicio Pantalla Led"
   ClientHeight    =   10110
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10110
   ScaleWidth      =   16710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   720
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   16455
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   1800
         TabIndex        =   19
         Top             =   7200
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   2715
         TabIndex        =   18
         Top             =   7200
         Width           =   7110
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   7005
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6735
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   15975
         _Version        =   393216
         _ExtentX        =   28178
         _ExtentY        =   11880
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
         SpreadDesigner  =   "M_EstructuraServicioPanLed.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   11775
      Begin VB.CommandButton Command1 
         Caption         =   "Historial"
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
         Left            =   10680
         TabIndex        =   20
         Top             =   360
         Width           =   855
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1995
         TabIndex        =   1
         Top             =   735
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1995
         TabIndex        =   3
         Top             =   1155
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   2
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
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   2010
         TabIndex        =   0
         Top             =   315
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   3285
         Picture         =   "M_EstructuraServicioPanLed.frx":1A49
         Top             =   1080
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3285
         Picture         =   "M_EstructuraServicioPanLed.frx":1D53
         Top             =   660
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3735
         TabIndex        =   13
         Top             =   1155
         Width           =   6735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3735
         TabIndex        =   12
         Top             =   735
         Width           =   6735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         Left            =   720
         TabIndex        =   11
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   720
         TabIndex        =   10
         Top             =   840
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3300
         Picture         =   "M_EstructuraServicioPanLed.frx":205D
         Top             =   240
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3750
         TabIndex        =   9
         Top             =   315
         Width           =   6735
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
         Index           =   0
         Left            =   735
         TabIndex        =   8
         Top             =   420
         Width           =   735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Height          =   270
         Index           =   0
         Left            =   3795
         TabIndex        =   16
         Top             =   360
         Width           =   6735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Height          =   270
         Index           =   1
         Left            =   3780
         TabIndex        =   14
         Top             =   780
         Width           =   6735
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
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
         Height          =   270
         Index           =   2
         Left            =   3780
         TabIndex        =   15
         Top             =   1200
         Width           =   6735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   16710
      _ExtentX        =   29475
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   300
   End
End
Attribute VB_Name = "M_EstructuraServicioPanLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Public modo   As String
Dim MsgTitulo As String

Private Sub Command1_Click()

On Error GoTo Man_Error

If Trim(LimpiaDato(fpText.text)) = "" Then

   fg_descarga
   MsgBox "Debe seleccionar un Ceco ", vbCritical, MsgTitulo
   Exit Sub

End If

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Listar_Lista"), CStr(Me.HelpContextID), "", "", "")

vg_codcasino = ""
B_HistorialEstructuraPanLed.LlenarHistorial (LimpiaDato(fpText.text))
B_HistorialEstructuraPanLed.Show 1

If Trim(vg_codcasino) <> "" Then

   fpText.text = vg_codcasino
   fpLongInteger1(0).text = vg_codregimen
   fpLongInteger1(1).text = vg_codservicio

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()
    
On Error GoTo Man_Error

    Call fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
    Call fg_carga("")
    Me.HelpContextID = vg_OpcM
    MsgTitulo = "Parametrizar Estructura Servicio Pantalla Led"
    Call fg_centra(Me)
    Let Me.Height = 10545
    Let Me.Width = 16800
    Frame1.Enabled = True
    modo = ""
    Gl_Mo_Botones Me, 1
    Gl_Ac_Botones Me, 1, 19, modo
    
    Call FormatearDatos
    Call fg_descarga
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub FormatearDatos()

On Error GoTo Man_Error

    Let vaSpread1.MaxRows = 0
    Let fpText.text = ""
    Let fpLongInteger1(0).Value = ""
    Let fpLongInteger1(1).Value = ""
    Let TextDet1(2).text = ""
    Let TextDet1(3).text = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Index
    
    Case 0
        
        Set RS = vg_db.Execute("sgpadm_Sel_RegimenBloque " & Val(fpLongInteger1(0).Value) & "")
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(1).Caption = ""
            Exit Sub
        
        End If
        fpayuda(1).Caption = Trim(RS!reg_nombre)
        RS.Close
        Set RS = Nothing
       
    Case 1
        
        Set RS = vg_db.Execute("sgpadm_Sel_ServicioBloque " & Val(fpLongInteger1(1).Value) & "")
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(2).Caption = ""
            Exit Sub
        
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close
        Set RS = Nothing

End Select
    
Call MoverGrilla(fpText.text)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Call MoverGrilla(fpText.text)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Man_Error
    
    Select Case KeyCode
        
        Case 120
            
            Image1_Click 0
    
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

    Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
    If RS.EOF Then
        
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        fpLongInteger1(0).Value = ""
        fpayuda(1).Caption = ""
        fpLongInteger1(1).Value = ""
        fpayuda(2).Caption = ""
        vaSpread1.MaxRows = 0
        Exit Sub
    
    End If
    fpayuda(0).Caption = Trim(RS!Cli_nombre)
    fpText.text = RS!Cli_codigo
    RS.Close
    Set RS = Nothing
 
    fpLongInteger1(0).Value = ""
    fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = ""
    fpayuda(2).Caption = ""
    Call MoverGrilla(fpText.text)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo


End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error
    
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub MoverGrilla(cencos As String)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Frame1.Enabled = True
Gl_Ac_Botones Me, 1, 19, modo

If Trim(cencos) = "" And Trim(fpLongInteger1(0).text) = "" And Trim(fpLongInteger1(1).text) = "" Then Exit Sub

Set RS = vg_db.Execute("sgpadm_Sel_ListarEstServicioxCecoPanLed '" & cencos & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & "")

Call fg_carga("")
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
vaSpread1.Row = -1: vaSpread1.Col = 1

Do While Not RS.EOF = True
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 1
   vaSpread1.text = RS!Sel
   
   vaSpread1.Col = 2
   vaSpread1.text = RS!ess_codigo
   
   vaSpread1.Col = 3
   vaSpread1.text = RS!ess_nombre
   
   vaSpread1.Col = 4
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = 0
   
   vaSpread1.Col = 5
   vaSpread1.text = RS!EstHomologacionClan
   
   vaSpread1.Col = 6
   vaSpread1.text = RS!NomEstHomologacionClan
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing
vaSpread1.Visible = True
Call fg_descarga

Exit Sub
Man_Error:
    Frame1.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
    Select Case Index
        
        Case 0
            
            vg_left = fpayuda(0).Left + 2300
            vg_nombre = ""
            vg_codigo = ""
            Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpText.text = vg_codigo
            fpayuda(0).Caption = vg_nombre
            fpLongInteger1(0).Value = ""
            Let fpayuda(1).Caption = ""
            fpLongInteger1(1).Value = ""
            Let fpayuda(2).Caption = ""
            fpLongInteger1(0).SetFocus
        
        Case 1
            
            vg_left = fpayuda(1).Left + 2300
            vg_nombre = ""
            vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_regimen", "", "Regimen", "RegBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(0).Value = Val(vg_codigo)
            fpLongInteger1(0).SetFocus
            fpayuda(1).Caption = vg_nombre
            fpLongInteger1(1).SetFocus
        
        Case 2
            
            Let vg_left = fpayuda(2).Left + 2300
            Let vg_nombre = ""
            Let vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_servicio", "", "Servicio", "SerBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(1).Value = Val(vg_codigo)
            fpLongInteger1(1).SetFocus
            fpayuda(2).Caption = vg_nombre
            
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
   
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

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 4
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread1.Visible = False
    
    If Trim(TextDet1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 1
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 4
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 4
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 4
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 4
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 1
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 4
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 4
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index
    
    Case 1
        
    Case 3 ' Modo Modificar
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), CStr(Me.HelpContextID), "", "", "")
        
        If modo = "" Then modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
        Frame1.Enabled = False
        
    Case 5 'Modo Eliminar
        
        If MsgBox("Elimina...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        Call EliminarGrilla
        
    Case 7

        Call MoverGrilla(fpText.text)
    
    Case 10
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Cancelar"), CStr(Me.HelpContextID), "", "", "")
        Call MoverGrilla(fpText.text)
    
    Case 12
        
        GrabaRegistro
    
    Case 15
        
        ExportarExcel

    Case 18
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "", "")
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ExportarExcel()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim Sql             As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

'-------> Validar cantidad registro se sobre pase hoja excel
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ExportarExcelPanLed ")

If Not RS.EOF Then
  
   If RS.RecordCount > 1020000 Then
      
      RS.Close
      Set RS = Nothing
      
      MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ceco", vbCritical
      Exit Sub
   
   End If
  
End If

'-------> Guardar nombre archivo excel
NomArchivoExcel = ""
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.Filter = "Todos los archivos *.xls,*.xlsx"
On Error Resume Next
CD.ShowSave
           
'-------> JPAZ Permite controlar Boton Cancelar
If Err.Number = 32755 Then
   
   MsgBox "Proceso cancelado"
   Exit Sub

End If
            
If CD.FileName = "" Then
   
   MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
   Exit Sub

Else
   
   Extension = ""
   Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
   
   If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
      MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
      Exit Sub
   End If
   
   NomArchivoExcel = CD.FileName

End If
          
fg_carga ""
  
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")
  
'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True
    
'-------> Check version of Excel
Call encabezado(RS, xlWs)
          
xlWs.Cells(2, 1).CopyFromRecordset RS

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'xlApp.Columns("A:A").Select
'xlApp.Selection.Delete Shift:=xlToLeft
  
xlWb.Close True, NomArchivoExcel

Dim XL As New excel.Application 'Crea el objeto excel
XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
'-------> Close ADO objects
RS.Close
Set RS = Nothing
    
' -- Cerrar Excel
xlApp.Quit
'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing
  
fg_descarga
MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub GrabaRegistro()

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim i          As Long
Dim isel       As Boolean
Dim Ceco       As String
Dim regimen    As Long
Dim Servicio   As Long
Dim Estructura As Long
Dim MyBuffer   As String

Ceco = fpText.text
regimen = fpLongInteger1(0).Value
Servicio = fpLongInteger1(1).Value

If Trim(fpayuda(0).Caption) = "" Then

   MsgBox "Debe ingresar Centro Costo...", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub

End If

If Trim(fpayuda(1).Caption) = "" Then

   MsgBox "Debe ingresar Regimen...", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub

End If

If Trim(fpayuda(2).Caption) = "" Then

   MsgBox "Debe ingresar Servicio...", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub

End If

isel = False
For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
    
       isel = True
       Exit For
    
    End If

Next i

If Not isel Then

   MsgBox "Debe haber selecionado al menos un datos de la grilla", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub

End If

fg_carga ""
Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<ParPanLed>"

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" Then
    
       vaSpread1.Col = 2
       Estructura = vaSpread1.text
    
       MyBuffer = MyBuffer & " <PanLed"
       MyBuffer = MyBuffer & " Est = " & Chr(34) & Estructura & Chr(34)
       MyBuffer = MyBuffer & "/>"

    End If

Next i

MyBuffer = MyBuffer & "</ParPanLed>"

Gl_Ac_Botones Me, 1, 19, modo
Set RS = vg_db.Execute("sgpadm_DelIns_XmlHomologacionEstructuraServicioPanLed '" & MyBuffer & "', '" & Ceco & "', " & regimen & ", " & Servicio & "")
If Not RS.EOF Then
            
   fg_descarga
   If RS(0) > 0 Then
                  
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), CStr(Me.HelpContextID), "", "", "")
      MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
      
   Else
   
              
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), CStr(Me.HelpContextID), "", "", "")
    MsgBox "Proceso Termino Correctamente ", vbInformation + vbOKOnly, MsgTitulo
               
   End If
            
End If
RS.Close
Set RS = Nothing
Frame1.Enabled = True
fg_descarga

Exit Sub
Man_Error:
Frame1.Enabled = True
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub EliminarGrilla()

On Error GoTo Man_Error

Dim RS       As New ADODB.Recordset
Dim Ceco     As String
Dim regimen  As Long
Dim Servicio As Long

fg_carga ("")
Ceco = fpText.text
regimen = fpLongInteger1(0).Value
Servicio = fpLongInteger1(1).Value

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), CStr(Me.HelpContextID), "", "", "")

Set RS = vg_db.Execute("sgpadm_Del_HomologacionEstructuraServicioPanLed '" & Ceco & "', " & regimen & ", " & Servicio & "")
If Not RS.EOF Then
            
   fg_descarga
   If RS(0) > 0 Then
                  
      'registrar Log sistema error Eliminacion
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
      MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
      
   Else
   
      'registrar Log sistema Eliminar
      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")
      Call MoverGrilla(fpText.text)
      MsgBox "Proceso Eliminaci n Termino Correctamente ", vbInformation + vbOKOnly, MsgTitulo
               
   End If
            
End If
RS.Close
Set RS = Nothing
Frame1.Enabled = True
fg_descarga

Exit Sub
Man_Error:
Frame1.Enabled = True
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1

    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1.MaxRows 'BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Frame1.Enabled = False

Exit Sub
Man_Error:
    Frame1.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Col = 2 Or Col = 3 Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Frame1.Enabled = False

Exit Sub
Man_Error:
    Frame1.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
