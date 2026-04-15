VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_TransUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "transacciones Usuarios"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame1 
         Height          =   645
         Index           =   1
         Left            =   1695
         TabIndex        =   11
         Top             =   240
         Width           =   6690
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   0
            Left            =   705
            TabIndex        =   12
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
            AutoAdvance     =   0   'False
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
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   1
            Left            =   4170
            TabIndex        =   13
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
            AutoAdvance     =   0   'False
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
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Timer1 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   14
            Top             =   225
            Width           =   1065
            _Version        =   196608
            _ExtentX        =   1879
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
            ButtonStyle     =   1
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
            AutoAdvance     =   0   'False
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
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "hh:nn:ss"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   2
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
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Timer1 
            Height          =   315
            Index           =   1
            Left            =   5505
            TabIndex        =   15
            Top             =   225
            Width           =   1065
            _Version        =   196608
            _ExtentX        =   1879
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
            ButtonStyle     =   1
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
            AutoAdvance     =   0   'False
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
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "hh:nn:ss"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   2
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
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   17
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   1
            Left            =   3675
            TabIndex        =   16
            Top             =   270
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Función del sistema"
         Height          =   915
         Index           =   2
         Left            =   1680
         TabIndex        =   7
         Top             =   1065
         Width           =   6705
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   465
            Width           =   6465
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Uno"
            Height          =   240
            Index           =   2
            Left            =   165
            TabIndex        =   9
            Top             =   210
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   240
            Index           =   3
            Left            =   2850
            TabIndex        =   8
            Top             =   225
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de operación"
         Height          =   915
         Index           =   3
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   6705
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   240
            Index           =   4
            Left            =   2850
            TabIndex        =   6
            Top             =   225
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Uno"
            Height          =   240
            Index           =   5
            Left            =   165
            TabIndex        =   5
            Top             =   210
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   465
            Width           =   6465
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   10335
      Begin VB.Frame Frame4 
         Height          =   435
         Index           =   1
         Left            =   3480
         TabIndex        =   23
         Top             =   3960
         Width           =   3555
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   24
            Top             =   135
            Width           =   3450
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   3960
         Width           =   1770
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   22
            Top             =   135
            Width           =   1665
         End
      End
      Begin VB.CommandButton Boton 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8565
         TabIndex        =   19
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton Boton 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   18
         Top             =   3960
         Width           =   1455
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   9855
         _Version        =   393216
         _ExtentX        =   17383
         _ExtentY        =   5953
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
         MaxCols         =   4
         SpreadDesigner  =   "I_TransUsuario.frx":0000
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00D9D9FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   240
         Top             =   4110
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bloqueado"
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   20
         Top             =   4080
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   1920
         Top             =   4080
         Visible         =   0   'False
         Width           =   300
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "I_TransUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String

Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
        Combo1(3).ListIndex = Combo1(0).ListIndex
        
    Case 3
    
        Combo1(0).ListIndex = Combo1(3).ListIndex
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i  As Long

fg_centra Me
Option1(3).Value = True
Option1(4).Value = True

'Función de usuario
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio")

i = 1
vaSpread1.MaxRows = 0
vaSpread1.MaxRows = RS.RecordCount

If Not RS.EOF Then
    
    Do While Not RS.EOF
        
        DoEvents
        vaSpread1.Row = i
                
        vaSpread1.Col = -1
        vaSpread1.BackColor = IIf(IsNull(RS!usu_activo) Or RS!usu_activo = 0, Shape1(1).FillColor, Shape1(0).FillColor)

        
        vaSpread1.Col = 2
        vaSpread1.text = RS!usu_codigo
        
        vaSpread1.Col = 3
        vaSpread1.TypeHAlign = 0
        vaSpread1.text = Trim(RS!usu_Nombre)
        
        RS.MoveNext
        i = i + 1
    Loop
    
End If

RS.Close
Set RS = Nothing


'Función del sistema
Combo1(1).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_OpSistema")
Do While Not RS.EOF
    
    Combo1(1).AddItem RS!opc_nombre & Space(150) & "(" & fg_pone_rchar(RS!opc_codigo, 14, " ") & ")"
    RS.MoveNext

Loop
RS.Close
Set RS = Nothing

'Tipo de operación
Combo1(2).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_LogConceptos")
Do While Not RS.EOF
    
    Combo1(2).AddItem RS!loc_descripcion & Space(150) & "(" & fg_pone_cero(RS!loc_codigo, 3) & ")"
    RS.MoveNext

Loop
RS.Close
Set RS = Nothing

'ControlAccesoGen Boton, "", "", "", "", "0"

MsgTitulo = Me.Caption

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Boton_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim cDateIni   As String
Dim cDateFin   As String
Dim cUsuario   As String
Dim cFunSis    As String
Dim cTipOpe    As Long
Dim CambiaPass As String
Dim OpMarcado  As Boolean
Dim i          As Long
Dim MyBuffer   As String
Dim Usu        As String

Dim xlApp      As Object
Dim xlWb       As Object
Dim xlWs       As Object

Select Case Index

    Case 0
        
        If Trim(Date1(0).text) = "" Or Trim(Date1(1).text) = "" Or Trim(Timer1(0).text) = "" Or Trim(Timer1(1).text) = "" Then
        
            MsgBox "Debe ingresar período.", vbExclamation + vbOKOnly, MsgTitulo
            Exit Sub
        
        End If
        
        cDateIni = Format(Date1(0).text, "yyyy-mm-dd") & " " & Format(Timer1(0).text, "hh:nn:ss")
        cDateFin = Format(Date1(1).text, "yyyy-mm-dd") & " " & Format(Timer1(1).text, "hh:nn:ss")
        
        If CDate(cDateIni) > CDate(cDateFin) Then
        
           MsgBox "Período de fechas no válida.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
    
        End If
        
        If Option1(2).Value And Combo1(1).ListIndex = -1 Then
           
           MsgBox "Debe seleccionar Función del sistema.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If Option1(5).Value And Combo1(2).ListIndex = -1 Then
        
           MsgBox "Debe seleccionar Tipo de operación.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        OpMarcado = False
        
        For i = 1 To vaSpread1.MaxRows
        
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" Then
            
               OpMarcado = True
            
            End If
        
        Next i
        
        If Not OpMarcado Then
        
           MsgBox "Debe seleccionar al menos un usuario.", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<Usuario>"
    
        For i = 1 To vaSpread1.MaxRows
        
            vaSpread1.Row = i
            vaSpread1.Col = 1
        
            If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
           
               vaSpread1.Col = 2
               Usu = vaSpread1.text
           
               Let MyBuffer = MyBuffer & " <Usu"
               Let MyBuffer = MyBuffer & " Usu = " & Chr(34) & Usu & Chr(34)
               Let MyBuffer = MyBuffer & "/>"
           
            End If
    
        Next i
    
        Let MyBuffer = MyBuffer & "</Usuario>"
        
        fg_carga ""
    
        cFunSis = Trim(fg_codigocbo(Combo1, 1, 14, ""))
        cTipOpe = fg_codigocbo(Combo1, 2, 3, 0)
   
        Set RS = vg_db.Execute("sgpadm_Sel_XmlTransccionesUsuarios '" & MyBuffer & "', '" & cDateIni & "', '" & cDateFin & "', '" & IIf(cFunSis = "0", "", cFunSis) & "', " & cTipOpe & "")
        If Not RS.EOF Then
            
           If RS.RecordCount > 1020000 Then
      
              RS.Close
              Set RS = Nothing
              MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos datos", vbCritical
              Exit Sub
   
           End If
           
           'Abrimos el Commondialog con ShowOpen
           CD.DialogTitle = "Seleccione un archivo excel"
           CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
           CD.DefaultExt = "*.xls|*.xlsx"
           CD.FilterIndex = 2
           CD.Flags = cdlOFNFileMustExist
           CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
           CD.FileName = ""
           CD.ShowSave

           'Si seleccionamos un archivo mostramos la ruta
           If CD.FileName <> "" Then

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
    
              xlWb.Close True, CD.FileName

              Dim XL As New excel.Application 'Crea el objeto excel
              XL.Workbooks.Open CD.FileName, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
              XL.Visible = True
              XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
              '-------> Close ADO objects
              RS.Close
              Set RS = Nothing
    
              '-- Cerrar Excel
              xlApp.Quit
              '-------> Release Excel references
              Set xlWs = Nothing
              Set xlWb = Nothing
              Set xlApp = Nothing
  
              fg_descarga
              MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
           Else
              'Si no mostramos un texto de advertencia de que no se seleccionó _
               ninguno, ya que FileName devuelve una cadena vacía
               
               MsgBox "No seleccionó ningún archivo", vbCritical

           End If
        
        
        Else
            
            RS.Close
            Set RS = Nothing
            
            MsgBox "No existen datos para los filtros. ", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        fg_descarga
    
    Case 1
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Unload(Cancel As Integer)

On Error GoTo Man_Error

'Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "")

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
        
        Combo1(0).Enabled = False: Combo1(3).Enabled = False
        Combo1(0).ListIndex = -1: Combo1(3).ListIndex = -1
    
    Case 1
        
        Combo1(0).Enabled = True: Combo1(3).Enabled = True
        If Combo1(0).ListCount > 0 Then Combo1(0).ListIndex = 0
        If Combo1(3).ListCount > 0 Then Combo1(3).ListIndex = 0
    
    Case 3
        
        Combo1(1).Enabled = False
        Combo1(1).ListIndex = -1
    
    Case 2
        
        Combo1(1).Enabled = True
        If Combo1(1).ListCount > 0 Then Combo1(1).ListIndex = 0
        
    Case 4
        
        Combo1(2).Enabled = False
        Combo1(2).ListIndex = -1
    
    Case 5
        
        Combo1(2).Enabled = True
        If Combo1(2).ListCount > 0 Then Combo1(2).ListIndex = 0
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 2 Then
   
   Text1(3).text = ""

ElseIf Index = 3 Then
   
   Text1(2).text = ""

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 4
    vaSpread1.text = 0

Next

Select Case Index

Case 2, 3
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 2
           
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
              vaSpread1.Col = 2
              
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
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 4
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Timer1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Dim i As Long

Select Case BlockCol

Case 1
    

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
