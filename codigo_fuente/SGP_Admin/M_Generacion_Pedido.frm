VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form M_Generacion_Pedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Filtro Generación Pedido"
   ClientHeight    =   9720
   ClientLeft      =   4155
   ClientTop       =   375
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9720
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   8745
      Left            =   120
      TabIndex        =   3
      Top             =   900
      Width           =   6285
      Begin VB.CheckBox CkArrastreSaldo 
         Caption         =   "Genera Arrastre de Saldo"
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
         Left            =   4320
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   360
         TabIndex        =   22
         Top             =   8340
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   19
         Top             =   7680
         Width           =   1260
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Procesar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3360
         TabIndex        =   17
         Top             =   7680
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   885
         Left            =   255
         TabIndex        =   9
         Top             =   6615
         Width           =   5865
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   1320
            TabIndex        =   25
            Top             =   960
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm"
            Format          =   62980099
            UpDown          =   -1  'True
            CurrentDate     =   41698.9993055556
            MinDate         =   40179
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1320
            TabIndex        =   10
            Top             =   345
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
            Left            =   3855
            TabIndex        =   11
            Top             =   375
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
         Begin VB.Label Label2 
            Caption         =   "Fecha Limite Confirmación"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Left            =   240
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   13
            Top             =   405
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta"
            Height          =   195
            Index           =   0
            Left            =   2745
            TabIndex        =   12
            Top             =   420
            Width           =   915
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Productos PAP"
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
         Index           =   1
         Left            =   285
         TabIndex        =   7
         Top             =   3645
         Value           =   1  'Checked
         Width           =   5730
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Incluir Productos CD"
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
         Index           =   0
         Left            =   285
         TabIndex        =   6
         Top             =   3285
         Width           =   5895
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1935
         Index           =   1
         Left            =   165
         TabIndex        =   15
         Top             =   1290
         Width           =   5955
         _Version        =   393216
         _ExtentX        =   10504
         _ExtentY        =   3413
         _StockProps     =   64
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
         MaxRows         =   0
         ScrollBars      =   2
         SpreadDesigner  =   "M_Generacion_Pedido.frx":0000
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1935
         Index           =   2
         Left            =   165
         TabIndex        =   16
         Top             =   4365
         Width           =   5955
         _Version        =   393216
         _ExtentX        =   10504
         _ExtentY        =   3413
         _StockProps     =   64
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
         MaxRows         =   0
         ScrollBars      =   2
         SpreadDesigner  =   "M_Generacion_Pedido.frx":02F6
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   210
         TabIndex        =   18
         Top             =   555
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
         AutoCase        =   1
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
         MaxLength       =   4
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
      Begin VB.Label lbl_proceso 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         TabIndex        =   23
         Top             =   9135
         Width           =   2070
      End
      Begin VB.Label lblcamchk 
         Height          =   165
         Left            =   2310
         TabIndex        =   21
         Top             =   1320
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rango de Fechas"
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
         Left            =   300
         TabIndex        =   14
         Top             =   6390
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         Index           =   4
         Left            =   255
         TabIndex        =   8
         Top             =   4125
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Familia Productos:"
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
         Left            =   195
         TabIndex        =   5
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras:"
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
         Left            =   180
         TabIndex        =   4
         Top             =   300
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   150
      TabIndex        =   0
      Top             =   45
      Width           =   6300
      Begin VB.OptionButton Option1 
         Caption         =   "Pedido PAP"
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
         Index           =   2
         Left            =   120
         TabIndex        =   20
         Top             =   270
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Proyectado"
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
         Index           =   1
         Left            =   4560
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pedido CD"
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
         Index           =   0
         Left            =   2520
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "M_Generacion_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub llena_grilla_familia()

On Error GoTo Man_Error

'´Rescata las FAmilia de Productos
Dim RS  As New ADODB.Recordset
Dim Sql As String

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_BuscaFamilia_V02 "
Set RS = vg_db.Execute(Sql)
    
'-------> Inicio LLenar grilla
   
vaSpread1(1).MaxRows = 0
 
Do While Not RS.EOF
    
   vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
   vaSpread1(1).Row = vaSpread1(1).MaxRows
        
   vaSpread1(1).Col = 1 ' Seleccion
   vaSpread1(1).text = 0
        
   vaSpread1(1).Col = 2 ' Codigo
   vaSpread1(1).text = RS(0)
        
   vaSpread1(1).Col = 3 ' Familia
   vaSpread1(1).text = RS(1)
        
        
   RS.MoveNext
   
Loop
RS.Close
Set RS = Nothing
  
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub llena_grilla_Centro_costo()

On Error GoTo Man_Error

Dim Sql As String
Dim RS  As New ADODB.Recordset

If Option1(1).Value = 0 Then
    
    vaSpread1(2).Enabled = False

Else
    
    vaSpread1(2).Enabled = True

End If

 Dim OrgCompras As String
 
 OrgCompras = fpText(1)
 
 ' Permite seleccionar klos Celos
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
 
 
Sql = " sgpadm_Sel_BuscaCecos "
Sql = Sql & " '" & OrgCompras & "'"
Set RS = vg_db.Execute(Sql)
    
'-------> Inicio LLenar grilla
vaSpread1(2).MaxRows = 0
 
Do While Not RS.EOF
    
   vaSpread1(2).MaxRows = vaSpread1(2).MaxRows + 1
   vaSpread1(2).Row = vaSpread1(2).MaxRows
        
   vaSpread1(2).Col = 1 ' Seleccion
   vaSpread1(2).text = 1
        
   vaSpread1(2).Col = 2 ' Codigo
   vaSpread1(2).text = RS(0) 'Val(RS(0))
        
   vaSpread1(2).Col = 3 ' Descripcion
   vaSpread1(2).text = RS(1)
        
   RS.MoveNext
   
Loop
RS.Close
Set RS = Nothing
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS           As New ADODB.Recordset
Dim fechalimite  As String
Dim fechasistema As String
Dim fechadesde   As String
Dim fechahasta   As String
Dim Sql          As String
Dim tipopedido   As Integer

If vaSpread1(2).MaxRows < 1 Then

    MsgBox "No existe cecos asociado a la Org. Compras, proceso cancelado...", vbExclamation
    Exit Sub

End If

fechalimite = Format(DTPicker1, "DD/MM/YYYY")
fechadesde = Format(fpDateTime1(0), "DD/MM/YYYY")
fechahasta = Format(fpDateTime1(1), "DD/MM/YYYY")

fechasistema = Date

'Tipo pedido
If Option1(0).Value = True Then
   
   tipopedido = 3

ElseIf Option1(2).Value = True Then
   
   tipopedido = 1

ElseIf Option1(1).Value = True Then
   
   tipopedido = 2

End If
'-------> validar fecha desde corresponda día lunes
If DatePart("w", fechadesde, 2) <> 1 And (tipopedido = 1 Or tipopedido = 3) Then '=1 lunes
    
    MsgBox "Fecha desde debe corresponder día [lunes]...", vbExclamation
    Exit Sub

End If

'-------> validar fecha hasta corresponda día domingo
If DatePart("w", fechahasta, 2) <> 7 And (tipopedido = 1 Or tipopedido = 3) Then '=7 Domingo
    
    MsgBox "Fecha hasta debe corresponde día [domingo]...", vbExclamation
    Exit Sub

End If

If Format(fechahasta, ("YYYYMMDD")) < Format(fechadesde, ("YYYYMMDD")) Then
    
    MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation
    Exit Sub

End If

'If Option1(1) = False Then

'If Format(fechalimite, ("YYYYMMDD")) >= Format(fechasistema, "YYYYMMDD") Then
'  If Format(fechalimite, ("YYYYMMDD")) < Format(FechaDesde, "YYYYMMDD") Then
'
' Else
'    MsgBox "La fecha de confirmación no puede ser mayor o igual a la fecha desde...", vbExclamation
'    Exit Sub
'  End If
'Else
'    MsgBox "La fecha de confirmación no puede ser menor a la fecha actual...", vbExclamation
'    Exit Sub
'End If
'End If


If Len(fpText(1)) < 4 Then
  
  MsgBox " Se debe ingresar un centro logistico", vbExclamation
  Exit Sub

End If

Dim seleccion As String

' Valida que haya una familia por lo menos cuando es proyectado


Dim Conta As Long
Dim i     As Long
Conta = 0
 
If Option1(1).Value = True Then
  
  For i = 1 To vaSpread1(1).MaxRows
    
    vaSpread1(1).Row = i
    vaSpread1(1).Col = 1 'Seleccion
    seleccion = IIf(vaSpread1(1).text = "", 0, vaSpread1(1).text)
    
    If seleccion = 1 Then
       
       Conta = Conta + 1
    
    End If
  
  Next i

  If Conta = 0 Then
     
     MsgBox " Se debe seleccionar una Familia por lo menos", vbExclamation
     Exit Sub
  
  End If

End If
      
'-------> Valida los Productos
If Check1(0).Value = False And Check1(1).Value = False Then
    
    MsgBox " Debe incluir al menos un tipo de Producto", vbExclamation
    Exit Sub

End If

'-------> Procesa la Generacion de Pedido siendo Proyectado o Pedido
Screen.MousePointer = 11

If Option1(1).Value = 0 Then
    
    '-------> Validar que fecha desde y hasta no hallan pedidos
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = ""
    Sql = Sql & "'" & LimpiaDato(Trim(fpText(1).text)) & "'"
    Sql = Sql & ",'" & Format(fechadesde, "YYYYMMDD") & "'"
    Sql = Sql & ",'" & Format(fechahasta, "YYYYMMDD") & "'"
    Sql = Sql & ",'" & tipopedido & "'"
    Set RS = vg_db.Execute("sgpadm_Sel_ValidarPedidos  " & Sql)
    
    If Not RS.EOF Then
       
       Screen.MousePointer = 1
       RS.Close
       Set RS = Nothing
       MsgBox " Existe pedido, para periodo indicado", vbExclamation
       Exit Sub
    
    End If
    RS.Close
    Set RS = Nothing
    
   Call Genera_Envio_pedido

Else
   
   Call Genera_Envio_pedido_Proyectado

End If
Screen.MousePointer = 1

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Genera_Envio_pedido_Proyectado()

On Error GoTo Man_Error
  Dim RS           As New ADODB.Recordset
  Dim RS1          As New ADODB.Recordset
  Dim Sql          As String
  Dim Prod_CD      As Integer
  Dim Prod_PAP     As Integer
  Dim xmlfamilia   As String
  Dim seleccion    As Integer
  Dim centrocosto  As String
  Dim idpedido     As Long
  Dim FechaInicial As String
  Dim FechaFinal   As String
  Dim tipopedido   As Integer
  Dim codfamilia   As String
  Dim OrgCompra    As String
  Dim cod          As String
  Dim Conta        As Integer
  Dim i            As Integer
  Dim strincecos   As String
  Dim Ceco         As String
  Dim FechaIni     As Date
  Dim FechaFin     As Date
  Dim IndFecha     As Date
  Dim FechaAux     As Date
  
  For i = 1 To vaSpread1(2).MaxRows
    
    vaSpread1(2).Row = i
    vaSpread1(2).Col = 1 'Seleccion
    seleccion = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
    
    If seleccion = 1 Then
       
       Conta = Conta + 1
    
    End If
    
  Next i
  
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = Conta
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0
  
  ' Rescata los parametros a procesafr
  
  FechaInicial = Format(fpDateTime1(0).text, "YYYYMMDD")
  FechaFinal = Format(fpDateTime1(1).text, "YYYYMMDD")
  idpedido = 0
  Prod_CD = Check1(0)
  Prod_PAP = Check1(1)
  OrgCompra = fpText(1)
  
  If Option1(1).Value = True Then
    
    tipopedido = 2
  
  End If
   
  ' Rescata la Familia de Producto Seleccionada
  
   xmlfamilia = ""
   xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
   xmlfamilia = xmlfamilia & "<Fa>"
  
  For i = 1 To vaSpread1(1).MaxRows
    
    vaSpread1(1).Row = i
    vaSpread1(1).Col = 1 'Seleccion
    seleccion = IIf(vaSpread1(1).text = "", 0, vaSpread1(1).text)
    
    If seleccion = 1 Then
    
       xmlfamilia = xmlfamilia & " <F"
       vaSpread1(1).Row = i
       vaSpread1(1).Col = 2 'Id Ruta de Compras
       codfamilia = IIf(vaSpread1(1).text = "", 0, vaSpread1(1).text)
       xmlfamilia = xmlfamilia & " cod = " & Chr(34) & codfamilia & Chr(34)
       xmlfamilia = xmlfamilia & "/>"
  
    End If
  
  Next i
  
  xmlfamilia = xmlfamilia & "</Fa>"
  
  If Len(xmlfamilia) > 20000 Then
     
     MsgBox "Son muchas las familias debe se sacarle algunas "
     Exit Sub
    
  End If
  
  ' Recorre los Cecos para ser gr4abados
  
  For i = 1 To vaSpread1(2).MaxRows
    
    DoEvents
    
    vaSpread1(2).Row = i
    vaSpread1(2).Col = 1 'Seleccion
    seleccion = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
    
    If seleccion = 1 Then
       
       vaSpread1(2).Row = i
       vaSpread1(2).Col = 2 'Id Ruta de Compras
       Ceco = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
       centrocosto = centrocosto & "," & Ceco
   
    End If
    
  Next i
    
  If Len(centrocosto) > 8000 Then
     
     MsgBox "Son muchas las familias debe se sacarle algunas "
     Exit Sub
  
  End If
  
  ' Graba EnCabezado del Pedido proyectado
                        
  Sql = "sgpadm_INS_GrabaEncabezadoPedido_V01 "
  Sql = Sql & " '" & centrocosto & "',"
  Sql = Sql & FechaInicial & ","
  Sql = Sql & FechaFinal & ","
  Sql = Sql & tipopedido & ","
  Sql = Sql & " '" & xmlfamilia & "',"
  Sql = Sql & Prod_CD & ","
  Sql = Sql & Prod_PAP & ","
  Sql = Sql & idpedido & ","
  Sql = Sql & " '" & OrgCompra & "', "
  Sql = Sql & " '" & IIf(CkArrastreSaldo.Value = 1, 1, 0) & "'"
  
  If RS1.State = 1 Then RS1.Close
  RS1.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  Set RS1 = vg_db.Execute(Sql)
   
  ' Rescata el Ultimo Pedido para Generar Detalle del Proyectasdo
  If Not RS1.EOF Then
     
     idpedido = RS1(0)
  
  End If
  RS1.Close: Set RS1 = Nothing
  ' Rescatamos los Cecos Seleccionado
 
  For i = 1 To vaSpread1(2).MaxRows
        
        vaSpread1(2).Row = i
        vaSpread1(2).Col = 1 'Seleccion
        seleccion = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
        
        If seleccion = 1 Then
            
            vaSpread1(2).Col = 2 'C.Costo
            centrocosto = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
                        
            Sql = ""
            Sql = " sgpadm_Sel_GeneracionPedido_FDespacho_V09 "
            Sql = Sql & " '" & centrocosto & "',"
            Sql = Sql & FechaInicial & ","
            Sql = Sql & FechaFinal & ","
            Sql = Sql & tipopedido & ","
            Sql = Sql & " '" & xmlfamilia & "',"
            Sql = Sql & Prod_CD & ","
            Sql = Sql & Prod_PAP & ","
            Sql = Sql & idpedido & ","
            Sql = Sql & " '" & OrgCompra & "',"
            Sql = Sql & " '" & "" & "', "
            Sql = Sql & " '" & IIf(CkArrastreSaldo.Value = 1, 1, 0) & "'"
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            Set RS = vg_db.Execute(Sql)
            
            RS.Close
            Set RS = Nothing
            ProgressBar1.Value = ProgressBar1.Value + 1
            lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
            DoEvents
            
        End If
        
  Next i
                       
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient

  Sql = " sgpadm_iu_VadidaEstadoPedido "
  Sql = Sql & idpedido
  Set RS = vg_db.Execute(Sql)
    
  If RS.State = 1 Then
        
     RS.Close
     Set RS = Nothing
    
  End If
  MsgBox "Se Genero el Pedido OK"
  Unload Me
                       
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub Genera_Envio_pedido()

On Error GoTo Man_Error
  
  Dim RS           As New ADODB.Recordset
  Dim Sql          As String
  Dim Prod_CD      As Integer
  Dim Prod_PAP     As Integer
  Dim xmlfamilia   As String
  Dim seleccion    As Integer
  Dim centrocosto  As String
  Dim idpedido     As Long
  Dim FechaInicial As String
  Dim FechaFinal   As String
  Dim tipopedido   As Integer
  Dim OrgCompra    As String
  Dim fechalimite  As String
  Dim FechaIni     As Date
  Dim FechaFin     As Date
  Dim IndFecha     As Date
  Dim FechaAux     As Date
  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = vaSpread1(2).MaxRows
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0

  ' Rescata los parametros a procesar
  fechalimite = Format(DTPicker1, "YYYYMMDD  HH:MM:SS")
  
  FechaInicial = Format(fpDateTime1(0).text, "YYYYMMDD")
  FechaFinal = Format(fpDateTime1(1).text, "YYYYMMDD")
  idpedido = 0
  Prod_CD = Check1(0)
  Prod_PAP = Check1(1)
  OrgCompra = fpText(1)
  
  If Option1(2).Value = True Then
    
    tipopedido = 1
  
  End If
  
  If Option1(0).Value = True Then
    
    tipopedido = 3
  
  End If
       
  ' Rescata la Familia de Producto Seleccionada
    
  xmlfamilia = ""
  xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  xmlfamilia = xmlfamilia & "<Familias>"
  xmlfamilia = xmlfamilia & " <Familia"
  xmlfamilia = xmlfamilia & " codfamilia  = " & Chr(34) & " " & Chr(34)
  xmlfamilia = xmlfamilia & "/>"
  xmlfamilia = xmlfamilia & "</Familias>"
 
 ' Rescatamos los Cecos Seleccionado
 
 Dim i As Integer
 Dim X As Long
 
 For i = 1 To vaSpread1(2).MaxRows
        
     vaSpread1(2).Row = i
     vaSpread1(2).Col = 1 'Seleccion
     seleccion = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
        
     DoEvents
     
     If seleccion = 1 Then
            
            vaSpread1(2).Col = 2 'C.Costo
            centrocosto = IIf(vaSpread1(2).text = "", 0, vaSpread1(2).text)
                        
            Dim estad As String
            
            FechaIni = Format(fpDateTime1(0).text, "dd/mm/yyyy")
            FechaFin = Format(fpDateTime1(1).text, "dd/mm/yyyy")
'                centrocosto = "20650"
            For IndFecha = FechaIni To FechaFin Step 7
                
                FechaAux = IndFecha
                fechalimite = Format((CDate(Traerfechadia(FechaAux, 1)) - 11) & " " & "23:59:00", "yyyymmdd  HH:MM:SS")

                Sql = ""
                Sql = " sgpadm_Sel_GeneracionPedido_FDespacho_V09 "
                Sql = Sql & " '" & centrocosto & "',"
                Sql = Sql & Format(Traerfechadia(FechaAux, 1), "yyyymmdd") & ","
                Sql = Sql & Format(Traerfechadia(FechaAux, 7), "yyyymmdd") & ","
                Sql = Sql & tipopedido & ","
                Sql = Sql & " '" & xmlfamilia & "',"
                Sql = Sql & Prod_CD & ","
                Sql = Sql & Prod_PAP & ","
                Sql = Sql & idpedido & ","
                Sql = Sql & " '" & OrgCompra & "',"
                Sql = Sql & " '" & fechalimite & "',"
                Sql = Sql & " '1'"
                        
                If RS.State = 1 Then RS.Close
                RS.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS = vg_db.Execute(Sql)
                RS.Close
                Set RS = Nothing
                
            Next IndFecha
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
            DoEvents
     End If
 Next i
   
MsgBox "Se Genero el Pedido OK"
Unload Me
                       
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Command2_Click()

Unload Me

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

MsgTitulo = "Generacion Pedido"
fg_centra Me
vaSpread1(1).Enabled = False
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")

fpDateTime1(0).text = Traerfechadia(Format(Date, "dd/mm/yyyy"), 1)
fpDateTime1(1).text = Traerfechadia(Format(fpDateTime1(0).text, "dd/mm/yyyy"), 7)

Dim Fecha As String
Fecha = Format(Date, "dd/mm/yyyy") + " " + "23:59"
DTPicker1 = Fecha

fpDateTime1(0).Enabled = True
Check1(0).Enabled = False
Check1(1).Enabled = False
ProgressBar1.Visible = False

'-------> permiso a los pedidos PAP y CD
Me.HelpContextID = 1191000
Option1(2).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
Option1(0).Value = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", True, False)
Me.HelpContextID = 1192000
Option1(0).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
Option1(1).Value = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", True, False)

If Option1(2).Enabled = True Then
   
   Option1(2).Value = True

ElseIf Option1(0).Enabled = True Then
   
   Option1(0).Value = True

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub fpText_Change(Index As Integer)
On Error GoTo Man_Error

Call busca_OrgCompra

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub busca_OrgCompra()

On Error GoTo Man_Error

Dim pedido As Integer
Dim RS     As New ADODB.Recordset
Dim Sql    As String

If Option1(0).Value = True Then
  
  pedido = 3

End If

If Option1(2).Value = True Then
  
  pedido = 1

End If

If Option1(0).Value = True Or Option1(2).Value = True Then
    
    fpDateTime1(0).Enabled = True 'False
    Sql = " sgpadm_Sel_leemaximofechaPedido "
    Sql = Sql & " '" & fpText(1) & "',"
    Sql = Sql & pedido
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute(Sql)
    If Not RS.EOF Then
     
       If IsNull(RS(0)) Then
      
          If Len(fpText(1)) = 4 Then
             fpDateTime1(0).Enabled = True
          End If
       
          fpDateTime1(0).text = Traerfechadia(Format(Date, "dd/mm/yyyy"), 1)
          fpDateTime1(1).text = Traerfechadia(Format(fpDateTime1(0).text, "dd/mm/yyyy"), 7) 'Format(Date, "dd/mm/yyyy")
     
       Else
        
           fpDateTime1(0).text = Traerfechadia(Format(RS(0), "dd/mm/yyyy"), 1)
           fpDateTime1(1).text = Traerfechadia(Format(RS(0), "dd/mm/yyyy"), 7)
           
       End If
    End If
    RS.Close
    Set RS = Nothing

Else
   
   fpDateTime1(0).Enabled = True
   Check1(0).Enabled = True
   Check1(1).Enabled = True
   fpDateTime1(0).text = Traerfechadia(Format(Date, "dd/mm/yyyy"), 1)
   fpDateTime1(1).text = Traerfechadia(Format(fpDateTime1(0).text, "dd/mm/yyyy"), 7)

End If
  
If Option1(1).Value = True Then

   Call llena_grilla_familia
   
End If
   
Call llena_grilla_Centro_costo
 
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

If Option1(2).Value = True Then
    
    vaSpread1(1).MaxRows = 0
    vaSpread1(1).Enabled = False
    fpDateTime1(0).Enabled = True 'False
    Label2.Visible = True
    DTPicker1.Visible = True
    Check1(0).Enabled = False
    Check1(1).Enabled = False
    Check1(0).Value = 0
    Check1(1).Value = 1
    CkArrastreSaldo.Visible = False

End If

If Option1(0).Value = True Then
    
    vaSpread1(1).MaxRows = 0
    vaSpread1(1).Enabled = False
    fpDateTime1(0).Enabled = True 'False
    Label2.Visible = True
    DTPicker1.Visible = True
    Check1(0).Enabled = False
    Check1(1).Enabled = False
    Check1(0).Value = 1
    Check1(1).Value = 0
    CkArrastreSaldo.Visible = False

End If

If Option1(1).Value = True Then
    
    vaSpread1(1).Enabled = True
    fpDateTime1(0).Enabled = True
    Check1(0).Enabled = True
    Check1(1).Enabled = True
    Label2.Visible = False
    DTPicker1.Visible = False
    Check1(0).Value = 1
    Check1(1).Value = 1
    CkArrastreSaldo.Visible = True

End If

If Option1(1).Value = 0 Then
    
    vaSpread1(2).Enabled = False

Else
    
    vaSpread1(2).Enabled = True

End If
Call busca_OrgCompra

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_BlockSelected(Index As Integer, ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Dim i As Long
vaSpread1(Index).Row = vaSpread1(Index).ActiveRow
vaSpread1(Index).Col = vaSpread1(Index).ActiveCol
If vaSpread1(Index).Col = 1 Then

For i = BlockRow To BlockRow2
    
    vaSpread1(Index).Row = i
    vaSpread1(Index).Value = IIf(vaSpread1(Index).Value = "1", "0", "1")

Next i

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
