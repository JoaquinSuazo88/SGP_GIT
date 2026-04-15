VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form M_ViuMSR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizar Estado Minuta Sitio Remoto"
   ClientHeight    =   5235
   ClientLeft      =   5520
   ClientTop       =   4395
   ClientWidth     =   9180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   585
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   8775
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   315
         Left            =   2670
         TabIndex        =   5
         Top             =   180
         Width           =   1300
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   855
         TabIndex        =   3
         Top             =   180
         Width           =   1300
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
         ButtonStyle     =   2
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
         Text            =   "08/10/2010"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   225
         Width           =   450
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4485
      Left            =   30
      TabIndex        =   0
      Top             =   630
      Width           =   8775
      Begin FPSpread.vaSpread SprLog 
         Height          =   4005
         Left            =   150
         TabIndex        =   1
         Top             =   270
         Width           =   8355
         _Version        =   393216
         _ExtentX        =   14737
         _ExtentY        =   7064
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
         MaxRows         =   1
         SpreadDesigner  =   "M_ViuMSR.frx":0000
      End
   End
End
Attribute VB_Name = "M_ViuMSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private RS As New Recordset

Private Sub CmdBuscar_Click()
    Set RS = vg_db.Execute("sgpadm_s_TraeLogSitRem " & fpDateTime1.text)
    Let SprLog.MaxRows = 0
    Do While Not RS.EOF
        SprLog.MaxRows = SprLog.MaxRows + 1
        Call SprLog.SetText(1, SprLog.MaxRows, RS!cencos)
        Call SprLog.SetText(2, SprLog.MaxRows, RS!cli_nombre)
        Call SprLog.SetText(3, SprLog.MaxRows, RS!fecpro)
        Call SprLog.SetText(4, SprLog.MaxRows, RS!mensaje)
        RS.MoveNext
    Loop
        
    
End Sub

Private Sub Form_Load()
    Set RS = vg_db.Execute("sgpadm_s_TraeLogSitRem '" & Mid(Now, 1, 10) & "'")
    Let SprLog.MaxRows = 0
    Do While Not RS.EOF
        SprLog.MaxRows = SprLog.MaxRows + 1
        Call SprLog.SetText(1, SprLog.MaxRows, RS!cencos)
        Call SprLog.SetText(2, SprLog.MaxRows, RS!cli_nombre)
        Call SprLog.SetText(3, SprLog.MaxRows, RS!fecpro)
        Call SprLog.SetText(4, SprLog.MaxRows, RS!mensaje)
        RS.MoveNext
    Loop
End Sub
