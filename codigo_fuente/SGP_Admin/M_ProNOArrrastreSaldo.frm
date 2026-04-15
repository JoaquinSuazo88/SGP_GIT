VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_ProNOArrrastreSaldo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Productos Que NO Arrastran Saldo"
   ClientHeight    =   5685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   1110
      TabIndex        =   13
      Top             =   5445
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmd_agrega_todos_izq 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6375
      TabIndex        =   8
      Top             =   3405
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.CommandButton cmd_Agrega_izq 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6375
      TabIndex        =   7
      Top             =   2910
      Width           =   450
   End
   Begin VB.CommandButton cmd_Agrega_der 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6375
      TabIndex        =   4
      Top             =   2280
      Width           =   450
   End
   Begin VB.CommandButton cmd_agrega_todos_der 
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6375
      TabIndex        =   3
      Top             =   1800
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Productos que Arrastra Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4980
      Left            =   315
      TabIndex        =   0
      Top             =   495
      Width           =   6000
      Begin EditLib.fpText TEXT1 
         Height          =   315
         Index           =   3
         Left            =   1965
         TabIndex        =   11
         Top             =   4545
         Width           =   3555
         _Version        =   196608
         _ExtentX        =   6271
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
         ButtonStyle     =   0
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
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4110
         Left            =   150
         TabIndex        =   5
         Top             =   345
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7250
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
         MaxRows         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "M_ProNOArrrastreSaldo.frx":0000
      End
      Begin EditLib.fpText TEXT1 
         Height          =   315
         Index           =   2
         Left            =   615
         TabIndex        =   10
         Top             =   4545
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Productos que NO Arrastra Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4950
      Left            =   6930
      TabIndex        =   1
      Top             =   495
      Width           =   5925
      Begin EditLib.fpText text2 
         Height          =   315
         Index           =   2
         Left            =   585
         TabIndex        =   9
         Top             =   4515
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
         ButtonStyle     =   0
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
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4065
         Left            =   120
         TabIndex        =   6
         Top             =   345
         Width           =   5685
         _Version        =   393216
         _ExtentX        =   10028
         _ExtentY        =   7170
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
         MaxRows         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "M_ProNOArrrastreSaldo.frx":045E
      End
      Begin EditLib.fpText text2 
         Height          =   315
         Index           =   3
         Left            =   1980
         TabIndex        =   12
         Top             =   4530
         Width           =   3540
         _Version        =   196608
         _ExtentX        =   6244
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
         ButtonStyle     =   0
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
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Label lbl_porcentaje 
      Height          =   180
      Left            =   5100
      TabIndex        =   14
      Top             =   6375
      Width           =   3375
   End
End
Attribute VB_Name = "M_ProNOArrrastreSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmd_Agrega_der_Click()

On Error GoTo Man_Error

Screen.MousePointer = 11
DoEvents
Dim codigo As String
Dim descripcion As String
Dim seleccion As String

' MOVER DE LA SPREAD DEL IZQUERDA A LA DERECHA

For i = 1 To vaSpread1.MaxRows
               vaSpread1.Row = i
               vaSpread1.Col = 1
               seleccion = vaSpread1.text
     If seleccion = "1" Then
               
               vaSpread1.Col = 2
               codigo = vaSpread1.text
               vaSpread1.Col = 3
               descripcion = vaSpread1.text
               
               '-------> Mover datos a vector de casino incluido en la ruta
               
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               
               vaSpread2.Col = 2
               vaSpread2.text = codigo
               vaSpread2.Col = 3
               vaSpread2.text = descripcion
               
    End If
Next i

' ELIMINAR DE LA SPREAD LAS MARCADAS
For i = 1 To vaSpread1.MaxRows
               vaSpread1.Row = i
               vaSpread1.Col = 1
               seleccion = vaSpread1.text
     If seleccion = "1" Then

                vaSpread1.DeleteRows vaSpread1.Row, 1
                vaSpread1.MaxRows = vaSpread1.MaxRows - 1
                i = i - 1
     End If

Next i
Screen.MousePointer = 0
DoEvents
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub cmd_Agrega_izq_Click()

On Error GoTo Man_Error

Screen.MousePointer = 11
DoEvents
Dim codigo As String
Dim descripcion As String
Dim seleccion As String

' MOVER DE LA SPREAD DEL IZQUERDA A LA DERECHA

For i = 1 To vaSpread2.MaxRows
               vaSpread2.Row = i
               vaSpread2.Col = 1
               seleccion = vaSpread2.text
     If seleccion = "1" Then
               
               vaSpread2.Col = 2
               codigo = vaSpread2.text
               vaSpread2.Col = 3
               descripcion = vaSpread2.text
               
               '-------> Mover datos a vector de casino incluido en la ruta
               
               vaSpread1.MaxRows = vaSpread1.MaxRows + 1
               vaSpread1.Row = vaSpread1.MaxRows
               
               vaSpread1.Col = 2
               vaSpread1.text = codigo
               vaSpread1.Col = 3
               vaSpread1.text = descripcion
               
    End If
Next i

' ELIMINAR DE LA SPREAD LAS MARCADAS
For i = 1 To vaSpread2.MaxRows
               vaSpread2.Row = i
               vaSpread2.Col = 1
               seleccion = vaSpread2.text
     If seleccion = "1" Then

                vaSpread2.DeleteRows vaSpread2.Row, 1
                vaSpread2.MaxRows = vaSpread2.MaxRows - 1
                i = i - 1
     End If

Next i
Screen.MousePointer = 0
DoEvents

Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub cmd_agrega_todos_der_Click()

On Error GoTo Man_Error

Screen.MousePointer = 11
DoEvents


Dim codigo As String
Dim descripcion As String


  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = vaSpread1.MaxRows
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0



For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Row = i
               
               vaSpread1.Col = 2
               codigo = vaSpread1.text
               vaSpread1.Col = 3
               descripcion = vaSpread1.text
               
               '-------> Mover datos a vector de casino incluido en la ruta
               
               vaSpread2.MaxRows = vaSpread2.MaxRows + 1
               vaSpread2.Row = vaSpread2.MaxRows
               
               vaSpread2.Col = 2
               vaSpread2.text = codigo
               vaSpread2.Col = 3
               vaSpread2.text = descripcion
               
               ProgressBar1.Value = ProgressBar1.Value + 1
               lbl_porcentaje.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
               DoEvents
               
          Next i
lbl_porcentaje = ""
ProgressBar1.Visible = False
vaSpread1.MaxRows = 0
Screen.MousePointer = 0
DoEvents
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub cmd_agrega_todos_izq_Click()

On Error GoTo Man_Error

Screen.MousePointer = 11
DoEvents
Dim codigo As String
Dim descripcion As String


  ProgressBar1.Scrolling = ccScrollingSmooth
  ProgressBar1.Max = vaSpread2.MaxRows
  ProgressBar1.Visible = True
  ProgressBar1.Value = 0

For i = 1 To vaSpread2.MaxRows
               vaSpread2.Row = i
               
               vaSpread2.Col = 2
               codigo = vaSpread2.text
               vaSpread2.Col = 3
               descripcion = vaSpread2.text
               
               '-------> Mover datos a vector de casino incluido en la ruta
               
               vaSpread1.MaxRows = vaSpread1.MaxRows + 1
               vaSpread1.Row = vaSpread1.MaxRows
               
               vaSpread1.Col = 2
               vaSpread1.text = codigo
               vaSpread1.Col = 3
               vaSpread1.text = descripcion
               
               ProgressBar1.Value = ProgressBar1.Value + 1
               lbl_porcentaje.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
               DoEvents
              
          Next i
lbl_porcentaje = ""
vaSpread2.MaxRows = 0
ProgressBar1.Visible = False
Screen.MousePointer = 0
DoEvents
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
  
  fg_centra Me
  Toolbar1.ImageList = Partida.IL1
  ProgressBar1.Visible = False
  
  Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Grabar Producto SAP "
  Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

  Call carga_productos_sap
  Call carga_productos_con_Arrastre

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub carga_productos_con_Arrastre()

 On Error GoTo Man_Error
    
    Sql = " sgpadm_sel_Producto_con_Arrastre_Saldo"
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
    vaSpread2.MaxRows = 0
 
    Do While Not RS.EOF
    
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 2 ' Codigo SAP
        vaSpread2.text = Val(RS(0))
        
        vaSpread2.Col = 3 ' Descipcion
        vaSpread2.text = RS(1)
        
        RS.MoveNext
    Loop
    
'        vaSpread1.SetActiveCell 1, rowanterior
RS.Close: Set RS = Nothing

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub carga_productos_sap()

 On Error GoTo Man_Error
    
    Sql = " sgpadm_sel_Producto_Sin_Arrastre_Saldo"
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
    vaSpread1.MaxRows = 0
 
    Do While Not RS.EOF
    
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 2 ' Codigo SAP
        vaSpread1.text = Val(RS(0))
        
        vaSpread1.Col = 3 ' Descipcion
        vaSpread1.text = RS(1)
        
        RS.MoveNext
    Loop
'        vaSpread1.SetActiveCell 1, rowanterior

     RS.Close: Set RS = Nothing
     
Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub genera_productos()

Dim Producto As String

On Error GoTo Man_Error

  Screen.MousePointer = 11
  
  If vaSpread2.MaxRows > 0 Then
    ProgressBar1.Scrolling = ccScrollingSmooth
    ProgressBar1.Max = vaSpread2.MaxRows
    ProgressBar1.Visible = True
    ProgressBar1.Value = 0
  End If

' Rescata la Familia de Producto Seleccionada
  
  xmlfamilia = ""
  xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  xmlfamilia = xmlfamilia & "<Productos>"
  
  For i = 1 To vaSpread2.MaxRows
    xmlfamilia = xmlfamilia & " <Producto"
    vaSpread2.Row = i
    vaSpread2.Col = 2 'Id Ruta de Compras
    Producto = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    xmlfamilia = xmlfamilia & " Producto = " & Chr(34) & Producto & Chr(34)
    xmlfamilia = xmlfamilia & "/>"
    ProgressBar1.Value = ProgressBar1.Value + 1
    lbl_porcentaje.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
    
    DoEvents
  Next i
  xmlfamilia = xmlfamilia & "</Productos>"
   
    Sql = " sgpadm_iu_Producto_con_Arrastre_Saldo "
    Sql = Sql & " '" & xmlfamilia & "'"
    Set RS = vg_db.Execute(Sql)
  If Not RS.EOF Then
    If RS(0) >= 0 Then
      MsgBox "La Generacion de Productos con Arrastre Termino Correctamente", vbExclamation
    Else
      MsgBox "La Generacion de Productos termino con Problema " + RS(1), vbExclamation
    End If
  End If
  
  Toolbar1.Buttons(2).Visible = True
  Toolbar1.Buttons(1).Visible = False
  ProgressBar1.Visible = False
  vaSpread1.MaxRows = 0
  vaSpread2.MaxRows = 0
  lbl_porcentaje = ""
  TEXT1(2) = ""
  TEXT1(3) = ""
  text2(2) = ""
  text2(3) = ""
  
  Call carga_productos_sap
  Call carga_productos_con_Arrastre

  Screen.MousePointer = 0
  
Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  Screen.MousePointer = 0

End Sub

Private Sub Text1_Change(Index As Integer)

Select Case Index

Case 2, 3
    
    vaSpread1.Visible = False
    If Trim(TEXT1(Index).text) <> "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(TEXT1(Index).text) & "*"
           vaSpread1.Col = 2
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Else
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           End If
        Next i
        vaSpread1.SetActiveCell 1, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(TEXT1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(TEXT1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(TEXT1(Index).text) = "" Then
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       Next
       vaSpread1.SetActiveCell 1, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TEXT1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell 1, 1
    End If
    vaSpread1.Visible = True

End Select

End Sub

Private Sub text2_Change(Index As Integer)

Select Case Index

Case 2, 3
    vaSpread2.Visible = False
    If Trim(text2(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index
           indactivo = UCase(Trim(vaSpread2.Value)) Like "*" & UCase(text2(Index).text) & "*"
           vaSpread2.Col = 2
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell 1, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread1.ColUserSortIndicator(IIf(Trim(text2(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(TEXT1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(text2(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell 1, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(text2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell 1, 1
    End If
    vaSpread2.Visible = True

End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error
    
    Select Case Button.Index
    
    Case 1 ' Genera los Productos co Arrastre
       Call genera_productos
       
    Case 4 ' Salir del Programa
        Me.Hide
        Unload Me
    End Select
  
Exit Sub
Man_Error:
    fg_descarga
    If Err = 438 Or Err = 70 Then
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    End If
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
