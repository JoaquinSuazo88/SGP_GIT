VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form B_BusVas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Recetas en Planificación"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Criterio"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      Begin EditLib.fpText Text1 
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   480
         Width           =   1980
         _Version        =   196608
         _ExtentX        =   3492
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
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
         MaxLength       =   50
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar &Siguiente"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "B_BusVas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srow As Long
Dim ret As Integer
Dim form1 As Form
Dim est As Boolean
Dim text As String
Private Sub Command1_Click()
ret = 0
With form1.vaSpread1
     For srow = 1 To .MaxRows - 1
         ret = .SearchRow(srow, 0, .MaxCols - 1, text, 2)
         If ret > -1 Then
            .SetActiveCell ret, srow
            srow1 = srow
            est = True
            Command1.Enabled = False
            Command3.Enabled = True
            Command3.SetFocus
            Exit Sub
'       Form1.mnusearchnext.Enabled = True
         End If
     Next srow
     If ret = -1 Then MsgBox "Texto no fue encontrado.": Exit Sub
End With
End Sub

Private Sub Command2_Click()
If ret < 0 Or ret = 0 Then
'Form1.mnusearchnext.Enabled = False
End If
Unload Me
End Sub

Private Sub Command3_Click()
Dim ret2 As Integer
ret2 = 0: est = False
With form1.vaSpread1
     For srow1 = srow To .MaxRows - 1
         If ret > -1 Then
            ret2 = .SearchRow(srow1, ret, .MaxCols - 1, text, 2)
            If ret2 > -1 Then
               .SetActiveCell ret2, srow1
               srow = srow1
               ret = ret2
               ret2 = -1
               Exit Sub
            Else
               ret = 0
            End If
         ElseIf srow1 <> (.MaxRows - 1) Then
            ret = 1
         ElseIf srow1 = (.MaxRows - 1) Then
            ret = ret2
            ret2 = -1
         End If
     Next srow1
     Command1.Enabled = True: Command3.Enabled = False: srow = 1: ret = 0: Command1.SetFocus
End With
End Sub

Private Sub Form_Load()
fg_centra Me
est = True
Text1.text = ""
text = ""
Command3.Enabled = False
est = False
End Sub

Private Sub Text1_Change()
'If est Then Exit Sub
text = Text1.text
Command1.Enabled = True: Command3.Enabled = False: srow = 1: ret = 0
End Sub

Sub Partidas(Form As Form)
Set form1 = Form
est = False
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
text = Text1.text
End Sub
