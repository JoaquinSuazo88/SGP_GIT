VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form B_CelEdi 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editar Estructura"
   ClientHeight    =   930
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930
   ScaleWidth      =   6555
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   255
      Left            =   5280
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin EditLib.fpText fpText 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   4905
      _Version        =   196608
      _ExtentX        =   8652
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre Estructura"
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "B_CelEdi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***--- Creado  :24-12-09
'***--- Autor   :Samuel Melendez
'***--- Este formulario aun quese puede llamar directamente
'***--- es recomendable llamarlo mas bien desde
'***--- la funcion "CellEdite" desde la clase "G_Proc"
'***--- la cual llama a este formulario y devuelve el texto ingresado

Private Sub Command1_Click()

    If Trim(fpText.text) = "" Then Unload Me: Exit Sub
    G_Proc.Txt = fpText.text
    Unload Me

End Sub

Private Sub Form_Activate()
    
    fpText.SelStart = 0
    fpText.SelLength = Len(fpText)
    fpText.SetFocus

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        
        Case 13
            
            Command1_Click
        
        Case 27
            
            G_Proc.Txt = "-1e"
            Unload Me
    
    End Select

End Sub




