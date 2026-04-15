VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form V_Acceso 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control Acceso SGPADM"
   ClientHeight    =   5970
   ClientLeft      =   2640
   ClientTop       =   1890
   ClientWidth     =   7470
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5970
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Debe ingresar nueva contraseña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1920
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   3960
      Width           =   7245
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   2
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   3
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   615
         Width           =   1395
      End
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   2
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2990
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
         NoSpecialKeys   =   0
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
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   "*"
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
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   3
         Left            =   1800
         TabIndex        =   5
         Top             =   480
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2990
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
         NoSpecialKeys   =   0
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
         MaxLength       =   20
         MultiLine       =   0   'False
         PasswordChar    =   "*"
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
      Begin VB.Label Label3 
         Caption         =   "La contraseña debe tener largo minimo 8 caracteres e incluir: Mayúsculas, minusculas, números y caracteres especiales"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1320
         Width           =   7095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Contraseña"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   840
      Top             =   2880
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Index           =   1
      Left            =   4260
      TabIndex        =   12
      Top             =   3000
      Width           =   3105
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
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
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
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
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2745
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   120
      Width           =   7335
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   1320
         Picture         =   "V_Acceso.frx":0000
         ScaleHeight     =   1215
         ScaleWidth      =   3135
         TabIndex        =   18
         Top             =   120
         Width           =   3135
      End
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   1800
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2990
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
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   "*"
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
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   1440
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2990
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
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
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
         MaxLength       =   20
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5895
         _Version        =   196608
         _ExtentX        =   10398
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
         ButtonStyle     =   2
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
         NoSpecialKeys   =   1
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
         ControlType     =   3
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Trabajo"
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
         Left            =   90
         TabIndex        =   13
         Top             =   2235
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña"
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
         Left            =   90
         TabIndex        =   11
         Top             =   1860
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
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
         Left            =   90
         TabIndex        =   10
         Top             =   1485
         Width           =   660
      End
   End
End
Attribute VB_Name = "V_Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim cPer As Long
Dim RS1  As New ADODB.Recordset

Select Case Index

Case 0
    
    vg_NUsr = LimpiaDato(Trim(LCase(Nombre(0).text)))
    vg_Pass = LimpiaDato(Trim(LCase(Nombre(1).text)))
    
    AbrirBase
    ActVersion
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & vg_NUsr & "'")
    If RS1.EOF Then
       
       MsgBox "Usuario no existe..."
       RS1.Close
       Set RS1 = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    If IsNull(RS1!usu_activo) Or RS1!usu_activo = 0 Then
    
        MsgBox "Usuario esta bloqueado, favor contactese con el Administrador del Sistema..."
    
       RS1.Close
       Set RS1 = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    If Not RS1.EOF And UCase(vg_Pass) <> UCase(fg_Desencripta(TipoDato(RS1!usu_password, ""))) Then
    
       MsgBox "La clave no corresponde al usuario..."
       RS1.Close
       Set RS1 = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
       
    End If
    
    If IsNull(RS1!usu_indppr) Or RS1!usu_indppr = 0 Then
        
        MsgBox "Estimado usuario por favor contactese con el Administrador del Sistema..."
        RS1.Close
        Set RS1 = Nothing
        Nombre(0).text = ""
        Nombre(1).text = ""
        Nombre(0).SetFocus
        
        Exit Sub
    
    Else
        
        vg_Indppr = RS1!usu_indppr
    
        If Trim(fpText1.text) = "" Then
    
           fpText1.text = IIf(IsNull(RS1!DirectorioTrabajo), "", RS1!DirectorioTrabajo)
           dir_trabajo = IIf(IsNull(RS1!DirectorioTrabajo), "", RS1!DirectorioTrabajo)
           
           If Trim(fpText1.text) = "" Then
           
              RS1.Close
              Set RS1 = Nothing
              
              MsgBox "Debe seleccionar la ruta de trabajo..."
              Exit Sub
       
           End If
           
         End If
        
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
'    If Not fg_ValidarUnidadDisco(Mid(dir_trabajo, 1, 3)) Then
'
'       MsgBox "Ruta trabajo no valida..."
'       Exit Sub
'
'    End If
    
    If Not fg_ValidarDirectorio(dir_trabajo) Then
    
       MsgBox "Ruta trabajo no valida..."
       Exit Sub
    
    
    End If
    
    '-------> actualizar ruta usuario
    vg_db.Execute ("sgpadm_Upd_UsuarioRutaTrabajo '" & vg_NUsr & "', '" & dir_trabajo & "'")
    
    '-------> validar usuario perfil
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_UsuarioPerfil '" & vg_NUsr & "'")
    
    If RS1.EOF Then
       
       MsgBox "Usuario no tiene asignado un perfil..."
       RS1.Close
       Set RS1 = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
       
    End If
    RS1.Close
    Set RS1 = Nothing
    
    vg_DPr = 0: vg_pais = ""
    vg_DCa = GetParametro("deccan")
    vg_pais = GetParametro("parpais")
    vg_PartePlani = False
    
    '-------> Nombre empresa
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS1 = vg_db.Execute("sgpadm_Sel_Parametros 'nomempresa'")
    
    If RS1.EOF Then
       
       vg_db.Execute "sgpadm_iu_param 'A', 'nomempresa', 'Parametro Nombre Empresa', 'C', 'Sodexo Chile S.A.'"
    
    End If
    RS1.Close
    Set RS1 = Nothing
    
    If Right(dir_trabajo, 1) <> "\" Then
    
       dir_trabajo = dir_trabajo & "\"
    
    End If
    
    If Not RevisaPassword Then Exit Sub
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ingreso_Correcto"), "SGPADM", "", "", "")
    
    
Case 1
    
    End

Case 2

    'par_pslong, par_psplaz, par_pscara, par_psante, par_psfall
    If Nombre(2).text = "" Or Nombre(3).text = "" Then
        
        MsgBox "Debe ingresar password...", vbCritical + vbOKOnly, "Ingreso al sistema"
        SendKeys "+{Tab}": SendKeys "+{Tab}"
        Exit Sub
    
    End If
    
    If Nombre(2).text = Nombre(3).text Then
        
        If Not fg_ValidaPassword(vg_NUsr, Trim(Nombre(2).text), MsgTitulo) Then
            
            Nombre(2).text = ""
            Nombre(3).text = ""
            SendKeys "+{Tab}": SendKeys "+{Tab}"
            Exit Sub
        
        End If
    
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("sgpadm_Upd_UsuarioContrasena '" & vg_NUsr & "', '" & fg_Encripta(Trim(Nombre(2).text)) & "'")
        If RS1.EOF Then
       
           If RS1(0) > 0 Then
          
              Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), RS1(0) & RS1(1), "", "")
              RS1.Close
              Set RS1 = Nothing
              MsgBox RS1(0) & " " & RS1(1), vbCritical + vbOKOnly, MsgTitulo
              Exit Sub
       
           End If
        
        End If
        RS1.Close
        Set RS1 = Nothing
        
        vg_Pass = Trim(Nombre(2).text)
        
        'INSERTA MODIFICACIÓN DE PASSWORD
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_CambiaPass"), "SGPADM", fg_Encripta(Trim(Nombre(2).text)), fg_Encripta(Trim(Nombre(1).text)), "")
        MsgBox "Su password fué cambiada...", vbInformation + vbOKOnly, "Ingreso al sistema"
        
        'INGRESO CORRECTO
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ingreso_Correcto"), "SGPADM", "", "", "")
        
        fg_carga ""
        fg_descarga
        Me.Hide
        Unload Me
    
    Else
        
        MsgBox "Las password ingresadas deben ser iguales...", vbCritical + vbOKOnly, "Ingreso al sistema"
        SendKeys "+{Tab}": SendKeys "+{Tab}"
        Exit Sub
    
    End If

Case 3

    Frame1(0).Enabled = True
    Frame1(1).Enabled = True
    Me.Height = 4180
    Nombre(2).text = ""
    Nombre(3).text = ""
    Frame1(2).Enabled = True
    SendKeys "+{Tab}"
    Exit Sub

End Select

Me.Hide
Unload Me

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Man_Error

'Command1(Index).Style = 1
Command1(0).BackColor = &H8000000F
Command1(1).BackColor = &H8000000F
Command1(Index).BackColor = &HFFC0C0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

Nombre(0).SetFocus
'Command1_Click 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 4180
Me.Width = 7560
fg_centra Me
Nombre(1).PasswordChar = "*"
Me.Caption = "Control Acceso - SGP ADM Chile v" & Trim(Str(App.Major)) & "." & Trim((App.Minor)) & "." & Trim(((App.Revision))) & " - " & "Servidor " & vg_SqlNSvr & " - " & "BBDD " & vg_SqlBase

'Nombre(0).Text = "ADM"
'Nombre(1).Text = "adm"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

On Error GoTo Man_Error

'M_ExpDir.Show 1, Me
'
'If vg_dir <> "" Then
'
'   fpText1.text = vg_dir
'   dir_trabajo = vg_dir
'
'End If

    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")

    If Trim(ret) <> "" Then
    
       fpText1.text = ret '& "\"
       
       fpText1.text = ret 'vg_dir
       dir_trabajo = ret 'vg_dir
   
       'Dir1_Change
       
    End If

Nombre(1).SetFocus

   
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Nombre_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii = 13 Then SendKeys "{TAB}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Nombre_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Index

    Case 0
        
        If Trim(fpText1.text) = "" Then
           
           Nombre(Index).text = UCase(Nombre(Index).text)

           AbrirBase

           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           Set RS = vg_db.Execute("sgpadm_Sel_Usuaurio_V01 '" & Nombre(Index).text & "'")
           If Not RS.EOF Then

              fpText1.text = IIf(IsNull(RS!DirectorioTrabajo), "", RS!DirectorioTrabajo)
              dir_trabajo = IIf(IsNull(RS!DirectorioTrabajo), "", RS!DirectorioTrabajo)

           End If
           RS.Close
           Set RS = Nothing

           vg_db.Close
           
        End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Function RevisaPassword() As Boolean

On Error GoTo Man_Error

Dim cPassPlazo As Long
Dim cDiasClave As Long
Dim RS         As New ADODB.Recordset

RevisaPassword = True

'REVISA LA FECHA DE LA ULTIMA PASSWORD
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Log_CambiaPass_V01 1, '" & vg_NUsr & "', " & fg_TraeLogConcepto("vg_logsis_CambiaPass") & "")
If Not RS.EOF Then

    cPassPlazo = GetParametro("plazpass")
'    cDiasClave = DateDiff("d", CDate(TipoDato(RS!Fecha, 0)), Date)
    cDiasClave = DateDiff("d", CDate(TipoDato(RS!Fecha, 0)), RS!fechasistema)
    
    If cDiasClave > cPassPlazo Then
        
        RevisaPassword = False
    
    End If

Else
    
    RevisaPassword = False

End If
RS.Close
Set RS = Nothing

'CONSULTA SI EL USUARIO INGRESO POR PRIMERA VEZ
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Log_CambiaPass_V01 2, '" & vg_NUsr & "', " & fg_TraeLogConcepto("vg_logsis_CambiaPass") & "")
If Not RS.EOF Then
    
    If TipoDato(RS!cuenta, 0) = 1 Then
        
        RevisaPassword = False
    
    End If

End If
RS.Close
Set RS = Nothing

If Not RevisaPassword Then
    
    Frame1(2).Visible = True
    Me.Height = 6330
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    Nombre(3).SetFocus
'    SendKeys "{TAB}"

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

