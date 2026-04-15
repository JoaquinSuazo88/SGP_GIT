VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form V_Acceso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Acceso SGP"
   ClientHeight    =   5850
   ClientLeft      =   4125
   ClientTop       =   3225
   ClientWidth     =   7485
   ControlBox      =   0   'False
   Icon            =   "V_Acceso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5850
   ScaleWidth      =   7485
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
      TabIndex        =   18
      Top             =   3840
      Width           =   7245
      Begin VB.CommandButton Command1 
         Caption         =   "Cancelar"
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
         Index           =   3
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   615
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
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
         Index           =   2
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   7
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
         AutoCase        =   3
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
         Index           =   6
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1575
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
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   840
         Width           =   810
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
         TabIndex        =   19
         Top             =   1320
         Width           =   7095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Index           =   1
      Left            =   4260
      TabIndex        =   12
      Top             =   2880
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
         TabIndex        =   3
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2625
      Index           =   0
      Left            =   45
      TabIndex        =   9
      Top             =   120
      Width           =   7335
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   1320
         Picture         =   "V_Acceso.frx":1CCA
         ScaleHeight     =   975
         ScaleWidth      =   3135
         TabIndex        =   22
         Top             =   240
         Width           =   3135
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2040
         Width           =   5805
      End
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   1
         Top             =   1680
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
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   0
         Left            =   1320
         TabIndex        =   0
         Top             =   1320
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1380
         TabIndex        =   17
         Top             =   2100
         Width           =   5790
      End
      Begin VB.Label Label1 
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
         Index           =   4
         Left            =   240
         TabIndex        =   16
         Top             =   2115
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   240
         TabIndex        =   11
         Top             =   1740
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login"
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
         Left            =   240
         TabIndex        =   10
         Top             =   1365
         Width           =   585
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
      _Version        =   393216
      _ExtentX        =   1085
      _ExtentY        =   450
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
      SpreadDesigner  =   "V_Acceso.frx":2761
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      Height          =   195
      Index           =   0
      Left            =   1560
      TabIndex        =   15
      Top             =   2760
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Pocesando : "
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   945
   End
End
Attribute VB_Name = "V_Acceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii = 13 Then SendKeys "{TAB}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim cPer As Long
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

Select Case Index

Case 0
    
    If Not Prametro_Generales Then
    
       Exit Sub
       
    End If
    
    Dim UsuarioAdm As String
    UsuarioAdm = GetParametro_Seguridad("CAdmsgpLoc")
    
    If UCase(vg_NUsr) <> UCase(UsuarioAdm) Then
       
       If Not RevisaPassword Then Exit Sub
    
    End If
        
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
        
        Set RS1 = vg_db.Execute("sgp_Upd_UsuarioContrasena '" & vg_NUsr & "', '" & fg_Encripta(Trim(Nombre(2).text)) & "'")
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
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_CambiaPass"), "SGP", fg_Encripta(Trim(Nombre(2).text)), fg_Encripta(Trim(Nombre(1).text)), "")
        MsgBox "Su password fue cambiada...", vbInformation + vbOKOnly, "Ingreso al sistema"
        
        'INGRESO CORRECTO
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ingreso_Correcto"), "SGP", "", "", "")
        
        fg_carga ""
        fg_descarga
        Me.Hide
        Unload Me
    
    Else
        
        MsgBox "Las password ingresadas deben ser iguales...", vbCritical + vbOKOnly, "Ingreso al sistema"
        SendKeys "+{Tab}": SendKeys "+{Tab}"
        Exit Sub
    
    End If

    If Not Prametro_Generales Then
    
       Exit Sub
       
    End If

Case 3

    Frame1(0).Enabled = True
    Frame1(1).Enabled = True
    Me.Height = 4070
    Nombre(2).text = ""
    Nombre(3).text = ""
    Frame1(2).Enabled = True
    On Error Resume Next
    SendKeys "+{Tab}"
    On Error Resume Next
    Exit Sub

End Select

Me.Hide
Unload Me

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()
'Nombre(0).SetFocus
'Command1_Click 0
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 4070
Me.Width = 7575
fg_centra Me

Me.Caption = "Control Acceso - SGP LOCAL Chile v" & Trim(Str(App.Major)) & "." & Trim((App.Minor)) & "." & Trim(((App.Revision))) & " - " & "BBDD " & vg_SqlBase

Nombre(1).PasswordChar = "*"
Nombre(0).text = IIf(Trim(vg_NUsr) = "", "", Trim(vg_NUsr))
Nombre(0).Enabled = IIf(Trim(vg_NUsr) = "", True, False)
Nombre(1).text = IIf(Trim(vg_Pass) = "", "", Trim(vg_Pass))
Nombre(1).Enabled = IIf(Trim(vg_Pass) = "", True, False)

If Not Nombre(0).Enabled Then

   Nombre_LostFocus 1

Else

   Nombre_LostFocus 0
   
End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Nombre_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii = 13 Then SendKeys "{TAB}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Public Sub Nombre_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

Select Case Index

Case 0

    Nombre(Index).text = UCase(Nombre(Index).text)

Case 1
       
    '------- Traer contrato asginado usuario
    If Trim(Nombre(0).text) = "" Then
    
       MsgBox "Debe ingresar nombre de usuario...", vbCritical + vbOKOnly, "Ingreso al sistema"
       Exit Sub
    
    End If
    
    vg_NUsr = LimpiaDato(Trim(LCase(Nombre(0).text)))
    vg_Pass = LimpiaDato(Trim((Nombre(1).text)))
    AbrirBase
    ActVersion
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgp_Sel_Usuario 1, '" & vg_NUsr & "'")
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

    If Not RS1.EOF And (UCase(vg_Pass) <> UCase(fg_Desencripta(TipoDato(RS1!usu_password, ""))) Or Trim(fg_Desencripta(TipoDato(RS1!usu_password, ""))) = "") Then
    
       MsgBox "La clave no corresponde al usuario..."
       RS1.Close
       Set RS1 = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
       
    End If
    
    If Not RS1.EOF And RS1!usu_perfil < 0 Then
    
       MsgBox "Usuario no tiene asignado un perfil..."
       RS1.Close
       Set RS1 = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    vg_CPer = RS1!usu_perfil
    RS1.Close
    Set RS1 = Nothing
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("sgp_Sel_UsuarioContratos 1, '" & vg_NUsr & "'")
    
    If RS1.EOF Then
       
       RS1.Close
       Set RS1 = Nothing
       MsgBox "Usuario no tiene asignado centro de costo..."
       Exit Sub
    
    End If
    
    Combo1(0).Clear
    Do While Not RS1.EOF
        
       Combo1(0).AddItem RS1!cli_nombre & Space(150) & "(" & fg_pone_espacio(RS1!uco_codcon, 10) & ")" 'fg_pone_espacio(RS1!cli_codigo, 10)
       RS1.MoveNext
    
    Loop
    RS1.Close
    Set RS1 = Nothing
    Combo1(0).ListIndex = 0
    
    If Combo1(0).listcount = 1 Then
    
       On Error Resume Next
       Command1(0).SetFocus
       On Error Resume Next
       Combo1(0).ListIndex = 0 'Else Combo1(0).SetFocus
    
    End If
    
    Dim vRet As Variant
    If Not ConsultaProcess("sgpsdx.exe") Then
    
       On Error Resume Next
       vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")

    End If
    
    Dim UsuarioAdm As String
    UsuarioAdm = GetParametro_Seguridad("CAdmsgpLoc")
    
    If UCase(vg_NUsr) <> UCase(UsuarioAdm) Then
       
       If Not RevisaPassword Then Exit Sub
    
    End If
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ingreso_Correcto"), "SGP", "", "", "")
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

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

Set RS = vg_db.Execute("sgp_Sel_Log_CambiaPass_V01 1, '" & vg_NUsr & "', " & fg_TraeLogConcepto("vg_logsis_CambiaPass") & "")
If Not RS.EOF Then

    cPassPlazo = GetParametro_Seguridad("plazpass")
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

'Set RS = vg_db.Execute("sgp_Sel_Log_CambiaPass_V01 2, '" & vg_NUsr & "', " & fg_TraeLogConcepto("vg_logsis_CambiaPass") & "")
Set RS = vg_db.Execute("sgp_Sel_Log_CambiaPass_V01 2, '" & vg_NUsr & "', " & fg_TraeLogConcepto("vg_logsis_Ingreso_Correcto") & "")

If Not RS.EOF Then
    
    If TipoDato(RS!cuenta, 0) = 0 Then '1 Then
        
        RevisaPassword = False
    
    End If

End If
RS.Close
Set RS = Nothing

If Not RevisaPassword Then
    
    Frame1(2).Visible = True
    Me.Height = 6270
    Frame1(0).Enabled = False
    Frame1(1).Enabled = False
    Nombre(3).SetFocus
'    SendKeys "{TAB}"

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Private Function Prametro_Generales() As Boolean

On Error GoTo Man_Error
    
Dim RS  As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
    
Prametro_Generales = True

    If Combo1(0).ListIndex = -1 Then
    
       MsgBox "Debe seleccionar centro costo...", vbCritical + vbOKOnly, "Ingreso al sistema"
       Prametro_Generales = False
       Exit Function
    
    End If
    
    '------- Mover centro de costo a variables globales
    vg_contra = Trim(fg_codigocbo(Combo1, 0, 10, ""))
    vg_nomcon = Trim(Left(Combo1(0).List(Combo1(0).ListIndex), Len(Combo1(0).List(Combo1(0).ListIndex)) - 20))
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgp_Sel_ClienteBodega 1, '" & vg_contra & "'")
    If RS1.EOF Then
       
       RS1.Close
       Set RS1 = Nothing
       MsgBox "No existe bodega asignado a este contrato, cancela ingreso al sistema...", vbCritical + vbOKOnly, "Ingreso al sistema"
       Prametro_Generales = False
       Exit Function
       
    End If
    
    vg_codbod = RS1!bod_codigo
    vg_nombod = Trim(RS1!bod_nombre)
    RS1.Close: Set RS1 = Nothing
    
    vg_pais = ""
    vg_pais = GetParametro("parpais")
    vg_tipmonsap = ""
    vg_tipmonsap = GetParametro("tipmonsap")
    
    Set RS1 = vg_db.Execute("Update DBO.b_minutadet Set mid_nummer = 0 Where mid_nummer Is Null")

    Set RS1 = vg_db.Execute("Update DBO.b_minutadet Set mid_numrac = 0 Where mid_numrac Is Null")

    Set RS1 = vg_db.Execute("Update DBO.b_minutadet Set mid_cosrec = 0 Where mid_cosrec Is Null")

    Set RS1 = vg_db.Execute("Update DBO.b_minutadet Set mid_cosdes = 0 Where mid_cosdes Is Null")
    
    SwSalir = 1
    
    '-------> Validar PC Servidor esta en blanco nombre de maquina
    Dim sEquipo As String * 255
    GetComputerName sEquipo, 255
    
    If ValidaPCServidorVacio = False Then
    
        If MsgBox("Desea que este PC realice Cierre Diario ?? ", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
    
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'SvrAppCont'")
           If RS.EOF Then
   
              vg_db.Execute ("sgp_Ins_Param 'SvrAppCont', 'Identifica PC Servidor.', 'C', '" & Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1) & "', '" & MuestraCasino(1) & "'")

           Else
        
              vg_db.Execute ("sgp_Upd_Param 2, '" & MuestraCasino(1) & "',  'SvrAppCont', '', '',  '" + Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1) + "'")
        
           End If
           RS.Close
           Set RS = Nothing
        
        End If
    
    End If

    '-------> Crear o bien Modificar caracteristica del MAC
    Dim DescripcionMaquina As String
    DescripcionMaquina = ""
    DescripcionMaquina = Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1) & ";" & ObtenerMACcomputadora & ";" & TipoDato(GetParametro("version"), 0) & ";" & DateTime.Now
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgp_Sel_Param 4, '" & MuestraCasino(1) & "', '" & Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1) & "'")
    If RS.EOF Then
   
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
   
       Set RS1 = vg_db.Execute("sgp_Ins_ParamMac '', 'Descripción PC', 'C', '" & DescripcionMaquina & "', '" & MuestraCasino(1) & "'")

       If Not RS1.EOF Then
   
          If RS1(0) > 0 Then
                       
             MsgBox RS1(0) & " " & RS1(1)
   
          End If
    
       End If

       RS1.Close: Set RS1 = Nothing

    Else
        
       vg_db.Execute ("sgp_Upd_Param 1, '" & MuestraCasino(1) & "', '" & RS!par_codigo & "', '', '', '" & DescripcionMaquina & "'")
        
    End If
    RS.Close: Set RS = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function
