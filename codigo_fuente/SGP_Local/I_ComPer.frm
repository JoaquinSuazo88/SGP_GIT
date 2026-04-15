VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_ComPer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Compras por Periodo"
   ClientHeight    =   4875
   ClientLeft      =   2640
   ClientTop       =   4290
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   5685
      _ExtentX        =   10028
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame6 
      Caption         =   "Filtro de Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   15
      TabIndex        =   13
      Top             =   345
      Width           =   5655
      Begin VB.CheckBox Check1 
         Caption         =   "Tipo de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3585
         TabIndex        =   3
         Top             =   360
         Width           =   1995
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2280
         TabIndex        =   2
         Top             =   345
         Width           =   1275
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Bodega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1155
         TabIndex        =   1
         Top             =   360
         Width           =   1035
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   195
         TabIndex        =   0
         Top             =   360
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   30
      TabIndex        =   8
      Top             =   1110
      Width           =   5610
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   90
         TabIndex        =   12
         Top             =   2760
         Width           =   5370
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   330
            Width           =   3135
         End
         Begin VB.Label Label1 
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   4
            Left            =   180
            TabIndex        =   17
            Top             =   375
            Width           =   1125
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   90
         TabIndex        =   11
         Top             =   1830
         Width           =   5385
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   105
            TabIndex        =   22
            Top             =   465
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
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
            MaxLength       =   20
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
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   1620
            Picture         =   "I_ComPer.frx":0000
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rut"
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
            Left            =   150
            TabIndex        =   21
            Top             =   225
            Width           =   315
         End
         Begin VB.Label label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2115
            TabIndex        =   20
            Top             =   465
            Width           =   3090
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   2160
            TabIndex        =   19
            Top             =   510
            Width           =   3090
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bodega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   90
         TabIndex        =   10
         Top             =   975
         Width           =   5430
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   2130
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   285
            Width           =   3105
         End
         Begin VB.Label Label1 
            Caption         =   "Bodega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   16
            Top             =   330
            Width           =   1125
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   90
         TabIndex        =   9
         Top             =   135
         Width           =   5445
         Begin EditLib.fpDateTime Fecha 
            Height          =   315
            Index           =   0
            Left            =   1275
            TabIndex        =   4
            Top             =   255
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
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Fecha 
            Height          =   315
            Index           =   1
            Left            =   3915
            TabIndex        =   5
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
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   2
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
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Hasta"
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
            Index           =   1
            Left            =   2880
            TabIndex        =   15
            Top             =   300
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha Desde"
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
            Left            =   105
            TabIndex        =   14
            Top             =   300
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "I_ComPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset

Private Sub Check1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    Case 0
        
        If Check1(0).Value = 1 Then
           
           Fecha(0).Enabled = True: Fecha(1).Enabled = True
           Fecha(0).text = Date: Fecha(1).text = Date
        
        Else
           
           Fecha(0).Enabled = False: Fecha(1).Enabled = False
           Fecha(0).text = "": Fecha(1).text = ""
        
        End If
    
    Case 1
    
    '    If Check1(1).Value = 1 Then
    '       Combo1(0).Enabled = True
    '    Else
    '       Combo1(0).Enabled = False: Combo1(0).ListIndex = -1
    '    End If
    
    Case 2
        
        If Check1(2).Value = 1 Then
           
           Image1(0).Enabled = True: fpText1(0).Enabled = True
           fpText1(0).SetFocus
        
        Else
           
           Image1(0).Enabled = False: fpText1(0).Enabled = True
           fpText1(0).text = "": Label2(0).Caption = ""
        
        End If
    
    Case 3
        
        If Check1(3).Value = 1 Then
           
           Combo1(1).Enabled = True
        
        Else
           
           Combo1(1).Enabled = False: Combo1(1).ListIndex = -1
        
        End If
        
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Select

End Sub

Private Sub fecha_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Fecha_LostFocus(Index As Integer)

On Error GoTo Man_Error

If IsDate(Fecha(0).text) = False Then
    
    Fecha(0).text = Date

ElseIf IsDate(Fecha(1).text) = False Then
    
    Fecha(1).text = Date

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Local Error GoTo Error_Partida
'***** Crear botones para botonera *****
'Dim btnX As Button

Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
EspFecha Fecha(0)
EspFecha Fecha(1)
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Me.Height = 5385
Me.Width = 5805
MsgTitulo = "Compras por Periodo"
fg_centra Me
'****** Carga Combos con Información Necesaria *****
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
'-------> Cargar Combo tipo Documento
CargarDatoCombo Combo1, 1, "a_tipodocumento", "tdo_", "Gen", "A"
'***** Limpia Objetos *****
Fecha(0).text = ""
Fecha(1).text = ""
fpText1(0).text = ""
Check1(0).Value = 0
Check1(1).Value = 1
Check1(1).Enabled = False
Check1(2).Value = 0
Check1(3).Value = 0
If Combo1(0).listcount > 0 Then Combo1(0).ListIndex = 0
'***** Deshabilita Objetos *****
Combo1(0).Enabled = False
Combo1(1).Enabled = False
Fecha(0).Enabled = False
Fecha(1).Enabled = False
fpText1(0).Enabled = False
Image1(0).Enabled = False
Label2(0).Enabled = False

Exit Sub
Error_Partida:
    MsgBox "Error: " & Err.Number & "-" & Err.Description, vbExclamation, MsgTitulo
End Sub

Private Sub fpText1_GotFocus(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    
    Case 0
        
        If Trim(fpText1(0).text) = "" Or vg_Dig = "N" Then Exit Sub
        fpText1(0).text = fg_DespintaRut(fpText1(0).text)
        fpText1(0).text = Mid(fpText1(0).text, 1, Len(Trim(fpText1(0).text)) - 1)

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_LostFocus(Index As Integer)

Select Case Index
    
    Case 0
        
        If fpText1(0).text = "" Then Exit Sub
        fpText1(0).text = fg_RutDig(Trim(fpText1(0).text))
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        RS1.Open RutinaLectura.Proveedor(1, Trim(LimpiaDato(fpText1(0).text)), ""), vg_db, adOpenStatic
        
        If Not RS1.EOF Then
            
            Label2(0).Caption = RS1!prv_nombre
        
        Else
            
            fpText1(0).text = "": Label2(0).Caption = ""
            fpText1(0).SetFocus: RS1.Close: Set RS1 = Nothing: Exit Sub
        
        End If
        RS1.Close
        Set RS1 = Nothing
        
        fpText1(0).text = fg_PintaRut(fpText1(0).text)

End Select

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

vg_codigo = 0
Select Case Index
    
    Case 0
        
        fpText1(0).text = ""
        vg_left = Label2(0).Left + 2300
        B_TabEst.LlenaDatos "b_proveedor", "prv_", "Proveedor", "Gen"
        B_TabEst.Show 1, Me
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        Label2(Index).Caption = vg_nombre
        fpText1(Index).text = fg_PintaRut(vg_codigo)
        Check1(2).SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim paso As Long
paso = 0

Select Case Button.Index

Case 1 '------- Previsualizar
    
    If Check1(0).Value = 1 Then  'Fecha
       
       If IsDate(Fecha(0)) = False Then MsgBox "Rango de fechas no valido...", vbExclamation, MsgTitulo: Exit Sub
       If IsDate(Fecha(1)) = False Then MsgBox "Rango de fechas no valido...", vbExclamation, MsgTitulo: Exit Sub
       If CDate(Fecha(0).text) > CDate(Fecha(1).text) Then MsgBox "Rango de fechas no valido...", vbExclamation, MsgTitulo: Exit Sub
       paso = paso + 1
    
    End If
    
    If Check1(1).Value = 1 Then
       
       If Combo1(0).ListIndex = -1 Then MsgBox "Bodega no valida...", vbExclamation, MsgTitulo: Exit Sub
       paso = paso + 1
    
    End If
    
    If Check1(2).Value = 1 Then
       
       If Len(fpText1(0).text) = 0 Then MsgBox "Proveedor no valido...", vbExclamation, MsgTitulo: Exit Sub
       paso = paso + 1
    
    End If
    
    If Check1(3).Value = 1 Then
       
       If Combo1(1).ListIndex = -1 Then MsgBox "Tipo de documento no valido...", vbExclamation, MsgTitulo: Exit Sub
       paso = paso + 1
    
    End If
    
    If paso > 0 Then I_ComprasPer Else MsgBox "Seleccione metodo de busqueda...", vbExclamation, MsgTitulo

Case 3 '------- Salir
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub
