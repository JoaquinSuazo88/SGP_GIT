VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_Perfil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfil de Acceso"
   ClientHeight    =   7275
   ClientLeft      =   1965
   ClientTop       =   2640
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   11880
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   476
      TabMaxWidth     =   4
      TabCaption(0)   =   "Perfil de Acceso"
      TabPicture(0)   =   "M_Perfil.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Perfil.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "vaSpread2"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5610
         Left            =   0
         TabIndex        =   15
         Top             =   960
         Width           =   7920
         _Version        =   393216
         _ExtentX        =   13970
         _ExtentY        =   9895
         _StockProps     =   64
         BackColorStyle  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   4
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         ScrollBars      =   2
         SpreadDesigner  =   "M_Perfil.frx":0038
         ScrollBarTrack  =   1
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -73785
         TabIndex        =   10
         Top             =   435
         Width           =   5370
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Perfil.frx":1C4B
            Left            =   1680
            List            =   "M_Perfil.frx":1C55
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   240
            Width           =   2415
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   16
            Top             =   600
            Width           =   2505
            _Version        =   196608
            _ExtentX        =   4410
            _ExtentY        =   870
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
            NoSpecialKeys   =   0
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
            AutoSize        =   -1  'True
            Caption         =   "Buscar Columna"
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
            Index           =   7
            Left            =   255
            TabIndex        =   14
            Top             =   345
            Width           =   1380
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Texto"
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
            Left            =   255
            TabIndex        =   13
            Top             =   645
            Width           =   1140
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
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
            Left            =   4260
            TabIndex        =   12
            Top             =   645
            Width           =   585
         End
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   5
         Left            =   -72900
         TabIndex        =   1
         Top             =   3075
         Width           =   3735
         _Version        =   196608
         _ExtentX        =   6588
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
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
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4845
         Left            =   -74640
         TabIndex        =   17
         Top             =   1440
         Width           =   7230
         _Version        =   393216
         _ExtentX        =   12753
         _ExtentY        =   8546
         _StockProps     =   64
         AutoCalc        =   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FormulaSync     =   0   'False
         MaxCols         =   2
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Perfil.frx":1C69
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin VB.Label Label3 
         Caption         =   "Perfil:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   600
         Width           =   6375
      End
      Begin VB.Label Label1 
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   -74280
         TabIndex        =   8
         Top             =   3180
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Oficina"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   -74280
         TabIndex        =   7
         Top             =   2865
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   -74280
         TabIndex        =   6
         Top             =   2550
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   -74280
         TabIndex        =   5
         Top             =   2235
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   -74280
         TabIndex        =   4
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   -74280
         TabIndex        =   3
         Top             =   1605
         Width           =   1395
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Label4"
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
         Left            =   -74760
         TabIndex        =   2
         Top             =   4560
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8100
      _ExtentX        =   14288
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim codigo As Integer
Dim Est As Boolean
Dim OpGr As Boolean, sw As Boolean

Private Sub MDIForm_Activate()

On Error GoTo Man_Error

vg_opimp = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long

OpGr = True
vaSpread1.Row = Fila

vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

vaSpread1.Col = 2
Nombre = Trim(LimpiaDato(vaSpread1.Value))

If Trim(Nombre) = "" Then
   
   MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 2
   vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub
   
End If

If modo = "A" Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_s_perfil 8, '','" & Trim(Mid(Nombre, 1, 30)) & "'")
   
   If Not RS.EOF Then
   
      RS.Close
      Set RS = Nothing
      MsgBox "Perfil ya existe...", vbCritical + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   RS.Close
   Set RS = Nothing
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS = vg_db.Execute("sgpadm_Ins_Perfil '" & Trim(Mid(Nombre, 1, 30)) & "'")
    
   If Not RS.EOF Then
    
      If RS(0) > 0 Then
          
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), CStr(Me.HelpContextID), RS(0) & RS(1), "", "")
         RS.Close
         Set RS = Nothing
         MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
         Exit Sub
       
      Else
      
         codigo = RS(3)
         
      End If
    
   End If
   RS.Close
   Set RS = Nothing

   vaSpread1.Col = 1
   vaSpread1.Value = codigo
   
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado"), CStr(Me.HelpContextID), "", "", "")

Else
    
    
   For i = 1 To vaSpread1.MaxRows
   
       vaSpread1.Row = i
       vaSpread1.Col = 2
       If Fila <> i And Trim(Nombre) = Trim(vaSpread1.text) Then
   
          MsgBox "Ya existe nombre Perfil ...", vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
          
   
       End If
       
   Next i
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS = vg_db.Execute("sgpadm_Upd_Perfil " & codigo & ", '" & Trim(Mid(Nombre, 1, 30)) & "'")
    
   If Not RS.EOF Then
    
      If RS(0) > 0 Then
          
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), CStr(Me.HelpContextID), RS(0) & RS(1), "", "")
         RS.Close
         Set RS = Nothing
         MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
         Exit Sub
       
      Else
      
         codigo = RS(2)
         
      End If
    
   End If
   RS.Close
   Set RS = Nothing

   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), CStr(Me.HelpContextID), "", "", "")
    
End If

Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True
fpText1(0).Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True

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

Me.HelpContextID = vg_OpcM
MsgTitulo = "Perfiles"
Label3.Visible = False
Label5.Visible = False
fg_centra Me
modo = ""
ibusca = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
'Me.vaSpread1.hwnd
MoverDatosGrillas
SSTab1.Tab = 0
OpGr = False
Est = True
sw = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

'If Me.WindowState = 0 Then
'   Frame1.Move 1695, 360, 6015, 971
'   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
'ElseIf Me.WindowState = 2 Then
'   Frame1.Move 4200, 360, 6015, 971
'   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
'End If
'Toolbar1.Refresh

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If LimpiaDato(Trim(fpText1(0).text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   
   Set RS = vg_db.Execute("sgpadm_s_perfil 2, 0, '%" & UCase(LimpiaDato(fpText1(0).text)) & "%'")
   If Not RS.EOF Then
      
      ibusca = RS!nReg
      vaSpread1.MaxRows = RS!nReg
   
   Else
      
      ibusca = 0
      vaSpread1.MaxRows = 0
   
   End If

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   
   Set RS = vg_db.Execute("sgpadm_s_perfil 3, 0, '%" & UCase(LimpiaDato(fpText1(0).text)) & "%'")
   
   If Not RS.EOF Then
      
      ibusca = RS!nReg
      vaSpread1.MaxRows = RS!nReg
   
   Else
      
      ibusca = 0
      vaSpread1.MaxRows = 0
   
   End If

End If
i = 1

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Value = RS!per_codigo
      
      vaSpread1.Col = 2
      vaSpread1.Value = Trim(RS!per_nombre)
      
      RS.MoveNext
   
   Loop
   modo = ""
   Gl_Ac_Botones Me, 1, 1, modo

End If
RS.Close
Set RS = Nothing

If fpText1(0).text = "" Then
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"

Else
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

'On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case SSTab1.Tab

    Case 0
        
        Label3.Visible = False
        Label5.Visible = False
        modo = ""
    
    Case 1
        
        modo = "M"
        Label3.Visible = True
        Label5.Visible = True
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 2
        Label5.Caption = vaSpread1.text
        vaSpread1.Col = 1
        codigo = vaSpread1.text
        vaSpread2.MaxRows = 1000
        i = 1
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        
        Set RS = vg_db.Execute("sgpadm_s_perfil 4, " & codigo & ", ''")
        Est = True
        
        Do While Not RS.EOF
            
            DoEvents
            i = i + 1
            vaSpread2.Row = i
            
            vaSpread2.Col = 1
            vaSpread2.text = IIf(IsNull(RS!opc_nombre), "", Trim(RS!opc_nombre))
            
            vaSpread2.Col = 3
            vaSpread2.text = IIf(IsNull(RS!dpe_deracc), "0", Trim(Str(RS!dpe_deracc)))
            
            vaSpread2.Col = 4
            vaSpread2.text = IIf(IsNull(RS!dpe_deragr), "0", RS!dpe_deragr)
            
            vaSpread2.Col = 5
            vaSpread2.text = IIf(IsNull(RS!dpe_dermod), "0", RS!dpe_dermod)
            
            vaSpread2.Col = 6
            vaSpread2.text = IIf(IsNull(RS!dpe_dereli), "0", RS!dpe_dereli)
            
            vaSpread2.Col = 7
            vaSpread2.text = IIf(IsNull(RS!dpe_derimp), "0", RS!dpe_derimp)
            
            vaSpread2.Col = 8
            vaSpread2.text = IIf(IsNull(RS!opc_codigo), "", Trim(RS!opc_codigo))
            
            RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        Est = False
        vaSpread2.MaxRows = i
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim codigo    As Long
Dim Nombre    As String
Dim NomCor    As String
Dim CodOpc    As String
Dim Action    As Long
Dim Agrega    As Long
Dim Modifica  As Long
Dim Elimina   As Long
Dim Imprime   As Long
Dim RS        As New ADODB.Recordset
Dim MyBuffer  As String
Dim EstExiste As String

Dim xlApp      As Object
Dim xlWb       As Object
Dim xlWs       As Object

Select Case Button.Index

    Case 1
        
        SSTab1.Tab = 0
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.TabEnabled(1) = False
        'vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        'vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.Row: vaSpread1.SetFocus
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 2
        vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
        vaSpread1.SetFocus
    
    Case 3
        
        modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.Tab = 1
        
    Case 5
        
        If vaSpread1.ActiveRow < 1 Then
           
           MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
        
           Exit Sub
           
        End If
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = Val(vaSpread1.Value)
        
        Set RS = vg_db.Execute("sgpadm_Del_PerfilDerechoPerfil '" & codigo & "'")

        If Not RS.EOF Then
   
           If RS(0) > 0 Then

              'registrar Log sistema error Eliminacion
              Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, RS(0) & RS(1), "", "")
                         
              MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
      
           Else
   
              'registrar Log sistema Eliminar
              Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")
      
        
              vaSpread1.DeleteRows vaSpread1.Row, 1
              vaSpread1.MaxRows = vaSpread1.MaxRows - 1
              Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
              modo = ""
              Gl_Ac_Botones Me, 1, 1, modo
              MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, MsgTitulo
           
           End If

        End If
        RS.Close
        Set RS = Nothing
        
    Case 7
        
        MoverDatosGrillas
        
    Case 10
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        'modo = "A"
        
        If modo = "A" Then
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.DeleteRows vaSpread1.Row, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
            SSTab1.Tab = 0
            MoverDatosGrillas
        
        Else
            
            OpGr = True
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
            
            Select Case SSTab1.Tab
                
                Case 0
                    
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    Set RS = vg_db.Execute("sgpadm_s_perfil 5, " & codigo & ", ''")
                    If Not RS.EOF Then
                    
                       DoEvents
                       vaSpread1.Col = 2
                       vaSpread1.Value = Trim(RS!per_nombre)
                       
                    End If
                    RS.Close
                    Set RS = Nothing
                
                Case 1
                    
                    Label3.Visible = True
                    Label5.Visible = True
                    vaSpread1.Row = vaSpread1.ActiveRow
                    vaSpread1.Col = 2
                    Label5.Caption = vaSpread1.text
                    vaSpread1.Col = 1
                    codigo = vaSpread1.text
                    vaSpread2.MaxRows = 3000: i = 1
                    
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    Set RS = vg_db.Execute("sgpadm_s_perfil 4, " & codigo & ", ''")
                    Est = True
                    Do While Not RS.EOF
                    
                        DoEvents
                        i = i + 1
                        vaSpread2.Row = i
                        
                        vaSpread2.Col = 1
                        vaSpread2.text = IIf(IsNull(RS!opc_nombre), "", Trim(RS!opc_nombre))
                        
                        vaSpread2.Col = 2
                        vaSpread2.text = "0"
                        
                        vaSpread2.Col = 3
                        vaSpread2.text = IIf(IsNull(RS!dpe_deracc), "0", Trim(Str(RS!dpe_deracc)))
                        
                        vaSpread2.Col = 4
                        vaSpread2.text = IIf(IsNull(RS!dpe_deragr), "0", RS!dpe_deragr)
                        
                        vaSpread2.Col = 5
                        vaSpread2.text = IIf(IsNull(RS!dpe_dermod), "0", RS!dpe_dermod)
                        
                        vaSpread2.Col = 6
                        vaSpread2.text = IIf(IsNull(RS!dpe_dereli), "0", RS!dpe_dereli)
                        
                        vaSpread2.Col = 7
                        vaSpread2.text = IIf(IsNull(RS!dpe_derimp), "0", RS!dpe_derimp)
                        
                        vaSpread2.Col = 8
                        vaSpread2.text = IIf(IsNull(RS!opc_codigo), "", Trim(RS!opc_codigo))
                        
                        RS.MoveNext
                    
                    Loop
                    RS.Close
                    Set RS = Nothing
                    
                    Est = False
                    vaSpread2.MaxRows = i
            
            End Select
            
            OpGr = False
        
        End If
        
        Gl_Ac_Botones Me, 1, 1, modo
        Combo1.Enabled = True: fpText1(0).Enabled = True
    
    Case 12
        
        If SSTab1.Tab = 0 Then
            
            GrabaRegistro vaSpread1.ActiveRow
        
        Else
            
            OpGr = True
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = 1
            codigo = Val(vaSpread1.Value)
            
            Let MyBuffer = ""
            Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
            Let MyBuffer = MyBuffer & "<DerechoPerfil>"
            
            
'            vg_db.Execute "DELETE FROM a_derechosperfil WHERE dpe_codper=" & codigo & ""
            
            For i = 2 To vaSpread2.MaxRows
                
                vaSpread2.Row = i
                vaSpread2.Col = 8
                CodOpc = Trim(vaSpread2.Value)
                
                vaSpread2.Col = 3
                accion = IIf(Val(vaSpread2.Value) = 1, 1, 0)
                
                vaSpread2.Col = 4
                Agrega = IIf(Val(vaSpread2.Value) = 1, 1, 0)
                
                vaSpread2.Col = 5
                Modifica = IIf(Val(vaSpread2.Value) = 1, 1, 0)
                
                vaSpread2.Col = 6
                Elimina = IIf(Val(vaSpread2.Value) = 1, 1, 0)
                
                vaSpread2.Col = 7
                Imprime = IIf(Val(vaSpread2.Value) = 1, 1, 0)
                
'                vg_db.Execute "INSERT INTO a_derechosperfil (dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp)  _ &"
'                VALUES (" & codigo & ", '" & CodOpc & "', " & accion & ", " & Agrega & ", " & Modifica & ", " & Elimina & ", " & Imprime & ")"
            
                MyBuffer = MyBuffer & " <Perfil"
                MyBuffer = MyBuffer & " codper= " & Chr(34) & codigo & Chr(34)
                MyBuffer = MyBuffer & " codopc= " & Chr(34) & CodOpc & Chr(34)
                MyBuffer = MyBuffer & " deracc= " & Chr(34) & accion & Chr(34)
                MyBuffer = MyBuffer & " deragr= " & Chr(34) & Agrega & Chr(34)
                MyBuffer = MyBuffer & " dermod= " & Chr(34) & Modifica & Chr(34)
                MyBuffer = MyBuffer & " dereli= " & Chr(34) & Elimina & Chr(34)
                MyBuffer = MyBuffer & " derimp= " & Chr(34) & Imprime & Chr(34)
                Let MyBuffer = MyBuffer & "/>"
            
            Next i
            
            Let MyBuffer = MyBuffer & "</DerechoPerfil>"
    
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
    
            EstExiste = 0
            Set RS = vg_db.Execute("sgpadm_s_perfil 4, " & codigo & ", ''")
                        
            If Not RS.EOF Then
            
               EstExiste = 1
            
            End If
            RS.Close
            Set RS = Nothing
            
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Ins_XmlDerechoPerfil '" & MyBuffer & "', " & codigo & "")
            If Not RS.EOF Then
       
               If RS(0) > 0 Then
          
                  Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto(IIf(EstExiste = "0", "vg_logsis_Error_Agregado", "vg_logsis_Error_Modificado")), CStr(Me.HelpContextID), RS(0) & RS(1), "", "")
          
                  RS.Close
                  Set RS = Nothing
                  MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
                  Exit Sub
       
               End If
    
            End If
            RS.Close
            Set RS = Nothing
            
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto(IIf(EstExiste = "0", "vg_logsis_Agregar_DerPerfil", "vg_logsis_Modificar_DerPerfil")), CStr(Me.HelpContextID), "", "", "")
            
            modo = "A"
            Gl_Ac_Botones Me, 1, 1, modo
            OpGr = False
            SSTab1.Tab = 0
        
        End If
    
    Case 15
        
        If vaSpread1.MaxRows < 1 Then
        
           MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), CStr(Me.HelpContextID), "", "", "")
       
'        If SSTab1.Tab = 0 Then
'
'            I_Perfil
'
'        Else
'
'            I_acceso
'
'        End If

        Set RS = vg_db.Execute("sgpadm_Sel_PerfilesAcceso")
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
        
            fg_descarga
            MsgBox "No existe información...", vbCritical
            RS.Close
            Set RS = Nothing
        
        
        End If
        fg_descarga
    
    Case 18
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "", "")
        
        Me.Hide
        Unload Me
        
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub
End If
If Err = 3034 Then Exit Sub
'Resume
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

'Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
'If Col <> 4 Or OpGr Then Exit Sub
'If modo = "" Then modo = "M"
'End Sub
    
'Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
'If modo = "" Then modo = "M"
'End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

OpGr = True
vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
Set RS = vg_db.Execute("sgpadm_s_perfil 6, 0, ''")

If Not RS.EOF Then

    Do While Not RS.EOF
    
        DoEvents
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1
        vaSpread1.Lock = True
        vaSpread1.Value = RS!per_codigo
        
        vaSpread1.Col = 2
        vaSpread1.Lock = False
        vaSpread1.Value = Trim(RS!per_nombre)
        
        RS.MoveNext
        
    Loop
    
Else
    
    Gl_Ac_Botones Me, 1, 2, modo

End If
RS.Close
Set RS = Nothing

vaSpread1.Visible = True
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Cancela()

On Error GoTo Man_Error

Dim codigo As Long
Dim RS     As New ADODB.Recordset

OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = Val(vaSpread1.text)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
RS.Open "SELECT * FROM a_grupocambioing WHERE gci_codigo = " & codigo & "", vg_db, adOpenStatic
DoEvents

If Not RS.EOF Then
   
   vaSpread1.Col = 2
   vaSpread1.Value = IIf(IsNull(RS!gci_nombre), "", Trim(RS!gci_nombre))

End If
RS.Close
Set RS = Nothing

OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

'Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
'If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") Then GrabaRegistro Row
'End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
'ElseIf Toolbar1.Buttons(12).Visible = False Then
'    Cancela
End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

Dim i As Long
If Est Then Exit Sub

If Row = 1 Then
    
    vaSpread2.Col = Col
    vaSpread2.Row = -1
    Est = True
    vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
    Est = False

End If

If Col = 2 Then
    
    vaSpread2.Row = Row
    
    For i = 3 To vaSpread2.maxcols - 1
        
        vaSpread2.Col = i
        Est = True
        vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
        Est = False
    
    Next i

End If

If Row = 1 And Col = 2 Then
    
    For i = 3 To vaSpread2.maxcols - 1
        
        vaSpread2.Col = i
        vaSpread2.Row = -1
        Est = True
        vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
        Est = False
    
    Next i

End If
Gl_Ac_Botones Me, 1, 0, modo
vaSpread2.Refresh

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
