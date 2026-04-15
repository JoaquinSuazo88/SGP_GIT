VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form M_Perfil 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Perfil de Acceso"
   ClientHeight    =   6000
   ClientLeft      =   4605
   ClientTop       =   2655
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8916
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
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Perfil.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "vaSpread2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4650
         Left            =   0
         TabIndex        =   16
         Top             =   360
         Width           =   7920
         _Version        =   393216
         _ExtentX        =   13970
         _ExtentY        =   8202
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
         TabIndex        =   11
         Top             =   435
         Width           =   5255
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Perfil.frx":1C4B
            Left            =   1680
            List            =   "M_Perfil.frx":1C55
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   2415
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   1680
            TabIndex        =   17
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         Height          =   3435
         Left            =   -74400
         TabIndex        =   10
         Top             =   1440
         Width           =   6255
         _Version        =   393216
         _ExtentX        =   11033
         _ExtentY        =   6059
         _StockProps     =   64
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
         MaxCols         =   2
         ScrollBars      =   2
         SpreadDesigner  =   "M_Perfil.frx":1C69
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
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
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
      TabIndex        =   19
      Top             =   5520
      Width           =   6375
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
      TabIndex        =   18
      Top             =   5520
      Width           =   495
   End
End
Attribute VB_Name = "M_Perfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim est As Boolean
Dim OpGr As Boolean, sw As Boolean

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim codigo As Long
Dim Nombre As String
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.Value))
If Trim(Nombre) = "" Then MsgBox "Favor ingresar nombre, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" Then
    vg_db.BeginTrans
    RS.Open "SELECT per_codigo FROM a_perfil ORDER BY per_codigo DESC", vg_db, adOpenStatic
    If Not RS.EOF Then RS.MoveFirst: codigo = RS!per_codigo + 1 Else codigo = 1
    RS.Close: Set RS1 = Nothing
    vg_db.Execute "INSERT INTO a_perfil (per_codigo, per_nombre) VALUES (" & codigo & ", '" & Trim(Nombre) & "')"
    vg_db.CommitTrans
    SSTab1.TabEnabled(1) = True
    vaSpread1.Col = 1: vaSpread1.Value = codigo
Else
    vg_db.BeginTrans
    vg_db.Execute "UPDATE a_perfil SET per_nombre = '" & Trim(Nombre) & "' WHERE per_codigo = " & codigo
    vg_db.CommitTrans
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Frame1.Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
'Me.Height = 6405
'Me.Width = 8310
MsgTitulo = "Perfiles"
Label3.Visible = False
Label5.Visible = False
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
SSTab1.Tab = 0
OpGr = False
est = True
sw = True
End Sub

Private Sub Form_Resize()
If Me.WindowState = 0 Then
   Frame1.Move 1095, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
ElseIf Me.WindowState = 2 Then
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Dim sql1 As String
If LimpiaDato(Trim(fpText1(0).text)) & Chr(KeyAscii) = "" Then Exit Sub
opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS.Open "SELECT per_codigo,per_nombre FROM a_perfil WHERE per_codigo LIKE '%" & UCase(LimpiaDato(fpText1(0).text)) & "%'", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    sql1 = IIf(vg_tipbase = "1", " UCASE(per_nombre) ", " UPPER(per_nombre) ")
    RS.Open "SELECT per_codigo, per_nombre FROM a_perfil WHERE " & sql1 & " LIKE '%" & UCase(LimpiaDato(fpText1(0).text)) & "%'", vg_db, adOpenStatic
End If
i = 1: vaSpread1.MaxRows = RS.RecordCount
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.Row = i
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS!per_codigo
      vaSpread1.Col = 2
      vaSpread1.Lock = opusu
      vaSpread1.Value = Trim(RS!per_nombre)
      RS.MoveNext: i = i + 1
   Loop
   modo = ""
   Gl_Ac_Botones Me, 1, 1, modo
End If
RS.Close: Set RS = Nothing
If fpText1(0).text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim sql1 As String
Dim RS As New ADODB.Recordset
Dim codigo As Long
Dim Nombre As String
On Error GoTo Man_Error
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
    
    vaSpread2.MaxRows = 0
    vaSpread2.MaxRows = 1000
    i = 1
    
    sql1 = IIf(vg_tipbase = "1", " mid(opc_codigo,1,1) ", " substring(convert(varchar(1),opc_codigo),1,1) ")
    
    vg_db.BeginTrans
    
    vg_db.Execute "INSERT INTO a_derechosperfil (dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT " & codigo & ", opc_codigo, 0, 0, 0, 0, 0 FROM a_opcsistema WHERE opc_codigo NOT IN (SELECT dpe_codopc FROM a_derechosperfil WHERE dpe_codper = " & codigo & ") AND " & sql1 & " " & IIf(vg_modpac, "IN", "NOT IN") & " ('5')"
    vg_db.Execute "INSERT INTO a_derechosperfil (dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT " & codigo & ", opc_codigo, 0, 0, 0, 0, 0 FROM a_opcsistema WHERE opc_codigo NOT IN (SELECT dpe_codopc FROM a_derechosperfil WHERE dpe_codper = " & codigo & ") AND " & sql1 & " NOT IN ('5')"
    
    vg_db.CommitTrans
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_opcsistema, a_derechosperfil WHERE opc_codigo = dpe_codopc AND dpe_codper = " & codigo & " ORDER BY opc_codigo", vg_db, adOpenStatic
    est = True
    
    vaSpread2.Col = -1
    vaSpread2.Row = -1
    vaSpread2.Lock = IIf(codigo = 0, True, False)
    
    Do While Not RS.EOF
        
        i = i + 1
        vaSpread2.Row = i
        
        vaSpread2.Col = 1
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = Trim(RS!opc_nombre)
        
        vaSpread2.Col = 2
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        
        vaSpread2.Col = 3
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = IIf(IsNull(RS!dpe_deracc), "", Trim(Str(RS!dpe_deracc)))
        
        vaSpread2.Col = 4
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = IIf(IsNull(RS!dpe_deragr), "0", RS!dpe_deragr)
        
        vaSpread2.Col = 5
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = IIf(IsNull(RS!dpe_dermod), "0", RS!dpe_dermod)
        
        vaSpread2.Col = 6
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = IIf(IsNull(RS!dpe_dereli), "0", RS!dpe_dereli)
        
        vaSpread2.Col = 7
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = IIf(IsNull(RS!dpe_derimp), "0", RS!dpe_derimp)
        
        vaSpread2.Col = 8
        vaSpread2.Lock = IIf(codigo = 0 Or RS!opc_codigo = 4800000 Or RS!opc_codigo = 4810000, True, False)
        vaSpread2.Value = IIf(IsNull(RS!opc_codigo), "", Trim(RS!opc_codigo))
        
        RS.MoveNext
    
    Loop
    
    RS.Close
    Set RS = Nothing
    
    est = False: vaSpread2.MaxRows = i
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim codigo As Long, Nombre As String, NomCor As String
Dim CodOpc As String, Action As Long, Agrega As Long, Modifica As Long, Elimina As Long, Imprime As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    SSTab1.Tab = 0
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.Row: vaSpread1.SetFocus
    Frame1.Enabled = False
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.Tab = 1
    Frame1.Enabled = False
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
    
    If codigo = 0 Then
    
       MsgBox "No es posible eliminar Perfil del SGP Administrador...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vg_db.BeginTrans
    vg_db.Execute "DELETE a_perfil FROM a_perfil WHERE per_codigo = " & codigo
    vg_db.Execute "DELETE a_derechosperfil FROM a_derechosperfil WHERE dpe_codper = " & codigo
    vg_db.CommitTrans
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
    modo = "": Gl_Ac_Botones Me, 1, 1, modo

Case 7
    
    MoverDatosGrillas

Case 10
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 0
        MoverDatosGrillas
    Else
        OpGr = True
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
        RS.Open "SELECT * FROM a_perfil WHERE per_codigo = " & codigo, vg_db, adOpenStatic
        If Not RS.EOF Then vaSpread1.Col = 2: vaSpread1.Value = Trim(RS!per_nombre)
        RS.Close: Set RS = Nothing
        vaSpread2.Visible = False
        vaSpread2.MaxRows = 0: vaSpread2.MaxRows = 1000: i = 1
        RS.Open "SELECT * FROM a_opcsistema, a_derechosperfil WHERE opc_codigo = dpe_codopc AND dpe_codper = " & codigo & " ORDER BY opc_codigo", vg_db, adOpenStatic
        est = True
        Do While Not RS.EOF
            i = i + 1: vaSpread2.Row = i
            vaSpread2.Col = 1: vaSpread2.Value = Trim(RS!opc_nombre)
            vaSpread2.Col = 3: vaSpread2.Value = IIf(IsNull(RS!dpe_deracc), "", Trim(Str(RS!dpe_deracc)))
            vaSpread2.Col = 4: vaSpread2.Value = IIf(IsNull(RS!dpe_deragr), "0", RS!dpe_deragr)
            vaSpread2.Col = 5: vaSpread2.Value = IIf(IsNull(RS!dpe_dermod), "0", RS!dpe_dermod)
            vaSpread2.Col = 6: vaSpread2.Value = IIf(IsNull(RS!dpe_dereli), "0", RS!dpe_dereli)
            vaSpread2.Col = 7: vaSpread2.Value = IIf(IsNull(RS!dpe_derimp), "0", RS!dpe_derimp)
            vaSpread2.Col = 8: vaSpread2.Value = IIf(IsNull(RS!opc_codigo), "", Trim(RS!opc_codigo))
            RS.MoveNext
        Loop
        RS.Close: Set RS = Nothing
        est = False: vaSpread2.MaxRows = i
        vaSpread2.Visible = True
        OpGr = False
        
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
    End If
    Gl_Ac_Botones Me, 1, 1, modo
    Frame1.Enabled = True
Case 12
    If SSTab1.Tab = 0 Then
       GrabaRegistro vaSpread1.ActiveRow
    Else
       OpGr = True
       vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
       vg_db.BeginTrans
       vg_db.Execute "DELETE a_derechosperfil FROM a_derechosperfil WHERE dpe_codper = " & codigo
       For i = 2 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = 8: CodOpc = Trim(vaSpread2.Value)
           vaSpread2.Col = 3: accion = IIf(Val(vaSpread2.Value) = 1, 1, 0)
           vaSpread2.Col = 4: Agrega = IIf(Val(vaSpread2.Value) = 1, 1, 0)
           vaSpread2.Col = 5: Modifica = IIf(Val(vaSpread2.Value) = 1, 1, 0)
           vaSpread2.Col = 6: Elimina = IIf(Val(vaSpread2.Value) = 1, 1, 0)
           vaSpread2.Col = 7: Imprime = IIf(Val(vaSpread2.Value) = 1, 1, 0)
           vg_db.Execute "INSERT INTO a_derechosperfil (dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) VALUES (" & codigo & ", '" & CodOpc & "', " & accion & ", " & Agrega & ", " & Modifica & ", " & Elimina & ", " & Imprime & ")"
       Next i
       vg_db.CommitTrans
       modo = "A": Gl_Ac_Botones Me, 1, 1, modo
       OpGr = False
       SSTab1.Tab = 0
    End If
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If SSTab1.Tab = 0 Then I_Perfil Else I_acceso
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Col <> 4 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
End Sub
    
Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Frame1.Enabled = False
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = False
End Sub

Private Sub MoverDatosGrillas()

Dim RS As New ADODB.Recordset
Dim opusu As Boolean
OpGr = True
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT * FROM a_perfil ORDER BY per_codigo", vg_db, adOpenStatic
opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = 1: vaSpread1.Value = RS!per_codigo
      vaSpread1.Col = 2: vaSpread1.Lock = opusu: vaSpread1.Value = Trim(RS!per_nombre)
      RS.MoveNext
   Loop
Else
    Gl_Ac_Botones Me, 1, 2, modo
End If
RS.Close: Set RS = Nothing
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
OpGr = False
Frame1.Enabled = True

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") Then GrabaRegistro Row
End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim j As Long
Dim i As Long
Dim opcion As String
If est Then Exit Sub

If Row = 1 Then
    
    vaSpread2.Col = Col
    vaSpread2.Row = -1
    est = True
    vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
    est = False

    'desactiva perfiles y usuario
    For i = 2 To vaSpread2.MaxRows
    
        vaSpread2.Row = i
        vaSpread2.Col = 8
        opcion = vaSpread2.text
        
        If opcion = "4800000" Or opcion = "4810000" Then
        
           For j = 2 To vaSpread2.MaxCols - 1
               
               vaSpread2.Col = j
               est = True
               vaSpread2.Lock = False
               vaSpread2.text = "0"
               vaSpread2.Lock = True
               est = False
           
           Next j
           
        End If
        
    Next i
    
End If

If Col = 2 Then
    
    vaSpread2.Row = Row
    
    For i = 3 To vaSpread2.MaxCols - 1
        
        
        vaSpread2.Col = i
        est = True
        vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
        est = False
    
    Next i

End If

If Row = 1 And Col = 2 Then
    
    For i = 3 To vaSpread2.MaxCols - 1
        
        vaSpread2.Col = i
        vaSpread2.Row = -1
        est = True
        vaSpread2.text = IIf(ButtonDown = 1, "1", "0")
        est = False
    
    
    Next i

    'desactiva perfiles y usuario
    For i = 2 To vaSpread2.MaxRows
    
        vaSpread2.Row = i
        vaSpread2.Col = 8
        opcion = vaSpread2.text
        
        If opcion = "4800000" Or opcion = "4810000" Then
        
           For j = 2 To vaSpread2.MaxCols - 1
               
               vaSpread2.Col = j
               est = True
               vaSpread2.Lock = False
               vaSpread2.text = "0"
               vaSpread2.Lock = True
               est = False
           
           Next j
           
        End If
        
    Next i

End If

SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
vaSpread2.Refresh

End Sub
