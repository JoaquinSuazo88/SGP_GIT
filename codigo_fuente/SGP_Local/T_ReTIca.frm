VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_RetIca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retención ICA"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6915
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6405
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   11520
      _ExtentX        =   20320
      _ExtentY        =   11298
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Retención Ica"
      TabPicture(0)   =   "T_ReTIca.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle Retención Ica"
      TabPicture(1)   =   "T_ReTIca.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblNOMBRE(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "vaSpread2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1420
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -72345
         TabIndex        =   2
         Top             =   510
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "T_ReTIca.frx":0038
            Left            =   2175
            List            =   "T_ReTIca.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Left            =   2175
            TabIndex        =   4
            Top             =   555
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
            Index           =   0
            Left            =   660
            TabIndex        =   7
            Top             =   300
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
            Left            =   660
            TabIndex        =   6
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
            Left            =   4755
            TabIndex        =   5
            Top             =   645
            Width           =   585
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5415
         Left            =   120
         TabIndex        =   8
         Top             =   810
         Width           =   11220
         _Version        =   393216
         _ExtentX        =   19791
         _ExtentY        =   9551
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
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
         FormulaSync     =   0   'False
         MaxCols         =   7
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_ReTIca.frx":0056
         ClipboardOptions=   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4575
         Left            =   -72990
         TabIndex        =   9
         Top             =   1620
         Width           =   7245
         _Version        =   393216
         _ExtentX        =   12779
         _ExtentY        =   8070
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         ButtonDrawMode  =   1
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
         FormulaSync     =   0   'False
         MaxCols         =   2
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_ReTIca.frx":1AB6
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   0
         Left            =   3300
         TabIndex        =   10
         Top             =   480
         Width           =   5280
      End
   End
End
Attribute VB_Name = "T_RetIca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean, irow As Long, itop As Long

Private Sub GrabaRegistro(Fila)
Dim Codigo As Long, Nombre As String
On Error GoTo Man_Error
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.Value))
If Trim(Nombre) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" And SSTab1.Tab = 0 Then
   Codigo = 0
   Set RS1 = vg_db.Execute("sgpadm_iu_retencionica 'A', 0, '" & Trim(Mid(Nombre, 1, 100)) & "'")
   If Not RS1.EOF Then
      Codigo = RS1!indice
      vaSpread1.Col = 1: vaSpread1.Value = Codigo
   End If
   RS1.Close: Set RS1 = Nothing
ElseIf modo = "M" And SSTab1.Tab = 0 Then
   vg_db.Execute "sgpadm_iu_retencionica 'M', " & Codigo & ", '" & Trim(Mid(Nombre, 1, 100)) & "'"
End If
Dim codmun As Long, porret As Double, codcta As String, tipret As String, indret As String
'------> DETALLE
If vaSpread2.MaxRows > 0 And SSTab1.Tab = 1 Then
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1: codmun = Val(vaSpread2.Value)
    vaSpread2.Col = 3: porret = vaSpread2.Value
    vaSpread2.Col = 4: codcta = Trim(LimpiaDato(vaSpread2.Value))
    vaSpread2.Col = 6: tipret = Trim(LimpiaDato(vaSpread2.Value))
    vaSpread2.Col = 7: indret = Trim(LimpiaDato(vaSpread2.Value))
    If porret = 0 Then
        MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread2.Col = 3: vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow: vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    End If
    If modo = "M" Then
       Set RS1 = vg_db.Execute("sgpadm_s_detretencionica 2, " & Codigo & ", " & codmun & "")
       If RS1.EOF Then
          vg_db.Execute "sgpadm_iu_detretencionica 'A', " & Codigo & ", " & codmun & ", " & porret & ", '" & codcta & "', '" & tipret & "', '" & indret & "'"
       Else
          vg_db.Execute "sgpadm_iu_detretencionica 'M', " & Codigo & ", " & codmun & ", " & porret & ", '" & codcta & "', '" & tipret & "', '" & indret & "'"
       End If
       RS1.Close: Set RS1 = Nothing
    End If
    vaSpread2.Col = 1: vaSpread2.Value = codmun
End If
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
If SSTab1.Tab = 0 Then Exit Sub
If Toolbar1.Buttons(1).Visible = True Then Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Command1_Click()
vg_left = Command1.Left + 3801
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cta. Contable", "Gen"
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then vaSpread2.Col = 4: vaSpread2.Row = irow: vaSpread2.SetActiveCell 4, irow: vaSpread2.EditMode = True: vaSpread2.EditModeReplace = True: vaSpread2.SetFocus: Exit Sub
vaSpread2.Row = irow
vaSpread2.Col = 4
vaSpread2.text = vg_codigo
vaSpread2.Col = 5
vaSpread2.text = vg_nombre
If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 7395
Me.Width = 11835
Msgtitulo = "Retención ICA"
fg_centra Me
modo = "": itop = 1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
SSTab1.Tab = 0
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
vaSpread1.Visible = False
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   Set RS2 = vg_db.Execute("sgpadm_s_retencionica 3, 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   Set RS2 = vg_db.Execute("sgpadm_s_retencionica 4, 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
End If
If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
i = 1
If Not RS2.EOF Then
   Do While Not RS2.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.Value = RS2!rei_codigo
      vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.Value = Trim(RS2!rei_nombre)
      RS2.MoveNext
   Loop
   SSTab1.TabEnabled(1) = True
   Gl_Ac_Botones Me, 1, 1, modo
Else
   SSTab1.TabEnabled(1) = False
   Gl_Ac_Botones Me, 1, 2, modo
End If
RS2.Close: Set RS2 = Nothing
If fpText1.text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
vaSpread1.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0, 1
    Set RS1 = vg_db.Execute("sgpadm_s_retencionica 2, 0, ''")
    If Not RS1.EOF Then
       Gl_Ac_Botones Me, 1, 1, modo
    Else
       Gl_Ac_Botones Me, 1, 2, modo
    End If
    RS1.Close: Set RS1 = Nothing
    If SSTab1.Tab = 0 Then Exit Sub
    If Toolbar1.Buttons(1).Visible = True Then Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
    Me.Refresh
    lblNOMBRE(0).Caption = ""
    vaSpread1.Col = 1: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(0).Caption = "(" & vaSpread1.Value & ") - "
    vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(0).Caption = lblNOMBRE(0).Caption & vaSpread1.Value
    vaSpread1.Col = 1
    MoverDatosGrillas2 Val(vaSpread1.text)
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Codigo As Long, Nombre As String, codmun As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False
    SSTab1.TabEnabled(1) = False
    vaSpread2.MaxRows = 0
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    If SSTab1.Tab = 0 Then
       SSTab1.TabEnabled(1) = False
    ElseIf SSTab1.Tab = 1 Then
       SSTab1.TabEnabled(0) = False
    End If
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If SSTab1.Tab = 0 Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
        vg_db.Execute "DELETE b_detretencionica FROM b_detretencionica WHERE dri_codigo = " & Codigo & ""
        vg_db.Execute "DELETE b_retencionica FROM b_retencionica WHERE rei_codigo = " & Codigo & ""
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    ElseIf SSTab1.Tab = 1 Then
        vaSpread2.Row = vaSpread2.ActiveRow: vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread2.Col = 1: codmun = Val(vaSpread2.Value): vaSpread1.Col = 1: Codigo = vaSpread1.Value
        vg_db.Execute "DELETE b_detretencionica FROM b_detretencionica WHERE dri_codigo = " & Codigo & " AND dri_codmun = '" & codmun & "'"
        vaSpread2.DeleteRows vaSpread2.Row, 1
        vaSpread2.MaxRows = vaSpread2.MaxRows - 1
    End If
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 7
    fpText1.text = ""
    If SSTab1.Tab = 0 Then
       MoverDatosGrillas
    ElseIf SSTab1.Tab = 1 Then
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1
       MoverDatosGrillas2 Val(vaSpread1.text)
    End If
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(1) = True
    If modo = "A" Then
       MoverDatosGrillas
    ElseIf modo = "Cancel" Then
       MoverDatosGrillas2
       modo = ""
       Cancela
    Else
       Cancela
    End If
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_RetencionIca
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 And modo <> "A" Then vaSpread1.Row = Row: vaSpread1.Col = 1: MoverDatosGrillas2 Val(vaSpread1.text)
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(1) = False
End Sub

Private Sub MoverDatosGrillas()
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
Set RS1 = vg_db.Execute("sgpadm_s_retencionica 2, 0, ''")
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      If vaSpread1.Row = 1 Then MoverDatosGrillas2 RS1!rei_codigo
      vaSpread1.Col = 1: vaSpread1.Value = RS1!rei_codigo
      vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!rei_nombre)
      RS1.MoveNext
   Loop
   SSTab1.TabEnabled(1) = True
Else
   SSTab1.TabEnabled(1) = False
End If
RS1.Close: Set RS1 = Nothing
vaSpread1.Visible = True
Gl_Ac_Botones Me, 1, 1, modo
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
End Sub

Public Sub MoverDatosGrillas2(Optional ByVal Codigo As Long)
vaSpread2.MaxRows = 0
itop = 1
Set RS2 = vg_db.Execute("sgpadm_s_detretencionica 1, " & Codigo & ", 0")
Do While Not RS2.EOF
    vaSpread2.MaxRows = vaSpread2.MaxRows + 1
    vaSpread2.Row = vaSpread2.MaxRows
    vaSpread2.Col = 1: vaSpread2.Value = RS2!mun_codigo
    vaSpread2.Col = 2: vaSpread2.Value = IIf(IsNull(RS2!mun_nombre), "", Trim(RS2!mun_nombre))
    vaSpread2.Col = 3: vaSpread2.Value = RS2!dri_portar
    vaSpread2.Col = 4: vaSpread2.Value = RS2!dri_codcta
    vaSpread2.Col = 5: vaSpread2.Value = IIf(IsNull(RS2!cta_nombre), "", RS2!cta_nombre)
    vaSpread2.Col = 6: vaSpread2.Value = IIf(IsNull(RS2!dri_tipret) Or Trim(RS2!dri_tipret) = "", "", RS2!dri_tipret)
    vaSpread2.Col = 7: vaSpread2.Value = IIf(IsNull(RS2!dri_indret) Or Trim(RS2!dri_indret) = "", "", RS2!dri_indret)
    RS2.MoveNext
Loop
RS2.Close: Set RS2 = Nothing
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If (Col <> 5) Or Row = 0 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case Is <> 4
    Command1.Visible = False
Case 4
    Command1.Top = IIf(Row = 1, 1420, 1420 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    irow = Row
    vaSpread2.Col = 4
    vaSpread2.TypeHAlign = TypeHAlignLeft
End Select
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
irow = Row
Command1.Top = IIf(Row = 1, 1420, 1420 + (240 * (Row - itop)))
Command1.Visible = True
If ChangeMade = False And Col <> 6 Then
   If Col <> 4 Then Command1.Visible = False
   Exit Sub
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
Select Case Col
Case Is <> 4
    Command1.Visible = False
Case 4
    Command1.Top = IIf(Row = 1, 1420, 1420 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread2.Row = Row
    vaSpread2.Col = Col
    RS2.Open "SELECT cta_nombre FROM a_ctacontable WHERE cta_codigo = '" & vaSpread2.Value & "'", vg_db, adOpenStatic
    If RS2.EOF Then RS2.Close: Set RS2 = Nothing: vaSpread2.text = "": vaSpread2.Col = 5: vaSpread2.text = "": Exit Sub
    vaSpread2.Col = 5: vaSpread2.text = Trim(RS2!cta_nombre)
    RS2.Close: Set RS2 = Nothing
    Command1.Visible = False
End Select
End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro vaSpread1.ActiveRow
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub vaSpread2_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
itop = NewTop
Command1.Visible = False
End Sub

Private Sub Cancela()
Dim codmun As Long, Codigo As Long
If SSTab1.Tab = 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
    Set RS1 = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & Codigo & "")
    If Not RS1.EOF Then
       vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!ser_nombre)
       vaSpread1.Col = 3: vaSpread1.Value = Trim(RS1!ser_orden)
       vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
       vaSpread1.Col = 5: vaSpread1.text = IIf(IsNull(RS1!ser_facturable), "0", Trim(RS1!ser_facturable))
    End If
    RS1.Close: Set RS1 = Nothing
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
ElseIf SSTab1.Tab = 1 Then
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: Codigo = Val(vaSpread1.Value)
    vaSpread2.Row = vaSpread2.ActiveRow: vaSpread2.Col = 1: codmun = Val(vaSpread2.Value)
    MoverDatosGrillas2 Codigo
    Set RS1 = vg_db.Execute("sgpadm_s_detretencionica 1, " & Codigo & ", 0")
'    If Not RS1.EOF Then
'       vaSpread2.Col = 2: vaSpread2.Value = Trim(RS1!ess_nombre)
'       vaSpread2.Col = 3: vaSpread2.Value = RS1!ess_orden
'    End If
'    RS1.Close: Set RS1 = Nothing
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
    If SSTab1.Tab = 0 Then Exit Sub
    If Toolbar1.Buttons(1).Visible = True Then Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
End If
End Sub
