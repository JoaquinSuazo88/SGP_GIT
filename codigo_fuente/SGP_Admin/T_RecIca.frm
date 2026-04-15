VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_RecIca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retención ICA"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   10410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      _ExtentX        =   18362
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
      Width           =   10200
      _ExtentX        =   17992
      _ExtentY        =   11298
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Servicio"
      TabPicture(0)   =   "T_RecIca.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Estructura de Servicio..."
      TabPicture(1)   =   "T_RecIca.frx":001C
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
         Left            =   5010
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
         Left            =   -73065
         TabIndex        =   2
         Top             =   510
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "T_RecIca.frx":0038
            Left            =   2175
            List            =   "T_RecIca.frx":0042
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
            IncHoriz        =   0,25
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
         Height          =   5420
         Left            =   120
         TabIndex        =   8
         Top             =   810
         Width           =   9900
         _Version        =   393216
         _ExtentX        =   17462
         _ExtentY        =   9560
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
         SpreadDesigner  =   "T_RecIca.frx":0056
         ClipboardOptions=   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4575
         Left            =   -73710
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
         SpreadDesigner  =   "T_RecIca.frx":1AA4
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
         Left            =   2220
         TabIndex        =   10
         Top             =   480
         Width           =   5280
      End
   End
End
Attribute VB_Name = "T_RecIca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Public CallForm As String


Private Sub GrabaRegistro(Fila)
Dim codigo As Long, nombre As String
On Error GoTo Man_Error
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread1.Col = 2: nombre = Trim(LimpiaDato(vaSpread1.Value))
If Trim(nombre) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" And SSTab1.Tab = 0 Then
   codigo = 0
   Set RS1 = vg_db.Execute("sgpadm_iu_retencionica 'A', 0, '" & Trim(Mid(nombre, 1, 100)) & "'")
   If Not RS1.EOF Then
      codigo = RS1!indice
      vaSpread1.Col = 1: vaSpread1.Value = codigo
   End If
   RS1.Close: Set RS1 = Nothing
ElseIf modo = "M" And SSTab1.Tab = 0 Then
   vg_db.Execute "sgpadm_iu_retencionica 'M', " & codigo & ", '" & Trim(Mid(nombre, 1, 100)) & "'"
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
       Set RS1 = vg_db.Execute("sgpadm_s_detretencionica 1, " & codmun & "")
       If RS1!.EOF Then
          vg_db.Execute "sgpadm_iu_detretencionica 'A', " & codmun & ", " & porret & ", '" & codcta & "', '" & tipret & "', '" & indret & "'"
       Else
          vg_db.Execute "sgpadm_iu_detretencionica 'M', " & codmun & ", " & porret & ", '" & codcta & "', '" & tipret & "', '" & indret & "'"
       End If
    End If
    vaSpread2.Col = 1: vaSpread2.Value = codmun
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 7395
Me.Width = 10500
Msgtitulo = "Retención ICA"
fg_centra Me
modo = ""
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
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS2!rei_codigo
      vaSpread1.Col = 2: vaSpread1.Value = Trim(RS2!rei_nombre)
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
    Set RS1 = vg_db.Execute("sgpadm_s_retencionica , 1, 0, ''")
    If Not RS1.EOF Then
       Gl_Ac_Botones Me, 1, 1, modo
    Else
       Gl_Ac_Botones Me, 1, 2, modo
    End If
    RS1.Close: Set RS1 = Nothing
    If SSTab1.Tab = 0 Then Exit Sub
    Me.Refresh
    vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(0).Caption = vaSpread1.Value
    MoverDatosGrillas2
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, nombre As String, codmun As String
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
        vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
        vg_db.Execute "DELETE a_retecionica FROM a_retencionica WHERE rei_codigo = " & codigo & ""
        vg_db.Execute "DELETE a_detretecionica FROM a_detretecionica WHERE dri_codigo = " & codigo & ""
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    ElseIf SSTab1.Tab = 1 Then
        vaSpread2.Row = vaSpread2.ActiveRow: vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread2.Col = 1: codigo = Val(vaSpread2.Value): vaSpread1.Col = 1: codmun = vaSpread1.Value
        vg_db.Execute "DELETE a_detretencionica FROM a_detretencionica WHERE dri_codigo = " & codigo & " AND dri_codmun = '" & codmun & "'"
        vaSpread2.DeleteRows vaSpread2.Row, 1
        vaSpread2.MaxRows = vaSpread2.MaxRows - 1
    End If
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 7
    fpText1.text = ""
    If SSTab1.Tab = 0 Then
       MoverDatosGrillas
    ElseIf SSTab1.Tab = 1 Then
       MoverDatosGrillas2
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
'    vaSpread1.Col = 1
'    For i = 1 To vaSpread1.MaxRows
'        vaSpread1.Row = i
'        If vaSpread1.text = lblNOMBRE(0).Caption Then
'           vaSpread1.SetActiveCell 1, vaSpread1.Row
'           Exit For
'        End If
'    Next i
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If SSTab1.Tab = 0 Then
       I_retencionica
    ElseIf SSTab1.Tab = 1 Then
       I_DetalleRetencionica
    End If
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
If vaSpread1.MaxRows > 0 And modo <> "A" Then MoverDatosGrillas2
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
If vg_Indppr = 1 Or vg_Indppr = 2 Then
  vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
  vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
End If
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
End Sub

Private Sub MoverDatosGrillas()

vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If vg_Indppr = "1" Or vg_Indppr = "2" Then
   Set RS1 = vg_db.Execute("sgpadm_s_servicio 5, '', " & vg_Indppr & ", '%%'")
Else
   Set RS1 = vg_db.Execute("sgpadm_s_servicio 7, '', 0, '%%'")
End If
   Do While Not RS1.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      If vaSpread1.Row = 1 Then MoverDatosGrillas2
      vaSpread1.Col = 1: vaSpread1.Value = RS1!ser_codigo
      vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!ser_nombre)
      vaSpread1.Col = 3: vaSpread1.Value = Trim(RS1!ser_orden)
      vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
      vaSpread1.Col = 5: vaSpread1.text = IIf(IsNull(RS1!ser_facturable), "0", Trim(RS1!ser_facturable))
      lisnom = "": liscod = "": cParam = "": encuentra = False
      For j = 1 To UBound(vTipoSer)
          If vTipoSer(j, 1) <> "" Then
             lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoSer(j, 2))
             liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoSer(j, 1)
          End If
      Next j
      vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = lisnom
      vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = liscod
    
      vaSpread1.Col = 7
      codaux = -1
      For z = 0 To vaSpread1.TypeComboBoxCount
          vaSpread1.TypeComboBoxCurSel = z
          If vaSpread1.text = IIf(RS1!ser_indppr = "1", "1", "2") Then codaux = z: Exit For
          codaux = -1
      Next z
      vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = codaux
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   vaSpread1.Visible = True
   Gl_Ac_Botones Me, 1, 1, modo
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
End Sub

Public Sub MoverDatosGrillas2(Optional ByVal servicio As Long)
Dim codigo As Long
vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1
If servicio < 1 Then
    codigo = Val(vaSpread1.Value)
Else
    codigo = servicio
    lblNOMBRE(0).Caption = G_Proc.fg_TraerNombre("a_servicio", "ser_codigo", servicio, "ser_nombre")
    lblNOMBRE(2).Caption = servicio
End If

vaSpread2.MaxRows = 0
Set RS2 = vg_db.Execute("SELECT * FROM a_estservicio WHERE ess_codser = " & codigo & " ORDER BY ess_orden, ess_nombre")
Do While Not RS2.EOF
    vaSpread2.MaxRows = vaSpread2.MaxRows + 1
    vaSpread2.Row = vaSpread2.MaxRows
    vaSpread2.Col = 1: vaSpread2.Value = RS2!ess_codigo
    vaSpread2.Col = 2: vaSpread2.Value = Trim(RS2!ess_nombre)
    vaSpread2.Col = 3: vaSpread2.Value = RS2!ess_orden
    vaSpread2.Col = 4: vaSpread2.Value = RS2!ess_racmin
    RS2.MoveNext
Loop
RS2.Close: Set RS2 = Nothing
End Sub



Private Sub MoverDatosGrillas3()
Dim codigo As Long
vaSpread3.Row = -1: vaSpread3.Col = -1:
vaSpread3.BackColor = &H80000018
vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread3.MaxRows = 0:   vaSpread3.MaxRows = 2
Set RS2 = vg_db.Execute("SELECT * FROM a_serviciorac WHERE sra_codser = " & codigo & " ORDER BY sra_coditem, sra_serdia")
If Not RS2.EOF Then
   Do While Not RS2.EOF
      vaSpread3.Row = RS2!sra_coditem
      vaSpread3.Col = RS2!sra_serdia: vaSpread3.text = IIf(RS2!sra_raciones = 0, "", RS2!sra_raciones)
      RS2.MoveNext
   Loop
End If
vaSpread3.MaxRows = (vaSpread3.MaxRows + 1)
vaSpread3.Row = vaSpread3.MaxRows
vaSpread3.Col = 1
vaSpread3.Col2 = vaSpread3.MaxCols
vaSpread3.Row2 = vaSpread3.MaxRows
vaSpread3.Lock = True
vaSpread3.BlockMode = True
' Lock cells
vaSpread3.Lock = True
' Protect the cells from being edited
vaSpread3.Protect = True
' Turn block mode off
vaSpread3.BlockMode = False
vaSpread3.Col = -1: vaSpread3.BackColor = &HE0E0E0
SumarTotales

RS2.Close: Set RS2 = Nothing
modo = "M"
Gl_Ac_Botones Me, 1, 7, modo
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vg_Indppr = 1 Or vg_Indppr = 2 Then
  vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = "": vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
  vaSpread1.Col = 7:  vaSpread1.TypeComboBoxList = "": vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
End If
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If (Col <> 5) Or Row = 0 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
If vg_Indppr = 1 Or vg_Indppr = 2 Then
  vaSpread1.Col = 6:  vaSpread1.TypeComboBoxList = "": vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
  vaSpread1.Col = 7:  vaSpread1.TypeComboBoxList = "": vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
End If
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Select Case Col
Case 6
    Dim indice As Long
    vaSpread1.Row = Row
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = "": vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
      vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = "": vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
    End If
    vaSpread1.Col = 6: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = indice
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.EditEnterAction = EditEnterActionNone
End Select
End Sub


Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    If Trim(lblNOMBRE(2).Caption) <> "" And Trim(CallForm) = "M_Plami2" Then
        vaSpread1.Col = 1
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            If vaSpread1.text = lblNOMBRE(2).Caption Then
                vaSpread1.SetActiveCell 1, vaSpread1.Row
                Exit For
            End If
        Next i
        
    End If
    GrabaRegistro vaSpread1.ActiveRow
    
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub Cancela()
If SSTab1.Tab = 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    Set RS1 = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo = " & codigo & "")
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
    vaSpread2.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codser = Val(vaSpread1.Value)
    vaSpread2.Row = vaSpread2.ActiveRow: vaSpread2.Col = 1: codigo = Val(vaSpread2.Value)
    Set RS1 = vg_db.Execute("SELECT * FROM a_estservicio WHERE ess_codser = " & codser & " AND ess_codigo = " & codigo & "")
    If Not RS1.EOF Then
       vaSpread2.Col = 2: vaSpread2.Value = Trim(RS1!ess_nombre)
       vaSpread2.Col = 3: vaSpread2.Value = RS1!ess_orden
    End If
    RS1.Close: Set RS1 = Nothing
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
ElseIf SSTab1.Tab = 2 Then
    Me.Refresh
    vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(1).Caption = vaSpread1.Value
    MoverDatosGrillas3
End If
End Sub

Private Sub vaSpread3_EditChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread3.MaxRows < 1 Then Exit Sub
vaSpread3.Row = Row
vaSpread3.Col = Col
If Val(vaSpread3.text) = 0 Then vaSpread3.text = ""
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(1) = False
SumarTotales
End Sub

Private Sub vaSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread3.MaxRows < 1 Then Exit Sub
vaSpread3.Row = Row
vaSpread3.Col = Col
If Val(vaSpread3.text) = 0 Then vaSpread3.text = ""
End Sub

Sub SumarTotales()
Dim i As Long, j As Long, nrorac As Long
For j = 1 To vaSpread3.MaxCols
    vaSpread3.Row = vaSpread3.MaxRows
    vaSpread3.Col = j: vaSpread3.text = ""
Next j
For i = 1 To (vaSpread3.MaxRows - 1)
    nrorac = 0
    For j = 1 To vaSpread3.MaxCols
        vaSpread3.Row = i
        vaSpread3.Col = j: nrorac = Val(vaSpread3.Value)
        vaSpread3.Row = vaSpread3.MaxRows
        If nrorac > 0 Then vaSpread3.Col = j: vaSpread3.Value = (Val(vaSpread3.Value) + nrorac)
    Next j
Next i
End Sub
