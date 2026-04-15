VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_LiPrCa 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Precio Cafetería"
   ClientHeight    =   6000
   ClientLeft      =   2910
   ClientTop       =   3795
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5160
      Left            =   -15
      TabIndex        =   1
      Top             =   390
      Width           =   7800
      _ExtentX        =   13758
      _ExtentY        =   9102
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Articulos de cafetería"
      TabPicture(0)   =   "T_LiPrCa.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Composición..."
      TabPicture(1)   =   "T_LiPrCa.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblNOMBRE(0)"
      Tab(1).Control(1)=   "vaSpread2"
      Tab(1).ControlCount=   2
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4365
         Left            =   -74925
         TabIndex        =   9
         Top             =   705
         Width           =   7620
         _Version        =   393216
         _ExtentX        =   13441
         _ExtentY        =   7699
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
         MaxCols         =   4
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_LiPrCa.frx":0038
         ClipboardOptions=   0
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3645
         Left            =   75
         TabIndex        =   8
         Top             =   1425
         Width           =   7620
         _Version        =   393216
         _ExtentX        =   13441
         _ExtentY        =   6429
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
         MaxCols         =   4
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_LiPrCa.frx":198C
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   825
         TabIndex        =   2
         Top             =   300
         Width           =   6015
         Begin VB.CheckBox Check1 
            Caption         =   "Emitir lista de precios con composición"
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
            Left            =   1275
            TabIndex        =   11
            Top             =   825
            Width           =   4185
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "T_LiPrCa.frx":331D
            Left            =   2175
            List            =   "T_LiPrCa.frx":3327
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   135
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Left            =   2175
            TabIndex        =   4
            Top             =   450
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
            Top             =   195
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
            Top             =   540
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
            Top             =   540
            Width           =   585
         End
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "Descripción"
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
         Left            =   -74220
         TabIndex        =   10
         Top             =   345
         Width           =   5280
      End
   End
End
Attribute VB_Name = "T_LiPrCa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long
Dim Msgtitulo As String
Dim est As Boolean

Private Sub GrabaRegistro(Fila)
Dim codmer As String, candet As Double, codenc As String, codser As Long, nomenc As String, precio As Double, i As Long, j As Long, activo As String
Dim nrorac As Long, sql1 As String
On Error GoTo Man_Error
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codenc = Trim(vaSpread1.Value)
vaSpread1.Col = 2: nomenc = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 3: precio = Format(vaSpread1.Value, fg_Pict(9, 2))
vaSpread1.Col = 4: activo = vaSpread1.Value
If Trim(nomenc) = "" Or precio = 0 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" And SSTab1.Tab = 0 Then
    '-------> ENCABEZADO
    MoverDatosGrillas2
    sql1 = IIf(vg_tipbase = "1", " MAX(val(tpc_codigo)) ", " MAX(convert(int,tpc_codigo)) ")
    RS1.Open "SELECT " & sql1 & " AS tpc_codigo FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then codenc = Trim(Str(TipoDato(RS1!tpc_codigo, 0) + 1)) Else codenc = "1"
    RS1.Close: Set RS1 = Nothing
    vg_db.BeginTrans
    vg_db.Execute "INSERT INTO b_totpreciocaf (tpc_codigo, tpc_nombre, tpc_precio, tpc_cencos, tpc_activo) VALUES ('" & codenc & "', '" & IIf(vg_tipbase = "1", Trim(nomenc), LTrim(nomenc)) & "', " & precio & ", '" & MuestraCasino(1) & "', '" & activo & "')"
    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.Value = codenc
ElseIf modo = "M" And SSTab1.Tab = 0 Then
    '-------> ENCABEZADO
    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_totpreciocaf SET tpc_nombre = '" & Trim(nomenc) & "', tpc_precio = " & precio & ", tpc_activo = '" & activo & "' WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & codenc & "'"
    vg_db.CommitTrans
End If

'-------> DETALLE
If vaSpread2.MaxRows > 0 And SSTab1.Tab = 1 Then
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1: codmer = Trim(vaSpread2.Value)
    vaSpread2.Col = 4: candet = Format(vaSpread2.Value, fg_Pict(9, IIf(vg_pais = "CL", 3, vg_DCa)))
    If candet = 0 Or Trim(codmer) = "" Then
        MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread2.Col = 4: vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow: vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    End If
    vg_db.BeginTrans
    If modo = "A" Then
        vg_db.Execute "INSERT INTO b_detpreciocaf (dpc_codigo, dpc_codmer, dpc_cantidad, dpc_cencos) VALUES ('" & codenc & "', '" & codmer & "', " & candet & ", '" & MuestraCasino(1) & "')"
    Else
        vg_db.Execute "UPDATE b_detpreciocaf SET dpc_cantidad = " & candet & " WHERE dpc_cencos = '" & MuestraCasino(1) & "' AND dpc_codigo = '" & codenc & "' AND dpc_codmer = '" & codmer & "'"
    End If
    vaSpread2.Col = 1: vaSpread2.Value = codmer
    vg_db.CommitTrans
End If

Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(1) = True
OpGr = False
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, Msgtitulo: Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6000
Me.Width = 7800
Msgtitulo = "Lista de Precio Cafetería"
fg_centra Me
est = True: modo = "": ibusca = 0
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
vaSpread1.Row = -1
vaSpread1.Col = 3: vaSpread1.TypeNumberShowSep = True: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = 2
vaSpread2.Col = 4: vaSpread2.TypeNumberShowSep = True: vaSpread2.TypeNumberSeparator = vg_CSep: vaSpread2.TypeNumberDecimal = vg_CDec: vaSpread2.TypeNumberDecPlaces = IIf(vg_pais = "CL", 3, vg_DCa)
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False
est = False
SSTab1.Tab = 0
End Sub

Private Sub fpText1_Change()
Dim sql1 As String
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS2.Open "SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_activo FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo LIKE '%" & UCase(LimpiaDato(fpText1.text)) & "%'", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    sql1 = IIf(vg_tipbase = "1", " UCASE(tpc_nombre) ", " UPPER(tpc_nombre) ")
    RS2.Open "SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_activo FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " LIKE '%" & UCase(LimpiaDato(fpText1.text)) & "%'", vg_db, adOpenStatic
End If
If ibusca <> RS2.RecordCount Then ibusca = RS2.RecordCount: vaSpread1.MaxRows = RS2.RecordCount
i = 1
If Not RS2.EOF Then
   Do While Not RS2.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS2!tpc_codigo
      vaSpread1.Col = 2
      vaSpread1.Value = Trim(RS2!tpc_nombre)
      vaSpread1.Col = 3
      vaSpread1.text = Format(RS2!tpc_precio, fg_Pict(9, 2))
      vaSpread1.Col = 4
      vaSpread1.Value = IIf(IsNull(RS2!tpc_activo) Or RS2!tpc_activo = "0" Or Trim(RS2!tpc_activo) = "", "0", "1")
      RS2.MoveNext
   Loop
   SSTab1.TabEnabled(1) = True
   Gl_Ac_Botones Me, 1, 1, modo
Else
   SSTab1.TabEnabled(1) = False
   Gl_Ac_Botones Me, 1, 2, modo
End If
RS2.Close: Set RS2 = Nothing
If fpText1.text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
    DoEvents
    Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Case 1
    vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(0).Caption = vaSpread1.Value
    MoverDatosGrillas2
    DoEvents
    Gl_Ac_Botones Me, 1, IIf(vaSpread2.MaxRows = 0, 2, 1), modo
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As String, Nombre As String, Orden As String, codpro As String
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If SSTab1.Tab = 0 Then
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.TabEnabled(1) = False
        vaSpread2.MaxRows = 0
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 3: vaSpread1.text = 0
        vaSpread1.Col = 4: vaSpread1.text = "1"
        vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        If vaSpread1.MaxRows < 1 Then Exit Sub
        '-------> Llama  a formulario de busqueda de productos y carga datos
        vg_left = Screen.Width \ 2 - B_TabEst.Width \ 2 'vaSpread2.Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
        B_TabEst.Show 1
        If Val(vg_codigo) = 0 Then Exit Sub
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 1
            If Trim(vg_codigo) = Trim(vaSpread2.text) And Row <> i And Trim(vaSpread2.text) <> "" Then
                MsgBox "Producto ya fué ingresado", vbCritical + vbOKOnly, Msgtitulo
                vaSpread2.SetActiveCell 1, Row: Exit Sub
            End If
        Next i
        RS1.Open "SELECT pro_codigo, pro_nombre, uni_nomcor FROM b_productos, a_unidad WHERE pro_coduni = uni_codigo AND pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then
            modo = "A"
            Gl_Ac_Botones Me, 1, 0, modo
            SSTab1.TabEnabled(0) = False
            vaSpread2.MaxRows = vaSpread2.MaxRows + 1
            vaSpread2.Row = vaSpread2.MaxRows
            vaSpread2.Col = 1: vaSpread2.text = Trim(RS1!pro_codigo)
            vaSpread2.Col = 2: vaSpread2.text = Trim(RS1!pro_nombre)
            vaSpread2.Col = 3: vaSpread2.text = Trim(RS1!uni_nomcor)
            vaSpread2.Col = 4: vaSpread2.text = 0
            If Me.Visible And vaSpread2.Enabled And vaSpread2.Visible Then vaSpread2.SetFocus
            vaSpread2.SetActiveCell 1, vaSpread2.MaxRows: vaSpread2.SetFocus
        Else
            RS1.Close: Set RS1 = Nothing
            MsgBox "Producto no existe", vbCritical + vbOKOnly, Msgtitulo
            vaSpread2.text = "": vaSpread2.SetActiveCell 1, vaSpread2.ActiveRow: Exit Sub
        End If
        RS1.Close: Set RS1 = Nothing
    End If
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False
Case 5
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If SSTab1.Tab = 0 Then
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: codigo = Trim(vaSpread1.Value)
        vg_db.BeginTrans
        vg_db.Execute "DELETE b_totpreciocaf FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & codigo & "'"
        vg_db.CommitTrans
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
        SSTab1.TabEnabled(1) = IIf(vaSpread1.MaxRows = 0, False, True)
        Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
    ElseIf SSTab1.Tab = 1 Then
        If vaSpread2.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        vaSpread2.Row = vaSpread2.ActiveRow: vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread2.Col = 1: codpro = Trim(vaSpread2.Value): vaSpread1.Col = 1: codigo = Trim(vaSpread1.Value)
        vg_db.BeginTrans
        vg_db.Execute "DELETE b_detpreciocaf FROM b_detpreciocaf WHERE dpc_cencos = '" & MuestraCasino(1) & "' AND dpc_codigo = '" & codigo & "' AND dpc_codmer = '" & codpro & "'"
        vg_db.CommitTrans
        vaSpread2.DeleteRows vaSpread2.Row, 1
        vaSpread2.MaxRows = vaSpread2.MaxRows - 1
        modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
    End If
Case 7
    fpText1.text = ""
    If SSTab1.Tab = 0 Then MoverDatosGrillas
    If SSTab1.Tab = 1 Then MoverDatosGrillas2
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(1) = True
    If modo = "A" Then
        If SSTab1.Tab = 0 Then
            MoverDatosGrillas
            modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
            SSTab1.TabEnabled(1) = IIf(vaSpread1.MaxRows = 0, False, True)
            Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
        ElseIf SSTab1.Tab = 1 Then
            MoverDatosGrillas2
            modo = "": Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
        End If
        Combo1.Enabled = True: fpText1.Enabled = True
    Else
        Cancela
    End If
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If SSTab1.Tab = 0 Then
       If Check1.Value = 0 Then
          I_LPCafeteria
       Else
          I_LPCafeteriaDet
       End If
    ElseIf SSTab1.Tab = 1 Then
       vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1
       I_CompLPCafeteria Val(vaSpread1.Value)
    End If
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, Msgtitulo: Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 And modo <> "A" Then
    MoverDatosGrillas2
End If
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(1) = False
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If (Col <> 4) Or Row = 0 Or OpGr Or est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(1) = False
End Sub

Private Sub MoverDatosGrillas()
vaSpread1.MaxRows = 0
RS1.Open "SELECT * FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' ORDER BY tpc_codigo, tpc_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    If vaSpread1.Row = 1 Then MoverDatosGrillas2
    vaSpread1.Col = 1: vaSpread1.Value = Trim(RS1!tpc_codigo)
    vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!tpc_nombre)
    vaSpread1.Col = 3: vaSpread1.text = Format(RS1!tpc_precio, fg_Pict(9, 2))
    vaSpread1.Col = 4: vaSpread1.Value = IIf(IsNull(RS1!tpc_activo) Or Trim(RS1!tpc_activo) = "" Or RS1!tpc_activo = "0", "0", "1")
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows = 0, 2, 1), modo
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
SSTab1.TabEnabled(1) = IIf(vaSpread1.MaxRows = 0, False, True)
End Sub

Private Sub MoverDatosGrillas2()
Dim codigo As String
vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Trim(vaSpread1.Value)
vaSpread2.MaxRows = 0
RS2.Open "SELECT dpc.*, pro.pro_nombre, uni.uni_nomcor FROM b_detpreciocaf dpc, b_productos pro, a_unidad uni WHERE pro.pro_codigo = dpc.dpc_codmer AND pro.pro_coduni = uni.uni_codigo AND dpc_cencos = '" & MuestraCasino(1) & "' AND dpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
Do While Not RS2.EOF
    vaSpread2.MaxRows = vaSpread2.MaxRows + 1
    vaSpread2.Row = vaSpread2.MaxRows
    vaSpread2.Col = 1: vaSpread2.Value = RS2!dpc_codmer
    vaSpread2.Col = 2: vaSpread2.Value = Trim(RS2!pro_nombre)
    vaSpread2.Col = 3: vaSpread2.Value = Trim(RS2!uni_nomcor)
    vaSpread2.Col = 4: vaSpread2.text = Format(RS2!dpc_cantidad, fg_Pict(9, vg_DCa))
    RS2.MoveNext
Loop
RS2.Close: Set RS2 = Nothing
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows = 0 Then Exit Sub
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
   GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
   Cancela
End If
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
Select Case Col
Case 1
    vaSpread2.Row = Row
    vaSpread2.Col = 2: vaSpread2.text = ""
    vaSpread2.Col = 3: vaSpread2.text = ""
    vaSpread2.Col = 4: vaSpread2.text = ""
End Select
End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
Dim codigo As Long, codpro As String
If Row < 1 Then Exit Sub
Select Case Col
Case 1
    If Not OpGr And Col <> NewCol And (modo = "A") And Toolbar1.Buttons(12).Visible = True Then
        vaSpread2.Col = Col: vaSpread2.Row = Row
        If Trim(vaSpread2.text) = "" Then
            Cancel = True
        Else
            vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
            vaSpread2.Col = 1: codpro = Trim(vaSpread2.Value)
            
            For i = 1 To vaSpread2.MaxRows
                vaSpread2.Row = i: vaSpread2.Col = 1
                If Trim(codpro) = Trim(vaSpread2.text) And Row <> i And Trim(vaSpread2.text) <> "" Then
                    Cancel = True
                    vaSpread2.Col = 1: vaSpread2.Row = Row: vaSpread2.text = ""
                    MsgBox "Producto ya fué ingresado", vbCritical + vbOKOnly, Msgtitulo
                    Exit Sub
                End If
            Next i
            RS1.Open "SELECT pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor FROM b_productos pro, a_unidad uni WHERE pro.pro_coduni = uni.uni_codigo AND pro.pro_codigo = '" & codpro & "'", vg_db, adOpenStatic
            If Not RS1.EOF Then
                vaSpread2.Row = Row
                vaSpread2.Col = 1: vaSpread2.Value = Trim(RS1!pro_codigo)
                vaSpread2.Col = 2: vaSpread2.Value = Trim(RS1!pro_nombre)
                vaSpread2.Col = 3: vaSpread2.Value = Trim(RS1!uni_nomcor)
                vaSpread2.SetActiveCell 4, Row
            Else
                Cancel = True
                vaSpread2.Col = 1: vaSpread2.text = ""
                MsgBox "Producto no existe", vbCritical + vbOKOnly, Msgtitulo
            End If
            RS1.Close: Set RS1 = Nothing
        End If
    End If
End Select
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro vaSpread1.ActiveRow
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub Cancela()
Dim codigo As String, codpro As String
If SSTab1.Tab = 0 Then
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = Trim(vaSpread1.Value)
    RS1.Open "SELECT * FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & codigo & "'", vg_db, adOpenStatic
    est = True
    If Not RS1.EOF Then
       vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!tpc_nombre)
       vaSpread1.Col = 3: vaSpread1.text = Format(RS1!tpc_precio, fg_Pict(9, 2))
       vaSpread1.Col = 4: vaSpread1.Value = IIf(IsNull(RS1!tpc_activo) Or Trim(RS1!tpc_activo) = "" Or RS1!tpc_activo = "0", "0", "1")
    End If
    RS1.Close: Set RS1 = Nothing
    est = False: modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
ElseIf SSTab1.Tab = 1 Then
    vaSpread2.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Trim(vaSpread1.Value)
    vaSpread2.Row = vaSpread2.ActiveRow: vaSpread2.Col = 1: codpro = Trim(vaSpread2.Value)
    RS1.Open "SELECT dpc.*, pro.pro_nombre, uni.uni_nomcor FROM b_detpreciocaf dpc, b_productos pro, a_unidad uni WHERE pro.pro_codigo = dpc.dpc_codmer AND pro.pro_coduni = uni.uni_codigo AND dpc.dpc_cencos = '" & MuestraCasino(1) & "' AND dpc.dpc_codigo = '" & codigo & "' AND dpc.dpc_codmer = '" & codpro & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        vaSpread2.Col = 1: vaSpread2.Value = Trim(RS1!dpc_codmer)
        vaSpread2.Col = 2: vaSpread2.Value = Trim(RS1!pro_nombre)
        vaSpread2.Col = 3: vaSpread2.Value = Trim(RS1!uni_nomcor)
        vaSpread2.Col = 4: vaSpread2.text = Format(RS1!dpc_cantidad, fg_Pict(9, vg_DCa))
    End If
    RS1.Close: Set RS1 = Nothing
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
End If
End Sub
