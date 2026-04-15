VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_TipDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Documento"
   ClientHeight    =   5850
   ClientLeft      =   2145
   ClientTop       =   1980
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5850
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   9128
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tipo Documento"
      TabPicture(0)   =   "T_TipDoc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Clase Documento SAP"
      TabPicture(1)   =   "T_TipDoc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   465
         TabIndex        =   3
         Top             =   600
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "T_TipDoc.frx":0038
            Left            =   2085
            List            =   "T_TipDoc.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2415
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Left            =   2085
            TabIndex        =   5
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
            Left            =   540
            TabIndex        =   8
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
            Left            =   540
            TabIndex        =   7
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
            Left            =   4665
            TabIndex        =   6
            Top             =   645
            Width           =   585
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4335
         Left            =   -74760
         TabIndex        =   1
         Top             =   600
         Width           =   6615
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3615
            Left            =   140
            TabIndex        =   2
            Top             =   360
            Width           =   6285
            _Version        =   393216
            _ExtentX        =   11077
            _ExtentY        =   6376
            _StockProps     =   64
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
            MaxCols         =   3
            MaxRows         =   20
            SpreadDesigner  =   "T_TipDoc.frx":0056
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3255
         Left            =   360
         TabIndex        =   9
         Top             =   1680
         Width           =   6495
         _Version        =   393216
         _ExtentX        =   11456
         _ExtentY        =   5741
         _StockProps     =   64
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
         MaxCols         =   4
         MaxRows         =   20
         SpreadDesigner  =   "T_TipDoc.frx":0436
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_TipDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean

Private Sub GrabaRegistro(Fila As Long)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim i  As Long
Dim codigo As String, Nombre As String, cladoc As String, cdosap As String
Dim Orden As Long, codreg As Long
OpGr = True
With vaSpread1
    .Row = Fila
    .Col = 1: codigo = Trim(.text)
    .Col = 2: Nombre = Trim(.text)
    .Col = 3: cladoc = Trim(.text)
    .Col = 4: Orden = Val(.text)
    If codigo = "" Or Trim(Nombre) = "" Or Orden < 1 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: .Row = Fila: .Col = 2: .SetActiveCell 2, .Row: .SetFocus: OpGr = False: Exit Sub
    If modo = "A" Then
        RS.Open RutinaLectura.TipoDocumento(1, codigo, ""), vg_db, adOpenStatic
        If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Tipo documento existe", vbExclamation + vbOKOnly, "Tipo Documento": Exit Sub
        RS.Close: Set RS = Nothing
        vg_db.Execute "INSERT INTO a_tipodocumento (tdo_codigo, tdo_nombre, tdo_cladoc, tdo_orden) VALUES ('" & codigo & "', '" & Trim(Nombre) & "', '" & Trim(cladoc) & "', " & Orden & ")"
        .Col = 1: .Lock = True: .Value = codigo
    Else
        Select Case SSTab1.Tab
        Case 0
            vg_db.Execute "UPDATE a_tipodocumento SET tdo_nombre = '" & Trim(Nombre) & "', tdo_cladoc = '" & Trim(cladoc) & "', tdo_orden  = " & Orden & " WHERE  tdo_codigo = '" & codigo & "'"
        Case 1
           vg_db.Execute "DELETE a_clasedocsap WHERE cds_coddoc = '" & codigo & "'"
           With vaSpread2
                For i = 1 To .MaxRows
                    .Row = i
                    .Col = 1
                    codreg = Val(.text)
                    .Col = 3
                    cdosap = Trim(.text)
                    If Trim(cdosap) <> "" Then
                       vg_db.Execute "INSERT INTO a_clasedocsap (cds_coddoc, cds_codreg, cds_cdosap) VALUES ('" & codigo & "', " & codreg & ", '" & cdosap & "')"
                    End If
                Next i
           End With
        End Select
    End If
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
End With
Combo1.Enabled = True: fpText1.Enabled = True
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
modo = "": Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
SSTab1.Tab = 0
Me.Height = 6255
Me.Width = 7470
MsgTitulo = "Tipo Documento"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
'-------> ocultar pestaña tipo documento sap
If vg_pais = "CL" Then
   SSTab1.TabVisible(1) = False
ElseIf vg_pais = "CO" Then
   vaSpread1.Row = 0
   vaSpread1.Col = 3
   vaSpread1.ColHidden = True
   vaSpread1.Col = 2
   vaSpread1.ColWidth(2) = 38.38
End If
MoverDatosGrillas
MoverTipoDocSAP
OpGr = False
End Sub

Private Sub Form_Resize()
'If Me.WindowState = 0 Then
'   Frame1.Move 0, 360, 6015, 971
'   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
'ElseIf Me.WindowState = 2 Then
'   Frame1.Move 4200, 360, 6015, 971
'   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440
'End If
Toolbar1.Refresh
End Sub

Private Sub fpText1_Change()
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
Dim RS As New ADODB.Recordset
vaSpread1.Visible = False
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   RS.Open RutinaLectura.TipoDocumento(2, "", UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   RS.Open RutinaLectura.TipoDocumento(3, "", UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
ibusca = RS.RecordCount: vaSpread1.MaxRows = RS.RecordCount: i = 1
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = Trim(RS!tdo_codigo)
      vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = Trim(RS!tdo_nombre)
      vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = Trim(RS!tdo_cladoc)
      vaSpread1.Col = 4: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.text = IIf(IsNull(RS!tdo_orden), "", Trim(RS!tdo_orden))
      RS.MoveNext
   Loop
   Gl_Ac_Botones Me, 1, 1, modo
End If
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
If fpText1.text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, IIf(vaSpread1.MaxRows > 0, 1, 2), 6), modo
If vg_modprod = False Then
   vaSpread1.Col = 1: vaSpread1.Col2 = vaSpread1.MaxCols: vaSpread1.Row = 1: vaSpread1.Row2 = vaSpread1.MaxRows
   vaSpread1.BlockMode = True: vaSpread1.Lock = True: vaSpread1.Protect = True: vaSpread1.BlockMode = False
   vaSpread2.Col = 1: vaSpread2.Col2 = vaSpread2.MaxCols: vaSpread2.Row = 1: vaSpread2.Row2 = vaSpread2.MaxRows
   vaSpread2.BlockMode = True: vaSpread2.Lock = True: vaSpread2.Protect = True: vaSpread2.BlockMode = False
End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
    Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, IIf(vaSpread1.MaxRows > 0, 1, 2), 6), modo
'    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 1
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As String, Nombre As String, cladoc As String, codreg As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = False
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = 1: vaSpread1.Lock = False: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    Select Case SSTab1.Tab
    Case 0
       SSTab1.TabEnabled(0) = True
       SSTab1.TabEnabled(1) = False
    Case 1
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = True
    End Select
Case 5
    MsgTitulo = IIf(SSTab1.Tab = 0, "Tipo Documento", "Clase Documento SAP")
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignLeft: codigo = vaSpread1.text
    vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignLeft: cladoc = Trim(LimpiaDato(vaSpread1.text))
    Select Case SSTab1.Tab
    Case 0
        vg_db.Execute "DELETE a_tipodocumento FROM a_tipodocumento WHERE tdo_codigo = '" & codigo & "'"
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Case 1
        vaSpread2.Row = vaSpread2.ActiveRow
        vaSpread2.Col = 1
        codreg = Val(vaSpread2.text)
        vg_db.Execute "DELETE a_clasedocsap FROM a_clasedocsap WHERE cds_coddoc = '" & codigo & "' and cds_codreg = " & codreg & ""
        MoverTipoDocSAP
    End Select
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    MsgTitulo = "Tipo Documento"
Case 7
    fpText1.text = ""
    Select Case SSTab1.Tab
    Case 0
        MoverDatosGrillas
    Case 1
        MoverTipoDocSAP
    End Select
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Else
        Cancela
    End If
    SSTab1.TabEnabled(0) = True
    If vaSpread1.MaxRows > 0 Then SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    GrabaRegistro vaSpread1.ActiveRow
    Select Case SSTab1.Tab
    Case 0
    Case 1
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
    End Select
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Select Case SSTab1.Tab
    Case 0
        I_TipoDocumento
    Case 1
        I_ClaseDocumentoSAP
    End Select
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 And Row > 0 And vg_pais = "CO" Then MoverTipoDocSAP
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
SSTab1.TabEnabled(1) = False
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub MoverDatosGrillas()
Dim RS As New ADODB.Recordset
With vaSpread1
    .Visible = False
    .MaxRows = 0
    RS.Open RutinaLectura.TipoDocumento(1, "", ""), vg_db, adOpenStatic
    Do While Not RS.EOF
        .MaxRows = vaSpread1.MaxRows + 1
        .Row = .MaxRows
        .Col = 1: .TypeHAlign = TypeHAlignLeft: .Lock = True: .text = Trim(RS!tdo_codigo)
        .Col = 2: .TypeHAlign = TypeHAlignLeft: .Lock = False: .text = Trim(IIf(IsNull(RS!tdo_nombre), "", RS!tdo_nombre))
        .Col = 3: .TypeHAlign = TypeHAlignLeft: .Lock = False: .text = Trim(IIf(IsNull(RS!tdo_cladoc), "", RS!tdo_cladoc))
        .Col = 4: .TypeHAlign = TypeHAlignRight: .Lock = False: .text = IIf(IsNull(RS!tdo_orden), "", Trim(RS!tdo_orden))
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    .Visible = True
    Gl_Ac_Botones Me, 1, 1, modo
    Label2.Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, IIf(.MaxRows > 0, 1, 2), 6), modo
    If vg_modprod = False Then
       .Col = 1: .Col2 = vaSpread1.MaxCols: .Row = 1: vaSpread1.Row2 = .MaxRows
       .BlockMode = True: .Lock = True: .Protect = True: .BlockMode = False
       vaSpread2.Col = 1: vaSpread2.Col2 = vaSpread2.MaxCols: vaSpread2.Row = 1: vaSpread2.Row2 = vaSpread2.MaxRows
       vaSpread2.BlockMode = True: vaSpread2.Lock = True: vaSpread2.Protect = True: vaSpread2.BlockMode = False
    End If
End With
End Sub

Private Sub MoverTipoDocSAP()
If vaSpread1.MaxRows < 1 Then Exit Sub
Dim RS As New ADODB.Recordset
Dim coddoc As String
Dim Nombre As String
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
coddoc = vaSpread1.text
vaSpread1.Col = 2
Nombre = vaSpread1.text
vaSpread2.MaxRows = 0
Frame2.Caption = "(" & coddoc & ") " & Trim(Nombre)
RS.Open RutinaLectura.Region(1, coddoc, ""), vg_db, adOpenStatic
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.text = RS!reg_codigo
   vaSpread2.Col = 2
   vaSpread2.text = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
   vaSpread2.Col = 3
   vaSpread2.text = IIf(IsNull(RS!cds_cdosap), "", RS!cds_cdosap)
   RS.MoveNext
Loop
Gl_Ac_Botones Me, 1, IIf(vg_modprod = True, IIf(vaSpread1.MaxRows > 0, 1, 2), 6), modo
If vg_modprod = False Then
   vaSpread1.Col = 1: vaSpread1.Col2 = vaSpread1.MaxCols: vaSpread1.Row = 1: vaSpread1.Row2 = vaSpread1.MaxRows
   vaSpread1.BlockMode = True: vaSpread1.Lock = True: vaSpread1.Protect = True: vaSpread1.BlockMode = False
   vaSpread2.Col = 1: vaSpread2.Col2 = vaSpread2.MaxCols: vaSpread2.Row = 1: vaSpread2.Row2 = vaSpread2.MaxRows
   vaSpread2.BlockMode = True: vaSpread2.Lock = True: vaSpread2.Protect = True: vaSpread2.BlockMode = False
End If
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub Cancela()
Dim RS As New ADODB.Recordset
Dim codigo As String
OpGr = True
With vaSpread1
    .Row = .ActiveRow
    .Col = 1
    codigo = .text
    Select Case SSTab1.Tab
    Case 0
        RS.Open RutinaLectura.TipoDocumento(1, codigo, ""), vg_db, adOpenStatic
        If Not RS.EOF Then
           .Col = 2: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS!tdo_nombre)
           .Col = 3: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS!tdo_cladoc)
           .Col = 4: .TypeHAlign = TypeHAlignRight: .text = IIf(IsNull(RS!tdo_orden), "", Trim(RS!tdo_orden))
        End If
        RS.Close: Set RS = Nothing
    Case 1
        MoverTipoDocSAP
    End Select
End With
OpGr = False
End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
End Sub
