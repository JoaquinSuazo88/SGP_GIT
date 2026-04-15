VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_ArtDie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Productos"
   ClientHeight    =   6495
   ClientLeft      =   2040
   ClientTop       =   1755
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6495
   ScaleWidth      =   8175
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   10821
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      TabCaption(0)   =   "Aporte Dietetico"
      TabPicture(0)   =   "M_ArtDie.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_ArtDie.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   5535
         Index           =   1
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   6255
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   4335
            Left            =   240
            TabIndex        =   10
            Top             =   840
            Width           =   5775
            _Version        =   393216
            _ExtentX        =   10186
            _ExtentY        =   7646
            _StockProps     =   64
            Enabled         =   0   'False
            DisplayRowHeaders=   0   'False
            EditEnterAction =   2
            EditModePermanent=   -1  'True
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
            MaxCols         =   7
            MaxRows         =   30
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_ArtDie.frx":0038
            VisibleCols     =   7
            VisibleRows     =   15
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   435
            TabIndex        =   11
            Top             =   465
            Width           =   45
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   0
         Left            =   -74040
         TabIndex        =   3
         Top             =   480
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "M_ArtDie.frx":0708
            Left            =   1680
            List            =   "M_ArtDie.frx":0712
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   1680
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
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
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
            Caption         =   "Buscar Columna"
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
            Left            =   260
            TabIndex        =   8
            Top             =   340
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Buscar Texto"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   255
            TabIndex        =   7
            Top             =   640
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   4260
            TabIndex        =   6
            Top             =   640
            Width           =   525
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   -74640
         TabIndex        =   1
         Top             =   1560
         Width           =   7185
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3900
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   6765
            _Version        =   393216
            _ExtentX        =   11933
            _ExtentY        =   6879
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
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
            MaxRows         =   20
            OperationMode   =   3
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_ArtDie.frx":0726
            VisibleCols     =   2
            VisibleRows     =   15
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_ArtDie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset
Dim ibusca As Long, codnutriente As Long
Dim valnutriente As Double
Dim i As Integer, j As Integer, itab As Integer
Dim modo As String, incluir As String, alterar As String, eliminar As String, imprimir As String, codigo As String
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()
Me.Height = 6870
Me.Width = 8265
fg_centra Me
SSTab1.Tab = 0: itab = 0
modo = "M"
Mover_Botones
Combo1.ListIndex = 1
ValidarOpcion
Ac_Boton 2
MoverDatosGrilla
End Sub
Private Sub fpText_Change()
If LimpiaDato(Trim(fpText.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 1 Then
   Set ConSql = vg_db.Execute("select count(pro_codigo) as nreg " & _
                "From b_productos " & _
                "Where Ucase(pro_nombre) like '%" + UCase(LimpiaDato(fpText.Text)) + "%' " & _
                "", , adCmdText)
   If ConSql.EOF Or ConSql!NReg = 0 Then ConSql.Close: Set ConSql = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado": SSTab1.TabEnabled(1) = False: Ac_Boton 2: Exit Sub
   If ibusca <> ConSql!NReg Then ibusca = ConSql!NReg: vaSpread1.MaxRows = ConSql!NReg: ConSql.Close: Set ConSql = Nothing
   Set ConSql = vg_db.Execute("select pro_codigo, pro_nombre " & _
                "From b_productos " & _
                "Where Ucase(pro_nombre) like '%" + UCase(LimpiaDato(fpText.Text)) + "%' " & _
                "order by pro_nombre", , adCmdText)
Else
   Set ConSql = vg_db.Execute("select count(pro_codigo) as nreg " & _
                "From b_productos " & _
                "Where pro_codigo like '%" + UCase(LimpiaDato(fpText.Text)) + "%' " & _
                "", , adCmdText)
   If ConSql.EOF Or ConSql!NReg = 0 Then ConSql.Close: Set ConSql = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado": SSTab1.TabEnabled(1) = False: Ac_Boton 2: Exit Sub
   If ibusca <> ConSql!NReg Then ibusca = ConSql!NReg: vaSpread1.MaxRows = ConSql!NReg: ConSql.Close: Set ConSql = Nothing
   Set ConSql = vg_db.Execute("select pro_codigo, pro_nombre " & _
                "From b_productos " & _
                "Where pro_codigo like '%" + UCase(LimpiaDato(fpText.Text)) + "%' " & _
                "order by pro_codigo", , adCmdText)
End If
i = 1
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.Row = i
      i = i + 1

      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Trim(ConSql!pro_codigo)

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(ConSql!pro_nombre)
      ConSql.MoveNext
   Loop
   SSTab1.TabEnabled(1) = True
   Ac_Boton 1
End If
ConSql.Close: Set ConSql = Nothing
If fpText.Text = "" Then
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Else
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Encontrado"
End If
vaSpread1.SetActiveCell 1, 1
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
  Case 1
    If vaSpread1.MaxRows > 0 And itab = 0 Then
       Validar_Nutrientes
       If vaSpread2.MaxRows > 0 Then
          modo = "M"
          SSTab1.TabEnabled(0) = True
          SSTab1.Tab = 1
          SSTab1.TabEnabled(1) = True
          M_ArtDie.Refresh
          Mover_Detalle
       Else
          SSTab1.Tab = 0
       End If
    End If
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    modo = "A": itab = 1
    Validar_Nutrientes
    If vaSpread2.MaxRows < 1 Then Exit Sub
    SSTab1.TabEnabled(0) = False
    SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    Ac_Boton 0
    M_ArtDie.Refresh
    itab = 0
  Case 3
    modo = "M"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Validar_Nutrientes
    If vaSpread2.MaxRows > 0 Then
       itab = 1
       SSTab1.TabEnabled(0) = False
       SSTab1.Tab = 1
       SSTab1.TabEnabled(1) = True
       Ac_Boton 0
       Mover_Detalle
       M_ArtDie.Refresh
       itab = 0
    End If
  Case 5
    Borra_Datos
  Case 7
    MoverDatosGrilla
  Case 10
    Cancela_Datos
  Case 12
    Actualiza_Datos
  Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Producto": Exit Sub
    I_ArtDie.Show 1
  Case 18
    Me.Hide
    Unload Me
End Select
End Sub
Private Sub MoverDatosGrilla()

On Error GoTo Man_Error

fg_carga (ss)
vaSpread1.MaxRows = 0
Set ConSql = vg_db.Execute("select pro_codigo, pro_nombre " & _
             "From b_productos " & _
             "order by pro_nombre", , adCmdText)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
              
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Trim(ConSql!pro_codigo)

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.Value = Trim(ConSql!pro_nombre)
             
      ConSql.MoveNext
   Loop
   Ac_Boton 1
Else
   Ac_Boton 2
End If
ConSql.Close: Set ConSql = Nothing
fpText.Text = ""
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Aporte Productos"
End Sub
Private Sub Validar_Nutrientes()
If vaSpread1.MaxRows < 1 Then Exit Sub
fg_carga (ss)
vaSpread2.MaxRows = 0
Set ConSql = vg_db.Execute("select * " & _
             "From a_nutriente " & _
             "order by nut_codigo", , adCmdText)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      vaSpread2.Col = 1
      vaSpread2.CellType = 5
      vaSpread2.TypeHAlign = 0
      vaSpread2.Text = Trim(ConSql!nut_nombre)
             
      vaSpread2.Col = 2
      vaSpread2.CellType = 2
      vaSpread2.TypeFloatDecimalPlaces = 3
      vaSpread2.TypeFloatMin = "-99999999.999"
      vaSpread2.TypeFloatMax = "99999999.999"
      vaSpread2.TypeFloatMoney = False
      vaSpread2.TypeFloatSeparator = True
      vaSpread2.TypeHAlign = 1
      vaSpread2.TypeFloatCurrencyChar = Asc("$")
      vaSpread2.TypeFloatDecimalChar = Asc(".")
      vaSpread2.TypeFloatSepChar = Asc(",")
      vaSpread2.Text = Format(0, "0.000")
      vaSpread2.ForeColor = &HFF0000
             
      vaSpread2.Col = 3
      vaSpread2.CellType = 5
      vaSpread2.TypeHAlign = 0
      vaSpread2.Text = Trim(ConSql!nut_nomuni)
        
      vaSpread2.Col = 4
      vaSpread2.Text = ConSql!nut_indpri
        
      vaSpread2.Col = 5
      vaSpread2.Text = 0
             
      vaSpread2.Col = 6
      vaSpread2.Text = 0
            
      vaSpread2.Col = 7
      vaSpread2.Text = ConSql!nut_codigo
       
      ConSql.MoveNext
   Loop
Else
   fg_descarga
   MsgBox "No existe Información de nutrientes ", vbCritical + vbOKOnly, "Aporte Productos"
End If
fg_descarga
ConSql.Close: Set ConSql = Nothing
End Sub
Private Sub Mover_Detalle()

j = 0
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.Text
vaSpread1.Col = 2: Label3.Caption = vaSpread1.Text
Set ConSql = vg_db.Execute("select b_productonut.pnu_codapo, b_productonut.pnu_canapo, " & _
             "a_nutriente.nut_secnro, a_nutriente.nut_indpri " & _
             "From a_nutriente, b_productonut " & _
             "where b_productonut.pnu_codapo=a_nutriente.nut_codigo " & _
             "and   b_productonut.pnu_codpro='" & codigo & "' " & _
             "order by a_nutriente.nut_secnro", , adCmdText)
If Not ConSql.EOF Then
   Do While Not ConSql.EOF
      For i = 1 To vaSpread2.MaxRows
          vaSpread2.Row = i
          vaSpread2.Col = 7
          If Val(vaSpread2.Text) = ConSql!pnu_codapo Then
              
             vaSpread2.Col = 2
             vaSpread2.CellType = 2
             vaSpread2.TypeFloatDecimalPlaces = 3
             vaSpread2.TypeFloatMin = "-99999999.999"
             vaSpread2.TypeFloatMax = "99999999.999"
             vaSpread2.TypeFloatMoney = False
             vaSpread2.TypeFloatSeparator = True
             vaSpread2.TypeHAlign = 1
             vaSpread2.TypeFloatCurrencyChar = Asc("$")
             vaSpread2.TypeFloatDecimalChar = Asc(".")
             vaSpread2.TypeFloatSepChar = Asc(",")
             vaSpread2.Text = ConSql!pnu_canapo
             vaSpread2.ForeColor = &HFF0000
              
             vaSpread2.Col = 5
             vaSpread2.Text = ConSql!pnu_canapo
             
             vaSpread2.Col = 6
             vaSpread2.Text = 1
              
             vaSpread2.Col = 7
             vaSpread2.Text = ConSql!pnu_codapo
             j = j + 1
            Exit For
          End If
      Next i
      ConSql.MoveNext
   Loop
End If
ConSql.Close: Set ConSql = Nothing
If vaSpread2.MaxRows > j Then Habdes 0: Ac_Boton 0
vaSpread2.Enabled = True
fg_descarga
End Sub
Private Sub Borra_Datos()

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.Text
TITLE = "Eliminar Dato"
Resp_Delete (TITLE)
If respuesta = vbYes Then
' ***      Borrando Nutriente *** '
   vg_db.BeginTrans
     vg_db.Execute "delete b_productonut from b_productonut where pnu_codpro='" & codigo & "'"
   vg_db.CommitTrans
'   vaSpread1.Row = vaSpread1.ActiveRow
'   vaSpread1.DeleteRows vaSpread1.Row, 1
'   vaSpread1.MaxRows = vaSpread1.MaxRows - 1
'   vaSpread1.Row = vaSpread1.MaxRows
   If vaSpread1.MaxRows < 1 Then
      SSTab1.TabEnabled(1) = False
      SSTab1.Tab = 0
   Else
      SSTab1.TabEnabled(1) = True
      SSTab1.Tab = 0
   End If
End If
fpText.SetFocus

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub Cancela_Datos()
TITLE = "Nutrientes"
msg = "Cancelar Operación"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
Select Case ws_respuesta
  Case Is = vbYes
    SSTab1.TabEnabled(0) = True
    Set ConSql = vg_db.Execute("select count(pro_codigo) as nreg " & _
                 "From b_productos " & _
                 "Where Ucase(pro_nombre) like '%" + UCase(("")) + "%' " & _
                 "", , adCmdText)
    If ConSql.EOF Or ConSql!NReg = 0 Then
       ConSql.Close: Set ConSql = Nothing
       SSTab1.TabEnabled(1) = False
       Ac_Boton 2
       SSTab1.Tab = 0
    ElseIf ConSql!NReg > 0 Then
       ConSql.Close: Set ConSql = Nothing
       If vaSpread1.MaxRows > 0 Then
          SSTab1.TabEnabled(1) = True
       Else
          SSTab1.TabEnabled(1) = False
       End If
       Ac_Boton 1
       SSTab1.Tab = 0
    End If
  Case Is = vbCancel
    Exit Sub
End Select
End Sub
Private Sub Actualiza_Datos()

On Error GoTo Man_Error

If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
   If modo = "A" Then
      vg_db.BeginTrans
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 2: valnutriente = 0: valnutriente = Val(vaSpread2.Text)
            vaSpread2.Col = 7: codnutriente = 0: codnutriente = Val(vaSpread2.Text)
            Set ConSql = vg_db.Execute("select * " & _
                         "from b_productonut " & _
                         "where pnu_codpro='" & codigo & "' " & _
                         "and   pnu_codapo=" & codnutriente & "", , adCmdText)
            If ConSql.EOF Then
               vg_db.Execute "insert into b_productonut (pnu_codpro, pnu_codapo, " & _
                             "pnu_canapo) values ('" & codigo & "', " & codnutriente & ", " & _
                             "" & valnutriente & ")"
            Else
               vg_db.Execute "update b_productonut set pnu_canapo=" & valnutriente & " " & _
                             "where pnu_codpro='" & codigo & "' " & _
                             "and   pnu_codapo=" & codnutriente & ""
            End If
            ConSql.Close: Set ConSql = Nothing
        Next i
      vg_db.CommitTrans
   ElseIf modo = "M" Then
      vg_db.BeginTrans
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = 2: valnutriente = 0: valnutriente = Val(vaSpread2.Text)
            vaSpread2.Col = 7: codnutriente = 0: codnutriente = Val(vaSpread2.Text)
            Set ConSql = vg_db.Execute("select * " & _
                         "from b_productonut " & _
                         "where pnu_codpro='" & codigo & "' " & _
                         "and   pnu_codapo=" & codnutriente & "", , adCmdText)
            If ConSql.EOF Then
               vg_db.Execute "insert into b_productonut (pnu_codpro, pnu_codapo, " & _
                             "pnu_canapo) values ('" & codigo & "', " & codnutriente & ", " & _
                             "" & valnutriente & ")"
            Else
               vg_db.Execute "update b_productonut set pnu_canapo=" & valnutriente & " " & _
                             "where pnu_codpro='" & codigo & "' " & _
                             "and   pnu_codapo=" & codnutriente & ""
            End If
            ConSql.Close: Set ConSql = Nothing
        Next i
      vg_db.CommitTrans
   End If
   
   SSTab1.TabEnabled(0) = True
   If vaSpread1.MaxRows < 1 Then
      SSTab1.TabEnabled(1) = False
   Else
      SSTab1.TabEnabled(1) = True
      SSTab1.Tab = 0
   End If
   Ac_Boton 1
End If

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Function Habdes(Opcion As Integer)
Select Case Opcion
  Case 0
    If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
       SSTab1.TabEnabled(0) = False
       SSTab1.Tab = 1
       SSTab1.TabEnabled(1) = True
    End If
End Select
End Function
Function Ac_Boton(Boton As Integer)
Select Case Boton
  Case 0
    If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
       Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
       Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(4).Visible = True
       Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
       Toolbar1.Buttons(7).Visible = False: Toolbar1.Buttons(8).Visible = True
       Toolbar1.Buttons(10).Visible = True: Toolbar1.Buttons(11).Visible = False
       Toolbar1.Buttons(12).Visible = True: Toolbar1.Buttons(13).Visible = False
       Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = True
    End If
  Case 1
    If incluir = "1" Then
       Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    ElseIf incluir = "0" Then
       Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
    End If
    
    If alterar = "1" Then
       Toolbar1.Buttons(3).Visible = True: Toolbar1.Buttons(4).Visible = False
    ElseIf alterar = "0" Then
       Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(4).Visible = True
    End If
    
    If eliminar = "1" Then
       Toolbar1.Buttons(5).Visible = True: Toolbar1.Buttons(6).Visible = False
    ElseIf eliminar = "0" Then
       Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    End If
    
    Toolbar1.Buttons(7).Visible = True: Toolbar1.Buttons(8).Visible = False
    Toolbar1.Buttons(10).Visible = False: Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False: Toolbar1.Buttons(13).Visible = True
    
    If imprimir = "1" Then
       Toolbar1.Buttons(15).Visible = True: Toolbar1.Buttons(16).Visible = False
    ElseIf imprimir = "0" Then
       Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = True
    End If
    
  Case 2
    If incluir = "1" Then
       Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
    End If
    Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(4).Visible = True
    Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False: Toolbar1.Buttons(8).Visible = True
    Toolbar1.Buttons(10).Visible = False: Toolbar1.Buttons(11).Visible = True
    Toolbar1.Buttons(12).Visible = False: Toolbar1.Buttons(13).Visible = True
'    Toolbar1.Buttons(14).Visible = True
    Toolbar1.Buttons(15).Visible = False: Toolbar1.Buttons(16).Visible = True
End Select
End Function
Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 65 To 90
    fpText.Text = Chr(KeyCode)
    fpText.SetFocus
  Case 48 To 57
    fpText.Text = Chr(KeyCode)
    fpText.SetFocus
  Case 97 To 122
    fpText.Text = Chr(KeyCode)
    fpText.SetFocus
  Case 27 And Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True
    Cancela_Datos
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
    Agrega_Datos
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    If vaSpread1.MaxRows < 1 Then Exit Sub
    Agrega_Datos
  Case 115 And Toolbar1.Buttons(5).Visible = True
    Borra_Datos
End Select
End Sub
Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)
Habdes 0
vaSpread2.Row = Row
vaSpread2.Col = 6
vaSpread2.Text = 1
Ac_Boton 0
End Sub
Private Sub vaSpread2_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27 And Toolbar1.Buttons(10).Visible = True And Toolbar1.Buttons(12).Visible = True
    Cancela_Datos
  Case 113 And Toolbar1.Buttons(1).Visible = True
    modo = "A"
'    Agrega_Datos
  Case 114 And Toolbar1.Buttons(3).Visible = True
    modo = "M"
    If vaSpread1.MaxRows < 1 Then Exit Sub
'    Agrega_Datos
  Case 115 And Toolbar1.Buttons(5).Visible = True
    Borra_Datos
End Select
End Sub
Sub ValidarOpcion()
incluir = "0": alterar = "0": eliminar = "0": imprimir = "0"
incluir = "0": alterar = "1": eliminar = "1": imprimir = "1"
'Set ConSql = vg_db.Execute("select Sdx_UsuCtrlAcceso.* from Sdx_UsuCtrlAcceso, Sdx_Programa " & _
'             "where Sdx_UsuCtrlAcceso.login='" & UCase(vg_NUsr) & "' " & _
'             "and   Sdx_UsuCtrlAcceso.programa=Sdx_Programa.codprograma " & _
'             "and   Sdx_Programa.nomprograma='M_ArtDie'", , adCmdText)
'If Not ConSql.EOF Then
'   incluir = ConSql!incluir
'   alterar = ConSql!alterar
'   eliminar = ConSql!eliminar
'   imprimir = ConSql!imprimir
'End If
'ConSql.Close: Set ConSql = Nothing
End Sub
Sub Mover_Botones()

   Toolbar1.ImageList = Partida.IL1
   Set btnX = Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): btnX.Visible = True: btnX.ToolTipText = "Incluir"
   Set btnX = Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): btnX.Visible = False: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): btnX.Visible = True: btnX.ToolTipText = "Alterar"
   Set btnX = Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): btnX.Visible = False: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): btnX.Visible = True: btnX.ToolTipText = "Borrar "
   Set btnX = Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): btnX.Visible = False: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): btnX.Visible = True: btnX.ToolTipText = "Actualizar Lista   "
   Set btnX = Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): btnX.Visible = False: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): btnX.Visible = False: btnX.ToolTipText = "Cancelar "
   Set btnX = Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): btnX.Visible = True: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = False: btnX.ToolTipText = "Confirmar "
   Set btnX = Toolbar1.Buttons.Add(, "I_Conformar ", , tbrDefault, "I_Confirmar "): btnX.Visible = True: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): btnX.Visible = True: btnX.ToolTipText = "Imprimir "
   Set btnX = Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): btnX.Visible = False: btnX.ToolTipText = ""
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

End Sub

