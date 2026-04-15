VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CoGA13 
   Caption         =   "Copiar Otros Costo A13"
   ClientHeight    =   5940
   ClientLeft      =   2445
   ClientTop       =   2265
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   5940
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   8655
      Begin VB.Frame Frame1 
         Caption         =   "Datos Destino"
         Height          =   975
         Index           =   1
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   8055
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   1350
            TabIndex        =   8
            Top             =   600
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1658
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
            ButtonStyle     =   1
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
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "11/2017"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/yyyy"
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   0
            IncHour         =   0
            IncMinute       =   0
            IncSecond       =   0
            ButtonColor     =   -2147483637
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   1350
            TabIndex        =   12
            Top             =   225
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
            AutoAdvance     =   0   'False
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
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2745
            Picture         =   "M_CoGA13.frx":0000
            Top             =   120
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3195
            TabIndex        =   14
            Top             =   225
            Width           =   4335
         End
         Begin VB.Label Label3 
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
            Index           =   6
            Left            =   120
            TabIndex        =   13
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   645
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "(mm/aaaa)"
            Height          =   225
            Index           =   1
            Left            =   2340
            TabIndex        =   9
            Top             =   660
            Width           =   840
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   3240
            TabIndex        =   15
            Top             =   270
            Width           =   4335
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Datos Origen"
         Enabled         =   0   'False
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   8055
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   1350
            TabIndex        =   5
            Top             =   360
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1658
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
            ButtonStyle     =   1
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
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   "11/2017"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/yyyy"
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
            ThreeDFrameColor=   -2147483637
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   0
            IncHour         =   0
            IncMinute       =   0
            IncSecond       =   0
            ButtonColor     =   -2147483637
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   405
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "(mm/aaaa)"
            Height          =   225
            Index           =   0
            Left            =   2340
            TabIndex        =   6
            Top             =   420
            Width           =   840
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3495
      Index           =   2
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   10215
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3135
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   9735
         _Version        =   393216
         _ExtentX        =   17171
         _ExtentY        =   5530
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
         MaxCols         =   7
         SpreadDesigner  =   "M_CoGA13.frx":030A
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5940
      Left            =   10365
      TabIndex        =   11
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   10478
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CoGA13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim auxcencos As String, Msgtitulo As String, est As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
fg_carga ""
Msgtitulo = "Copiar Otros Costo A13"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
est = True
fpText1(1).Enabled = ModCasino
Image1(1).Enabled = ModCasino
fpText1(1).text = MuestraCasino(1)
fpayuda(1).Caption = MuestraCasino(2)
vaSpread1.MaxRows = 0
est = False
fg_descarga
End Sub

Private Sub fpText1_Change(Index As Integer)
RS.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(1).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
fpayuda(1).Caption = RS!cli_nombre
RS.Close: Set RS = Nothing
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = ""
Select Case Index
Case 1
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpDateTime1(1).SetFocus
End Select
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    Dim isel As Integer, i As Integer, codigo As Long, descri As String, valor As Double, Orden As Long, ctacon As String
    If Trim(fpayuda(1).Caption) = "" Then MsgBox "Debe seleccionar contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    With vaSpread1
        If .MaxRows < 1 Then Exit Sub
        isel = 0
        For i = 1 To .MaxRows
            .Row = i: .Col = 1: If .text = "1" Then isel = 1: Exit For
        Next i
        If isel = 0 Then MsgBox "Debe seleccionar a lo menor un item", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If MsgBox("Graba A13...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        fg_carga ""
        vg_db.BeginTrans
         For i = 1 To .MaxRows
            .Row = i
            .Col = 1
            If .text = "1" Then
               .Col = 2: codigo = .text
               .Col = 3: descri = .text
               .Col = 4: valor = .text
               .Col = 5: ctacon = .text
               .Col = 7: Orden = .text
               vg_db.Execute "DELETE b_gastosa13 FROM b_gastosa13 WHERE gas_cencos = '" & auxcencos & "' AND gas_anomes = " & Format(fpDateTime1(1).text, "yyyymm") & " AND gas_codigo = " & codigo & ""
               vg_db.Execute "INSERT INTO b_gastosa13 (gas_cencos, gas_anomes, gas_codigo, gas_descri, gas_valor, gas_orden, gas_ctacon) " & _
                            "VALUES ('" & auxcencos & "', " & Format(fpDateTime1(1).text, "yyyymm") & ", " & codigo & ", '" & descri & "', " & valor & ", " & Orden & ", '" & ctacon & "')"
            End If
        Next i
        vg_db.CommitTrans
        fg_descarga
        MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    End With
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub LlenarDatos(cencos As String, Fecha As Long)
With vaSpread1
    .MaxRows = 0
    est = True
    auxcencos = cencos
    fpDateTime1(0).text = Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)
    RS.Open "SELECT b_gastosa13.*, b_gastosa13.gas_cencos, b_gastosa13.gas_anomes, a_ctacontable.cta_nombre " & _
            "FROM b_gastosa13 LEFT JOIN a_ctacontable ON b_gastosa13.gas_ctacon = a_ctacontable.cta_codigo " & _
            "WHERE b_gastosa13.gas_cencos = '" & cencos & "' And b_gastosa13.gas_anomes = " & Fecha & " ORDER BY b_gastosa13.gas_orden", vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = vaSpread1.MaxRows + 1
          .Row = vaSpread1.MaxRows
          .Col = 1: .CellType = 10: .TypeCheckText = "": .TypeCheckCenter = True: .text = ""
          .Col = 2: .text = RS!gas_codigo
          .Col = 3: .text = RS!gas_descri
          .Col = 4: .text = Format(RS!gas_valor, fg_Pict(9, 2))
          .Col = 5: .text = RS!gas_ctacon
          .Col = 6: .text = IIf(IsNull(RS!cta_nombre), "", Trim(RS!cta_nombre))
          .Col = 7: .text = RS!gas_orden
          RS.MoveNext
       Loop
       .SetActiveCell 1, 1
    '   .SetFocus
    End If
    RS.Close: Set RS = Nothing
    est = False
End With
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")
End Sub
