VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_MovRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mover Recetas"
   ClientHeight    =   5325
   ClientLeft      =   2310
   ClientTop       =   2190
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6495
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1410
         TabIndex        =   5
         Top             =   240
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpText fpayuda 
         Height          =   315
         Index           =   0
         Left            =   2715
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   3645
         _Version        =   196608
         _ExtentX        =   6429
         _ExtentY        =   556
         Enabled         =   0   'False
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
         BackColor       =   -2147483638
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
         ControlType     =   3
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
         Caption         =   "Cat. Dietetica"
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
         Left            =   135
         TabIndex        =   7
         Top             =   300
         Width           =   1185
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2280
         Picture         =   "M_MovRec.frx":0000
         Top             =   165
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6495
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3855
         Left            =   330
         TabIndex        =   1
         Top             =   600
         Width           =   5775
         _Version        =   393216
         _ExtentX        =   10186
         _ExtentY        =   6800
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
         MaxCols         =   3
         MaxRows         =   20
         SelectBlockOptions=   0
         SpreadDesigner  =   "M_MovRec.frx":030A
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   8
         Left            =   1560
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Plato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   360
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5325
      Left            =   6525
      TabIndex        =   8
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   9393
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_MovRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim fillcatdie As Long, imarca As Long, i As Long
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Msgtitulo = "Mover Recetas"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 120
    If Index = 1 Then Image1_Click 0
End Select
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(0).text = "": Exit Sub
RS1.Open "SELECT * FROM a_recetacatdie WHERE car_codigo = " & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(0).text = "": Exit Sub
RS1.Close: Set RS1 = Nothing
fpayuda(0).text = fg_BuscaenArbol(Val(fpLongInteger1(0).Value), "a_recetacatdie", "car_codigo")
End Sub

Private Sub Image1_Click(Index As Integer)
vg_codigo = "": vg_nombre = ""
vg_left = fpayuda(0).Left + 2400
B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
B_ArbEst.Show 1
If vg_codigo = "" Then Exit Sub
fpLongInteger1(0).Value = Val(vg_codigo)
fpayuda(0).text = vg_nombre
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    Dim codreceta As Long
    RS1.Open "SELECT * FROM a_recetacatdie WHERE car_codigo = " & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No Existe Categoria dietetica", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS1.Close: Set RS1 = Nothing
    imarca = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then imarca = 1: Exit For
    Next i
    If imarca = 0 Then MsgBox "Debe Seleccionar A lo menor una Receta", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Mover registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           vaSpread1.SetActiveCell 1, vaSpread1.Row
           codreceta = 0: vaSpread1.Col = 2: codreceta = Val(vaSpread1.text)
           RS1.Open "SELECT rec_codigo FROM b_receta WHERE rec_codigo = " & codreceta & "", vg_db, adOpenStatic
           If Not RS1.EOF Then
              vg_db.Execute "UPDATE b_receta SET rec_catdie=" & Val(fpLongInteger1(0).Value) & " WHERE rec_codigo=" & codreceta & " "
              RS1.Close: Set RS1 = Nothing
              vg_swmovreceta = 1
           Else
              RS1.Close: Set RS1 = Nothing
              MsgBox "Información no puede moverse, puede que halla sido eliminado ó movido de recetario ", vbInformation + vbOKOnly, Msgtitulo
           End If
        End If
    Next i
    If vg_swmovreceta = 1 Then fg_descarga: MsgBox "Proceso Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    vg_db.CommitTrans
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

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
With vaSpread1
    If .MaxRows > 0 And Col = 1 And Row = 0 Then
       If imarca = 0 Then
          For i = 1 To .MaxRows
              .Row = i
              .Col = 1
              .CellType = 10
              .TypeCheckText = ""
              .TypeCheckCenter = True
              .Value = "1" ' checked
          Next i
          imarca = 1
       Else
          For i = 1 To .MaxRows
              .Row = i
              .Col = 1
              .CellType = 10
              .TypeCheckText = ""
              .TypeCheckCenter = True
              .Value = "" ' checked
          Next i
          imarca = 0
       End If
    End If
End With
End Sub

Sub LlenarRecetas(ruttippla As String, filcatdie As Long, filtippla As Long)
fillcatdie = filcatdie
With vaSpread1
    .MaxRows = 0
    RS1.Open "SELECT rec_codigo, rec_nombre FROM b_receta WHERE (rec_catdie = " & filcatdie & " OR " & filcatdie & " = 0) AND (rec_tippla = " & filtippla & " OR " & filtippla & " = 0) AND rec_tiprec = 0 ORDER BY rec_nombre", vg_db, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
         
          .Col = 2
          .CellType = 5
          .TypeHAlign = TypeHAlignRight
          .text = RS1!rec_codigo
          
          .Col = 3
          .CellType = 5
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS1!rec_nombre)
          
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
End With
Label2(8).Caption = ruttippla
End Sub
