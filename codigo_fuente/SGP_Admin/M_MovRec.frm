VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_MovRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mover Recetas"
   ClientHeight    =   6420
   ClientLeft      =   2310
   ClientTop       =   2190
   ClientWidth     =   7335
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   675
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6735
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2730
         TabIndex        =   8
         Top             =   240
         Width           =   3840
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
         TabIndex        =   6
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
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2775
         TabIndex        =   9
         Top             =   285
         Width           =   3840
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   6735
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4795
         Left            =   210
         TabIndex        =   1
         Top             =   480
         Width           =   6340
         _Version        =   393216
         _ExtentX        =   11183
         _ExtentY        =   8458
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
         MaxCols         =   4
         MaxRows         =   20
         SelectBlockOptions=   0
         SpreadDesigner  =   "M_MovRec.frx":030A
         ScrollBarTrack  =   3
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Recetas No Vigentes"
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   11
         Top             =   5400
         Width           =   1515
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00D9D9FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1320
         Top             =   5430
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Recetas Vigentes"
         Height          =   195
         Index           =   0
         Left            =   3825
         TabIndex        =   10
         Top             =   5400
         Width           =   1260
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   3465
         Top             =   5430
         Width           =   300
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   5055
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
         Height          =   195
         Index           =   12
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6420
      Left            =   6795
      TabIndex        =   7
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   11324
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
Dim MsgTitulo As String

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

fg_centra Me
MsgTitulo = "Mover Recetas"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

Dim RS1 As New ADODB.Recordset

If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(0).Caption = "": Exit Sub

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT * FROM a_recetacatdie with (nolock) WHERE car_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic

If RS1.EOF Then

   RS1.Close
   Set RS1 = Nothing
   fpayuda(0).Caption = ""
   Exit Sub

End If

RS1.Close
Set RS1 = Nothing

fpayuda(0).Caption = fg_BuscaenArbol(Val(fpLongInteger1(0).Value), "a_recetacatdie", "car_codigo")

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

Private Sub Image1_Click(Index As Integer)

vg_codigo = "": vg_nombre = ""
vg_left = fpayuda(0).Left + 2400
B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
B_ArbEst.Show 1
If vg_codigo = "" Then Exit Sub
fpLongInteger1(0).Value = Val(vg_codigo)
fpayuda(0).Caption = Trim(vg_nombre)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

Select Case Button.Index

Case 1
    
    Dim CodReceta As Long
    Dim coddi1 As Long, coddi2 As Long, codti1 As Long, codti2 As Long, codti3 As Long, tippla As Long
    Dim StrFamb As String, StrFam As String, nomrec As String, nomfan As String
    '------- Validar categoria dietetica
'    StrFam = fg_BuscaCodArbol(Val(fpLongInteger1(0).Value), "a_recetacatdie", "car_codigo")
'    If Len(StrFam) <> 0 Then
'
'       Do While InStr(StrFam, ";") <> 0
'
'          StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
'          StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
'          coddi1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
'          If Val(Mid(StrFamb, 1)) = 0 Then MsgBox "Debe seleccionar un nivel superior, en categoria dietetica...", vbCritical, MsgTitulo: Exit Sub
'          coddi2 = Val(Mid(StrFamb, 1))
'
'       Loop
'
'    End If
    
    '------- Fin validar categoria dietetica
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT * FROM a_recetacatdie with (nolock) WHERE car_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No Existe Categoria dietetica", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS1.Close: Set RS1 = Nothing
    imarca = 0
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then imarca = 1: Exit For
    
    Next i
    
    If imarca = 0 Then MsgBox "Debe Seleccionar A lo menor una Receta", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Mover registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i: vaSpread1.Col = 1
        
        If vaSpread1.text = "1" Then
           
           vaSpread1.SetActiveCell 1, vaSpread1.Row
           CodReceta = 0
           vaSpread1.Col = 2
           CodReceta = Val(vaSpread1.text)
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           RS1.Open "SELECT * FROM b_receta with (nolock) WHERE rec_codigo = " & CodReceta & "", vg_db, adOpenStatic
           
           If Not RS1.EOF Then
              
              nomrec = Trim(RS1!rec_nombre)
              nomfan = Trim(RS1!rec_nomfan)
              tippla = RS1!rec_tippla
              vg_db.Execute "sgpadm_iu_receta_V05 'M4', " & CodReceta & ", " & Val(fpLongInteger1(0).Value) & ", 0, '', '', '', '', '', 0, '', 0, '', '', 0, '', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,''"
              RS1.Close: Set RS1 = Nothing
              vg_swmovreceta = 1
           
           Else
              
              RS1.Close: Set RS1 = Nothing
              MsgBox "Información no puede moverse, puede que halla sido eliminado ó movido de recetario ", vbInformation + vbOKOnly, MsgTitulo
           
           End If
        
        End If
        
    Next i
    
    If vg_swmovreceta = 1 Then
    
       fg_descarga
       MsgBox "Proceso Finalizado Sin Problema", vbInformation + vbOKOnly, MsgTitulo
    
    End If
    
Case 3
    
    Me.Hide
    Unload Me

End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then
   'vg_db.RollbackTrans:
   MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
   Exit Sub
End If
'tecfood If Err = -2147467259 Then vg_dbtec.RollbackTrans: vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then
   'vg_db.RollbackTrans
   Exit Sub
End If
'tecfood If Err = 3034 Then vg_dbtec.RollbackTrans: vg_db.RollbackTrans: Exit Sub
'vg_db.RollbackTrans
'tecfood vg_dbtec.RollbackTrans: vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

If vaSpread1.MaxRows > 0 And Col = 1 And Row = 0 Then
   
   If imarca = 0 Then
      
      For i = 1 To vaSpread1.MaxRows
          
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = 10
          vaSpread1.TypeCheckText = ""
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "1" ' checked
      
      Next i
      
      imarca = 1
   
   Else
      
      For i = 1 To vaSpread1.MaxRows
          
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.CellType = 10
          vaSpread1.TypeCheckText = ""
          vaSpread1.TypeCheckCenter = True
          vaSpread1.Value = "" ' checked
      
      Next i
      
      imarca = 0
   
   End If

End If

End Sub

Sub LlenarRecetas(ruttippla As String, FilCatDie As Long, FilTipPla As Long)

Dim RS1 As New ADODB.Recordset

fillcatdie = FilCatDie
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT rec_codigo, rec_nombre, rec_fecvig FROM b_receta with (nolock) WHERE (rec_catdie = " & FilCatDie & " OR " & FilCatDie & "=0) AND (rec_tippla = " & FilTipPla & " OR " & FilTipPla & "=0) AND rec_tiprec='0' ORDER BY rec_nombre", vg_db, adOpenForwardOnly ', adOpenStatic
If Not RS1.EOF Then

   Do While Not RS1.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
     
      vaSpread1.Col = -1
      If RS1!rec_fecvig <= Val(Format(Date, "yyyymmdd")) And RS1!rec_fecvig > 0 Then vaSpread1.BackColor = Shape1(1).FillColor
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = RS1!rec_codigo
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS1!rec_nombre)
      
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
'      vaSpread1.CellType = CellTypeDate
'      vaSpread1.TypeDateCentury = False
'      vaSpread1.TypeHAlign = TypeHAlignCenter
'      vaSpread1.TypeSpin = False
'      vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
'      vaSpread1.TypeDateMin = "01011973":  vaSpread1.TypeDateMax = "31125000"
      vaSpread1.text = IIf(RS1!rec_fecvig > 0, Mid(RS1!rec_fecvig, 7, 2) & "/" & Mid(RS1!rec_fecvig, 5, 2) & "/" & Mid(RS1!rec_fecvig, 1, 4), "")
      vaSpread1.TypeDateCentury = True
      
      RS1.MoveNext
   
   Loop

End If

RS1.Close
Set RS1 = Nothing
Label2(8).Caption = ruttippla

End Sub

