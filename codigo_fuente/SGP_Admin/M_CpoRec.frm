VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_CpoRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Receta Destino"
   ClientHeight    =   1935
   ClientLeft      =   2550
   ClientTop       =   3255
   ClientWidth     =   8595
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   8595
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1935
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8055
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Caption         =   "Fecha Vencimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   14
         Top             =   1635
         Width           =   1950
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   885
         Width           =   5140
         _Version        =   196608
         _ExtentX        =   9066
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
         BackColor       =   -2147483624
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1800
         TabIndex        =   3
         Top             =   1230
         Width           =   5140
         _Version        =   196608
         _ExtentX        =   9066
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
         BackColor       =   -2147483624
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2340
         TabIndex        =   15
         Top             =   1560
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   556
         Enabled         =   0   'False
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
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
         ButtonStyle     =   3
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
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
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
         Text            =   "12/10/2004"
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   1
         IncHour         =   1
         IncMinute       =   1
         IncSecond       =   1
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
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
         Index           =   1
         Left            =   2190
         TabIndex        =   10
         Top             =   540
         Width           =   4710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Index           =   9
         Left            =   165
         TabIndex        =   9
         Top             =   600
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Categoria Dietetica"
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
         Index           =   8
         Left            =   165
         TabIndex        =   8
         Top             =   285
         Width           =   1650
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(Opcional)"
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
         Left            =   7005
         TabIndex        =   7
         Top             =   280
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "(Opcional)"
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
         Left            =   7005
         TabIndex        =   6
         Top             =   580
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Receta"
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
         Left            =   165
         TabIndex        =   5
         Top             =   975
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Fantasia"
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
         Index           =   2
         Left            =   165
         TabIndex        =   4
         Top             =   1305
         Width           =   1440
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1740
         Picture         =   "M_CpoRec.frx":0000
         Top             =   120
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1740
         Picture         =   "M_CpoRec.frx":030A
         Top             =   435
         Width           =   480
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
         Left            =   2190
         TabIndex        =   12
         Top             =   210
         Width           =   4710
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2235
         TabIndex        =   13
         Top             =   260
         Width           =   4710
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2235
         TabIndex        =   11
         Top             =   585
         Width           =   4710
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   1935
      Left            =   8055
      TabIndex        =   0
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   3413
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CpoRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim catdieprevio As Long, tipplaprevio As Long, catdie As Long, tippla As Long, indice As Long, i As Long
Dim MsgTitulo As String

Private Sub Check1_Click(Index As Integer)

If Check1(0).Value = 1 Then fpDateTime1(0).Enabled = True: fpDateTime1(0).text = Format(Date, "dd/mm/yyyy") Else fpDateTime1(0).Enabled = False: fpDateTime1(0).text = "  /  /    "

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

Dim RS As New ADODB.Recordset

fg_centra Me
MsgTitulo = "Copiar Recetas"

catdieprevio = 0
tipplaprevio = 0
catdie = 0
tippla = 0
indice = 0

Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ConsultaRecetaOrigenCopiar " & vg_codreceta & "")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   MsgBox "Receta seleccionada no existe proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
   Me.Hide
   Unload Me

End If

catdieprevio = RS!car_previo
catdie = RS!rec_catdie
fpayuda(0).Caption = RS!car_nombre
tipplaprevio = RS!tip_previo
tippla = RS!rec_tippla
fpayuda(1).Caption = RS!tip_nombre
fpText1(0).text = RS!rec_nombre
fpText1(1).text = RS!rec_nomfan
fpDateTime1(0).text = IIf(IsNull(RS!rec_fecvig) Or RS!rec_fecvig = 0, "  /  /    ", Mid(RS!rec_fecvig, 7, 2) & "/" & Mid(RS!rec_fecvig, 5, 2) & "/" & Mid(RS!rec_fecvig, 1, 4))
Check1(0).Value = IIf(IsNull(RS!rec_fecvig) Or RS!rec_fecvig = 0, 0, 1)

RS.Close
Set RS = Nothing
'------- Buscar raiz categoria dietetica
For i = 1 To 10
    
    If catdieprevio = 0 Then Exit For
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_ConsultaCategoriaDiateticaCopiarReceta " & catdieprevio & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit For
    fpayuda(0).Caption = RS!car_nombre & "\" & fpayuda(0).Caption
    codcatdieprevio = RS!car_previo
    If RS!car_previo = 0 Then RS.Close: Set RS = Nothing: Exit For
    
    RS.Close
    Set RS = Nothing

Next
'------- Buscar raiz tipo plato
For i = 1 To 10
    
    If tipplaprevio = 0 Then Exit For

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_ConsultaTipoPlatoCopiarReceta " & tipplaprevio & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit For
    fpayuda(1).Caption = RS!tip_nombre & "\" & fpayuda(1).Caption
    tipplaprevio = RS!tip_previo
    If RS!tip_previo = 0 Then RS.Close: Set RS = Nothing: Exit For
    
    RS.Close
    Set RS = Nothing

Next

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

If IsDate(fpDateTime1(Index).text) = False Then Exit Sub

End Sub

Private Sub fpText1_Change(Index As Integer)

fpText1(1).text = fpText1(0).text

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Image1_Click(Index As Integer)

Select Case Index

Case 0
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(0).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
    B_ArbEst.Show 1
    If vg_codigo = "" Then Exit Sub
    catdie = Val(vg_codigo)
    fpayuda(0).Caption = Trim(vg_nombre)

Case 1
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(1).Left + 2400
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
    B_ArbEst.Show 1
    tippla = Val(vg_codigo)
    If tippla = 0 Then Exit Sub
    fpayuda(1).Caption = Trim(vg_nombre)

End Select

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim RS  As New ADODB.Recordset

Select Case Button.Index

Case 1
    
    If Trim(fpText1(0).text) = "" Or Trim(fpText1(1).text) = "" Then
    
       MsgBox "Falta Información ...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    Dim coddi1 As Long, coddi2 As Long, codti1 As Long, codti2 As Long, codti3 As Long, fecvig As Long
    Dim StrFamb As String, StrFam As String
    Dim unidamedida As Long
    
'    '------- Validar categoria dietetica
'    StrFam = fg_BuscaCodArbol(catdie, "a_recetacatdie", "car_codigo")
'    If Len(StrFam) <> 0 Then
'
'       Do While InStr(StrFam, ";") <> 0
'
'          StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
'          StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
'          coddi1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
'          If Val(Mid(StrFamb, 1)) = 0 Then MsgBox "Debe seleccionar un nivel superior, en categoria dietetica...", vbCritical, MsgTitulo: Exit Sub
'          coddi2 = Val(Mid(StrFamb, 1)): catdie = Val(Mid(StrFamb, 1))
'
'       Loop
'
'    End If
    '------- Fin validar categoria dietetica
    
    If catdie = 0 Then
       
       MsgBox "Debe seleccionar un nivel inferior más, en categoria receta", vbInformation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    '------- Validar tipo de plato
    
    StrFam = fg_BuscaCodArbol(tippla, "a_recetatippla", "tip_codigo")
    
    If Len(StrFam) <> 0 Then
       
       Do While InStr(StrFam, ";") <> 0
          
          StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
          StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
          codti1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
          codti2 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
          If codti2 = 0 Then MsgBox "Debe seleccionar un nivel superior, en tipo de plato...", vbCritical, MsgTitulo: Exit Sub
          codti3 = Val(Mid(StrFamb, 1)): tippla = IIf(codti3 < 1, codti2, codti3)
       
       Loop
    
    End If
    
    If tippla = 0 Then
       
       MsgBox "Debe seleccionar un nivel inferior más, en tipo de plato", vbInformation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    '------- Fin validar tipo de plato
    If MsgBox("Copia registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    fg_carga ""
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgpadm_Sel_ConsultaCopiaReceta_V03 " & vg_codreceta & "")
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: Exit Sub
    indice = 0: fecvig = 0
    fecvig = IIf(Check1(0).Value = 0, 0, Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2) & Mid(fpDateTime1(0).text, 1, 2))
    
    unidamedida = RS!cod_uniReceta
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS1 = vg_db.Execute("sgpadm_iu_receta_V05 'A', 0, " & catdie & ", " & tippla & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "', " & _
                            "'" & LimpiaDato(Trim(fpText1(1).text)) & "', '" & RS!rec_metpre & "', '" & RS!rec_conche & "', " & _
                            "'" & RS!rec_sugere & "', " & RS!rec_basrac & ", '" & RS!rec_tiprec & "', " & fecvig & ", '" & RS!rec_gruvul & "', " & RS!rec_indppr & ", " & RS!rec_canser & ", '" & RS!rec_hipali & "', " & _
                            "" & RS!IdSellos & ", " & RS!IdCosto & ", " & RS!IdTipoIngPrincipal & ", " & RS!IdMetodoCoccion & ", " & RS!IdCategorizacionCompleja & ", " & RS!IdIngCruceGarnitura & ", " & RS!IdEfectoMeteorizante & ", " & RS!IdTiempoCoccion & ", " & RS!IdTiempoHh & ", " & RS!IdColor & ", " & RS!IdEtiquetadoSello & ", " & RS!IdSegundoIngredientePrincipal & ", " & RS!IdEquipamientoCoccion & ", " & RS!IdParametroSalsa & ", '" & RS!IdIntegraAMD & "'")
    If Not RS1.EOF Then
       
       indice = RS1!indice
    
    End If
    RS1.Close
    Set RS1 = Nothing
    
    RS.Close
    Set RS = Nothing
    
    Sql = " sgpadm_iu_codUnidadReceta "
    Sql = Sql & Trim(indice) & ","
    Sql = Sql & unidamedida
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute(Sql)
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
'    RS.Open "SELECT * FROM b_recetadet WHERE red_codigo = " & vg_codreceta & "", vg_db, adOpenStatic
    Set RS = vg_db.Execute("sgpadm_Sel_ConsultaCopiaRecetaDetalle " & vg_codreceta & "")
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          vg_db.Execute "sgpadm_iu_recetadet 'A', " & indice & ", " & RS!red_nroite & ", " & RS!red_codpro & ", " & _
                        "" & RS!red_canpro & ", " & RS!red_cospro & ", " & RS!red_pctapr & ", " & _
                        "" & RS!red_pctcoc & ", " & RS!red_pctnut & ", '', '" & RS!red_IndentificadorIngSumaTablaGramaje & "'"
          
          RS.MoveNext
       
       Loop
    
    End If
    RS.Close
    Set RS = Nothing
    
    'copia parametro de recetas
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Ins_CopiaParametroReceta " & vg_codreceta & ", '" & indice & "', '" & UCase(vg_NUsr) & "'")

    If Not RS.EOF Then
            
       If RS(0) > 0 Then
                   
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
      
'       Else
   
'          MsgBox "Proceso Termino Correctamente ", vbInformation + vbOKOnly, MsgTitulo
               
       End If
            
    End If
    RS.Close
    Set RS = Nothing
    
    fg_descarga
    vg_swpegreceta = 1
    MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, MsgTitulo

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub

If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
