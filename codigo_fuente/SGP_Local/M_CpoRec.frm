VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_CpoRec 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Receta Patrón"
   ClientHeight    =   2595
   ClientLeft      =   1575
   ClientTop       =   3030
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   30
      TabIndex        =   1
      Top             =   120
      Width           =   8445
      Begin VB.Frame Frame2 
         Height          =   885
         Left            =   180
         TabIndex        =   4
         Top             =   1020
         Width           =   8055
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   1110
            TabIndex        =   5
            Top             =   390
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
            ButtonDefaultAction=   0   'False
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   3
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
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            MaxValue        =   "2147483647"
            MinValue        =   "-2147483647"
            NegFormat       =   0
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
            AutoMenu        =   -1  'True
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Regimen"
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
            Index           =   26
            Left            =   240
            TabIndex        =   7
            Top             =   450
            Width           =   750
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   2010
            Picture         =   "M_CpoRec.frx":0000
            Top             =   300
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   2520
            TabIndex        =   6
            Top             =   390
            Width           =   5355
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   7
            Left            =   2535
            TabIndex        =   8
            Top             =   405
            Width           =   5385
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "por Regimen"
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
         Index           =   1
         Left            =   5250
         TabIndex        =   3
         Top             =   630
         Width           =   2625
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Local"
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
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   630
         Value           =   -1  'True
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
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
         Left            =   240
         TabIndex        =   9
         Top             =   270
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   2595
      Left            =   8520
      TabIndex        =   0
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   4577
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
Dim RS As New ADODB.Recordset
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Msgtitulo = "Copiar Recetas"
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
RS.Open "SELECT rec_nombre FROM  b_receta WHERE rec_codigo = " & vg_codreceta & "", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Receta seleccionada no existe proceso cancelado", vbExclamation + vbOKOnly, Msgtitulo: Me.Hide: Unload Me
Label1.Caption = RS!rec_nombre
RS.Close: Set RS = Nothing
Frame2.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Option1(0).Enabled = IIf(("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Or vg_5etapas = False), True, False)
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_codigo < 10000", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "": Exit Sub
fpayuda(6).Caption = Trim(RS!reg_nombre)
RS.Close: Set RS = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
vg_left = fpayuda(6).Left + 3000
vg_nombre = "": vg_codigo = ""
'B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", IIf(vg_5etapas = False, "No5etapas", "Gen")
B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", IIf(vg_5etapas = False, "No5etapas", "5etapas")
B_TabEst.Show 1
Me.Refresh
If vg_codigo = "" Then Exit Sub
fpLongInteger1(1).Value = Val(vg_codigo)
fpayuda(6).Caption = vg_nombre
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Frame2.Enabled = False
    fpLongInteger1(1).Value = ""
    fpayuda(6).Caption = ""
Case 1
    Frame2.Enabled = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If Option1(1).Value = True And Trim(fpLongInteger1(1).Value) = "" Then MsgBox "Falta Información ...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Copia registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    RS.Open "SELECT * FROM b_receta WHERE rec_codigo = " & vg_codreceta & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga ""
    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_receta SET rec_tiprec = " & IIf(Option1(0).Value = True, -1, Val(fpLongInteger1(1).Value)) & " WHERE rec_codigo = " & vg_codreceta & "" ' AND (rec_tiprec = 0 OR rec_tiprec = -1)"
    RS.Open "SELECT DISTINCT red_codigo FROM b_recetadet WHERE red_codigo = " & vg_codreceta & " AND red_tiprec = " & IIf(Option1(0).Value = True, -1, Val(fpLongInteger1(1).Value)) & " AND red_cencos = '" & MuestraCasino(1) & "' AND " & Val(fpLongInteger1(1).Value) & " >= 10000", vg_db, adOpenStatic
    If Not RS.EOF Then
       RS.Close: Set RS = Nothing
       RS.Open "SELECT * FROM b_recetadet WHERE red_codigo = " & vg_codreceta & " AND red_tiprec = 0 AND red_cencos = '0'", vg_db, adOpenStatic
       If Not RS.EOF Then
          Do While Not RS.EOF
             vg_db.Execute "UPDATE b_recetadet SET red_canpro = " & RS!red_canpro & " WHERE red_codigo = " & vg_codreceta & " AND red_tiprec = " & IIf(Option1(0).Value = True, -1, Val(fpLongInteger1(1).Value)) & " AND red_cencos = '" & MuestraCasino(1) & "' AND " & Val(fpLongInteger1(1).Value) & " >= 10000 AND red_codpro = '" & RS!red_codpro & "' AND red_nroite = " & RS!red_nroite & ""
             RS.MoveNext
          Loop
       End If
       RS.Close: Set RS = Nothing
    Else
       vg_db.Execute "DELETE b_recetadet FROM b_recetadet WHERE red_codigo=" & vg_codreceta & " AND red_tiprec=" & IIf(Option1(0).Value = True, -1, Val(fpLongInteger1(1).Value)) & " AND red_cencos='" & MuestraCasino(1) & "'"
       vg_db.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos) SELECT red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, " & IIf(Option1(0).Value = True, -1, Val(fpLongInteger1(1).Value)) & ", '" & MuestraCasino(1) & "' FROM b_recetadet WHERE red_codigo=" & vg_codreceta & " AND red_tiprec=0 AND red_cencos='0'"
       RS.Close: Set RS = Nothing
    End If
    vg_db.CommitTrans
    fg_descarga
    vg_swpegreceta = 1
    MsgBox "Copia Finalizada Sin Problema", vbInformation + vbOKOnly, Msgtitulo
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
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
