VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CpRPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copiar Receta en Planificación Teórica"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2325
      Left            =   0
      TabIndex        =   1
      Top             =   420
      Width           =   4335
      _Version        =   393216
      _ExtentX        =   7646
      _ExtentY        =   4101
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
      MaxRows         =   6
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_CpRPla.frx":0000
      UserResize      =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   2745
      Left            =   4350
      TabIndex        =   0
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   4842
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Diciembre 2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   4275
   End
End
Attribute VB_Name = "M_CpRPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nrodia As String
Dim VecDia() As String
Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    vg_codigo = ""
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        For j = 1 To vaSpread1.maxcols
            vaSpread1.Col = j
            If vaSpread1.CellType = CellTypeButton And vaSpread1.TypeButtonColor = 13741485 Then vg_codigo = vg_codigo & Trim(vaSpread1.TypeButtonText) & ";"
        Next j
    Next i
    If Trim(vg_codigo) = "" Then MsgBox "No existen días seleccionado", vbInformation + vbOKOnly, "Copiado Especial": Exit Sub
Case 3
    vg_codigo = ""
End Select
Me.Hide
Unload Me
End Sub

Sub Inicio(tfor As String, op As String, fecpla As String, tdia As String)
Me.Caption = tfor
Msgtitulo = tfor
'fecmin = TipM
nrodia = tdia

'------- Armar calendario
vaSpread1.Row = -1: vaSpread1.Col = -1:
vaSpread1.BackColor = Label1(0).BackColor
Label1(0).Caption = Meses("01/" & Mid(fg_pone_cero(fecpla, 6), 5, 2) & "/" & Mid(fg_pone_cero(fecpla, 6), 1, 4)) & " " & Mid(fg_pone_cero(fecpla, 6), 1, 4)
diafin = fg_mes(Mid(fg_pone_cero(fecpla, 6), 5, 2) & Mid(fg_pone_cero(fecpla, 6), 1, 4))
nrosem = 1
For i = 1 To diafin
    Select Case fg_Dia(fecpla & fg_pone_cero(i, 2))
    Case 1
        vaSpread1.Row = nrosem
        vaSpread1.Col = 7
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeCheckType = TypeCheckTypeNormal
'        vaSpread1.TypeButtonPicture = 0
        vaSpread1.TypeButtonText = CStr(i)
        nrosem = nrosem + 1
    Case 2
        vaSpread1.Row = nrosem
        vaSpread1.Col = 1
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeButtonText = CStr(i)
    Case 3
        vaSpread1.Row = nrosem
        vaSpread1.Col = 2
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeButtonText = CStr(i)
    Case 4
        vaSpread1.Row = nrosem
        vaSpread1.Col = 3
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeButtonText = CStr(i)
    Case 5
        vaSpread1.Row = nrosem
        vaSpread1.Col = 4
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeButtonText = CStr(i)
    Case 6
        vaSpread1.Row = nrosem
        vaSpread1.Col = 5
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeButtonText = CStr(i)
    Case 7
        vaSpread1.Row = nrosem
        vaSpread1.Col = 6
        vaSpread1.CellType = CellTypeButton
        vaSpread1.TypeButtonType = TypeButtonTypeTwoState
        vaSpread1.TypeButtonTextColor = &H800000
        vaSpread1.TypeButtonAlign = TypeButtonAlignBottom
        vaSpread1.TypeButtonText = CStr(i)
    End Select
Next i
vaSpread1.RetainSelBlock = False
'------- mover días no permitidos
ReDim Preserve VecDia(0)
ValLcntH = "": i = 0
For j = 1 To Len(tdia)
    If Asc(Mid(tdia, j, 1)) <> 59 Then
       ValLcntH = ValLcntH + Mid(tdia, j, 1)
    Else
       ReDim Preserve VecDia(i): VecDia(i) = ValLcntH: ValLcntH = "": i = i + 1
    End If
Next j
If Trim(ValLcntH) <> "" Then ReDim Preserve VecDia(i): VecDia(i) = ValLcntH
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
vaSpread1.Row = Row: vaSpread1.Col = Col
For i = 0 To UBound(VecDia)
    If Trim(vaSpread1.TypeButtonText) = VecDia(i) Then vaSpread1.TypeButtonType = TypeButtonTypeNormal: Exit Sub
Next i
If ButtonDown = 1 Then
   vaSpread1.TypeButtonColor = &HD1ADAD
ElseIf ButtonDown = 0 Then
   vaSpread1.TypeButtonColor = Label1(0).BackColor
End If
End Sub
