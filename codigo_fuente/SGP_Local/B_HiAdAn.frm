VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_HiAdAn 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historico Pedidos Adicionales y Anulaciones"
   ClientHeight    =   4815
   ClientLeft      =   3300
   ClientTop       =   2730
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5085
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4545
      _Version        =   393216
      _ExtentX        =   8017
      _ExtentY        =   8493
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
      MaxRows         =   30
      OperationMode   =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_HiAdAn.frx":0000
      ScrollBarTrack  =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4815
      Left            =   4545
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   8493
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_HiAdAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim Msgtitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
Msgtitulo = "Historico pedidos adicionales y anulaciones"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Sub LlenarDatos(cencos As String)
fg_carga ""
With vaSpread1
    .MaxRows = 0
    RS1.Open "SELECT distinct ped_anomes, ped_tipped, ped_fecped " & _
             "FROM b_minutapedido " & _
             "WHERE ped_codcas = '" & cencos & "' " & _
             "AND  (ped_tipped = 2 or ped_tipped = 3) " & _
             "AND   ped_fecenv > 0 ORDER BY ped_anomes, ped_tipped, ped_fecped", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
    Do While Not RS1.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1: .TypeHAlign = 2: .text = Mid(RS1!ped_anomes, 5, 2) & "/" & Mid(RS1!ped_anomes, 1, 4)
       .Col = 2: .text = IIf(RS1!ped_tipped = 2, "Adicionales", "Anulaciones")
       .Col = 3: .TypeHAlign = 2: .text = Mid(RS1!ped_fecped, 7, 2) & "/" & Mid(RS1!ped_fecped, 5, 2) & "/" & Mid(RS1!ped_fecped, 1, 4)
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
End With
fg_descarga
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If vaSpread1.MaxRows < 1 Then Exit Sub
    MoverDatos
Case 3
    vg_anomes = 0: vg_tipped = 0: vg_fecval = 0
    Cerrar
End Select
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
MoverDatos
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos
End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 27
    Cerrar
End Select
End Sub

Sub MoverDatos()
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    .Row = .ActiveRow
    .Col = 1: vg_anomes = Val(Mid(.text, 4, 4) & Mid(.text, 1, 2))
    .Col = 2: vg_tipped = Val(IIf(.text = "Adicionales", 2, 3))
    .Col = 3: vg_fecval = Val(Mid(.text, 7, 4) & Mid(.text, 4, 2) & Mid(.text, 1, 2))
End With
Cerrar
End Sub

Sub Cerrar()
Me.Hide
Unload Me
End Sub
