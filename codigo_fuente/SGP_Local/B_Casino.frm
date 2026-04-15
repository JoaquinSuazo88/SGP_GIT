VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_Casino 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Casino"
   ClientHeight    =   4815
   ClientLeft      =   2520
   ClientTop       =   1530
   ClientWidth     =   5085
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5085
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4540
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "B_Casino.frx":0000
         Left            =   1800
         List            =   "B_Casino.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   555
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   225
         Index           =   0
         Left            =   255
         TabIndex        =   4
         Top             =   345
         Width           =   1320
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
         Left            =   285
         TabIndex        =   3
         Top             =   660
         Width           =   1035
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3615
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   4545
      _Version        =   393216
      _ExtentX        =   8017
      _ExtentY        =   6376
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
      MaxRows         =   30
      OperationMode   =   3
      ScrollBars      =   2
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_Casino.frx":001E
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4815
      Left            =   4545
      TabIndex        =   6
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
Attribute VB_Name = "B_Casino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As ADODB.Recordset
Dim i As Long, ibusca As Long, irow As Long
Dim icombo As Integer
Private Sub Combo1_Click()
If icombo = 0 Then
   Text1.Text = ""
End If
End Sub
Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
Me.Left = vg_left
fg_carga (ss)
Mover_Botones
icombo = 1
Combo1.ListIndex = 1
vaSpread1.MaxRows = 0
RS1.Open "select Codigo_Casino, Nombre_Casino From Sdx_Casino " & _
         "Where Sdx_Casino.IndBorrado = 0 order by Nombre_Casino", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not ConSql.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
              
      vaSpread1.Col = 1
      vaSpread1.Text = RS1!Codigo_Casino

      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(RS1!Nombre_Casino)
             
      RS1.MoveNext
   Loop
   irow = 1
End If
RS1.Close: Set RS1 = Nothing
icombo = 0
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Casino"
End Sub
Private Sub Text1_Change()
If LimpiaDato(Trim(Text1.Text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   RS1.Open "select count(*) as nreg From Sdx_Casino Where IndBorrado = 0 " & _
            "and   Ucase(Codigo_Casino)) like '%" + "00000" + UCase(LimpiaDato(Text1.Text)) + "%'", vg_db, adOpenStatic
   If RS1.EOF Or RS1!NReg = 0 Then RS1.Close: Set RS1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
   If ibusca <> RS1!NReg Then ibusca = RS1!NReg: vaSpread1.MaxRows = RS1!NReg
   RS1.Close: Set RS1 = Nothing
   RS1.Open "select Codigo_Casino, Nombre_Casino From Sdx_Casino Where IndBorrado = 0 " & _
            "and   Ucase(Codigo_Casino) like '%" + "00000" + UCase(LimpiaDato(Text1.Text)) + "%' " & _
            "order by Nombre_Casino", vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   RS1.Open "select count(*) as nreg From Sdx_Casino Where IndBorrado = 0 " & _
            "and   Ucase(Nombre_Casino) like '%" + UCase(LimpiaDato(Text1.Text)) + "%'", vg_db, adOpenStatic
   If RS1.EOF Or RS1!NReg = 0 Then RS1.Close: Set RS1 = Nothing: ibusca = 0: vaSpread1.MaxRows = 0: Exit Sub
   If ibusca <> RS1!NReg Then ibusca = RS1!NReg: vaSpread1.MaxRows = RS1!NReg
   RS1.Close: Set RS1 = Nothing
   RS1.Open "select Codigo_Casino, Nombre_Casino From Sdx_Casino Where IndBorrado = 0 " & _
            "and   Ucase(Nombre_Casino) like '%" + UCase(LimpiaDato(Text1.Text)) + "%' " & _
            "order by Nombre_Casino", vg_db, adOpenStatic
End If
i = 1
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = RS1!Codigo_Casino
      
      vaSpread1.Col = 2
      vaSpread1.Text = Trim(RS1!Nombre_Casino)
      
      RS1.MoveNext
   Loop
   irow = 1
End If
RS1.Close: Set RS1 = Nothing
vaSpread1.SetActiveCell 1, 1
End Sub
Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
  Case 1
    MoverDatos
  Case 3
    Cerrar
End Select
End Sub
Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
irow = vaSpread1.ActiveRow
End Sub
Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
MoverDatos
End Sub
Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
If vaSpread1.MaxRows < 1 Then Exit Sub
irow = vaSpread1.ActiveRow
MoverDatos
End Sub
Sub MoverDatos()
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = irow
vaSpread1.Col = 1
vg_auxcodcasino = vaSpread1.Text
Cerrar
End Sub
Sub Cerrar()
Me.Hide
Unload Me
End Sub
Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
  Case 27
    Cerrar
End Select
End Sub
Sub Mover_Botones()

   Toolbar1.ImageList = Partida.IL1
   Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
   Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
   Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

End Sub

