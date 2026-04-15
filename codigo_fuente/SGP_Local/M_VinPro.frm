VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_VinPro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculo Ingrediente & Productos"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9705
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   9705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8895
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2895
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   8415
         _Version        =   393216
         _ExtentX        =   14843
         _ExtentY        =   5106
         _StockProps     =   64
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
         MaxRows         =   10
         SpreadDesigner  =   "M_VinPro.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4155
      Left            =   9165
      TabIndex        =   2
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   7329
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_VinPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, titulo
End Sub


Sub LlenaDatos(coding As String)
Set RS = vg_db.Execute("SELECT DISTINCT a.ing_nombre, b.pro_codigo, b.pro_nombre " & _
        "FROM b_ingrediente a, b_productos b, b_productosing c " & _
        "WHERE a.ing_codigo = c.pri_coding " & _
        "AND   c.pri_codpro = b.pro_codigo " & _
        "AND   a.ing_codigo = '" & coding & "'")
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
Do While Not RS.EOF
   Frame1.Caption = coding & " - " & Trim(RS!ing_nombre)
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = Trim(RS!pro_codigo)
   vaSpread1.Col = 2
   vaSpread1.text = Trim(RS!pro_nombre)
   RS.MoveNext
Loop
vaSpread1.Visible = True
RS.Close: Set RS = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    Me.Hide
    Unload Me
End Select

End Sub
