VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_InvRot 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario Rotativo"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   7455
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   5850
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_InvRot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String
Dim RS As New ADODB.Recordset

Private Sub fecha_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Toolbar1.ImageList = Partida.IL1
'--------------------------- Crea Botones de la toolbar
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa": btnX.Enabled = True
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
'--------------------------- Da dimensiones al formulario para que no se descentre
Me.Height = 1920
Me.Width = 7830
Msgtitulo = "Inventario Rotativo"
fg_centra Me
Combo1(0).Clear
Combo1(0).AddItem "Inventario Rotativo"
Combo1(0).AddItem "Invetario Full"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Combo1(0).ListIndex = -1 Then MsgBox "Debe seleccionar tipo Informe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Combo1(0).ListIndex = 1 Then
       I_TomaInventarioFull
    Else
    End If
Case 3
    Me.Hide
    Unload Me
End Select
End Sub
