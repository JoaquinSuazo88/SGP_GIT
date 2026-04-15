VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form I_ForCom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Formato de Compras"
   ClientHeight    =   1470
   ClientLeft      =   5715
   ClientTop       =   2100
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   5295
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "I_ForCom.frx":0000
         Left            =   1080
         List            =   "I_ForCom.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   3795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informes"
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
         TabIndex        =   3
         Top             =   315
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_ForCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Form_Load()
fg_centra Me
Me.HelpContextID = vg_OpcM
Msgtitulo = "Formato de Compras"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Combo1.Clear
Combo1.AddItem "Formato Compras SAC"
Combo1.AddItem "Formato Compras SAP"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    If Combo1.ListIndex = -1 Then MsgBox "Debe seleccionar Informe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_FormatoCompras IIf(Combo1.ListIndex = 0, "sac", "sap")
Case 3
    Me.Hide
    Unload Me
End Select
End Sub
