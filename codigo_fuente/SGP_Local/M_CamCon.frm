VERSION 5.00
Begin VB.Form M_CamCon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de Contrato"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   1
         Top             =   1800
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   2640
         TabIndex        =   2
         Top             =   1800
         Width           =   1425
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   2925
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   675
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1260
         TabIndex        =   4
         Top             =   660
         Width           =   2910
      End
   End
End
Attribute VB_Name = "M_CamCon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim op As String
Dim MsgTitulo As String

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then SendKeys "{TAB}"
End Sub

Private Sub Command1_Click(Index As Integer)
Dim cPer As Long

Select Case Index

Case 0
    

    If vg_contra <> Trim(fg_codigocbo(Combo1, 0, 20, "")) And ((Forms.count - 2) > 0) Then
    
       MsgBox "Existen procesos activos, se recomienda cerrar, luego realice el cambio del contrato...", vbCritical + vbOKOnly, "Ingreso al sistema"
       Me.Hide
       Unload Me
       Exit Sub
    
    End If

'       If MsgBox("Existen ventanas activas...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
'       Dim frm As Form
'       If (Forms.count > 1) Then
'           For Each frm In Forms
'               If (frm.Name <> Me.Name And frm.Name <> "Partida") Then
'                   Unload frm
'               End If
'           Next
'       End If
'    End If
    '------- Mover centro de costo a variables globales
    vg_contra = Trim(fg_codigocbo(Combo1, 0, 10, ""))
    vg_nomcon = Trim(Left(Combo1(0).List(Combo1(0).ListIndex), Len(Combo1(0).List(Combo1(0).ListIndex)) - 10))
    RS.Open "SELECT DISTINCT b.bod_codigo, b.bod_nombre FROM b_clientes a, a_bodega b WHERE a.cli_codbod = b.bod_codigo AND a.cli_codigo = '" & vg_contra & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe bodega asignado a este contrato, cancela ingreso al sistema...", vbCritical + vbOKOnly, "Ingreso al sistema": Exit Sub
    vg_codbod = RS!bod_codigo
    vg_nombod = Trim(RS!bod_nombre)
    RS.Close: Set RS = Nothing
    Me.Hide
    Unload Me
    Partida.MDIForm_Load

Case 1
    
    Me.Hide
    Unload Me

End Select

End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
fg_centra Me
fg_carga ""
MsgTitulo = "Cambio de Contrato"
RS.Open "SELECT b.cli_codigo, b.cli_nombre, a.uco_codcon FROM b_usuariocontratos a, b_clientes b WHERE a.uco_codcon = b.cli_codigo AND b.cli_tipo = 0 AND a.uco_codusu = '" & vg_NUsr & "'", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS1 = Nothing: MsgBox "Usuario no tiene asignado centro de costo...": Exit Sub
Combo1(0).Clear
Do While Not RS.EOF
   Combo1(0).AddItem RS!cli_nombre & Space(150) & "(" & fg_pone_espacio(RS!uco_codcon, 10) & ")" 'fg_pone_espacio(RS!cli_codigo, 10)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
If Combo1(0).listcount > -1 Then Combo1(0).ListIndex = fg_buscacbo(Combo1, 0, 10, vg_contra)
fg_descarga
End Sub
