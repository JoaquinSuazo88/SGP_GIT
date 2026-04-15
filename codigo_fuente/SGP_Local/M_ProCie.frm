VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_ProCie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reproceso Cierre Diario"
   ClientHeight    =   3900
   ClientLeft      =   4980
   ClientTop       =   2430
   ClientWidth     =   8085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   8085
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   7695
      Begin MSComctlLib.ProgressBar PBar1 
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   2400
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fin Proceso :  "
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
         Left            =   4680
         TabIndex        =   6
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inicio Proceso :  "
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
         Left            =   720
         TabIndex        =   5
         Top             =   360
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Un momento Procesando Información Día : "
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
         Left            =   1080
         TabIndex        =   3
         Top             =   1320
         Width           =   3735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   6720
      TabIndex        =   1
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   5280
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "M_ProCie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case 0
    procesardia
    Unload Me
Case 1
   Me.Hide
   Unload Me
End Select
End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga
procesardia
Unload Me

Exit Sub
Man_Error:
fg_descarga
If Err = 380 Then Resume Next
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Form_Load()

Me.HelpContextID = vg_OpcM
MsgTitulo = "Proceso Cierre Díario"
Command1(0).Visible = False
fg_centra Me

End Sub

Sub procesardia()

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim fecpro    As Date
Dim fecter    As Date
Dim fecdia    As String
Dim EstCierre As Boolean

Label1.Caption = ""
Command1(0).Enabled = False
Command1(1).Enabled = False


Label1.Caption = False
RS.Open "SELECT DISTINCT par_nombre, par_valor FROM a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
If Not RS.EOF Then
   fecter = fg_Desencripta(TipoDato(RS!par_valor, ""))
End If
RS.Close: Set RS = Nothing

If Trim(fecter) = "" Then MsgBox "No esta activo la fecha cierre día, Comunicase con departamento de informatica" & VgLinea & Space(40) & "Proceso cancelado ...", vbCritical + vbOKOnly, "Menú Principal": End
If fg_Desencripta(GetParametro("fecrprodia")) = "" Then MsgBox "Fecha reproceso cierre diario esta blanco, Comunicase con departamento de informatica" & VgLinea & Space(40) & "Proceso cancelado ...", vbCritical + vbOKOnly, "Menú Principal": End

Dim i As Long

fecter = fecter - 1
fecpro = IIf(Text1.text <> "", Text1.text, fg_Desencripta(TipoDato(GetParametro("fecrprodia"), "")))

If (fecter - fecpro) < 0 Then

   PBar1.max = (fecpro - fecter)

Else

   PBar1.max = (fecter - fecpro)

End If


PBar1.Min = 0
i = 0
Label2(0).Caption = "Inicio Proceso :  " & Format(Now, "hh:mm:ss")
Do While fecpro <= fecter
   
   EstCierre = True
   
   fecdia = fecpro
   Label1.Caption = "Un momento Procesando Información Día : " & fecpro
   DoEvents
   
   If vg_tipbase = "1" Then
      
      ProcesoCierreAccess Me, True, fecdia
   
   Else
      
      If Not ProcesoCierreSql(Me, True, fecdia) Then
      
         EstCierre = False
         
      End If
   
   End If
   
   If EstCierre Then
      
      fecpro = fecpro + 1
      i = i + 1
   
      If i <= PBar1.max Then
      
         PBar1.Value = i
      
      End If
   End If
   Label2(1).Caption = "Fin Proceso : " & Format(Now, "hh:mm:ss")
Loop
'-------> Actualizar opción tabla a_param reporceso cierre diario
vg_db.Execute "update a_param set par_valor = 'S' where par_codigo = 'rprociedia' and par_cencos = '" & MuestraCasino(1) & "'"
MsgBox "Proceso Finalizado", vbInformation + vbOKOnly, MsgTitulo

Exit Sub
Man_Error:
fg_descarga
If Err = 380 Then Resume Next: Exit Sub
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Select Case Cancel
Case 1
   vg_db.Execute "update a_param set par_valor = '' where par_codigo = 'fecrprodia' and par_cencos = '" & MuestraCasino(1) & "'"
End Select
End Sub
