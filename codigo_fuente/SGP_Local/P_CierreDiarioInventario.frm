VERSION 5.00
Begin VB.Form P_CierreDiarioInventario 
   BorderStyle     =   0  'None
   Caption         =   "Proceso Cierre Inventario"
   ClientHeight    =   1440
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00C0FFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5895
      End
   End
End
Attribute VB_Name = "P_CierreDiarioInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FecCie As Date


Private Sub Form_Activate()

On Error GoTo Man_Error

Proceso

Me.Hide
Unload Me

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me

Label1.Caption = "Un momento, generando cierre diario de un precierre..." & VgLinea & "Espere hasta que finalice el proceso..."

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub Proceso()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

If GeneraMDBInventario(Me, FecCie) Then

   'Actualizar estado tin_ADMSGP
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS1 = vg_db.Execute("sgp_Upd_TomaInventarioEnvioADMSGP '" & MuestraCasino(1) & "', " & Format(FecCie, "yyyymmdd") & "")
   If Not RS1.EOF Then

      If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
    
         RS1.Close
         Set RS1 = Nothing
                
         MsgBox "Existe error actualizar estado envio inventario..", vbExclamation + vbOKOnly, MsgTitulo
         Exit Sub
 
      End If

   End If
   RS1.Close
   Set RS1 = Nothing

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub GeneraMDBInven(Fecha As Date)

On Error GoTo Man_Error

FecCie = Fecha
Label1.Caption = "Un momento, generando cierre diario de un precierre..." & VgLinea & "Espere hasta que finalice el proceso..."


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub


