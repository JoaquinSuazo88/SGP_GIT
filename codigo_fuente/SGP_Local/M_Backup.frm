VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_Backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup Base de Datos SGP"
   ClientHeight    =   1215
   ClientLeft      =   6915
   ClientTop       =   3510
   ClientWidth     =   4545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   4545
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   165
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
      Begin VB.Label Label1 
         Caption         =   "Un Momento, Respaldando Información"
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
         Left            =   420
         TabIndex        =   2
         Top             =   180
         Width           =   3525
      End
   End
   Begin ACTIVEZIPLib.ActiveZip AZ1 
      Left            =   0
      Top             =   0
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
   End
End
Attribute VB_Name = "M_Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub AZ1_Status(ByVal FileName As String, ByVal progress As Long)

Bar1(0).Value = progress

End Sub


Private Sub Form_Activate()
On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_centra Me
Me.Refresh

'-------> Compactar base de datos - validar si esta abierta abierta base
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
AbrirBase
If vg_tipbase = "1" Then
   
   vg_db.Close
   ''If Dir(dir_trabajo & "xxx.mdb") <> "" Then Kill dir_trabajo & "xxx.mdb"
   
   If Dir(dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 3)) & "ldb") <> "" Then
      
      MsgBox "El sistema necesita respaldar la base de dato que esta abierta." & Chr(13) & Chr(13) & "No se ejecutara hasta cerrar la Base o los programas relacionados", vbExclamation + vbOKOnly, Msgtitulo: End
   
   End If

   'Dim fso As New FileSystemObject
   If Dir(dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 3)) & "ldb") = "" And Trim(Environ("OS")) <> "" Then
      
      If Dir(dir_trabajo & "Backup\" & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip") <> "" Then Exit Sub
      DBEngine.CompactDatabase dir_trabajo & BaseDeDato, dir_trabajo & "xxx.mdb", dbLangGeneral
      Kill dir_trabajo & BaseDeDato
      fso.MoveFile dir_trabajo & "xxx.mdb", dir_trabajo & BaseDeDato
   
   End If
   '-------> Fin compactar base de datos - validar si esta abierta abierta base
   Label1.Visible = True: Bar1(0).Min = 0: Bar1(0).Value = 0: Bar1(0).max = 100: Bar1(0).Visible = True
   AZ1_Status dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip", 0
   AZ1.CreateZip dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip", "": AZ1.AddFile dir_trabajo & BaseDeDato, "", True, "": AZ1.Close
   If Dir(dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip") <> "" Then
      
      fso.MoveFile dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip", dir_trabajo & "backup\" & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip"
   
   End If

Else
   
   '-------> Fin compactar base de datos - validar si esta abierta abierta base
   Label1.Visible = True: Bar1(0).Min = 0: Bar1(0).Value = 0: Bar1(0).max = 100: Bar1(0).Visible = True
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute("select par_valor from a_param where par_codigo = 'version'")
   
   If Not RS.EOF Then
      
      If Val(RS!par_valor) < 189 Then
         
         RS.Close: Set RS = Nothing
         vg_db.Execute "sgp_s_bkpsgp '" & dir_bkpsql & "', '" & vg_SqlBase & "'"
      
      ElseIf Val(RS!par_valor) > 188 Then
         
         RS.Close: Set RS = Nothing
         If RS.State = 1 Then RS.Close
         RS.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         Set RS = vg_db.Execute("sgp_s_bkpsgp '" & dir_bkpsql & "', '" & vg_SqlBase & "'")
         
         If Not RS.EOF Then
            
            If RS(0) > 0 Then
               
               MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Msgtitulo
            
            End If
         
         End If
         
         RS.Close: Set RS = Nothing
      
      End If
   
   Else
      
      RS.Close: Set RS = Nothing
   
   End If
   
   AZ1_Status dir_bkpsql & vg_SqlBase & Format(Date, "yyyymmdd") & ".zip", 0
   AZ1.CreateZip dir_bkpsql & vg_SqlBase & Format(Date, "yyyymmdd") & ".zip", ""
   AZ1.AddFile dir_bkpsql & vg_SqlBase & ".bak", "", True, ""
   AZ1.Close
   Kill dir_bkpsql & vg_SqlBase & ".bak"

   If Dir(dir_bkpsql & vg_SqlBase & Format(Date, "yyyymmdd") & ".zip") <> "" Then
      
      fso.MoveFile dir_bkpsql & vg_SqlBase & Format(Date, "yyyymmdd") & ".zip", dir_trabajo & "backup\" & vg_SqlBase & Format(Date, "yyyymmdd") & ".zip"
   
   End If

End If

vg_db.Close
Bar1(0).Visible = False: Label1.Visible = False
Me.Hide
Unload Me

Exit Sub
Man_Error:

Select Case Err

Case 3049, 3204
    
    DoEvents
    MsgBox "El sistema esta respaldando información por otro usuario. " & Chr(13) & Chr(13) & "Inténtelo en unos minutos mas tarde.", vbExclamation + vbOKOnly, Msgtitulo: End

Case 35764
    
    DoEvents
    For i = 1 To 1000000
    
    Next i
    Resume

Case 76 Or -2147217900
    
    Resume Next
    Exit Sub

Case 58, 53
   
   Resume Next
   Exit Sub

Case -2147217900
   
   vg_db.Close
   Bar1(0).Visible = False: Label1.Visible = False
   Me.Hide
   Unload Me
   Exit Sub

Case -2147467259
    
    MsgBox "El sistema esta respaldando información por otro usuario. " & Chr(13) & Chr(13) & "Inténtelo en unos minutos mas tarde.", vbExclamation + vbOKOnly, Msgtitulo: End
'    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
'    End
'    Exit Sub

End Select

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Form_Load()

If vg_tipbase = "1" Then
   
   Msgtitulo = "Backup Base de Datos Access"
   Me.Caption = "Backup Base de Datos Access"

Else
   
   Msgtitulo = "Backup Base de Datos Sql Server"
   Me.Caption = "Backup Base de Datos Sql Server"

End If

End Sub

