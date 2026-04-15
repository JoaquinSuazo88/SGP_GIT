Attribute VB_Name = "SGP_Update"
Option Explicit

'Funcion API URLDownloadToFile
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

'Variable donde colocaremos la version
Dim Version   As String
Dim MsgTitulo As String

' Estructura SHFILEOPSTRUCT o para usar con el Api
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

'Declaración Api SHFileOperation
Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                                (lpFileOp As SHFILEOPSTRUCT) As Long

'Constantes
Private Const FO_COPY = &H2
Private Const FOF_ALLOWUNDO = &H40

Sub SGP_Update()

On Error GoTo Man_Error

    Dim RS                As New ADODB.Recordset
    Dim RS1               As New ADODB.Recordset
    Dim i                 As Long
    Dim Largo             As Long
    Dim NombreArchivo     As String
    Dim RutaActualizacion As String
    Dim VersionSgpUpdate  As Long
    Dim fso               As Object
    Dim Descarga()        As String
    
    
    'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
   vg_SqlNSvr = MiFunc("SQL SERVER", "Gestion.ini", "Servidor")
   vg_SqlBase = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "DataBase"), ""))
   vg_SqlNUsr = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Usuario"), ""))
   vg_SqlPass = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Password"), ""))
   AbrirBase
    
    ' Traer parametro actualización
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("select isnull(par_valor,'') as par_valor from a_param where par_codigo = 'DescargaR'")
    If Not RS.EOF Then
   
       If Trim(RS!par_valor) <> "" Then
   
          Descarga = Split(Trim(RS!par_valor), ";")
          RutaActualizacion = Descarga(3)
          VersionSgpUpdate = Descarga(4)
 
       End If

    End If
    RS.Close
    Set RS = Nothing
    
    Largo = Len(RutaActualizacion)
    NombreArchivo = ""
    
    'Sacar nombre de archivo
    For i = Largo To 1 Step -1
    
        If Mid$(RutaActualizacion, i, 1) = "/" Then Exit For
        
        NombreArchivo = Mid$(RutaActualizacion, i, 1) & NombreArchivo
    
    Next i
    
    DoEvents
    For i = 1 To 1000000
    
    Next i
    
    If Not isNetwork(NETWORK_ALIVE_LAN) Then
    
       Exit Sub
       'MsgBox "No hay conexión a internet, intentelo mas tarde. Proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: End
    
    End If
    
    'validar si existe NombreArchivo se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo) Then
   
       'Borrar archivo zip
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo)
    
    End If
    
    'validar si existe sgp.exe, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & "SGP_Update.exe") Then
    
       'Borrar archivo sgp.exe
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & "SGP_Upate.exe")
    
    End If
    
    
    
    'LLamamos a la APi de Win, y le pasamos los parametros, Url y la ruta de descarga ("C:/....")
'    Call URLDownloadToFile(0, Url, "c:\Temp\sgp.zip", 0, 0)
    Call URLDownloadToFile(0, Url, dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo, 0, 0)

    'validar si existe sgp.exe, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo) Then
    
      'Una ves descargado lo abrimos de forma Binaria
      AZ1.OpenZip dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo
      
      For i = 0 To AZ1.FileCount
      
          AZ1.ExtractFile AZ1.FileName(i), dir_trabajo_Inf & "SGP_UPDATE\", ""
      
      Next i
      
      AZ1.Close

    Else
    
'       MsgBox "Falla en la descarga de actualización, intentelo mas tarde. Proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: End
    
        Exit Sub
        
    End If
    
    Dim NomRenombradoSgp_Update As String
    
    'validar si existe sgp.exe versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & "SGP_Update.exe") Then
    
       'Generar backup sgp.exe
       NomRenombradoSgp_Update = fg_Archivo(dir_trabajo_Inf, "SGP_Update.exe")
       
       If fso.FileExists(dir_trabajo_Inf & "SGP_Update.exe") Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "SGP_Update.exe" As NomRenombradoSgp_Update
       
       End If
       
    End If
    
    'validar si existe sgp.exe, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & "SGP_Update.exe") Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & "SGP_Update.exe", dir_trabajo_Inf)
     
       'Borrar archivo sgp.exe
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & "SGP_Update.exe")
    
    End If
    
    'validar si existe nombrearchivo se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo) Then
   
       'Borrar archivo zip
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo)
    
    End If

    
    ' Mover versión SGP_Update
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("SELECT distinct * FROM b_clientes as bc with (nolock) WHERE bc.cli_tipo = 0 AND bc.cli_codbod > 0 and bc.cli_activo = 1")

    If Not RS.EOF Then

       Do While Not RS.EOF

          If RS1.State = 1 Then RS1.Close
          RS1.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient

          Set RS1 = vg_db.Execute("sgp_InsUpd_Param 'VersionSUP', 'Versión del Sistema SGP_Update', 'N', '" & VersionSgpUpdate & "', '" & RS!cli_codigo & "', '0'")

          If Not RS1.EOF Then

             If RS1(0) > 0 Then

                MsgBox RS1(0) & " " & RS1(1), vbCritical + vbOKOnly, MsgTitulo

                GoTo Man_Error:

              End If

          End If

          RS1.Close
          Set RS1 = Nothing

          RS.MoveNext

       Loop

    End If
    RS.Close
    Set RS = Nothing


    'Borrar archivos renombrado
    If fso.FileExists(NomRenombradoSgp_Update) Then
    
       'Borrar versión antigua
        Kill (NomRenombradoSgp_Update)
    
    End If
   
   Set fso = Nothing
    
   vg_db.Close

   'MsgBox ("Fin del proceso, se actualizo correctamente...")
   'End
          
Exit Sub
Man_Error:

    'Devolver todo atras en caso de error

    If fso.FileExists(NomRenombradoSgp_Update) And fso.FileExists(dir_trabajo_Inf & "SGP_Update.exe") Then
    
       'Borrar versión nueva por error
       Kill (dir_trabajo_Inf & "SGP_Update.exe")
    
       'renombrar versión antigua
       Name NomRenombradoSgp_Update As dir_trabajo_Inf & "SGP_Update.exe"
 
    ElseIf fso.FileExists(NomRenombradoSgp_Update) And Not fso.FileExists(dir_trabajo_Inf & "SGP_Update.exe") Then
 
       'renombrar versión antigua
       Name NomRenombradoSgp_Update As dir_trabajo_Inf & "SGP_Update.exe"
    
    End If
    

Set fso = Nothing

vg_db.Close
'MsgBox "Falla en la descarga de actualización, intentelo mas tarde. Proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
'End

End Sub

