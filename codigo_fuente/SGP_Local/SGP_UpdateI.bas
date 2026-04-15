Attribute VB_Name = "SGP_UpdateI"
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

    Dim SFTP              As New ChilkatSFtp
    Dim Success           As Long
    Dim RS                As New ADODB.Recordset
    Dim RS1               As New ADODB.Recordset
    Dim i                 As Long
    Dim Largo             As Long
    Dim RutaActualizacion As String
    Dim VersionSgpUpdate  As Long
    Dim fso               As Object
    Dim Descarga()        As String
    Dim Url               As String
    Dim Puerto            As Long
    Dim HostName          As String
    Dim UserName          As String
    Dim Password          As String
    Dim DirDescarga       As String
    Dim NombreArchivo     As String
    Dim NombreArchivoEti  As String
    Dim AuxCasino         As String
    Dim EstadoError       As String
    
'    Dim oFTP              As New chilkatftp
'    Dim obj As Object
'    Set obj = CreateObject("chilkatftp")
'    Dim oFTP  As CHILKATFTPLibCtl.ChilkatFTP
    'Set oFTP = New chilkatftp 'CreateObject("CHILKATFTPLibCtl")
    'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
    Set fso = CreateObject("Scripting.FileSystemObject")
    
   vg_SqlNSvr = MiFunc("SQL SERVER", "Gestion.ini", "Servidor")
   vg_SqlBase = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "DataBase"), ""))
   vg_SqlNUsr = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Usuario"), ""))
   vg_SqlPass = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Password"), ""))
   
   AbrirBase
    
    'Traer parametro actualización
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

'    Set RS = vg_db.Execute("select isnull(par_valor,'') as par_valor from a_param where par_codigo = 'DescargaR'")
'    If Not RS.EOF Then
'
'       If Trim(RS!par_valor) <> "" Then
'
'          Descarga = Split(Trim(RS!par_valor), ";")
'          RutaActualizacion = Descarga(3)
'          VersionSgpUpdate = Descarga(4)
'
'       End If
'
'    Else
'
'        Exit Sub
'
'    End If
'    RS.Close
'    Set RS = Nothing
'
'    Url = RutaActualizacion
'    Largo = Len(RutaActualizacion)
'    NombreArchivo = ""
'
'    'Sacar nombre de archivo
'    For i = Largo To 1 Step -1
'
'        If Mid$(RutaActualizacion, i, 1) = "/" Then Exit For
'
'        NombreArchivo = Mid$(RutaActualizacion, i, 1) & NombreArchivo
'
'    Next i
    
    
    Set RS = vg_db.Execute("SELECT distinct isnull(par_cencos, '') par_cencos, isnull(par_codigo, '') par_codigo, isnull(par_valor, '') as par_valor " & _
                           "FROM a_param as a with (nolock) " & _
                           "inner join b_clientes as b with (nolock) on a.par_cencos = b.cli_codigo " & _
                           "                                        and b.cli_tipo = 0 AND b.cli_codbod > 0 and b.cli_activo = 1 " & _
                           "WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("ftp")) & "%' order by par_cencos, par_codigo")
    If RS.EOF Then
       
       fg_descarga
       RS.Close
       Set RS = Nothing
       Exit Sub
    
    End If
    
    Dim success1 As Long
    
    AuxCasino = ""
    Puerto = 0
    HostName = ""
    UserName = ""
    Password = ""
    DirDescarga = ""
    NombreArchivo = ""
    NombreArchivoEti = ""
    
    Do While Not RS.EOF
          
          If AuxCasino <> RS!par_cencos Then

            If Trim(AuxCasino) <> "" _
               And Trim(Puerto) <> "" And Trim(HostName) <> "" _
               And Trim(UserName) <> "" And Trim(Password) <> "" _
               And Trim(DirDescarga) <> "" And Trim(NombreArchivo) <> "" Then
                       
               ' Set some timeouts, in milliseconds:
               SFTP.ConnectTimeoutMs = 5000
               SFTP.IdleTimeoutMs = 15000
               
               ' Connect to the SSH server.
               ' The standard SSH port = 22
               ' The hostname may be a hostname or IP address.
               Success = SFTP.Connect(HostName, Puerto)
               If (Success <> 1) Then
                    
'                    MsgBox sftp.LastErrorText
                   SFTP.Disconnect
                   Puerto = 0
                   HostName = ""
                   UserName = ""
                   Password = ""
                   DirDescarga = ""
                   NombreArchivo = ""
                
                End If
               
                ' Authenticate with the SSH server.  Chilkat SFTP supports
                ' both password-based authenication as well as public-key
                ' authentication.  This example uses password authenication.
                Success = SFTP.AuthenticatePw(UserName, Password)
                If (Success = 1) Then
                    
                  SFTP.Disconnect
                  Exit Do
                
                Else
                
                   SFTP.Disconnect
                   Puerto = 0
                   HostName = ""
                   UserName = ""
                   Password = ""
                   DirDescarga = ""
                   NombreArchivo = ""

                End If
                            
            End If
             
            AuxCasino = RS!par_cencos
          
          End If
          
          If RS!par_codigo = "ftpdirp" Then
          
             DirDescarga = fg_Desencripta(TipoDato(RS!par_valor, ""))
                 
          End If
          
          If RS!par_codigo = "ftpser" Then
          
             HostName = fg_Desencripta(TipoDato(RS!par_valor, ""))
          
          End If
          
          If RS!par_codigo = "ftpusu" Then
          
             UserName = fg_Desencripta(TipoDato(RS!par_valor, ""))
          
          End If
          
          If RS!par_codigo = "ftppas" Then
          
             Password = fg_Desencripta(TipoDato(RS!par_valor, ""))
          
          End If
          
          If RS!par_codigo = "ftppue" Then
          
             Puerto = fg_Desencripta(TipoDato(RS!par_valor, ""))
                     
          End If
          
          If RS!par_codigo = "ftpnarchp" Then
          
             NombreArchivo = fg_Desencripta(TipoDato(RS!par_valor, ""))
                     
          End If
          
          If RS!par_codigo = "ftpnarchle" Then
          
             NombreArchivoEti = fg_Desencripta(TipoDato(RS!par_valor, ""))
                     
          End If
          
          RS.MoveNext
       
    Loop
       
    RS.Close
    Set RS = Nothing
    
    If Trim(HostName) = "" Then
    
       fg_descarga
       Exit Sub
    
    End If
    
    If Trim(DirDescarga) = "" Then
    
       fg_descarga
       Exit Sub
    
    End If
    
    If Trim(UserName) = "" Then
    
       fg_descarga
       Exit Sub
    
    End If
    
    If Trim(Password) = "" Then
    
       fg_descarga
       Exit Sub
       
    End If
    
    If Trim(Puerto) = "" Then
    
       fg_descarga
       Exit Sub

    End If
    
    If Trim(NombreArchivo) = "" Then
    
       fg_descarga
       Exit Sub

    End If
    
    If Trim(NombreArchivoEti) = "" Then
    
       fg_descarga
       Exit Sub

    End If
    
    DoEvents
    
    If Not isNetwork(NETWORK_ALIVE_LAN) Then
    
       Exit Sub
       'MsgBox "No hay conexión a internet, intentelo mas tarde. Proceso cancelado", vbCritical + vbOKOnly, MsgTitulo: End
    
    End If
    
    M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "SGP_UPDATE\", "*.*", 0
    
    'validar si existe NombreArchivo se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo) Then
   
       'Borrar archivo zip
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo)
    
    End If
    
    'validar si existe push.exe, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & "Push.exe") Then
    
       'Borrar archivo Push.exe
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & "Push.exe")
    
    End If
    
    'LLamamos a la APi de Win, y le pasamos los parametros, Url y la ruta de descarga ("C:/....")
''    Call URLDownloadToFile(0, Url, "c:\Temp\sgp.zip", 0, 0)
'    Call URLDownloadToFile(0, Url, dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo, 0, 0)

    ' Set some timeouts, in milliseconds:
    SFTP.ConnectTimeoutMs = 5000
    SFTP.IdleTimeoutMs = 15000

    ' Connect to the SSH server.
    ' The standard SSH port = 22
    ' The hostname may be a hostname or IP address.
    Success = SFTP.Connect(HostName, Puerto)
    If (Success <> 1) Then
        
       fg_descarga
       SFTP.Disconnect
       Exit Sub
    
    End If

    ' Authenticate with the SSH server.  Chilkat SFTP supports
    ' both password-based authenication as well as public-key
    ' authentication.  This example uses password authenication.
    Success = SFTP.AuthenticatePw(UserName, Password)
    If (Success <> 1) Then
        
       fg_descarga
       SFTP.Disconnect
       Exit Sub
    
    End If
    
    ' After authenticating, the SFTP subsystem must be initialized:
    Success = SFTP.InitializeSftp()
    If (Success <> 1) Then
        
       fg_descarga
       SFTP.Disconnect
       Exit Sub
            
    End If
    
    Dim LocalFileName As String
    LocalFileName = dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo '"ActualizadorSGP.zip" '"c:\temp\ActualizadorSGP.zip"
    Dim RemoteFileName As String
    RemoteFileName = NombreArchivo

    ' Download the file:
    ' Note: The remote filepath may be an absolute filepath,
    ' a relative filepath, or simply a filename.
    ' Relative filepaths are always relative to the home directory
    ' of the SFTP/SSH user account.  There is no such thing
    ' as "current remote directory" in the SFTP protocol.
    ' A filename with no path implies that the file is located
    ' in the SFTP user account's home directory.
    Success = SFTP.DownloadFileByName(DirDescarga & "/" & RemoteFileName, LocalFileName)
    If (Success <> 1) Then
    
       fg_descarga
       SFTP.Disconnect
       Exit Sub
            
    End If

    SFTP.Disconnect


'    'conectar a sitio ftp
'    P_Push.oFTP.Port = Puerto
'    P_Push.oFTP.UseIEProxy = False
'    P_Push.oFTP.HostName = HostName
'    P_Push.oFTP.UserName = UserName
'    P_Push.oFTP.password = password
'    P_Push.oFTP.Passive = 1
'
'    Dim Success As Long
'    Success = P_Push.oFTP.Connect()
'    If (Success <> 1) Then
'
'       fg_descarga
'       P_Push.oFTP.Disconnect
'       Exit Sub
'
'    End If
'
''     Change to the remote directory where the file is located.
''     This step is only necessary if the file is not in the home directory
''     of the FTP account.
'    Success = P_Push.oFTP.ChangeRemoteDir(DirDescarga)
'    If (Success <> 1) Then
'
'       fg_descarga
'       P_Push.oFTP.Disconnect
'       Exit Sub
'
'    End If
'
'    Dim LocalFileName As String
'    LocalFileName = dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo
'    Dim RemoteFileName As String
'    RemoteFileName = NombreArchivo
'
'    ' Download the file.
'    Success = P_Push.oFTP.GetFile(RemoteFileName, LocalFileName)
'    If (Success <> 1) Then
'
'       fg_descarga
'       P_Push.oFTP.Disconnect
'       Exit Sub
'
'    End If
'
'    P_Push.oFTP.Disconnect
    
    EstadoError = "1"
    
    'validar si existe Push.exe, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo) Then
    
      'Una ves descargado lo abrimos de forma Binaria
      M_Backup.AZ1.OpenZip dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivo
      
      For i = 0 To M_Backup.AZ1.FileCount
      
          M_Backup.AZ1.ExtractFile M_Backup.AZ1.FileName(i), dir_trabajo_Inf & "SGP_UPDATE\", ""
      
      Next i
      
      M_Backup.AZ1.Close

    Else
    
        Exit Sub
        
    End If
    
    Dim NomRenombradoSgp_Update As String
    
    'validar si existe Push.exe versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & "Push.exe") Then
    
       'Generar backup Push.exe
       NomRenombradoSgp_Update = fg_Archivo(dir_trabajo_Inf, "Push.exe")
       
       If fso.FileExists(dir_trabajo_Inf & "Push.exe") Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "Push.exe" As NomRenombradoSgp_Update
       
       End If
       
    End If
    
    'validar si existe Push.exe, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & "Push.exe") Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & "Push.exe", dir_trabajo_Inf)
     
       'Borrar archivo Push.exe
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & "Push.exe")
    
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
      
   '-------
   '------- Inicio : Descargar Etiquetado Recetas
   '-------
   
    Dim Azucares As String
    Dim Calorias As String
    Dim Grasas   As String
    Dim Sodio    As String
    Dim Logo     As String
    
    If Not isNetwork(NETWORK_ALIVE_LAN) Then
    
       Exit Sub
    
    End If
    
    ' Mover versión SGP_Update
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    Set RS = vg_db.Execute("sgp_Sel_NomEtiquetadoNutricional")

    If Not RS.EOF Then

       Do While Not RS.EOF

          Azucares = RS(0)
          Calorias = RS(1)
          Grasas = RS(2)
          Sodio = RS(3)
          Logo = RS(4)
          
          RS.MoveNext

       Loop

    Else
    
       RS.Close
       Set RS = Nothing
       Exit Sub
       
    End If
    RS.Close
    Set RS = Nothing
    
    EstadoError = "2"
    M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "SGP_UPDATE\", "*.*", 0
    
    'validar si existe NombreArchivo se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti) Then
   
       'Borrar archivo zip
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti)
    
    End If
    
    'validar si existe Azucares, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Azucares) Then
    
       'Borrar archivo Azucares
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Azucares)
    
    End If
       
    'validar si existe Calorias, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Calorias) Then
    
       'Borrar archivo Calorias
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Calorias)
    
    End If
       
    'validar si existe Grasas, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Grasas) Then
    
       'Borrar archivo Grasas
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Grasas)
    
    End If
       
    'validar si existe Sodio, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Sodio) Then
    
       'Borrar archivo Sodio
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Sodio)
    
    End If
    
    'validar si existe Logo, se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Logo) Then
    
       'Borrar archivo Logo
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Logo)
    
    End If
       
    ' Set some timeouts, in milliseconds:
    SFTP.ConnectTimeoutMs = 5000
    SFTP.IdleTimeoutMs = 15000

    ' Connect to the SSH server.
    ' The standard SSH port = 22
    ' The hostname may be a hostname or IP address.
    Success = SFTP.Connect(HostName, Puerto)
    If (Success <> 1) Then
        
       fg_descarga
       SFTP.Disconnect
       Exit Sub
    
    End If

    ' Authenticate with the SSH server.  Chilkat SFTP supports
    ' both password-based authenication as well as public-key
    ' authentication.  This example uses password authenication.
    Success = SFTP.AuthenticatePw(UserName, Password)
    If (Success <> 1) Then
        
       fg_descarga
       SFTP.Disconnect
       Exit Sub
    
    End If
    
    ' After authenticating, the SFTP subsystem must be initialized:
    Success = SFTP.InitializeSftp()
    If (Success <> 1) Then
        
       fg_descarga
       SFTP.Disconnect
       Exit Sub
            
    End If
    
    Dim LocalFilenameEti As String
    LocalFilenameEti = dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti
    Dim RemoteFileNameEti As String
    RemoteFileNameEti = NombreArchivoEti

    ' Download the file:
    ' Note: The remote filepath may be an absolute filepath,
    ' a relative filepath, or simply a filename.
    ' Relative filepaths are always relative to the home directory
    ' of the SFTP/SSH user account.  There is no such thing
    ' as "current remote directory" in the SFTP protocol.
    ' A filename with no path implies that the file is located
    ' in the SFTP user account's home directory.
    Success = SFTP.DownloadFileByName(DirDescarga & "/" & RemoteFileNameEti, LocalFilenameEti)
    If (Success <> 1) Then
    
       fg_descarga
       SFTP.Disconnect
       Exit Sub
            
    End If

    SFTP.Disconnect
    
    
'    'conectar a sitio ftp
'    P_Push.oFTP.Port = Puerto
'    P_Push.oFTP.UseIEProxy = False
'    P_Push.oFTP.HostName = HostName
'    P_Push.oFTP.UserName = UserName
'    P_Push.oFTP.password = password
'    P_Push.oFTP.Passive = 1
'
'    Dim success2 As Long
'    success2 = P_Push.oFTP.Connect()
'    If (success2 <> 1) Then
'
'       fg_descarga
'       P_Push.oFTP.Disconnect
'       Exit Sub
'
'    End If
'
''     Change to the remote directory where the file is located.
''     This step is only necessary if the file is not in the home directory
''     of the FTP account.
'    success2 = P_Push.oFTP.ChangeRemoteDir(DirDescarga)
'    If (success2 <> 1) Then
'
'       fg_descarga
'       P_Push.oFTP.Disconnect
'       Exit Sub
'
'    End If
'
'    Dim LocalFilenameEti As String
'    LocalFilenameEti = dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti
'    Dim RemoteFileNameEti As String
'    RemoteFileNameEti = NombreArchivoEti
'
'    ' Download the file etiquetado.
'    success2 = P_Push.oFTP.GetFile(RemoteFileNameEti, LocalFilenameEti)
'    If (success2 <> 1) Then
'
'       fg_descarga
'       P_Push.oFTP.Disconnect
'       Exit Sub
'
'    End If
'
'    P_Push.oFTP.Disconnect
    
    'validar si existe sello etiquetado, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti) Then
    
      'Una ves descargado lo abrimos de forma Binaria
      M_Backup.AZ1.OpenZip dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti
      
      For i = 0 To M_Backup.AZ1.FileCount
      
          M_Backup.AZ1.ExtractFile M_Backup.AZ1.FileName(i), dir_trabajo_Inf & "SGP_UPDATE\", ""
      
      Next i
      
      M_Backup.AZ1.Close

    Else
    
        Exit Sub
        
    End If
    
    Dim NomRenombradoAzucares As String
    Dim NomRenombradoCalorias As String
    Dim NomRenombradoGrasas As String
    Dim NomRenombradoSodio As String
    Dim NomRenombradoLogo As String
    
    'validar si existe Azucares versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Azucares) Then
    
       'Generar backup Azucares
       NomRenombradoAzucares = fg_Archivo(dir_trabajo_Inf & "Etiquetado\", Azucares)
       
       If fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Azucares) Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "Etiquetado\" & Azucares As NomRenombradoAzucares
       
       End If
       
    End If
    
    'validar si existe Azucares, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Azucares) Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & Azucares, dir_trabajo_Inf & "Etiquetado\")
     
       'Borrar archivo Azucares
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Azucares)
    
    End If
    
    'Borrar archivos renombrado
    If fso.FileExists(NomRenombradoAzucares) Then
    
       'Borrar versión antigua
        Kill (NomRenombradoAzucares)
    
    End If
   
    'validar si existe Calorias versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Calorias) Then
    
       'Generar backup Calorias
       NomRenombradoCalorias = fg_Archivo(dir_trabajo_Inf & "Etiquetado\", Calorias)
       
       If fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Calorias) Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "Etiquetado\" & Calorias As NomRenombradoCalorias
       
       End If
       
    End If
    
    'validar si existe Calorias, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Calorias) Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & Calorias, dir_trabajo_Inf & "Etiquetado\")
     
       'Borrar archivo Calorias
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Calorias)
    
    End If
    
    'Borrar archivos renombrado Calorias
    If fso.FileExists(NomRenombradoCalorias) Then
    
       'Borrar versión antigua
        Kill (NomRenombradoCalorias)
    
    End If
    
    'validar si existe Grasas versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Grasas) Then
    
       'Generar backup Grasas
       NomRenombradoGrasas = fg_Archivo(dir_trabajo_Inf & "Etiquetado\", Grasas)
       
       If fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Grasas) Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "Etiquetado\" & Grasas As NomRenombradoGrasas
       
       End If
       
    End If
    
    'validar si existe Grasas, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Grasas) Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & Grasas, dir_trabajo_Inf & "Etiquetado\")
     
       'Borrar archivo Grasas
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Grasas)
    
    End If
    
    'Borrar archivos renombrado Grasas
    If fso.FileExists(NomRenombradoGrasas) Then
    
       'Borrar versión antigua
        Kill (NomRenombradoGrasas)
    
    End If
    
    'validar si existe Sodio versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Sodio) Then
    
       'Generar backup Sodio
       NomRenombradoSodio = fg_Archivo(dir_trabajo_Inf & "Etiquetado\", Sodio)
       
       If fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Sodio) Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "Etiquetado\" & Sodio As NomRenombradoSodio
       
       End If
       
    End If
    
    'validar si existe Sodio, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Sodio) Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & Sodio, dir_trabajo_Inf & "Etiquetado\")
     
       'Borrar archivo Sodio
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Sodio)
    
    End If
    
    'Borrar archivos renombrado Sodio
    If fso.FileExists(NomRenombradoSodio) Then
    
       'Borrar versión antigua
        Kill (NomRenombradoSodio)
    
    End If
    
    'validar si existe Logo versión antigua, si existe borrar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Logo) Then
    
       'Generar backup Logo
       NomRenombradoSodio = fg_Archivo(dir_trabajo_Inf & "Etiquetado\", Logo)
       
       If fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Logo) Then
       
          'renombrar nuevo archivo x archivo original
          Name dir_trabajo_Inf & "Etiquetado\" & Logo As NomRenombradoLogo
       
       End If
       
    End If
    
    'validar si existe Logo, si existe copiar
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & Logo) Then
    
       'Se envía el path origen y path destino
       Call Copiar_Archivo(dir_trabajo_Inf & "SGP_UPDATE\" & Logo, dir_trabajo_Inf & "Etiquetado\")
     
       'Borrar archivo Sodio
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & Logo)
    
    End If
    
    'Borrar archivos renombrado Logo
    If fso.FileExists(NomRenombradoLogo) Then
    
       'Borrar versión antigua
        Kill (NomRenombradoLogo)
    
    End If
    
    'validar si existe nombrearchivo se borra
    If fso.FileExists(dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti) Then
   
       'Borrar archivo zip
       Kill (dir_trabajo_Inf & "SGP_UPDATE\" & NombreArchivoEti)
    
    End If
   
   '-------
   '------- Fin : Descargar Etiquetado Recetas
   '-------
   
   EstadoError = "0"
   Set fso = Nothing
   
   vg_db.Close
         
Exit Sub
Man_Error:

    If EstadoError = "1" Then
    
       'Devolver todo atras en caso de error
       If fso.FileExists(NomRenombradoSgp_Update) And fso.FileExists(dir_trabajo_Inf & "Push.exe") Then
       
          'Borrar versión nueva por error
          Kill (dir_trabajo_Inf & "Push.exe")
       
          'renombrar versión antigua
          Name NomRenombradoSgp_Update As dir_trabajo_Inf & "Push.exe"
    
       ElseIf fso.FileExists(NomRenombradoSgp_Update) And Not fso.FileExists(dir_trabajo_Inf & "Push.exe") Then
    
          'renombrar versión antigua
          Name NomRenombradoSgp_Update As dir_trabajo_Inf & "Push.exe"
       
       End If
    
    ElseIf EstadoError = "2" Then
    
       'Devolver todo atras en caso de error Azucares
       If fso.FileExists(NomRenombradoAzucares) And fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Azucares) Then
       
          'Borrar versión nueva por error Azucares
          Kill (dir_trabajo_Inf & "Etiquetado\" & Azucares)
       
          'renombrar versión antigua Azucares
          Name NomRenombradoAzucares As dir_trabajo_Inf & "Etiquetado\" & Azucares
    
       ElseIf fso.FileExists(NomRenombradoAzucares) And Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Azucares) Then
    
          'renombrar versión antigua Azucares
          Name NomRenombradoAzucares As dir_trabajo_Inf & "Etiquetado\" & Azucares
       
       End If
    
      'Devolver todo atras en caso de error Calorias
       If fso.FileExists(NomRenombradoCalorias) And fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Calorias) Then
       
          'Borrar versión nueva por error Calorias
          Kill (dir_trabajo_Inf & "Etiquetado\" & Calorias)
       
          'renombrar versión antigua Calorias
          Name NomRenombradoCalorias As dir_trabajo_Inf & "Etiquetado\" & Calorias
    
       ElseIf fso.FileExists(NomRenombradoCalorias) And Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Calorias) Then
    
          'renombrar versión antigua Calorias
          Name NomRenombradoCalorias As dir_trabajo_Inf & "Etiquetado\" & Calorias
       
       End If
    
       'Devolver todo atras en caso de error Grasas
       If fso.FileExists(NomRenombradoGrasas) And fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Grasas) Then
       
          'Borrar versión nueva por error Grasas
          Kill (dir_trabajo_Inf & "Etiquetado\" & Grasas)
       
          'renombrar versión antigua Grasas
          Name NomRenombradoGrasas As dir_trabajo_Inf & "Etiquetado\" & Grasas
    
       ElseIf fso.FileExists(NomRenombradoGrasas) And Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Grasas) Then
    
          'renombrar versión antigua Grasas
          Name NomRenombradoGrasas As dir_trabajo_Inf & "Etiquetado\" & Grasas
       
       End If
    
       'Devolver todo atras en caso de error Sodio
       If fso.FileExists(NomRenombradoSodio) And fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Sodio) Then
       
          'Borrar versión nueva por error Sodio
          Kill (dir_trabajo_Inf & "Etiquetado\" & Sodio)
       
          'renombrar versión antigua Sodio
          Name NomRenombradoSodio As dir_trabajo_Inf & "Etiquetado\" & Sodio
    
       ElseIf fso.FileExists(NomRenombradoSodio) And Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Sodio) Then
    
          'renombrar versión antigua Sodio
          Name NomRenombradoSodio As dir_trabajo_Inf & "Etiquetado\" & Sodio
       
       End If
    
       'Devolver todo atras en caso de error Logo
       If fso.FileExists(NomRenombradoLogo) And fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Logo) Then
       
          'Borrar versión nueva por error Logo
          Kill (dir_trabajo_Inf & "Etiquetado\" & Logo)
       
          'renombrar versión antigua Logo
          Name NomRenombradoSodio As dir_trabajo_Inf & "Etiquetado\" & Logo
    
       ElseIf fso.FileExists(NomRenombradoLogo) And Not fso.FileExists(dir_trabajo_Inf & "Etiquetado\" & Logo) Then
    
          'renombrar versión antigua Logo
          Name NomRenombradoSodio As dir_trabajo_Inf & "Etiquetado\" & Logo
       
       End If
    
    End If

Set fso = Nothing

vg_db.Close
'MsgBox "Falla en la descarga de actualización, intentelo mas tarde. Proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
'End

End Sub

' Subrutina que copia el archivo
Public Sub Copiar_Archivo(ByVal Origen As String, ByVal Destino As String)

Dim t_Op As SHFILEOPSTRUCT

    With t_Op
        
        .hWnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
        
    End With

    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
 
End Sub

