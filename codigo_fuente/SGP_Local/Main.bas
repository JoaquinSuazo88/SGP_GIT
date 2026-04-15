Attribute VB_Name = "Inicio"

Option Explicit

Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wmsg As Long, ByVal wparam As Long, lparam As Long) As Long
Declare Function IsNetworkAlive Lib "SENSAPI.DLL" (ByRef lpdwFlags As Long) As Long
Public Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal dwReserved As Long) As Boolean

Global ConnStr                  As String
Global vg_dir                   As String
Global Dir_Bmp                  As String
Global vg_DirLog                As String
Global BaseDeDato               As String
Global dir_trabajo              As String
Global dir_trabajo_Inf          As String
Global dir_bkpsql               As String
Global Provider                 As String
Global vg_db                    As ADODB.Connection
Global dbI                      As ADODB.Connection
Global dbE                      As ADODB.Connection
Global db                       As Database
Global MsgTitulo                As String
'Global vg_Area                  As Workspace
Global vg_ArchTxt               As String
Global VgLinea                  As String
Global ICGrilla                 As Integer
Global Tpag                     As Long
Global ws_respuesta             As String
Global IndErr                   As Integer
Global FechaErr                 As String
Global respuesta                As String
Global vg_CSep                  As String
Global vg_CDec                  As String
Global vg_DPr                   As Integer
Global vg_DCa                   As Integer
Global vg_RDCa                  As Integer
Global vg_RDC                   As String
Global vg_TDC                   As String
Global vg_FDC                   As String
Global vg_NDC                   As Long
Global vg_NSOL                  As Long
Global vg_Consulta              As String
Global vg_Archxls               As String
'-------> Variables de apertura de recordset
Global vg_ModoOpen             As Integer
Global vg_Acc()                As Variant
Global vg_Prev                 As Variant
Global vg_NUsr                 As String
Global vg_Pass                 As String
Global vg_CPer                 As Long
Global vg_OpcM                 As Long
Global vg_Para                 As String
Global vg_Base                 As String
Global vg_DSN                  As String
Global vg_SVR                  As String
Global vg_Login                As String
Global vg_filtippla            As Long
Global vg_filnomtippla         As String
Global vg_filcatdie            As Long
Global vg_filnomcatdie         As String
Global vg_Dig                  As String '-------> Indica si calcula digito verificador
Global vg_codigo               As String
Global vg_codigo2              As String
Global vg_codigo3              As String
Global vg_codigo4              As String
Global vg_Guias                As String '-------> Obtiene Guías de Despacho
Global vg_GuiasTipo            As String
Global vg_GuiaCD               As String 'Guias CD
Global vg_FechaEmision_GGD     As Date
Global vg_nombre               As String
Global vg_dbndecimal           As Integer
Global vg_fecha                As String
Global vg_auxfecha             As String
Global vg_newcodrec            As Long
Global vg_newnomrec            As String
Global vg_newestrec            As Boolean
Global vg_fecval               As Long
Global vg_codcasino            As String
Global vg_codregimen           As Long
Global vg_codservicio          As Long
Global vg_tipped               As Long
Global vg_anomes               As Long
Global vg_bodega               As Long
Global vg_codbod               As Long
Global vg_nombod               As String
Global vg_contra               As String
Global vg_nomcon               As String
Global vg_Aux                  As String
Global vg_csapiva              As String
Global vg_csapotros            As String
Global vg_claencsap            As Long
Global vg_cladetsap            As Long
Global vg_docexento            As String
Global vg_docafecto            As String
Global vg_tipmonsap            As String
Global vg_reporte              As String
Global vg_codreg               As String
Global vg_codser               As String
Global vg_invrot               As String
Global vg_ptoate               As String

Global vg_opcion               As Integer
Global vg_codreceta            As Long
Global vg_tiprec               As Long
Global vg_auxtiprec            As Long
Global vg_left                 As Long
Global vg_swpegreceta          As Integer
Global vg_swmovreceta          As Integer
Global vg_modprod              As Boolean
Global vg_modrec               As Boolean
Global vg_modprove             As Boolean
Global vg_5etapas              As Boolean
Global vg_modpac               As Boolean '-------> modulo paciente
Global vg_tipser               As Boolean '-------> varible identifica tipo de servicio
Global vg_opgra                As Long '-------> opción de tipo gráficos
Global vg_fecini               As Long '-------> fecha inicial
Global vg_fecfin               As Long '-------> fecha final
Global RutaGif                 As String
Global RS_Dato                 As New ADODB.Recordset
Global vg_ciedia               As String '-------> cierre de día
Global vg_op1                  As String
Global vg_op2                  As String
Global vg_op3                  As String
Global vg_tipbase              As String
Global vg_pais                 As String
Global vg_IDBloque             As Long
Global vg_SqlBase              As String
Global vg_SqlNSvr              As String
Global vg_SqlNUsr              As String
Global vg_SqlPass              As String
Global vg_RutaActualizacion    As String
Global Version                 As Long
Global VersionSGPSDX           As Long
Global VersionSGPSDXPar        As Long

Global Vg_FechaDesde           As String
Global Vg_FechaHasta           As String
Global Vg_MinSre               As Boolean
Global vg_tipmin               As Boolean
Global vg_bloenv               As Long
Global vg_Block_Botton_Actua_Receta_MVI As Boolean 'MVA - MVI - BLOQUEO DE BOTON ACTUALIZAR RECETA - 2013-01-18
Global vg_Clave_MVI As String 'MVA - MVI - BLOQUEO DE BOTON ACTUALIZAR RECETA - 2013-01-18
Global vg_bloqueo_opciones As Boolean
Global RSTempCheck As New ADODB.Recordset
Global RSTem As New ADODB.Recordset
Global RSinsert As New ADODB.Recordset

'---------------> Clases globales
Global RutinaLectura As New RutinaLectura
Global G_Proc        As New G_Proc
'---------------> 0 < ----------

Global lpApplicationName As String, _
          lpKeyName As String, _
          lpDefault As String, _
          lpReturnString As String, _
          lpFileName As String, _
          lpString As String
   
Global nSize As Long, _
       Valid As Long, _
       Path, _
       Succ
Dim Retorno As String

Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
              (ByVal lpApplicationName As String, _
               ByVal lpKeyName As Any, _
               ByVal lpDefault As String, _
               ByVal lpReturnedString As String, _
               ByVal nSize As Long, _
               ByVal lpFileName As String) As Long
Global cHost As String
Global cUser As String
Global cPass As String

Sub Main()

Dim RS As New ADODB.Recordset
Dim cpar As String
Dim xpar As String
Dim cOpt As String, cDato As String
Dim Err1 As String

On Error GoTo Man_Error

VerConfReg

vg_ModoOpen = dbSQLPassThrough
vg_tipbase = IIf(Trim(MiFunc("Tipo Base Dato", "Gestion.Ini", "Base")) = "", "1", MiFunc("Tipo Base Dato", "Gestion.Ini", "Base"))
dir_trabajo = MiFunc("Path", "Gestion.Ini", "Ruta")
BaseDeDato = MiFunc("Base_de_datos", "Gestion.Ini", "Mdb")
Provider = MiFunc("Provider", "Gestion.Ini", "Jet")
RutaGif = MiFunc("Gif", "Gestion.Ini", "Ruta")
dir_trabajo_Inf = Environ("PROGRAMFILES") & "\SGP\"
vg_DirLog = dir_trabajo & "LogoSdx.jpg"
vg_CSep = ","
vg_CDec = "."

If App.PrevInstance Then
    
    MsgBox App.EXEName + " ya se encuentra en ejecución. Ud. no puede mantener dos copias del mismo programa en memoria."
    End

End If

'-------> Crear directorio SGP
If Dir(dir_trabajo_Inf, vbDirectory) = "" Then

   MkDir dir_trabajo_Inf
   
End If
'-------> Fin crear directorio SGP

'-------> Crear directorio Etiquetado
If Dir(dir_trabajo_Inf & "\" & "Etiquetado", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "Etiquetado"
   
End If
'-------> Fin crear directorio Etiquetado

'-------> Crear directorio Actualizar
If Dir(dir_trabajo_Inf & "\" & "SGP_Update", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "SGP_Update"

End If

If isNetwork(NETWORK_ALIVE_LAN) And Not ConsultaProcess("Push.exe") Then

   SGP_Update

End If

'-------> Crear directorio Actualizar
If Dir(dir_trabajo_Inf & "\" & "Actualizar", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "Actualizar"
   
End If
'-------> Fin crear directorio Actualizar

'-------> Crear directorio CFC
If Dir(dir_trabajo_Inf & "\" & "Cfc", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "Cfc"

End If
'-------> Fin crear directorio CFC

'-------> Crear directorio Update Versión
If Dir(dir_trabajo_Inf & "\" & "Upd", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "Upd"
   
End If
'-------> Fin crear directorio Update Versión

'-------> Crear directorio ExcelSGP
If Dir(dir_trabajo_Inf & "\" & "ExcelSGP", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "ExcelSGP"
   
End If
'-------> Fin crear directorio Excel Versión

'-------> Crear directorio formatorequesicion
If Dir(dir_trabajo_Inf & "\" & "FormatoRequisicion", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "FormatoRequisicion"
   
End If
'-------> Fin crear directorio formato requesicion

'-------> Crear directorio guias logistico
If Dir(dir_trabajo_Inf & "\" & "GuiaLogistico", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "GuiaLogistico"
   
End If
'-------> Fin crear directorio guias logistico

'-------> Crear directorio carga masiva ventas diarias
If Dir(dir_trabajo_Inf & "\" & "SPRS_Plantilla_Carga_Masiva", vbDirectory) = "" Then

   MkDir dir_trabajo_Inf & "\" & "SPRS_Plantilla_Carga_Masiva"
   
End If
'-------> Fin crear directorio masiva ventas diarias

''-------> Borrar Backup dejar solo los ultimos 10
'If Dir(dir_trabajo & "\" & "Backup", vbDirectory) = "Backup" Then
'
'End If

Dim Fecha As Date, diasem As String, estbac As Boolean
If vg_tipbase = "1" Then
   
   '-------> Verificar si existe proceso activo sgpsdx
    If ConsultaProcess("sgpsdx.exe") Then
       
       KillProcess ("sgpsdx.exe")
    
    End If
    '-------> Crear directorio Backup Access
    If Dir(dir_trabajo & "\" & "Backup", vbDirectory) = "" Then
    
       MkDir dir_trabajo & "\" & "Backup"
       
    End If
    
    'If fg_NumDia(Trim(Left(fg_Fecha_Dia(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(Day(Now)), 2), 2), Len(fg_Fecha_Dia(Year(Now) & fg_pone_cero(Str(Month(Now)), 2) & fg_pone_cero(Str(Day(Now)), 2), 2)) - 2))) = 2 Then
    Fecha = Format(Date - DatePart("w", Date, 2) + 1, "dd/mm/yyyy"): diasem = DatePart("ww", Fecha, 2)
    estbac = False
    
    Do While DatePart("ww", Fecha, 2) = diasem
    
    '   If Dir(dir_trabajo & "Backup\" & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip") = "" Then M_Backup.Show 1: Exit Do
       If Dir(dir_trabajo & "Backup\" & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Fecha, "yyyymmdd") & ".zip") <> "" Then estbac = True: Exit Do
       Fecha = Fecha + 1
    
    Loop
    
    If Not estbac Then If Dir(dir_trabajo & "Backup\" & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 4)) & Format(Date, "yyyymmdd") & ".zip") = "" Then M_Backup.Show 1
    '-------> Fin crear directorio Backup
Else
   
   'fg_Desencripta (TipoDato(RS!par_valor, ""))
   '-------> Verificar si existe proceso activo sgpsdx
   If ConsultaProcess("sgpsdx.exe") Then
      
      KillProcess ("sgpsdx.exe")
   
   End If
   
   vg_SqlNSvr = MiFunc("SQL SERVER", "Gestion.ini", "Servidor")
   vg_SqlBase = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "DataBase"), ""))
   vg_SqlNUsr = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Usuario"), ""))
   vg_SqlPass = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Password"), ""))
   AbrirBase
   dir_bkpsql = "C:\bkpsgp\"
   '-------> Crear directorio Backup Access
   If Dir(dir_trabajo & "\" & "Backup", vbDirectory) = "" Then
   
      MkDir dir_trabajo & "\" & "Backup"
      
   End If
   
   If Dir(dir_bkpsql, vbDirectory) = "" Then
      
      MkDir dir_bkpsql
   
   End If
   
   Fecha = Format(Date - DatePart("w", Date, 2) + 1, "dd/mm/yyyy")
   diasem = DatePart("ww", Fecha, 2)
   estbac = False
   
   Do While DatePart("ww", Fecha, 2) = diasem
      
      If Dir(dir_trabajo & "Backup\" & vg_SqlBase & Format(Fecha, "yyyymmdd") & ".zip") <> "" Then
      
         estbac = True
         Exit Do
         
      End If
      
      Fecha = Fecha + 1
   
   Loop
   
   If Not estbac Then
      
      '------->Validar PC Servidor
      If ValidaPCServidor Then
        
         If Dir(dir_trabajo & "Backup\" & vg_SqlBase & Format(Date, "yyyymmdd") & ".zip") = "" Then
         
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            Set RS = vg_db.Execute("sgp_Sel_ReducirLog '" & dir_bkpsql & "', '" & vg_SqlBase & "'")
         
            If Not RS.EOF Then
            
              If RS(0) > 0 Then
               
                 MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
            
              End If
         
            End If
         
            RS.Close: Set RS = Nothing
            
            
            vg_db.Close
            M_BorrarArchivos.InicioBorrado dir_trabajo & "Backup\", vg_SqlBase & "*.zip", 4
            
            M_Backup.Show 1
      
         End If
         
      End If
   
   End If
    '-------> Fin crear directorio Backup
    
End If

If Dir(dir_trabajo_Inf & "txt*.txt") <> "" Then Kill dir_trabajo_Inf & "txt*.txt"
If Dir(dir_trabajo_Inf & "reporte*.rtf") <> "" Then Kill dir_trabajo_Inf & "reporte*.rtf"

Open dir_trabajo & "sdxftp.ini" For Input As #1

Do While Not EOF(1)
    
    Input #1, cOpt, cDato
    
    Select Case cOpt
    
    Case "A"
        
        cHost = fg_Desencripta(cDato)
    
    Case "B"
        
        cUser = fg_Desencripta(cDato)
    
    Case "C"
        
        cPass = fg_Desencripta(cDato)
    
    End Select

Loop

Close #1
VgLinea = Chr(13) & Chr(10)  '-------> inserta retorno de carro a los mensages
V_Acceso.Refresh
V_Acceso.Show 1
AbrirBase
Err1 = "0"

If GetParametro("rprociedia") = "N" Or fg_Desencripta(GetParametro("fecrprodia")) = "" Then
   
   If Err1 = "0" Then M_ProCie.Show 1

End If

'-------> Insertar tipo base de datos tabla a_param
Set RS = vg_db.Execute("sgp_Sel_Param 1 , '" & vg_contra & "', 'tipobase'")

If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'tipobase', 'Tipo de Base Dato', 'N', '" & vg_tipbase & "', '" & vg_contra & "'")

Else

   vg_db.Execute ("sgp_Upd_Param 1, '" & vg_contra & "', 'tipobase', '', '', '" & vg_tipbase & "'")

End If
RS.Close: Set RS = Nothing

Partida.Show

Exit Sub
Man_Error:
If Err = 94 Then Err1 = "1"
If Err = 70 Or 94 Or 53 Then Resume Next
If Err = -2147217843 Then Resume Next
MsgBox Err & ":  " & error$(Err), vbCritical, "Mantención sistema SGP"

End Sub
