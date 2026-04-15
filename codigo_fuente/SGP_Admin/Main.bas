Attribute VB_Name = "Inicio"
Option Explicit
Option Compare Text

Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wmsg As Long, ByVal wparam As Long, lparam As Long) As Long
Global ConnStr                  As String
Global vg_dir                   As String
Global Dir_Bmp                  As String
Global vg_DirLog                As String
Global BaseDeDato               As String
Global dir_trabajo              As String
Global mdirpc                   As String
Global Provider                 As String
Global vg_db                    As ADODB.Connection
Global vg_dbpedweb              As ADODB.Connection
Global vg_dbtec                 As ADODB.Connection
Global vg_dbsac                 As ADODB.Connection
Global dbi                      As ADODB.Connection
Global db                       As Database
Global db1                      As Database
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
Global vg_RowEnd                As Long

' *** Variables de apertura de recordset *** '
Global vg_ModoOpen             As Integer
Global vg_Acc()                As Variant
Global vg_Prev                 As Variant
Global vg_NSvr                 As String
Global vg_NUsr                 As String
Global vg_Pass                 As String
Global vg_SqlBase              As String
Global vg_SqlNSvr              As String
Global vg_SqlNUsr              As String
Global vg_SqlPass              As String

Global vg_SqlBaseW             As String
Global vg_SqlNSvrW             As String
Global vg_SqlNUsrW             As String
Global vg_SqlPassW             As String
Global vg_estopen              As Boolean

Global vgtec_NSvr              As String
Global vgtec_NUsr              As String
Global vgtec_Pass              As String
'-------> Variables Sac
Global vgsac_NSvr              As String
Global vgsac_NUsr              As String
Global vgsac_Pass              As String

Global vg_VAccess              As String
Global vg_CPer                 As Long
Global vg_OpcM                 As Long
Global vg_Para                 As String
Global vg_Base                 As String
Global vg_DSN                  As String
Global vg_SVR                  As String
Global vg_Login                As String
Global vg_filtippla            As Long
Global vg_filtipplaMin         As Long
Global vg_filnomtippla         As String
Global vg_filcatdie            As Long
Global vg_filcatdieMin         As Long
Global vg_filnomcatdie         As String
Global vg_Dig                  As String 'Indica si calcula digito verificador
Global vg_codigo               As String
Global Vg_Codigo2              As String
Global Vg_Codigo3              As String
Global Vg_Codigo4              As String
Global Vg_Mes1                 As Long
Global Vg_Mes2                 As Long
Global Vg_Mes3                 As Long
Global Vg_Mes4                 As Long
Global vg_auxcod               As String
Global vg_anomes               As Long
Global vg_Guias                As String 'Obtiene Guías de Despacho
Global vg_nombre               As String
Global vg_Calorias             As Double
Global vg_Valor                As Double
Global vg_dbndecimal           As Integer
Global vg_fecha                As String
Global vg_auxfecha             As String
Global vg_newcodrec            As Long
Global vg_newnomrec            As String
Global vg_newestrec            As Boolean
Global vg_fecval               As Long
Global vg_cencos               As String
Global vg_IDBloque             As Long
Global vg_codcasino            As String
Global vg_codsubseg            As Long
Global vg_codregimen           As Long
Global vg_codservicio          As Long
Global vg_codlpr               As Long
Global vg_tipped               As Long
Global vg_tiprec               As Long
Global vg_ames                 As String
Global vg_tmisgp               As String ' Indicador que identifica el tipo minuta sgp casino
Global vg_Indppr               As String ' Indicador de Usuario
Global vg_IndpprSelec          As String ' Indicador de Selección visualización
Global vg_Zona                 As String ' Indicador de zona minuta
Global vg_ActCalorias          As Boolean
Global vg_RecetaReal           As Integer
Global vg_PartePlani           As Boolean
Global vg_nomsubseg            As String
Global vg_nomreg               As String
Global vg_nomser               As String
Global vg_pais                 As String
' Cantidad de minutos para el intervalo del timer _
  en este caso para 5 minutos
Global vg_IntMin               As Long
Global Vg_PlaSer               As String
Global vg_TemSeg               As Long

Private SheetName              As String
Private filepath               As String

Global vg_opcion               As Integer
Global vg_opimp                As Long
Global vg_codreceta            As Long
Global vg_left                 As Long
Global vg_swpegreceta          As Integer
Global vg_swmovreceta          As Integer
Global vg_modreceta            As Boolean
Global vg_GlosaEnvioCorreo     As String
Global RutaGif                 As String
Global SeleccTipoMinuta_MVI    As String 'MVA - MVI - VAR. PARA SABER EL TIPO DE SELEC. DE MINUTA EN INTERFAZ m_copia_min_seg
Global Sql_MVI                 As String 'MVA - MVI - VAR. PARA EJECUCION DE SQL
Global swEsCopia               As Boolean 'MVA - MVI - VAR. PARA SABER SI LA APLICACION CAE EN EL FORM O NO, POR TEMA DE OTRAS QUERIES COMPARTIDAS
Global swUp_CECO               As Boolean 'MVA - MVI - VAR. PARA EJECUCION GRAL.
Global swUp_REG                As Boolean 'MVA - MVI - VAR. PARA EJECUCION GRAL.
Global swUp_SERV               As Boolean 'MVA - MVI - VAR. PARA EJECUCION GRAL.
Global vg_BorradoDatos         As Boolean 'Estado borrado de datos

Global swEsMinBloque           As Boolean 'MVA - MVI - VAR. PARA EJECUCION GRAL.
Global vg_opcionmenubloque    As String
Public valida As String
Public pedidos As Integer

Global RS_Dato                 As New ADODB.Recordset
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
       
Global Vg_FechaDesde           As String
Global Vg_FechaHasta           As String
'---------------> Clases globales
Global G_Proc As New G_Proc
'---------------> 0 < ----------

'JC
Global vg_auxtiprec             As Long
Global vg_5etapas               As Boolean
Global vg_tipbase               As String
Global vg_ciedia                As String '-------> cierre de día
Global vg_codbod                As Long
Global VarSitioRemoto           As Boolean
'FIN JC

Dim Retorno As String

Declare Function GetPrivateProfileString Lib "kernel32" _
        Alias "GetPrivateProfileStringA" _
              (ByVal lpApplicationName As String, _
               ByVal lpKeyName As Any, _
               ByVal lpDefault As String, _
               ByVal lpReturnedString As String, _
               ByVal nSize As Long, _
               ByVal lpFileName As String) As Long
               
Public vg_CallForm As String ' variable para saber de que formulario se gatillo la ultima llamada
Public vg_CallFormDato As String ' variable para pasar un dato del ultimo formulario gatillado

Declare Function GetVolumeInformation Lib "kernel32.dll" Alias _
"GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal _
lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, _
lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal _
nFileSystemNameSize As Long) As Long



Sub Main()

On Error GoTo Man_Error

Dim cpar As String
Dim xpar As String

VerConfReg
'-------> Mover intervalo grabado
vg_IntMin = 3
vg_TemSeg = 0

vg_ModoOpen = dbSQLPassThrough
dir_trabajo = MiFunc("Path", "Gestion.Ini", "Ruta")
'BaseDeDato = MiFunc("Base_de_datos", "Gestion.Ini", "Mdb")
Provider = MiFunc("Provider", "Gestion.Ini", "Jet")
RutaGif = MiFunc("Gif", "Gestion.Ini", "Ruta")

vg_SqlNSvr = MiFunc("SQL SERVER", "Gestion.ini", "Servidor")
'vg_SqlBase = MiFunc("SQL SERVER", "Gestion.ini", "DataBase")
'vg_SqlNUsr = MiFunc("SQL SERVER", "Gestion.ini", "Usuario")
'vg_SqlPass = MiFunc("SQL SERVER", "Gestion.ini", "Password")
'vg_SqlNSvr = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Servidor"), ""))
vg_SqlBase = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "DataBase"), ""))
vg_SqlNUsr = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Usuario"), ""))
vg_SqlPass = fg_Desencripta(TipoDato(MiFunc("SQL SERVER", "Gestion.ini", "Password"), ""))

'vg_SqlNSvrW = MiFunc("SQL SERVER WEB", "Gestion.ini", "Servidor")
'vg_SqlBaseW = MiFunc("SQL SERVER WEB", "Gestion.ini", "DataBase")
'vg_SqlNUsrW = MiFunc("SQL SERVER WEB", "Gestion.ini", "Usuario")
'vg_SqlPassW = MiFunc("SQL SERVER WEB", "Gestion.ini", "Password")
vg_SqlNSvrW = fg_Desencripta(TipoDato(MiFunc("SQL SERVER WEB", "Gestion.ini", "Servidor"), ""))
vg_SqlBaseW = fg_Desencripta(TipoDato(MiFunc("SQL SERVER WEB", "Gestion.ini", "DataBase"), ""))
vg_SqlNUsrW = fg_Desencripta(TipoDato(MiFunc("SQL SERVER WEB", "Gestion.ini", "Usuario"), ""))
vg_SqlPassW = fg_Desencripta(TipoDato(MiFunc("SQL SERVER WEB", "Gestion.ini", "Password"), ""))


'tecffod vgtec_NSvr = MiFunc("ORACLE", "Gestion.ini", "Servidor")
'tecfood vgtec_NUsr = MiFunc("ORACLE", "Gestion.ini", "Usuario")
'tecfood vgtec_Pass = MiFunc("ORACLE", "Gestion.ini", "Password")

'vgsac_NSvr = MiFunc("ORACLE", "Gestion.ini", "Servidor")
'vgsac_NUsr = MiFunc("ORACLE", "Gestion.ini", "Usuario")
'vgsac_Pass = MiFunc("ORACLE", "Gestion.ini", "Password")
vgsac_NSvr = fg_Desencripta(TipoDato(MiFunc("ORACLE", "Gestion.ini", "Servidor"), ""))
vgsac_NUsr = fg_Desencripta(TipoDato(MiFunc("ORACLE", "Gestion.ini", "Usuario"), ""))
vgsac_Pass = fg_Desencripta(TipoDato(MiFunc("ORACLE", "Gestion.ini", "Password"), ""))

'-------20110603> Ttraer versión Access
vg_VAccess = MiFunc("Version Access", "Gestion.ini", "VAccess")

vg_CSep = ","
vg_CDec = "."

If App.PrevInstance Then

    MsgBox App.EXEName + " ya se encuentra en ejecución. Ud. no puede mantener dos copias del mismo programa en memoria."
    End

End If

vg_estopen = False
VgLinea = Chr(13) & Chr(10)  'inserta retorno de carro a los mensages
V_Acceso.Show 1

vg_DirLog = dir_trabajo & "LogoSdx.jpg"

If Dir(dir_trabajo & "txt*.txt") <> "" Then Kill dir_trabajo & "txt*.txt"
If Dir(dir_trabajo & "*.ss6") <> "" Then Kill dir_trabajo & "*.ss6"
    
'------- Crear directorio PC local, para generar archivos de envio información a los contratos.
'If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"
'If Dir(Mid(dir_trabajo, 1, 3) & "Temp", vbDirectory) = "" Then MkDir Mid(dir_trabajo, 1, 3) & "Temp"
If Dir(dir_trabajo & "Temp", vbDirectory) = "" Then MkDir (dir_trabajo) & "Temp"
'mdirpc = Dir("C:\Temp\" & "Actualizarsgp", vbDirectory)
'mdirpc = Dir(Mid(dir_trabajo, 1, 3) & "Temp\" & "Actualizarsgp", vbDirectory)
mdirpc = Dir(dir_trabajo & "Temp\" & "Actualizarsgp", vbDirectory)
'If mdirpc = "" Then MkDir Mid(dir_trabajo, 1, 3) & "Temp\Actualizarsgp"
If mdirpc = "" Then MkDir (dir_trabajo) & "Temp\Actualizarsgp"
'mdirpc = Mid(dir_trabajo, 1, 3) & "Temp\" & "Actualizarsgp" & "\"
mdirpc = dir_trabajo & "Temp\" & "Actualizarsgp" & "\"
'------- Fin crear directorio PC local, para generar archivos de envio información a los contratos.

'-------> Crear directorio Errores
If Dir(dir_trabajo & "\" & "Errores", vbDirectory) = "" Then MkDir dir_trabajo & "\" & "Errores"
'-------> Fin crear directorio Errores

'-------> Crear directorio ExcelSGP
If Dir(dir_trabajo & "\" & "ExcelMinutaSGP", vbDirectory) = "" Then MkDir dir_trabajo & "\" & "ExcelMinutaSGP"
'-------> Fin crear directorio Excel Versión

'AbrirBase
'AbrirBasetec
Partida.Show

Exit Sub
Man_Error:
If Err = 70 Or Err = 75 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, "Administrador SGP"

End Sub
