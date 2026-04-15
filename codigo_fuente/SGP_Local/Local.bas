Attribute VB_Name = "Local"

'Declarations
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Const LOCALE_ICENTURY = &H24
Public Const LOCALE_ICOUNTRY = &H5
Public Const LOCALE_ICURRDIGITS = &H19
Public Const LOCALE_ICURRENCY = &H1B
Public Const LOCALE_IDATE = &H21
Public Const LOCALE_IDAYLZERO = &H26
Public Const LOCALE_IDEFAULTCODEPAGE = &HB
Public Const LOCALE_IDEFAULTCOUNTRY = &HA
Public Const LOCALE_IDEFAULTLANGUAGE = &H9
Public Const LOCALE_IDIGITS = &H11
Public Const LOCALE_IINTLCURRDIGITS = &H1A
Public Const LOCALE_ILANGUAGE = &H1
Public Const LOCALE_ILDATE = &H22
Public Const LOCALE_ILZERO = &H12
Public Const LOCALE_IMEASURE = &HD
Public Const LOCALE_IMONLZERO = &H27
Public Const LOCALE_INEGCURR = &H1C
Public Const LOCALE_INEGSEPBYSPACE = &H57
Public Const LOCALE_INEGSIGNPOSN = &H53
Public Const LOCALE_INEGSYMPRECEDES = &H56
Public Const LOCALE_IPOSSEPBYSPACE = &H55
Public Const LOCALE_IPOSSIGNPOSN = &H52
Public Const LOCALE_IPOSSYMPRECEDES = &H54
Public Const LOCALE_ITIME = &H23
Public Const LOCALE_ITLZERO = &H25
Public Const LOCALE_NOUSEROVERRIDE = &H80000000
Public Const LOCALE_S1159 = &H28
Public Const LOCALE_S2359 = &H29
Public Const LOCALE_SABBREVCTRYNAME = &H7
Public Const LOCALE_SABBREVDAYNAME1 = &H31
Public Const LOCALE_SABBREVDAYNAME2 = &H32
Public Const LOCALE_SABBREVDAYNAME3 = &H33
Public Const LOCALE_SABBREVDAYNAME4 = &H34
Public Const LOCALE_SABBREVDAYNAME5 = &H35
Public Const LOCALE_SABBREVDAYNAME6 = &H36
Public Const LOCALE_SABBREVDAYNAME7 = &H37
Public Const LOCALE_SABBREVLANGNAME = &H3
Public Const LOCALE_SABBREVMONTHNAME1 = &H44
Public Const LOCALE_SCOUNTRY = &H6
Public Const LOCALE_SCURRENCY = &H14
Public Const LOCALE_SDATE = &H1D
Public Const LOCALE_SDAYNAME1 = &H2A
Public Const LOCALE_SDAYNAME2 = &H2B
Public Const LOCALE_SDAYNAME3 = &H2C
Public Const LOCALE_SDAYNAME4 = &H2D
Public Const LOCALE_SDAYNAME5 = &H2E
Public Const LOCALE_SDAYNAME6 = &H2F
Public Const LOCALE_SDAYNAME7 = &H30
Public Const LOCALE_SDECIMAL = &HE
Public Const LOCALE_SENGCOUNTRY = &H1002
Public Const LOCALE_SENGLANGUAGE = &H1001
Public Const LOCALE_SGROUPING = &H10
Public Const LOCALE_SINTLSYMBOL = &H15
Public Const LOCALE_SLANGUAGE = &H2
Public Const LOCALE_SLIST = &HC
Public Const LOCALE_SLONGDATE = &H20
Public Const LOCALE_SMONDECIMALSEP = &H16
Public Const LOCALE_SMONGROUPING = &H18
Public Const LOCALE_SMONTHNAME1 = &H38
Public Const LOCALE_SMONTHNAME10 = &H41
Public Const LOCALE_SMONTHNAME11 = &H42
Public Const LOCALE_SMONTHNAME12 = &H43
Public Const LOCALE_SMONTHNAME2 = &H39
Public Const LOCALE_SMONTHNAME3 = &H3A
Public Const LOCALE_SMONTHNAME4 = &H3B
Public Const LOCALE_SMONTHNAME5 = &H3C
Public Const LOCALE_SMONTHNAME6 = &H3D
Public Const LOCALE_SMONTHNAME7 = &H3E
Public Const LOCALE_SMONTHNAME8 = &H3F
Public Const LOCALE_SMONTHNAME9 = &H40
Public Const LOCALE_SMONTHOUSANDSEP = &H17
Public Const LOCALE_SNATIVECTRYNAME = &H8
Public Const LOCALE_SNATIVEDIGITS = &H13
Public Const LOCALE_SNATIVELANGNAME = &H4
Public Const LOCALE_SNEGATIVESIGN = &H51
Public Const LOCALE_SPOSITIVESIGN = &H50
Public Const LOCALE_SSHORTDATE = &H1F
Public Const LOCALE_STHOUSAND = &HF
Public Const LOCALE_STIME = &H1E
Public Const LOCALE_STIMEFORMAT = &H1003
Public Const NETWORK_ALIVE_AOL = &H4
Public Const NETWORK_ALIVE_LAN = &H1
Public Const NETWORK_ALIVE_WAN = &H2
'Local system uses a LAN to connect to the Internet.
Public Const INTERNET_CONNECTION_LAN = &H2
'Local system uses a modem to connect to the Internet.
Public Const INTERNET_CONNECTION_MODEM = &H1
'Local system uses a proxy server to connect to the Internet.
Public Const INTERNET_CONNECTION_PROXY = &H4
'Local system has RAS installed.
Public Const INTERNET_RAS_INSTALLED = &H10

Public Type Archivo
    FileName  As String
    FileTitle As String
    Success As Boolean
End Type

Type usuarios
  
  codprog As String * 5
  descripcion As String * 40
  codprog_anterior As String * 5
  indice As String * 3
  activo As String * 1

End Type

Dim vecusuarios() As usuarios

'Code:
Function Get_locale(Variable As String) As String ' Retrieve the regional setting

Dim iRet1 As Long
Dim iRet2 As Long
Dim lpLCDataVar As String
Dim Pos As Integer
Dim Locale As Long
      
Locale = GetUserDefaultLCID()
iRet1 = GetLocaleInfo(Locale, Variable, lpLCDataVar, 0)
Get_locale = String$(iRet1, 0)
iRet2 = GetLocaleInfo(Locale, Variable, Get_locale, iRet1)
Pos = InStr(Get_locale, Chr$(0))

If Pos > 0 Then
    
    Get_locale = Left$(Get_locale, Pos - 1)
'    MsgBox "Regional Setting = " + Symbol

End If

End Function

Sub Set_locale() 'Change the regional setting

      Dim Symbol As String
      Dim iRet As Long
      Dim Locale As Long
      
'LOCALE_SDATE is the constant for the date separator
'as stated in declarations
'for any other locale just change the contant in the Function

      Locale = GetUserDefaultLCID() 'Get user Locale ID
      Symbol = "-" 'New character for the locale
      iRet = SetLocaleInfo(Locale, LOCALE_SDATE, Symbol)
      
End Sub

Sub VerConfReg()

Dim sDeMC As String, sDeMS As String
Dim sThMC As String, sThMS As String
Dim sDeNC As String, sDeNS As String
Dim sThNC As String, sThNS As String
Dim sDatC As String, sDatS As String
Dim sSepF As String, sStr As String
sDeNC = Get_locale(LOCALE_SDECIMAL)
sThNC = Get_locale(LOCALE_STHOUSAND)
sDeMC = Get_locale(LOCALE_SMONDECIMALSEP)
sThMC = Get_locale(LOCALE_SMONTHOUSANDSEP)
sDatC = Get_locale(LOCALE_SSHORTDATE)
sStr = Get_locale(LOCALE_SSHORTDATE)
sSepF = Get_locale(LOCALE_SDATE)
Do While InStr(sDatC, "y") <> 0
    Mid(sDatC, InStr(sDatC, "y"), 1) = "a"
Loop
sDeNS = "."
sThNS = ","
sDeMS = "."
sThMS = ","
sDatS = "dd" & sSepF & "MM" & sSepF & "aaaa"

If sDeNC <> sDeNS Or sThNC <> sThNS Or sDeMC <> sDeMS Or sThMC <> sThMS Or sDatC <> sDatS Then
    iRet = SetLocaleInfo(Locale, LOCALE_SDECIMAL, ".")
    iRet = SetLocaleInfo(Locale, LOCALE_STHOUSAND, ",")
    iRet = SetLocaleInfo(Locale, LOCALE_SMONDECIMALSEP, ".")
    iRet = SetLocaleInfo(Locale, LOCALE_SMONTHOUSANDSEP, ",")
    If InStr(sStr, "y") <> 0 Then
        
        iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, "dd/MM/yyyy")
    
    Else
        
        iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, "dd/MM/aaaa")
    
    End If
    iRet = SetLocaleInfo(Locale, LOCALE_SDATE, "/")
'    P_ConReg.Show 1

End If

End Sub

