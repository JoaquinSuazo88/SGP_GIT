Attribute VB_Name = "Rutinas"
Option Explicit
Option Compare Text

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Private Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wmsg As Long, ByVal wparam As Long, lparam As Any) As Long
Const WM_PASTE = &H302

Private Declare Function GetLogicalDrives Lib "kernel32" () As Long


'-----------------------------jc
'Function TraerFechaCierre()
'Dim RS As New ADODB.Recordset
'vg_ciedia = ""
'RS.Open "SELECT DISTINCT par_nombre, par_valor FROM cas_a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & vg_codcasino & "'", vg_db, adOpenStatic
'If Not RS.EOF Then
'   vg_ciedia = fg_Desencripta(TipoDato(RS!par_valor, ""))
'   Partida.StatusBar1.Panels(8).text = Trim(RS!par_nombre) & " : " & CDate(vg_ciedia) - 1
'End If
'RS.Close: Set RS = Nothing
'If Trim(vg_ciedia) = "" Then MsgBox "No esta activo la fecha cierre día, Comunicase con departamento de informatica" & VgLinea & Space(40) & "Proceso cancelado ...", vbCritical + vbOKOnly, "Menú Principal": End
'End Function
'-----------------------------------------


Function ValidarUsuarioAcceso(CodOpc As Long, Usu As String) As String

Dim RS1 As New ADODB.Recordset
Dim acceso As String, incluir As String, alterar As String, eliminar As String, imprimir
'-----------------------------VALIDAR USUARIO-----------------
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_ValidarUsuarioAcceso '" & Usu & "', " & CodOpc & "")
ValidarUsuarioAcceso = "00000"
If Not RS1.EOF Then
    acceso = Trim(RS1!dpe_deracc)
    incluir = Trim(RS1!dpe_deragr)
    alterar = Trim(RS1!dpe_dermod)
    eliminar = Trim(RS1!dpe_dereli)
    imprimir = Trim(RS1!dpe_derimp)
    ValidarUsuarioAcceso = acceso & incluir & alterar & eliminar & imprimir
End If
RS1.Close
Set RS1 = Nothing
'--------------------------------------------------------------

End Function


Public Function fg_BuscaenArbol(codigo As Long, Tabla As String, CampoBus As String) As String

Dim RS1     As New ADODB.Recordset
Dim Nombre  As String
Dim i       As Long
    '-------> Buscar raiz en un TreeView
    Nombre = ""
    For i = 1 To 4
        
        If codigo = 0 Then Exit For
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS1 = vg_db.Execute("SELECT * FROM " & Tabla & " WHERE " & CampoBus & " = " & codigo & "") ', vg_db, adOpenForwardOnly ', adOpenStatic
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit For
        Nombre = Trim(RS1(1)) & "\" & Nombre
        codigo = RS1(2)
        If RS1(0) = 0 Then RS1.Close: Set RS1 = Nothing: Exit For
        RS1.Close
        Set RS1 = Nothing
    
    Next
    If Trim(Nombre) <> "" Then fg_BuscaenArbol = Mid(Nombre, 1, Len(Nombre) - 1) Else fg_BuscaenArbol = ""

End Function

Function fg_BuscaCodArbol(codigo As Long, Tabla As String, CampoBus As String) As String

Dim codarb As String
Dim i As Long
'-------> Buscar raiz en un TreeView
codarb = ""
For i = 1 To 5
    If codigo = 0 Then Exit For
    RS1.Open "select * from " & Tabla & " where " & CampoBus & " = " & codigo & "", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit For
    codarb = RS1(0) & "&" & codarb: codigo = RS1(2)
    If RS1(0) = 0 Then RS1.Close: Set RS1 = Nothing: Exit For
    RS1.Close: Set RS1 = Nothing
Next
If i = 2 Then codarb = codarb & "&"
If Trim(codarb) <> "" Then fg_BuscaCodArbol = codarb & ";" Else fg_BuscaCodArbol = ""

End Function

Sub fg_centra(frm As Form)
frm.Top = ((Screen.Height - IIf(frm.MDIChild, 1100, 0)) \ 2 - frm.Height \ 2)
frm.Left = Screen.Width \ 2 - frm.Width \ 2
'    frm.Top = Screen.Height \ 2 - frm.Height \ 2
'    frm.Left = Screen.Width \ 2 - frm.Width \ 2

'Formater fecha
Dim c As Control
For Each c In frm.Controls
        
        If TypeOf c Is fpDateTime Then
           
           c.CalFirstDay (1)
           c.ShortDayName(1) = "Dom"
           c.ShortDayName(2) = "Lun"
           c.ShortDayName(3) = "Mar"
           c.ShortDayName(4) = "Mie"
           c.ShortDayName(5) = "Jue"
           c.ShortDayName(6) = "Vie"
           c.ShortDayName(7) = "Sab"
           c.LongMonthName(1) = "Enero"
           c.LongMonthName(2) = "Febrero"
           c.LongMonthName(3) = "Marzo"
           c.LongMonthName(4) = "Abril"
           c.LongMonthName(5) = "Mayo"
           c.LongMonthName(6) = "Junio"
           c.LongMonthName(7) = "Julio"
           c.LongMonthName(8) = "Agosto"
           c.LongMonthName(9) = "Septiembre"
           c.LongMonthName(10) = "Octubre"
           c.LongMonthName(11) = "Noviembre"
           c.LongMonthName(12) = "Diciembre"
        
        End If
    
    Next

End Sub

Sub validar_respuesta(Title As String)
Dim msg As String, Style, Help, Ctxt
msg = "               Esta Seguro ?"
Style = vbYesNoCancel + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, Title, Help, Ctxt)
End Sub

Function fg_carga(texto As String)
    
    Screen.MousePointer = 11
    DoEvents

End Function

Function fg_descarga()

    Screen.MousePointer = 0
    
End Function

Sub MostrarPorcentaje()

'Paso.Gauge.value = Val((ContRegistro / TotalRegistro) * 100)
'Paso.Gauge.Refresh

'PantPorc.Porcentaje.value = Val((ContRegistro / TotalRegistro) * 100)
'PantPorc.Porcentaje.Refresh

End Sub

Function fg_RutDig(ByVal crut As String) As String

Dim d2 As String, Rut2 As String

    If vg_Dig = "N" Then fg_RutDig = crut: Exit Function
    Rut2 = Trim(crut): crut = Padl(crut, 9, "0")
    d2 = Mid("0K987654321", (Val(Mid(crut, 1, 1)) * 4 + Val(Mid(crut, 2, 1)) * 3 + Val(Mid(crut, 3, 1)) * 2 + Val(Mid(crut, 4, 1)) * 7 + Val(Mid(crut, 5, 1)) * 6 + Val(Mid(crut, 6, 1)) * 5 + Val(Mid(crut, 7, 1)) * 4 + Val(Mid(crut, 8, 1)) * 3 + Val(Mid(crut, 9, 1)) * 2) Mod 11 + 1, 1)
    fg_RutDig = Rut2 + d2

End Function

Function fg_Check_Rut(ByVal crut As String) As Integer

Dim d1$, d2$, L%
    
    If vg_Dig = "N" Then fg_Check_Rut = True: Exit Function
    fg_Check_Rut = False
    crut = Trim(crut)
    If crut = "" Then Exit Function
    d1 = UCase(Mid(crut, Len(crut), 1))
    L = InStr(crut, "-")
    L = IIf(L = 0, Len(crut) - 1, L - 1)
    crut = Mid(crut, 1, L)
    crut = Padl(crut, 9, "0")
    d2 = Mid("0K987654321", (Val(Mid(crut, 1, 1)) * 4 + Val(Mid(crut, 2, 1)) * 3 + Val(Mid(crut, 3, 1)) * 2 + Val(Mid(crut, 4, 1)) * 7 + Val(Mid(crut, 5, 1)) * 6 + Val(Mid(crut, 6, 1)) * 5 + Val(Mid(crut, 7, 1)) * 4 + Val(Mid(crut, 8, 1)) * 3 + Val(Mid(crut, 9, 1)) * 2) Mod 11 + 1, 1)
    fg_Check_Rut = IIf(d1 = d2, True, False)

End Function

Function Padl(ByVal lcCadena As String, lnCuantos As Integer, lcConque As String) As String
  
  Dim lcCadena1$
  lcCadena1 = lcCadena

  lcCadena1 = LTrim(RTrim(lcCadena1))
  If Len(lcCadena1) >= lnCuantos Then
     Padl = lcCadena1
     Exit Function
  End If
  If lcConque <> "" Then
    lcCadena1 = String$(lnCuantos - Len(lcCadena1), lcConque) + lcCadena1
  Else
    lcCadena1 = String$(lnCuantos - Len(lcCadena1), " ") + lcCadena1
  End If
  Padl = lcCadena1

End Function

Function fg_DespintaRut(ByVal rut As String) As String

fg_DespintaRut = Trim(fg_Quitachar(fg_Quitachar(rut, "."), "-"))

End Function

Function fg_PintaRut(ByVal lcRut As String) As String
Dim lcNewRut$
Dim lcRutFin$, d1$
Dim j%, i%, L%

If lcRut = "" Then
   fg_PintaRut = ""
   Exit Function
End If
If vg_Dig = "N" Then fg_PintaRut = lcRut: Exit Function
lcRut = fg_Quitachar(lcRut, "-")
lcRut = fg_Quitachar(lcRut, ".")
lcRut = fg_SacaCeros(lcRut)
lcNewRut = Trim(lcRut)
If lcNewRut = "" Then
   fg_PintaRut = ""
   Exit Function
End If

d1 = UCase(Mid(lcNewRut, Len(lcNewRut), 1))
L = InStr(lcNewRut, "-")

If (L = 0) Then
    L = Len(lcNewRut) - 1
Else
    L = L - 1
End If

lcNewRut = Mid(lcNewRut, 1, L)
lcRutFin = "-" + d1
j = 1

For i = Len(lcNewRut) To 1 Step -1
    If j = 3 Then
        j = 1
        If i > 1 Then
          lcRutFin = "." & Mid(lcNewRut, i, 1) & lcRutFin
        Else
          lcRutFin = Mid(lcNewRut, i, 1) & lcRutFin
        End If
    Else
        j = j + 1
        lcRutFin = Mid(lcNewRut, i, 1) & lcRutFin
    End If
Next i
fg_PintaRut = lcRutFin
End Function

Function fg_Quitachar(ByVal numero As String, ByVal Caracter As String) As String
'OBJETIVO : Elimina un Caracter espacifico de una cadena.
'ENTRADAS : La Cadena.
'           El Caracter a eliminar.
'SALIDAS  : La Cadena sin el caracter.
'USO      : QuitaChar(RUT$,"-")
Do While InStr(numero, Caracter) > 0
    
    numero = Mid(numero, 1, InStr(numero, Caracter) - 1) + Mid(numero, InStr(numero, Caracter) + 1, Len(numero))

Loop
fg_Quitachar = numero

End Function

Function fg_CambiaChar(ByVal numero As String, ByVal Caracter As String, ByVal Nuevo As String) As String

Do While InStr(numero, Caracter) > 0
    
    numero = Mid(numero, 1, InStr(numero, Caracter) - 1) & Nuevo & Mid(numero, InStr(numero, Caracter) + 1, Len(numero))

Loop
fg_CambiaChar = numero

End Function

Function fg_SacaCeros(ByVal cadena As String) As String

Dim i%
fg_SacaCeros = ""
If cadena <> "" Then
   
   i = 1
   Do While Mid(cadena, i, 1) = "0"
      
      i = i + 1
   
   Loop
   fg_SacaCeros = Mid(cadena, i)

End If

End Function

Function Fg_SacaParentesis(ByVal Parentesis As String) As String

Dim i%
Dim ValLcntH$
ValLcntH = ""
For i = 1 To Len(Parentesis)
    
    If Asc(Mid(Parentesis, i, 1)) = 40 Or Asc(Mid(Parentesis, i, 1)) = 41 Then
       
       Exit For
    
    Else
       
       ValLcntH = ValLcntH + Mid(Parentesis, i, 1)
    
    End If

Next i
Fg_SacaParentesis = ValLcntH

End Function

Function Fg_Sacacremilla(ByVal cremilla As String) As String
Dim i%
Dim ValLcntH$
ValLcntH = ""

For i = 1 To Len(cremilla)
    If Asc(Mid(cremilla, i, 1)) = 39 Then
'       Exit For
    Else
       ValLcntH = ValLcntH + Mid(cremilla, i, 1)
    End If
Next i
Fg_Sacacremilla = ValLcntH
End Function

Sub ins_log_error(Err As String)
End Sub

Function LimpiaDato(cString As String)

Do While InStr(cString, "'") <> 0
    
    Clipboard.SetText cString
    Mid(cString, InStr(cString, "'"), 1) = " "

Loop

LimpiaDato = cString

End Function

Function LimpiaDatoExcel(cString As String) As String

LimpiaDatoExcel = Replace(Trim(cString), ":", "")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "\", " ")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "/", " ")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "?", " ")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "*", " ")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "[", " ")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "]", " ")
LimpiaDatoExcel = Replace(Trim(LimpiaDatoExcel), "'", " ")
'LimpiaDato = cString
End Function

Function fg_CreaRS(Base As Connection, Sql As String, tipo As Integer, modo As Integer) As Recordset
If modo = dbSQLPassThrough Then
    Clipboard.SetText Sql
    Set fg_CreaRS = Base.OpenRecordset(Sql, tipo) ', Modo)
Else
    Set fg_CreaRS = Base.OpenRecordset(Sql, tipo)
End If
End Function

Function AbrirBase()
Dim cDrv As String
On Error GoTo ManError
Set vg_db = New ADODB.Connection
'vg_db.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDeDato + "' ;Persist Security Info=False"
'-------> Abrir sgpadm
vg_db.ConnectionString = "Driver={SQL Server};SERVER=" + vg_SqlNSvr + ";OLE DB Services = -2;UID=" + vg_SqlNUsr + ";PWD=" + vg_SqlPass + ";DATABASE=" + vg_SqlBase + ""
vg_db.ConnectionTimeout = 3600
vg_db.CommandTimeout = 3600
vg_db.Open

Exit Function
ManError:
'    If Err.Number = -2147467259 Then cDrv = "{Microsoft ODBC para SqlServer}": Resume
    MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Mantención sistema SGP"

End Function

Function AbrirBaseWebPed()
Dim cDrv As String
On Error GoTo ManError
If Trim(vg_SqlNSvrW) = "" Then vg_estopen = False: Exit Function
'-------> Abrir Pedido Web
Set vg_dbpedweb = New ADODB.Connection
vg_dbpedweb.ConnectionString = "Driver={SQL Server};SERVER=" + vg_SqlNSvrW + ";UID=" + vg_SqlNUsrW + ";PWD=" + vg_SqlPassW + ";DATABASE=" + vg_SqlBaseW + ""
vg_dbpedweb.ConnectionTimeout = 3600
vg_dbpedweb.CommandTimeout = 3600
vg_dbpedweb.Open
vg_estopen = True
vg_tipbase = 2
Exit Function
ManError:
If Err.Number = -2147467259 Then
   vg_estopen = False
   MsgBox Err & ":  " & "No se puede abrir base de datos. " & Chr(13) & Chr(13) & "Inténtelo en unos minutos mas tarde.", vbCritical + vbOKOnly, "Sistema Administrador SGP"
   cDrv = "{Microsoft ODBC para SqlServer}": Resume Next
   Exit Function
ElseIf Err.Number = -2147217843 Then
   vg_estopen = False
   cDrv = "{Microsoft ODBC para SqlServer}": Resume Next
   Exit Function
End If
End Function

Function AbrirBasetec()
Dim cDrv As String
Set vg_dbtec = New ADODB.Connection
On Error GoTo ManError
vg_dbtec.ConnectionTimeout = 3600
vg_dbtec.CommandTimeout = 3600
cDrv = "{Microsoft ODBC for Oracle}"
vg_dbtec.Open "DRIVER=" & cDrv & ";SERVER=" + vgtec_NSvr + ";uid=" + vgtec_NUsr + ";pwd=" + vgtec_Pass + ""
vg_dbpedweb.Open

Exit Function
ManError:
If Err.Number = -2147467259 Then cDrv = "{Microsoft ODBC para Oracle}": Resume
End Function

Function AbrirBaseSac()
Dim cDrv As String
Set vg_dbsac = New ADODB.Connection
On Error GoTo ManError
vg_dbsac.ConnectionTimeout = 3600
vg_dbsac.CommandTimeout = 3600
cDrv = "{Microsoft ODBC for Oracle}"
vg_dbsac.Open "DRIVER=" & cDrv & ";SERVER=" + vgsac_NSvr + ";uid=" + vgsac_NUsr + ";pwd=" + vgsac_Pass + ""

ManError:
If Err.Number = -2147467259 Then cDrv = "{Microsoft ODBC para Oracle}": Resume
End Function

Function fg_Fecha_Escrita(Fecha As String) As String
'Escribe la fecha en el siguiente formato:
'Jueves 16 de Junio de 1994
    Dim MiFecha$, dia%, mes%, Año%, mes1$, FechaP$, DiaSem%
    
    If Fecha = "" Then Exit Function

    MiFecha = Format$(Fecha, "mmm dd yyyy")
    dia = Day(DateValue(MiFecha))
    DiaSem = Weekday(DateValue(MiFecha))
    mes = Month(DateValue(MiFecha))
    Año = Year(DateValue(MiFecha))
    FechaP = ""
    Select Case DiaSem
        Case 1
            FechaP = "Domingo,"
        Case 2
            FechaP = "Lunes,"
        Case 3
            FechaP = "Martes,"
        Case 4
            FechaP = "Miércoles,"
        Case 5
            FechaP = "Jueves,"
        Case 6
            FechaP = "Viernes,"
        Case 7
            FechaP = "Sábado,"
    End Select

    FechaP = FechaP + " " + Format$(dia) + " de "
    Select Case mes
        Case 1
            FechaP = FechaP + "Enero"
        Case 2
            FechaP = FechaP + "Febrero"
        Case 3
            FechaP = FechaP + "Marzo"
        Case 4
            FechaP = FechaP + "Abril"
        Case 5
            FechaP = FechaP + "Mayo"
        Case 6
            FechaP = FechaP + "Junio"
        Case 7
            FechaP = FechaP + "Julio"
        Case 8
            FechaP = FechaP + "Agosto"
        Case 9
            FechaP = FechaP + "Septiembre"
        Case 10
            FechaP = FechaP + "Octubre"
        Case 11
            FechaP = FechaP + "Noviembre"
        Case 12
            FechaP = FechaP + "Diciembre"
    End Select
    fg_Fecha_Escrita = FechaP + " de " + Format$(Año)
End Function

Function fg_Pict(vEnt As Integer, vDec As Integer) As String

Select Case vEnt

Case 6
    
    fg_Pict = "###" & vg_CSep & "##0"

Case 9
    
    fg_Pict = "###" & vg_CSep & "###" & vg_CSep & "##0"

Case Else
    
    fg_Pict = "##" & vg_CSep & "###" & vg_CSep & "##0"

End Select

If vDec > 0 Then fg_Pict = fg_Pict & vg_CDec & String(vDec, "0")

End Function

Function fg_Fecha_Dia(Fecha As String, opcion As Integer) As String
'-------> Escribe la fecha en el siguiente formato:
'-------> Jueves 16 de Junio de 1994
    Dim MiFecha$, dia%, FechaP$, DiaSem%
    If Fecha = "" Then Exit Function
    MiFecha = Format$(fg_Ctod1(Fecha), "mmm dd yyyy")
    dia = Day(DateValue(MiFecha))
    DiaSem = Weekday(DateValue(MiFecha))
    FechaP = ""
    Select Case opcion
      Case 1
        Select Case DiaSem
          Case 1
            FechaP = "Dom."
          Case 2
            FechaP = "Lun."
          Case 3
            FechaP = "Mar."
          Case 4
            FechaP = "Mié."
          Case 5
            FechaP = "Jue."
          Case 6
            FechaP = "Vie."
          Case 7
            FechaP = "Sáb."
        End Select
      Case 2
        Select Case DiaSem
          Case 1
            FechaP = "Domingo"
          Case 2
            FechaP = "Lunes"
          Case 3
            FechaP = "Martes"
          Case 4
            FechaP = "Miércoles"
          Case 5
            FechaP = "Jueves"
          Case 6
            FechaP = "Viernes"
          Case 7
            FechaP = "Sábado"
        End Select
    End Select
    FechaP = FechaP + " " + Format$(dia)
    fg_Fecha_Dia = FechaP
End Function

Function fg_Ctod(Fecha As Variant) As String
If IsNull(Fecha) Then Exit Function
If Trim(Fecha) = "" Then Exit Function
'    fg_Ctod = CDate(Mid(fecha, 1, 4) + "-" + Mid(fecha, 5, 2) + "-" + Mid(fecha, 7, 2))
fg_Ctod = Mid(Fecha, 7, 2) + "-" + Mid(Fecha, 5, 2) + "-" + Mid(Fecha, 1, 4)
End Function

Function fg_Ctod1(Fecha As Variant) As String

If IsNull(Fecha) Then Exit Function
If Trim(Fecha) = "" Then Exit Function
fg_Ctod1 = Trim(Mid(Fecha, 7, 2)) + "/" + Trim(Mid(Fecha, 5, 2)) + "/" + Trim(Mid(Fecha, 1, 4))

End Function

Function Meses(Fecha As String) As String
Dim mes As Variant
Meses = ""
If Trim(Fecha) = "" Then Exit Function
mes = Array("", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre")
Meses = mes(Month(CDate(Fecha)))
End Function

Function fg_ArchivoTxt()

Dim i As Long
i = 1

For i = 1 To 99999
    
    If Dir(dir_trabajo & "txt" & fg_pone_cero(Trim(Str(i)), 5) & ".txt") = "" Then
        
        fg_ArchivoTxt = dir_trabajo & "txt" & fg_pone_cero(Trim(Str(i)), 5) & ".txt": Exit Function
    
    End If

Next i

End Function

Function fg_ArchivoXls(NombreArchivo As String) As String

Dim i As Long

i = 1

'-------> Crear directorio Errores
If Dir(dir_trabajo & "\" & "Errores", vbDirectory) = "" Then MkDir dir_trabajo & "\" & "Errores"
'-------> Fin crear directorio Errores

For i = 1 To 99999
    
    If Dir(dir_trabajo & "\" & "Errores\" & NombreArchivo & fg_pone_cero(Trim(Str(i)), 5) & ".Xls") = "" Then
        
        fg_ArchivoXls = dir_trabajo & "Errores\" & NombreArchivo & fg_pone_cero(Trim(Str(i)), 5) & ".Xls": Exit Function
    
    End If

Next i

End Function

Function fg_Archivospread()
Dim i As Long
i = 1
For i = 1 To 9999999
    If Dir(dir_trabajo & "ss6" & i & ".ss6") = "" Then
        fg_ArchivoTxt = dir_trabajo & "ss6" & i & ".ss6": Exit Function
    End If
Next i
End Function

Function fg_pone_cero(ByVal cadena As String, ByVal cuanto As Integer) As String
'-------> pone ceros a la izquierda
fg_pone_cero = ""
If cadena <> "" Then
   
   Do While Len(Trim(cadena)) < cuanto
      
      cadena = "0" + Trim(cadena)
   
   Loop
   fg_pone_cero = Trim(cadena)

End If

End Function

Function fg_pone_espacio(ByVal cadena As String, ByVal cuanto As Integer) As String
'-------> pone ceros a la izquierda
fg_pone_espacio = ""

If cadena <> "" Then
   
   Do While Len(cadena) < cuanto
      cadena = " " + cadena
   Loop
   fg_pone_espacio = cadena

End If
End Function

Function Fg_Sacasaltolinea(ByVal Parentesis As String) As String
Dim X%
Dim ValLcntH$
ValLcntH = ""

For X = 1 To Len(Parentesis)
    If Asc(Mid(Parentesis, X, 1)) <> 13 Then
       ValLcntH = ValLcntH + Mid(Parentesis, X, 1)
    End If
Next X
Fg_Sacasaltolinea = ValLcntH
End Function

Function Resp_Cancel(Title As String)
Dim msg As String, Style, Help, Ctxt
msg = "Cancelar Operación ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
'respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
respuesta = MsgBox(msg, Style, Title) ', Help, Ctxt)
End Function

Function Resp_Delete(Title As String)
Dim msg As String, Style, Help, Ctxt
msg = "Confirma Eliminar ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
respuesta = MsgBox(msg, Style, Title, Help, Ctxt)
End Function

Function Resp_Habilitar(Title As String)
Dim msg As String, Style, Help, Ctxt
msg = "Todos las opciones subordinadas al selecionar serán habilidata. ¿ Cancelar Operación ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
respuesta = MsgBox(msg, Style, Title)
End Function

Function Resp_Deshabilitar(Title As String)
Dim msg As String, Style, Help, Ctxt
msg = "Todos las opciones subordinadas serán deshabilidata. ¿ Cancelar Operación ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
respuesta = MsgBox(msg, Style, Title)
End Function

Function Fg_Sacapuntocoma(ByVal Parentesis As String) As String
Dim X%
Dim ValLcntH$
ValLcntH = ""

For X = 1 To Len(Parentesis)
    If Asc(Mid(Parentesis, X, 1)) <> 59 Then
       ValLcntH = ValLcntH + Mid(Parentesis, X, 1)
    End If
Next X
Fg_Sacapuntocoma = ValLcntH
End Function

Function Fg_Sacapunto(ByVal Punto As String) As String
Dim X%
Dim ValLcntH$
ValLcntH = ""

For X = 1 To Len(Punto)
    If Asc(Mid(Punto, X, 1)) <> 46 Then
       ValLcntH = ValLcntH + Mid(Punto, X, 1)
    Else
       Exit For
    End If
Next X
Fg_Sacapunto = ValLcntH
End Function

Function fg_buscacbo(Combo As Object, Index As Integer, Largo As Integer, cBusca As String)
Dim i As Integer
fg_buscacbo = -1
For i = 0 To Combo(Index).ListCount - 1
    Combo(Index).ListIndex = i
    If Mid(Trim(Combo(Index).text), Len(Trim(Combo(Index).text)) - Largo, Largo) = Val(cBusca) Then
        fg_buscacbo = i
        Exit For
    End If
Next
End Function

Function fg_buscacbostring(Combo As Object, Index As Integer, Largo As Integer, cBusca As String)

Dim i As Integer
fg_buscacbostring = -1

For i = 0 To Combo(Index).ListCount - 1
    
    Combo(Index).ListIndex = i
    
    If Mid(Trim(Combo(Index).text), Len(Trim(Combo(Index).text)) - Largo, Largo) = (cBusca) Then
        
        fg_buscacbostring = i
        Exit For
    
    End If

Next

End Function

Function fg_codigocbo(Combo As Object, Index As Integer, Largo As Integer, cDefa As Variant)

'If Len(Trim(Combo(Index).Text)) > Largo Then
    If Combo(Index).ListIndex = -1 Then fg_codigocbo = "0": Exit Function
    fg_codigocbo = Mid(Trim(Combo(Index).text), Len(Trim(Combo(Index).text)) - Largo, Largo)
'    fg_codigocbo = Combo(Index).List(Combo(Index).ListIndex)
'Else
'    fg_codigocbo = cDefa
'End If

End Function

Function fg_codigolista(List As Object, Index As Integer, Largo As Integer, cDefa As Variant)
If Len(Trim(List.text)) > Largo Then
    fg_codigolista = Mid(Trim(List.text), Len(Trim(List.text)) - Largo, Largo)
Else
    fg_codigolista = cDefa
End If
End Function

Function fg_Busca_En_Lista(List As Control, dato As String, Largo As Integer) As Integer
Dim i%
fg_Busca_En_Lista = 0
For i = 0 To List.ListCount - 1
    If fg_pone_cero(Trim(fg_Quitachar(Right(List.List(i), Largo), "(")), Largo) = dato Then
        fg_Busca_En_Lista = 1
        Exit For
    End If
Next
End Function

Function ValidarCampos(TvwDir As Object, Node As Node)
Dim ivalidar As Integer
ivalidar = 0
SendKeys "{enter}"
If ivalidar = 0 And LimpiaDato(Trim(TvwDir.Nodes(Node.Index).text)) = "" Then
   ivalidar = 1: MsgBox "Descripción Debe ser Informado", vbExclamation + vbOKOnly, "Tipo De Platos"
   Set TvwDir.SelectedItem = Node 'nd
'   Set Node = TvwDir.SelectedItem
'   Set TvwDir.SelectedItem = dest 'nd
   TvwDir.StartLabelEdit
   Exit Function
End If
End Function

Function fg_mes(Fecha As String)
Dim mes As Long, Año As Long
mes = Mid(Fecha, 1, 2)
Año = Mid(Fecha, 4, 4)
Select Case mes
  Case 1, 3, 5, 7, 8, 10, 12
    mes = 31
  Case 2
    If (Año Mod 4) = 0 Then
       mes = 29
    Else
       mes = 28
    End If
  Case 4, 6, 9, 11
    mes = 30
End Select
fg_mes = mes
End Function

Function fg_bcoenter(ByVal bcoenter As String) As String
Dim X%
Dim ValLcntH$
ValLcntH = ""

For X = 1 To Len(bcoenter)
    If Asc(Mid(bcoenter, X, 1)) <> 13 Then
       ValLcntH = ValLcntH + Mid(bcoenter, X, 1)
    End If
Next X
fg_bcoenter = ValLcntH
End Function

Function fg_Dia(Fecha As String) As String
    Dim MiFecha$, dia%, FechaP$, DiaSem%
    If Fecha = "" Then Exit Function
    MiFecha = Format$(fg_Ctod1(Fecha), "mmm dd yyyy")
    dia = Day(DateValue(MiFecha))
    DiaSem = Weekday(DateValue(MiFecha))
    fg_Dia = DiaSem
End Function

Function fg_NomDia(dia As Long) As String
    Select Case dia
    Case 1
        fg_NomDia = "Lunes"
    Case 2
        fg_NomDia = "Martes"
    Case 3
        fg_NomDia = "Miércoles"
    Case 4
        fg_NomDia = "Jueves"
    Case 5
        fg_NomDia = "Viernes"
    Case 6
        fg_NomDia = "Sábado"
    Case 7
        fg_NomDia = "Domingo"
    End Select
End Function

Function fg_NumDia(dia As String) As Integer
    Select Case dia
    Case "Lunes"
        fg_NumDia = 1
    Case "Martes"
        fg_NumDia = 2
    Case "Miércoles"
        fg_NumDia = 3
    Case "Jueves"
        fg_NumDia = 4
    Case "Viernes"
        fg_NumDia = 5
    Case "Sábado"
        fg_NumDia = 6
    Case "Domingo"
        fg_NumDia = 7
    End Select
End Function

Function MiFunc(Ficha As String, Ini As String, Concepto As String) As String
   On Error GoTo FileError
   lpApplicationName = Ficha
   lpDefault = ""
   lpReturnString = Space(128)
   nSize = Len(lpReturnString)
   lpFileName = LCase(App.Path) & "\" & Ini
   lpKeyName = Concepto
   Valid = GetPrivateProfileString(lpApplicationName, _
                                   lpKeyName, _
                                   lpDefault, _
                                   lpReturnString, _
                                   nSize, _
                                   lpFileName)
   MiFunc = Left(lpReturnString, Valid)
   Exit Function
FileError:
   Exit Function
End Function

Function LogoEmp()
If Not fg_ValidaArchivo(Trim(vg_DirLog)) Then Exit Function
With Preview.VSPrinter
    Preview.P1.Picture = LoadResPicture(101, vbResBitmap)
'    .X1 = 500: .X2 = 1745 + .X1: .Y1 = 500: .Y2 = 955 + .Y1
    .x1 = 500: .X2 = Preview.P1.Width + .x1: .Y1 = 500: .Y2 = Preview.P1.Height + .Y1
    .Picture = Preview.P1.Picture
    ExportPicture Preview.VSPrinter, Trim(vg_DirLog), True
    .CurrentX = 500: .CurrentY = .Y2 + 50
End With
End Function

Function fg_ValidaArchivo(nArc As String) As Boolean
Dim fso As Object
fg_ValidaArchivo = True
Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(nArc) Then fg_ValidaArchivo = False
End Function

Sub ExportPicture(vp As VSPrinter, sPicFile As String, bEmbed As Boolean)
'
' This function inserts a reference to a picture into a
' VSPrinter export file (HTML or RTF). Parameters are:
'
' vp:       VSPrinter control. If the ExportFile property is empty, there is
'           no export file and the function returns immediately.
'
' sPicFile: Name of the file that contains the picture to be inserted in
'           the document. Callers may use the SavePicture to create this file
'           if they have to.
'
' bEmbed:   If True, the picture is embedded in the RTF file.
'           If False, a link to the picture is inserted in the RTF file.
'           Embedding the picture makes the RTF file larger, but in this
'           case you no longer need the original picture file.
'
' Note: Pictures are never embedded in HTML files.
'
Dim rtfPic As String
    ' no export file? no work!
    If Len(vp.ExportFile) = 0 Then Exit Sub
    
    ' linked HTML
    If vp.ExportFormat < vpxRTF Then
        Dim ml!
        ml = vp.IndentLeft / 1440
        vp.ExportRaw = vbCrLf & _
            "<img style='margin-left:" & ml & "in;' src=" & sPicFile & ">" & vbCrLf
        Exit Sub
    End If
    
    ' linked RTF
    If Not bEmbed Then
        
        ' escape backslashes
        Dim sPicFileEsc$
        sPicFileEsc = Replace(sPicFile, "\", "\\\\")
        
        ' save it
        vp.ExportRaw = vbCrLf & _
            "{\field{\*\fldinst { INCLUDEPICTURE " & sPicFileEsc & " \\* MERGEFORMAT \\d }}}" & vbCrLf
        Exit Sub
    End If
    
    ' embedded RTF
    ' paste the picture into a rich edit control, then get the RTF out of it.
    Clipboard.Clear
    Clipboard.SetData LoadPicture(sPicFile)
    rtfPic = ""
    sendmessage Preview.rtfPic.hwnd, WM_PASTE, 0, 0
    Preview.rtfPic.SelIndent = vp.IndentLeft
    vp.ExportRaw = vbCrLf & Preview.rtfPic.TextRTF & vbCrLf
    Clipboard.Clear
End Sub

Sub Gl_Mo_Botones(Form As Object, op As Integer)

Dim BtnX As Object, btnX1 As Object
    
    With Form.Toolbar1.Buttons
        
        Select Case op
        
        Case 1 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 2 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR DATO-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Datos"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 3 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR DATO-FILTRAR-BUSQUEDA-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDropdown, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Datos": BtnX.ButtonMenus.Add text:="Copiar Recetas": BtnX.ButtonMenus.Add text:="Mover Recetas": BtnX.ButtonMenus.Add text:="Bach - Input Receta": BtnX.ButtonMenus.Add text:="Bach - Input Método Receta" 'btnX.ButtonMenus.Add Text:="Pegar Recetas": btnX.ButtonMenus.Add Text:="Mover Recetas"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
'            Set BtnX = .Add(, "A_BuscarPro", , tbrDefault, "A_BuscarPro"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar Productos"
'            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_BuscarPro", , tbrDropdown, "A_BuscarPro"): BtnX.Visible = True: BtnX.ToolTipText = "Reemplazar datos detalle receta": BtnX.ButtonMenus.Add text:="Reemplazar Ingredientes": BtnX.ButtonMenus.Add text:="Reemplazar % Ingredientes"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "Vinculo Ingrediente", , tbrDefault, "Vinculo Ingrediente"): BtnX.Visible = True:: BtnX.Enabled = False: BtnX.ToolTipText = "Ver Vinculos"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 4 '-------> INCLUIR-GRABAR-CANCELAR(ANULAR)-HISTORICO-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo "
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = "Grabar "
            Set BtnX = .Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Anular "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar "
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "
        
        Case 5 '-------> GRABAR-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = "Grabar "
            Set BtnX = .Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "
        
        Case 6 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico "
            Set BtnX = .Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Ajuste", , tbrDefault, "A_Ajuste"): BtnX.Visible = True: BtnX.ToolTipText = "Ajustar Inventario "
            Set BtnX = .Add(, "A_AnulaAjuste", , tbrDefault, "A_AnulaAjuste"): BtnX.Visible = True: BtnX.ToolTipText = "Anular Ajuste Inventario "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 7 '-------> INCLUIR-GRABAR-BORRAR-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo "
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = "Grabar "
            Set BtnX = .Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "
        
        Case 8 '-------> INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
            
            Form.Toolbar2.ImageList = Partida.IL1
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir Ingrediente"
            Set BtnX = Form.Toolbar2.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar Ingrediente"
            Set BtnX = Form.Toolbar2.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar2.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = Form.Toolbar2.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = Form.Toolbar2.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar2.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Consultar Ingrediente"
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Aportes Nutricionales"
            Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir Ingrediente"
            Set BtnX = Form.Toolbar2.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
        
        Case 9 '-------> INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
            
            Form.Toolbar3.ImageList = Partida.IL1
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar"
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar3.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar3.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Consultar"
            Set BtnX = Form.Toolbar3.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir"
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
        
        Case 10 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR DATO-EXPORTAR PRECIO-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Lista Precio"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_ImportarPrecio", , tbrDropdown, "A_ImportarPrecio"): BtnX.Visible = False: BtnX.ToolTipText = "Importar Precio": BtnX.ButtonMenus.Add text:="Desde SAC": BtnX.ButtonMenus.Add text:="Desde Excel"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
            Form.Toolbar2.ImageList = Partida.IL1
            Set btnX1 = Form.Toolbar2.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): btnX1.Visible = True: btnX1.ToolTipText = "Incluir Nuevo Mes"
            Set btnX1 = Form.Toolbar2.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): btnX1.Visible = False: btnX1.ToolTipText = ""
        
        Case 11 '-------> INCLUIR(1)-BORRAR(3)-CANCELAR(6)-
            
            Form.Toolbar4.ImageList = Partida.IL1
            Set BtnX = Form.Toolbar4.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir Formato de Compras"
            Set BtnX = Form.Toolbar4.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar4.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar Formato de Compras"
            Set BtnX = Form.Toolbar4.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar4.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Vincular Formato de Compras"
            Set BtnX = Form.Toolbar4.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar4.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = Form.Toolbar4.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar4.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = Form.Toolbar4.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar4.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar4.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = Form.Toolbar4.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
        
        Case 12 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-FIJAR VINCULO-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = IIf(vg_auxcod = "sap", False, True): BtnX.ToolTipText = "Borrar Vinculo"
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = IIf(vg_auxcod = "sap", True, False): BtnX.ToolTipText = ""
            Set BtnX = .Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar a Excel "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Visible = False
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = IIf(vg_auxcod = "sap", False, True): BtnX.ToolTipText = "Vincular Formato de Compras"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            If vg_auxcod = "sap" Then
               
'               Set BtnX = .Add(, "A_ImportarDatos", , tbrDefault, "A_ImportarDatos"): BtnX.Visible = True: BtnX.ToolTipText = "Importar Formato"
               Set BtnX = .Add(, "A_ImportarDatos", , tbrDropdown, "A_ImportarDatos"): BtnX.Visible = True: BtnX.ToolTipText = "Importar Formato": BtnX.ButtonMenus.Add text:="Desde Formato SAP": BtnX.ButtonMenus.Add text:="Desde Formato SAP Justicia"
               Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
               Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
               Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
               Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            
            End If
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 13 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Días Feriados"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 14 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR DATO-EXPORTAR PRECIO-IMPRIMIR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDropdown, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar y Importar Datos": BtnX.ButtonMenus.Add text:="Copiar Datos": BtnX.ButtonMenus.Add text:="Importar Datos"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_ImportarPrecio", , tbrDropdown, "A_ImportarPrecio"): BtnX.Visible = False: BtnX.ToolTipText = "Importar Precio": BtnX.ButtonMenus.Add text:="Desde SAC": BtnX.ButtonMenus.Add text:="Desde Excel"
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 15
            
            Form.Toolbar1.ImageList = Partida.IL1
'            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
'            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 16 'M_SsllPrecioRef(CARGA EXCEL-BORRAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR)
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Cargar Excel"
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 17 '-------> INCLUIR(1)-BORRAR(3)-CANCELAR(6)-
            
            Form.Toolbar5.ImageList = Partida.IL1
            Set BtnX = Form.Toolbar5.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir Formato de Compras"
            Set BtnX = Form.Toolbar5.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar5.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar Formato de Compras"
            Set BtnX = Form.Toolbar5.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar5.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Vincular Formato de Compras"
            Set BtnX = Form.Toolbar5.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar5.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = Form.Toolbar5.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar5.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = Form.Toolbar5.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar5.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar5.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = Form.Toolbar5.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
        
        Case 18
            
            Form.Toolbar3.ImageList = Partida.IL1
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar3.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = Form.Toolbar3.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
        
        Case 19 '-------> INCLUIR-ALTERAR-BORRAR-SALIR
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Modificar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            
'            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
'            Set BtnX = .Add(, "A_Reporcesar", , tbrDefault, "A_Reporcesar"): BtnX.Visible = True: BtnX.ToolTipText = "Ejecución "
            
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 20 ''''''acp(02102013) '-------> Confirmar-Imprimir-Copiar-Salir-
            
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
            Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_CopiarD", , tbrDropdown, "A_CopiarD"): BtnX.Visible = True: BtnX.ButtonMenus.Add text:="Copiar Formato de Compras": BtnX.ButtonMenus.Add text:="Bach - Input Ingresar": BtnX.ButtonMenus.Add text:="Bach - Input Modificar": BtnX.ButtonMenus.Add text:="Bach - Input Eliminar"
            'Set BtnX = .Add(, "A_ImportarPrecio", , tbrDropdown, "A_ImportarPrecio"): BtnX.Visible = False: BtnX.ToolTipText = "Importar Precio": BtnX.ButtonMenus.Add text:="Desde SAC": BtnX.ButtonMenus.Add text:="Desde Excel"
            
            Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
        
        Case 21
        
            Form.Toolbar1.ImageList = Partida.IL1
            Set BtnX = .Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
            Set BtnX = .Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
            Set BtnX = .Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
            Set BtnX = .Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
            
            Set BtnX = .Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
            Set BtnX = .Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
            Set BtnX = .Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
            Set BtnX = .Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Imprimir ", , tbrDropdown, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir ": BtnX.ButtonMenus.Add text:="Imprimir Usuario Perfil": BtnX.ButtonMenus.Add text:="Transacciones Usuarios"
            Set BtnX = .Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
            Set BtnX = .Add(, , "", tbrDefault, 0): BtnX.Enabled = False
            Set BtnX = .Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
                    
'            Set BtnX = .Add(, "A_CopiarD", , tbrDropdown, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Datos": BtnX.ButtonMenus.Add text:="Copiar Recetas": BtnX.ButtonMenus.Add text:="Mover Recetas": BtnX.ButtonMenus.Add text:="Bach - Input Receta" 'btnX.ButtonMenus.Add Text:="Pegar Recetas": btnX.ButtonMenus.Add Text:="Mover Recetas"

        End Select
    End With
End Sub

Function Gl_Ac_Botones(Form As Form, Op1 As Integer, Op2 As Integer, modo As String)

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir As String
'-----------------------------VALIDAR USUARIO-----------------
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_ValidarUsuario_02 '" & vg_NUsr & "', " & Form.HelpContextID & " ")
    
    If Not RS1.EOF Then
        
        Do While Not RS1.EOF
            
            incluir = RS1!dpe_deragr
            alterar = RS1!dpe_dermod
            eliminar = RS1!dpe_dereli
            imprimir = RS1!dpe_derimp
            
            RS1.MoveNext
        
        Loop
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '--------------------------------------------------------------
    With Form.Toolbar1
    Select Case Op1
    
    Case 1
        
        Select Case Op2
            
            Case 0 'CANCELAR-CONFIRMAR
                
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    .Buttons(1).Visible = False: .Buttons(2).Visible = True
                    .Buttons(3).Visible = False: .Buttons(4).Visible = True
                    .Buttons(5).Visible = False: .Buttons(6).Visible = True
                    .Buttons(7).Visible = False: .Buttons(8).Visible = True
                    .Buttons(10).Visible = True: .Buttons(11).Visible = False
                    .Buttons(12).Visible = True: .Buttons(13).Visible = False
                    .Buttons(15).Visible = False: .Buttons(16).Visible = True
                End If
            
            Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 2 'INCLUIR
                
                If incluir = "1" Then
                    
                    .Buttons(1).Visible = True: .Buttons(2).Visible = False
                
                End If
                
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 3 'Habilitar Solamente Salir
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 4 'ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 5 '-------> ALTERAR-ACTUALIZAR-IMPRIMIR
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 7 '-------> ALTERAR-IMPRIMIR-SALIR
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                If alterar = "1" Then .Buttons(3).Visible = True: .Buttons(4).Visible = False
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 8 '-------> ALTERAR-BORRAR-SALIR
                
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = True: .Buttons(6).Visible = False
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 9
                
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = True: .Buttons(6).Visible = False
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = True: .Buttons(11).Visible = False
                .Buttons(12).Visible = True: .Buttons(13).Visible = False
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 10 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 11 '-------> ACTUALIZAR-IMPRIMIR-SALIR
                
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 12 'BORRAR-IMPRIMIR
                
                .Buttons(1).Visible = False
                .Buttons(2).Visible = True
                .Buttons(3).Visible = False
                .Buttons(4).Visible = True
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 13 'Alterar-ACTUALIZAR-IMPRIMIR
                
                .Buttons(1).Visible = False
                .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 14 'INCLUIR-IMPRIMIR
                
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = False
                .Buttons(4).Visible = True
                .Buttons(5).Visible = False
                .Buttons(6).Visible = True
                '.Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 15 'INCLUIR-BORRAR-ACTUALIZAR-IMPRIMIR
                
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = False
                .Buttons(4).Visible = True
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
            
            Case 16 'INCLUIR-ALTERAR-BORRAR
                
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(6).Visible = IIf(eliminar = "1", True, False)
                .Buttons(7).Visible = IIf(eliminar = "1", False, True)
            
            Case 17 'acp
                
                Form.Toolbar1.Buttons(2).Visible = True: Form.Toolbar1.Buttons(1).Visible = False
                Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
                Form.Toolbar1.Buttons(7).Visible = IIf(incluir = "1", True, False)
                Form.Toolbar1.Buttons(9).Visible = True
                
            Case 18 '-------> ALTERAR-BORRAR-SALIR
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(16).Visible = True
            
            Case 19 '-------> ALTERAR-BORRAR-IMPRIMIR-SALIR
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = IIf(imprimir = "1", True, False)
                .Buttons(16).Visible = IIf(imprimir = "1", False, True)
        
        End Select
    
    Case 2
        Select Case Op2
            Case 0 '-------> CANCELAR-CONFIRMAR
                
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    
                    .Buttons(1).Visible = False: .Buttons(2).Visible = True
                    .Buttons(3).Visible = False: .Buttons(4).Visible = True
                    .Buttons(5).Visible = False: .Buttons(6).Visible = True
                    .Buttons(7).Visible = False: .Buttons(8).Visible = True
                    .Buttons(10).Visible = True: .Buttons(11).Visible = False
                    .Buttons(12).Visible = True: .Buttons(13).Visible = False
                    .Buttons(15).Enabled = False
                    .Buttons(17).Visible = False: .Buttons(18).Visible = True
                
                End If
            
            Case 1 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = IIf(incluir = "1", True, False)
                .Buttons(17).Visible = IIf(imprimir = "1", True, False)
                .Buttons(18).Visible = IIf(imprimir = "1", False, True)
            Case 2 '-------> INCLUIR
                
                If incluir = "1" Then
                   
                    .Buttons(1).Visible = True: .Buttons(2).Visible = False
                
                End If
                
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Visible = False: .Buttons(18).Visible = True
            
            Case 3 '-------> Habilitar Solamente Salir
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Visible = False: .Buttons(18).Visible = True
        
        End Select
    
    Case 3
        Select Case Op2
            Case 0 'CANCELAR-CONFIRMAR
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    .Buttons(1).Visible = False: .Buttons(2).Visible = True
                    .Buttons(3).Visible = False: .Buttons(4).Visible = True
                    .Buttons(5).Visible = False: .Buttons(6).Visible = True
                    .Buttons(7).Visible = False: .Buttons(8).Visible = True
                    .Buttons(10).Visible = True: .Buttons(11).Visible = False
                    .Buttons(12).Visible = True: .Buttons(13).Visible = False
                    .Buttons(15).Enabled = False
                    .Buttons(17).Enabled = False
                    .Buttons(19).Enabled = False
                    .Buttons(21).Visible = False: .Buttons(22).Visible = True
                    .Buttons(24).Enabled = False
                End If
            Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = IIf(incluir = "1", True, False)
                .Buttons(17).Enabled = True 'IIf(incluir = "1", True, False)
                .Buttons(19).Enabled = IIf(incluir = "1", True, False)
                .Buttons(21).Visible = IIf(imprimir = "1", True, False)
                .Buttons(22).Visible = IIf(imprimir = "1", False, True)
            Case 2 'INCLUIR
                If incluir = "1" Then
                    .Buttons(1).Visible = True: .Buttons(2).Visible = False
                End If
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Enabled = True
                .Buttons(19).Enabled = False
                .Buttons(21).Visible = False: .Buttons(22).Visible = True
            Case 3 'Habilitar Solamente Salir
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Enabled = False
                .Buttons(19).Enabled = False
                .Buttons(21).Visible = False: .Buttons(22).Visible = True
        End Select
    Case 4
        .Refresh
        Select Case Op2
            Case 1 'Ninguno
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = True
                .Buttons(9).Visible = False: .Buttons(10).Visible = True
                .Buttons(11).Visible = True: .Buttons(12).Visible = True
            Case 2, 5 'Grabar
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = IIf(Op2 = 2, IIf(incluir = 1, True, False), False)
                .Buttons(4).Visible = IIf(Op2 = 2, IIf(incluir = 0, True, False), True)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = True
                .Buttons(9).Visible = False: .Buttons(10).Visible = True
                .Buttons(11).Visible = True: .Buttons(12).Visible = True
            Case 3, 4 'Anular - Imprimir
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = IIf(Op2 = 3, IIf(eliminar = 1, True, False), False)
                .Buttons(6).Visible = IIf(Op2 = 3, IIf(eliminar = 0, True, False), True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = True
                .Buttons(9).Visible = IIf(imprimir = 1, True, False)
                .Buttons(10).Visible = IIf(imprimir = 0, True, False)
                .Buttons(11).Visible = True: .Buttons(12).Visible = True
        End Select
    Case 5
        .Refresh
        Select Case Op2
            Case 1 'Grabar
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
            Case 2 'Imprimir
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = True: .Buttons(4).Visible = False
        End Select
    Case 7 'INCLUIR-GRABAR-BORRAR-IMPRIMIR-SALIR
        .Refresh
        Select Case Op2
            Case 1 'Incluir - Grabar
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = IIf(incluir = 1, True, False)
                .Buttons(4).Visible = IIf(incluir = 0, True, False)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(8).Visible = False: .Buttons(9).Visible = True
                .Buttons(11).Visible = True
            Case 2 'Eliminar - Imprimir
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = IIf(eliminar = 1, True, False)
                .Buttons(6).Visible = IIf(eliminar = 0, True, False)
                .Buttons(8).Visible = IIf(imprimir = 1, True, False)
                .Buttons(9).Visible = IIf(imprimir = 0, True, False)
                .Buttons(11).Visible = True
        End Select
    Case 8 'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-COPIAR(12)-IMPRIMIR(13)
        Form.Toolbar2.Refresh
        Select Case Op2
            Case 0 'CANCELAR-CONFIRMAR
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    Form.Toolbar2.Buttons(1).Visible = False: Form.Toolbar2.Buttons(2).Visible = True
                    Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
                    Form.Toolbar2.Buttons(6).Visible = True: Form.Toolbar2.Buttons(7).Visible = False
                    Form.Toolbar2.Buttons(8).Visible = True: Form.Toolbar2.Buttons(9).Visible = False
                    Form.Toolbar2.Buttons(11).Enabled = False
                    Form.Toolbar2.Buttons(12).Enabled = False
                    Form.Toolbar2.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
                End If
            Case 1 'INCLUIR-BORRAR-BUSCAR-COPIAR-IMPRIMIR
                Form.Toolbar2.Buttons(1).Visible = IIf(incluir = "1", True, False)
                Form.Toolbar2.Buttons(2).Visible = IIf(incluir = "0", True, False)
                Form.Toolbar2.Buttons(3).Visible = IIf(eliminar = "1", True, False)
                Form.Toolbar2.Buttons(4).Visible = IIf(eliminar = "0", True, False)
                Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
                Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
                Form.Toolbar2.Buttons(11).Enabled = True
                Form.Toolbar2.Buttons(12).Enabled = True
                Form.Toolbar2.Buttons(13).Visible = IIf(imprimir = "1", True, False)
                Form.Toolbar2.Buttons(14).Visible = IIf(imprimir = "0", True, False)
            Case 2 'INCLUIR
                If incluir = "1" Then
                    Form.Toolbar2.Buttons(1).Visible = True: Form.Toolbar2.Buttons(2).Visible = False
                End If
                Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
                Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
                Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
                Form.Toolbar2.Buttons(11).Enabled = False
                Form.Toolbar2.Buttons(12).Enabled = False
                Form.Toolbar2.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
        End Select
    Case 9 'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(13)
        Form.Toolbar3.Refresh
        Select Case Op2
            Case 0 'CANCELAR-CONFIRMAR
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    Form.Toolbar3.Buttons(1).Visible = False: Form.Toolbar3.Buttons(2).Visible = True
                    Form.Toolbar3.Buttons(3).Visible = False: Form.Toolbar3.Buttons(4).Visible = True
                    Form.Toolbar3.Buttons(6).Visible = True: Form.Toolbar3.Buttons(7).Visible = False
                    Form.Toolbar3.Buttons(8).Visible = True: Form.Toolbar3.Buttons(9).Visible = False
                    Form.Toolbar3.Buttons(11).Enabled = False
                    Form.Toolbar3.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
                End If
            Case 1 'INCLUIR-BORRAR-BUSCAR-COPIAR-IMPRIMIR
                Form.Toolbar3.Buttons(1).Visible = IIf(incluir = "1", True, False)
                Form.Toolbar3.Buttons(2).Visible = IIf(incluir = "0", True, False)
                Form.Toolbar3.Buttons(3).Visible = IIf(eliminar = "1", True, False)
                Form.Toolbar3.Buttons(4).Visible = IIf(eliminar = "0", True, False)
                Form.Toolbar3.Buttons(6).Visible = False: Form.Toolbar3.Buttons(7).Visible = True
                Form.Toolbar3.Buttons(8).Visible = False: Form.Toolbar3.Buttons(9).Visible = True
                Form.Toolbar3.Buttons(11).Enabled = True
                Form.Toolbar3.Buttons(13).Visible = IIf(imprimir = "1", True, False)
                Form.Toolbar3.Buttons(14).Visible = IIf(imprimir = "0", True, False)
            Case 2 'INCLUIR
                If incluir = "1" Then
                    Form.Toolbar3.Buttons(1).Visible = True: Form.Toolbar3.Buttons(2).Visible = False
                End If
                Form.Toolbar3.Buttons(3).Visible = False: Form.Toolbar3.Buttons(4).Visible = True
                Form.Toolbar3.Buttons(6).Visible = False: Form.Toolbar3.Buttons(7).Visible = True
                Form.Toolbar3.Buttons(8).Visible = False: Form.Toolbar3.Buttons(9).Visible = True
                Form.Toolbar3.Buttons(11).Enabled = False
                Form.Toolbar3.Buttons(13).Visible = False: Form.Toolbar3.Buttons(14).Visible = True
        End Select
    Case 10
        Select Case Op2
            Case 0 '-------> CANCELAR-CONFIRMAR
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    .Buttons(1).Visible = False: .Buttons(2).Visible = True
                    .Buttons(3).Visible = False: .Buttons(4).Visible = True
                    .Buttons(5).Visible = False: .Buttons(6).Visible = True
                    .Buttons(7).Visible = False: .Buttons(8).Visible = True
                    .Buttons(10).Visible = True: .Buttons(11).Visible = False
                    .Buttons(12).Visible = True: .Buttons(13).Visible = False
                    .Buttons(15).Enabled = False
                    .Buttons(17).Enabled = False
                    .Buttons(19).Visible = False
                    .Buttons(20).Visible = True ': .Buttons(21).Visible = True
                End If
            Case 1 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = IIf(incluir = "1", True, False)
        '        .Buttons(17).Enabled = IIf(incluir = "1", True, False)
                .Buttons(19).Visible = IIf(imprimir = "1", True, False)
                .Buttons(20).Visible = IIf(imprimir = "0", True, False)
            
                Form.Toolbar2.Buttons(1).Visible = False
                Form.Toolbar2.Buttons(2).Visible = True
   
            Case 2 '-------> INCLUIR
                If incluir = "1" Then
                    .Buttons(1).Visible = True: .Buttons(2).Visible = False
                End If
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Enabled = True
                .Buttons(19).Enabled = False
                .Buttons(20).Visible = False: .Buttons(21).Visible = True
            Case 3 '-------> Habilitar Solamente Salir
                
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Enabled = False
                .Buttons(19).Enabled = False
                .Buttons(20).Visible = False: .Buttons(21).Visible = True
            
            Case 4
                
                Form.Toolbar2.Buttons(1).Visible = IIf(modo = "A", False, True)
                Form.Toolbar2.Buttons(2).Visible = IIf(modo = "A", True, False)
            
            Case 5
                
                Form.Toolbar2.Buttons(1).Visible = False
                Form.Toolbar2.Buttons(2).Visible = True
        
        End Select
    
    Case 11 'INCLUIR(1)-BORRAR(3)-COPIAR(5)
        Form.Toolbar4.Refresh
        Select Case Op2
            
            Case 0
                
                Form.Toolbar4.Buttons(7).Visible = True: Form.Toolbar4.Buttons(8).Visible = False
                Form.Toolbar4.Buttons(9).Visible = True: Form.Toolbar4.Buttons(10).Visible = False
                Form.Toolbar4.Buttons(12).Visible = False: Form.Toolbar4.Buttons(13).Visible = True
            
            Case 1 'INCLUIR-BORRAR-VINCULAR-IMPRIMIR
                
                Form.Toolbar4.Buttons(1).Visible = IIf(incluir = "1", True, False)
                Form.Toolbar4.Buttons(2).Visible = IIf(incluir = "0", True, False)
                Form.Toolbar4.Buttons(3).Visible = IIf(alterar = "1", True, False)
                Form.Toolbar4.Buttons(4).Visible = IIf(alterar = "0", True, False)
                Form.Toolbar4.Buttons(5).Visible = IIf(alterar = "0", False, True)
                Form.Toolbar4.Buttons(7).Visible = False: Form.Toolbar4.Buttons(8).Visible = True
                Form.Toolbar4.Buttons(9).Visible = False: Form.Toolbar4.Buttons(10).Visible = True
                Form.Toolbar4.Buttons(12).Visible = IIf(imprimir = "1", True, False)
                Form.Toolbar4.Buttons(13).Visible = IIf(imprimir = "0", True, False)
        
        End Select
    
    Case 12
        
        Select Case Op2
            
            Case 0 '-------> ALTERAR-BORRAR-SALIR
                
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = True: .Buttons(6).Visible = False
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = True
            
            Case 1
                
                .Buttons(1).Visible = True: .Buttons(2).Visible = False
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = True: .Buttons(6).Visible = False
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = True: .Buttons(11).Visible = False
                .Buttons(12).Visible = True: .Buttons(13).Visible = False
                .Buttons(15).Visible = True
        
        End Select
    Case 13
        
        Select Case Op2
            Case 0 'CANCELAR-CONFIRMAR
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    .Buttons(1).Visible = False: .Buttons(2).Visible = True
                    .Buttons(3).Visible = False: .Buttons(4).Visible = True
                    .Buttons(5).Visible = False: .Buttons(6).Visible = True
                    .Buttons(7).Visible = False: .Buttons(8).Visible = True
                    .Buttons(10).Visible = True: .Buttons(11).Visible = False
                    .Buttons(12).Visible = True: .Buttons(13).Visible = False
                    .Buttons(15).Enabled = False
                    .Buttons(17).Visible = False: .Buttons(18).Visible = True
                End If
            Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = True: .Buttons(15).Enabled = True
                .Buttons(17).Visible = IIf(imprimir = "1", True, False)
                .Buttons(18).Visible = IIf(imprimir = "1", False, True)
            Case 2 'INCLUIR
                If incluir = "1" Then
                    .Buttons(1).Visible = True: .Buttons(2).Visible = False
                End If
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Visible = False: .Buttons(18).Visible = True
            Case 3 'ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = False
                .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = False: .Buttons(15).Enabled = False
                .Buttons(17).Visible = IIf(imprimir = "1", True, False)
                .Buttons(18).Visible = IIf(imprimir = "1", False, True)
                .Buttons(20).Visible = False: .Buttons(20).Enabled = False
        End Select
    Case 14
        Select Case Op2
            Case 0 '-------> CANCELAR-CONFIRMAR
                If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                    .Buttons(1).Visible = False: .Buttons(2).Visible = True
                    .Buttons(3).Visible = False: .Buttons(4).Visible = True
                    .Buttons(5).Visible = False: .Buttons(6).Visible = True
                    .Buttons(7).Visible = False: .Buttons(8).Visible = True
                    .Buttons(10).Visible = True: .Buttons(11).Visible = False
                    .Buttons(12).Visible = True: .Buttons(13).Visible = False
                    .Buttons(15).Enabled = False
                    .Buttons(17).Enabled = False
                    .Buttons(19).Visible = False
                    .Buttons(20).Visible = True ': .Buttons(21).Visible = True
                End If
            Case 1 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = IIf(incluir = "1", True, False)
                .Buttons(2).Visible = IIf(incluir = "1", False, True)
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = IIf(eliminar = "1", True, False)
                .Buttons(6).Visible = IIf(eliminar = "1", False, True)
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = IIf(incluir = "1", True, False)
        '        .Buttons(17).Enabled = IIf(incluir = "1", True, False)
                .Buttons(19).Visible = IIf(imprimir = "1", True, False)
                .Buttons(20).Visible = IIf(imprimir = "0", True, False)
    
            Case 2 '-------> INCLUIR
                If incluir = "1" Then
                    .Buttons(1).Visible = True: .Buttons(2).Visible = False
                End If
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Enabled = True
                .Buttons(19).Enabled = False
                .Buttons(20).Visible = False: .Buttons(21).Visible = True
            Case 3 '-------> Habilitar Solamente Salir
                .Buttons(1).Visible = False: .Buttons(2).Visible = True
                .Buttons(3).Visible = False: .Buttons(4).Visible = True
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = False: .Buttons(8).Visible = True
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Enabled = False
                .Buttons(17).Enabled = False
                .Buttons(19).Enabled = False
                .Buttons(20).Visible = False: .Buttons(21).Visible = True
            Case 4 'Alterar-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = False
                .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(19).Visible = IIf(imprimir = "1", True, False)
                .Buttons(20).Visible = IIf(imprimir = "1", False, True)
            Case 5 'Alterar-ACTUALIZAR-IMPRIMIR
                .Buttons(1).Visible = False
                .Buttons(2).Visible = True
                .Buttons(3).Visible = IIf(alterar = "1", True, False)
                .Buttons(4).Visible = IIf(alterar = "1", False, True)
                .Buttons(5).Visible = False: .Buttons(6).Visible = True
                .Buttons(7).Visible = True: .Buttons(8).Visible = False
                .Buttons(10).Visible = False: .Buttons(11).Visible = True
                .Buttons(12).Visible = False: .Buttons(13).Visible = True
                .Buttons(15).Visible = True: .Buttons(15).Enabled = True
                .Buttons(19).Visible = IIf(imprimir = "1", True, False)
                .Buttons(20).Visible = IIf(imprimir = "1", False, True)
        Case 15
            Select Case Op2
                Case 0 'CANCELAR-CONFIRMAR
                    If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
                        .Buttons(2).Enabled = True
                        .Buttons(1).Enabled = True
                        .Buttons(2).Visible = True
                        .Buttons(1).Visible = True
                    End If
            End Select
        End Select
    Case 15 'INCLUIR(1)-BORRAR(3)-COPIAR(5)
        Form.Toolbar5.Refresh
        Select Case Op2
            Case 0
                Form.Toolbar5.Buttons(7).Visible = True: Form.Toolbar5.Buttons(8).Visible = False
                Form.Toolbar5.Buttons(9).Visible = True: Form.Toolbar5.Buttons(10).Visible = False
                Form.Toolbar5.Buttons(12).Visible = False: Form.Toolbar5.Buttons(13).Visible = True
            Case 1 'INCLUIR-BORRAR-VINCULAR-IMPRIMIR
                Form.Toolbar5.Buttons(1).Visible = IIf(incluir = "1", True, False)
                Form.Toolbar5.Buttons(2).Visible = IIf(incluir = "0", True, False)
                Form.Toolbar5.Buttons(3).Visible = IIf(alterar = "1", True, False)
                Form.Toolbar5.Buttons(4).Visible = IIf(alterar = "0", True, False)
                Form.Toolbar5.Buttons(5).Visible = IIf(alterar = "0", False, True)
                Form.Toolbar5.Buttons(7).Visible = False: Form.Toolbar5.Buttons(8).Visible = True
                Form.Toolbar5.Buttons(9).Visible = False: Form.Toolbar5.Buttons(10).Visible = True
                Form.Toolbar5.Buttons(12).Visible = IIf(imprimir = "1", True, False)
                Form.Toolbar5.Buttons(13).Visible = IIf(imprimir = "0", True, False)
        End Select

'Case 8
'    Form.Toolbar2.Refresh
'    Select Case Op2
'    Case 1 'Ninguno
'        Form.Toolbar2.Buttons(1).Visible = True: Form.Toolbar2.Buttons(2).Visible = False
'        Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
'        Form.Toolbar2.Buttons(5).Visible = False: Form.Toolbar2.Buttons(6).Visible = True
'        Form.Toolbar2.Buttons(7).Visible = True: Form.Toolbar2.Buttons(8).Visible = True
'        Form.Toolbar2.Buttons(9).Visible = False: Form.Toolbar2.Buttons(10).Visible = True
'    Case 2 'Grabar
'        Form.Toolbar2.Buttons(1).Visible = True: Form.Toolbar2.Buttons(2).Visible = False
'        Form.Toolbar2.Buttons(3).Visible = IIf(incluir = 1, True, False)
'        Form.Toolbar2.Buttons(4).Visible = IIf(incluir = 0, True, False)
'        Form.Toolbar2.Buttons(5).Visible = False: Form.Toolbar2.Buttons(6).Visible = True
'        Form.Toolbar2.Buttons(7).Visible = True: Form.Toolbar2.Buttons(8).Visible = True
'        Form.Toolbar2.Buttons(9).Visible = False: Form.Toolbar2.Buttons(10).Visible = True
'    Case 3, 4 'Grabar - Eliminar - Imprimir
'        Form.Toolbar2.Buttons(1).Visible = True: Form.Toolbar2.Buttons(2).Visible = False
'        Form.Toolbar2.Buttons(3).Visible = IIf(Op2 = 3, IIf(incluir = 1, True, False), False)
'        Form.Toolbar2.Buttons(4).Visible = IIf(Op2 = 3, IIf(incluir = 0, True, False), True)
'        'Form.Toolbar2.Buttons(3).Visible = IIf(incluir = 1, True, False)
'        'Form.Toolbar2.Buttons(4).Visible = IIf(incluir = 0, True, False)
'        Form.Toolbar2.Buttons(5).Visible = IIf(eliminar = 1, True, False)
'        Form.Toolbar2.Buttons(6).Visible = IIf(eliminar = 0, True, False)
'        Form.Toolbar2.Buttons(7).Visible = True: Form.Toolbar2.Buttons(8).Visible = True
'        Form.Toolbar2.Buttons(9).Visible = IIf(imprimir = 1, True, False)
'        Form.Toolbar2.Buttons(10).Visible = IIf(imprimir = 0, True, False)
'    End Select
    End Select
    End With
End Function

Function Gl_Ac_BotonesRealPropuesta(Form As Form, Op1 As Integer, Op2 As Integer, modo As String, Indppr As String, ProducInppr As String)
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir As String
'-----------------------------VALIDAR USUARIO-----------------
RS1.Open "SELECT distinct dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
         "INNER JOIN a_usuarioperfil usp ON per.per_codigo = usp.usp_codper " & _
         "WHERE usp.usp_codusu = '" & vg_NUsr & "' and dpe.dpe_codopc = " & Form.HelpContextID, vg_db, adOpenStatic
If Not RS1.EOF Then
    Do While Not RS1.EOF
        incluir = RS1!dpe_deragr
        alterar = RS1!dpe_dermod
        eliminar = RS1!dpe_dereli
        imprimir = RS1!dpe_derimp
        RS1.MoveNext
    Loop
End If
RS1.Close: Set RS1 = Nothing
'--------------------------------------------------------------

Select Case Op1
Case 1
    Select Case Op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            If Indppr = ProducInppr Then
              Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
              Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
              Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
              Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
              Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
              Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
              Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
            Else
              Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = False
              Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = False
              Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = False
              Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = False
              Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
              Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
              Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = False
            End If
        End If
    Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        If Indppr = ProducInppr Then
          Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
          Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
          Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
          Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
          Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
          Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
          Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
          Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
          Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
          Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
        Else
          Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
          Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
          Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", False, True)
          Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", True, False)
          Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", False, True)
          Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", True, False)
          Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
          Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
          Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
          Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
        End If
    End Select
Case 8 'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-COPIAR(12)-IMPRIMIR(13)
    Form.Toolbar2.Refresh
    Select Case Op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
          If Indppr = ProducInppr Then
            Form.Toolbar2.Buttons(1).Visible = False: Form.Toolbar2.Buttons(2).Visible = True
            Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
            Form.Toolbar2.Buttons(6).Visible = True: Form.Toolbar2.Buttons(7).Visible = False
            Form.Toolbar2.Buttons(8).Visible = True: Form.Toolbar2.Buttons(9).Visible = False
            Form.Toolbar2.Buttons(11).Enabled = False
            Form.Toolbar2.Buttons(12).Enabled = False
            Form.Toolbar2.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
          Else
            
            Form.Toolbar2.Buttons(1).Visible = IIf(incluir = "1", True, False)
            Form.Toolbar2.Buttons(2).Visible = IIf(incluir = "0", False, True)
            Form.Toolbar2.Buttons(3).Visible = IIf(eliminar = "1", False, True)
            Form.Toolbar2.Buttons(4).Visible = IIf(eliminar = "0", True, False)
            Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = False
            Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = False
            Form.Toolbar2.Buttons(11).Enabled = True
            Form.Toolbar2.Buttons(12).Enabled = True
            Form.Toolbar2.Buttons(13).Visible = IIf(imprimir = "1", True, False)
            Form.Toolbar2.Buttons(14).Visible = IIf(imprimir = "0", False, True)
            
            
'            Form.Toolbar2.Buttons(1).Visible = False: Form.Toolbar2.Buttons(2).Visible = False
'            Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = False
'            Form.Toolbar2.Buttons(6).Visible = True: Form.Toolbar2.Buttons(7).Visible = True
'            Form.Toolbar2.Buttons(8).Visible = True: Form.Toolbar2.Buttons(9).Visible = True
'            Form.Toolbar2.Buttons(11).Enabled = False
'            Form.Toolbar2.Buttons(12).Enabled = False
'            Form.Toolbar2.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
          End If
        End If
    Case 1 'INCLUIR-BORRAR-BUSCAR-COPIAR-IMPRIMIR
        If Indppr = ProducInppr Then
          Form.Toolbar2.Buttons(1).Visible = IIf(incluir = "1", True, False)
          Form.Toolbar2.Buttons(2).Visible = IIf(incluir = "0", True, False)
          Form.Toolbar2.Buttons(3).Visible = IIf(eliminar = "1", True, False)
          Form.Toolbar2.Buttons(4).Visible = IIf(eliminar = "0", True, False)
          Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
          Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
          Form.Toolbar2.Buttons(11).Enabled = True
          Form.Toolbar2.Buttons(12).Enabled = True
          Form.Toolbar2.Buttons(13).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar2.Buttons(14).Visible = IIf(imprimir = "0", True, False)
        Else
          Form.Toolbar2.Buttons(1).Visible = IIf(incluir = "1", False, True)
          Form.Toolbar2.Buttons(2).Visible = IIf(incluir = "0", True, False)
          Form.Toolbar2.Buttons(3).Visible = IIf(eliminar = "1", False, True)
          Form.Toolbar2.Buttons(4).Visible = IIf(eliminar = "0", True, False)
          Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
          Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
          Form.Toolbar2.Buttons(11).Enabled = True
          Form.Toolbar2.Buttons(12).Enabled = True
          Form.Toolbar2.Buttons(13).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar2.Buttons(14).Visible = IIf(imprimir = "0", False, True)
        End If
    End Select
Case 9 'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(13)
    Form.Toolbar3.Refresh
    Select Case Op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
          If Indppr = ProducInppr Then
            Form.Toolbar3.Buttons(1).Visible = False: Form.Toolbar3.Buttons(2).Visible = True
            Form.Toolbar3.Buttons(3).Visible = False: Form.Toolbar3.Buttons(4).Visible = True
            Form.Toolbar3.Buttons(6).Visible = True: Form.Toolbar3.Buttons(7).Visible = False
            Form.Toolbar3.Buttons(8).Visible = True: Form.Toolbar3.Buttons(9).Visible = False
            Form.Toolbar3.Buttons(11).Enabled = False
            Form.Toolbar3.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
          Else
            Form.Toolbar3.Buttons(1).Visible = False: Form.Toolbar3.Buttons(2).Visible = False
            Form.Toolbar3.Buttons(3).Visible = False: Form.Toolbar3.Buttons(4).Visible = False
            Form.Toolbar3.Buttons(6).Visible = True: Form.Toolbar3.Buttons(7).Visible = True
            Form.Toolbar3.Buttons(8).Visible = True: Form.Toolbar3.Buttons(9).Visible = True
            Form.Toolbar3.Buttons(11).Enabled = False
            Form.Toolbar3.Buttons(13).Visible = False: Form.Toolbar2.Buttons(14).Visible = True
          End If
        End If
    Case 1 'INCLUIR-BORRAR-BUSCAR-COPIAR-IMPRIMIR
      If Indppr = ProducInppr Then
          Form.Toolbar3.Buttons(1).Visible = IIf(incluir = "1", True, False)
          Form.Toolbar3.Buttons(2).Visible = IIf(incluir = "0", True, False)
          Form.Toolbar3.Buttons(3).Visible = IIf(eliminar = "1", True, False)
          Form.Toolbar3.Buttons(4).Visible = IIf(eliminar = "0", True, False)
          Form.Toolbar3.Buttons(6).Visible = False: Form.Toolbar3.Buttons(7).Visible = True
          Form.Toolbar3.Buttons(8).Visible = False: Form.Toolbar3.Buttons(9).Visible = True
          Form.Toolbar3.Buttons(11).Enabled = True
          Form.Toolbar3.Buttons(13).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar3.Buttons(14).Visible = IIf(imprimir = "0", True, False)
        Else
          Form.Toolbar3.Buttons(1).Visible = IIf(incluir = "1", False, True)
          Form.Toolbar3.Buttons(2).Visible = IIf(incluir = "0", True, False)
          Form.Toolbar3.Buttons(3).Visible = IIf(eliminar = "1", False, True)
          Form.Toolbar3.Buttons(4).Visible = IIf(eliminar = "0", True, False)
          Form.Toolbar3.Buttons(6).Visible = False: Form.Toolbar3.Buttons(7).Visible = True
          Form.Toolbar3.Buttons(8).Visible = False: Form.Toolbar3.Buttons(9).Visible = True
          Form.Toolbar3.Buttons(11).Enabled = True
          Form.Toolbar3.Buttons(13).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar3.Buttons(14).Visible = IIf(imprimir = "0", True, False)
      End If
    End Select
Case 11 'INCLUIR(1)-BORRAR(3)-COPIAR(5)
    Form.Toolbar4.Refresh
    Select Case Op2
    Case 0
      If Indppr = ProducInppr Then
        Form.Toolbar4.Buttons(7).Visible = True: Form.Toolbar4.Buttons(8).Visible = False
        Form.Toolbar4.Buttons(9).Visible = True: Form.Toolbar4.Buttons(10).Visible = False
        Form.Toolbar4.Buttons(12).Visible = False: Form.Toolbar4.Buttons(13).Visible = True
      Else
        Form.Toolbar4.Buttons(7).Visible = True: Form.Toolbar4.Buttons(8).Visible = False
        Form.Toolbar4.Buttons(9).Visible = True: Form.Toolbar4.Buttons(10).Visible = False
        Form.Toolbar4.Buttons(12).Visible = False: Form.Toolbar4.Buttons(13).Visible = True
      End If
    Case 1 'INCLUIR-BORRAR-VINCULAR-IMPRIMIR
      If Indppr = ProducInppr Then
          Form.Toolbar4.Buttons(1).Visible = IIf(incluir = "1", True, False)
          Form.Toolbar4.Buttons(2).Visible = IIf(incluir = "0", True, False)
          Form.Toolbar4.Buttons(3).Visible = IIf(alterar = "1", True, False)
          Form.Toolbar4.Buttons(4).Visible = IIf(alterar = "0", True, False)
          Form.Toolbar4.Buttons(5).Visible = IIf(alterar = "0", False, True)
          Form.Toolbar4.Buttons(7).Visible = False: Form.Toolbar4.Buttons(8).Visible = True
          Form.Toolbar4.Buttons(9).Visible = False: Form.Toolbar4.Buttons(10).Visible = True
          Form.Toolbar4.Buttons(12).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar4.Buttons(13).Visible = IIf(imprimir = "0", True, False)
        Else
          Form.Toolbar4.Buttons(1).Visible = IIf(incluir = "1", False, True)
          Form.Toolbar4.Buttons(2).Visible = IIf(incluir = "0", True, False)
          Form.Toolbar4.Buttons(3).Visible = IIf(alterar = "1", False, True)
          Form.Toolbar4.Buttons(4).Visible = IIf(alterar = "0", True, False)
          Form.Toolbar4.Buttons(5).Visible = IIf(alterar = "0", False, True)
          Form.Toolbar4.Buttons(7).Visible = False: Form.Toolbar4.Buttons(8).Visible = True
          Form.Toolbar4.Buttons(9).Visible = False: Form.Toolbar4.Buttons(10).Visible = True
          Form.Toolbar4.Buttons(12).Visible = IIf(imprimir = "1", True, False)
          Form.Toolbar4.Buttons(13).Visible = IIf(imprimir = "0", True, False)
      End If
    End Select
End Select
End Function

Function ValidarUsuario(Formu As Form) As String
'--- Esta funcion ha sido reemplazada por la funcion ValidaPerfil
'--- la cual es una copia de esta pero fue mejorada
'--- no utilizar mas esta funcion sino mas bien la funcion ValidaPerfil
'--- no borrar esta funcion ya que aun esta siendo llamada
Dim RS1 As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir
'-----------------------------VALIDAR USUARIO-----------------
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_ValidarUsuario '" & vg_NUsr & "', " & Formu.HelpContextID & " ")
ValidarUsuario = "0000"
DoEvents
If Not RS1.EOF Then
    
    incluir = Trim(RS1!dpe_deragr)
    alterar = Trim(RS1!dpe_dermod)
    eliminar = Trim(RS1!dpe_dereli)
    imprimir = Trim(RS1!dpe_derimp)
    ValidarUsuario = incluir & alterar & eliminar & imprimir

End If
RS1.Close
Set RS1 = Nothing

'--------------------------------------------------------------
End Function

Function ValidaPerfil(Formu As Form) As String
'--- Esta funcion devuelve los accesos asignados al usuario
'--- de manera de deshabilitar o habilitar funciones en el formulario
'--- segun corresponda a las autorizaciones asignadas al usuario actual

Dim RS1 As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir As String, acceder As String
'-----------------------------VALIDAR USUARIO-----------------
'RS1.Open "SELECT dpe.dpe_deracc, dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
'         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
'         "INNER JOIN a_usuarios usu ON per.per_codigo = usu.usu_perfil " & _
'         "WHERE usu.usu_codigo='" & vg_NUsr & "' and dpe.dpe_codopc=" & Formu.HelpContextID, vg_db, adOpenStatic
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
'RS1.Open "SELECT distinct dpe.dpe_deracc, dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
'         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
'         "INNER JOIN a_usuarioperfil usp ON per.per_codigo = usp.usp_codper " & _
'         "WHERE usp.usp_codusu = '" & vg_NUsr & "' and dpe.dpe_codopc=" & Formu.HelpContextID & " order by dpe_deragr desc,dpe.dpe_dermod desc, dpe.dpe_dereli desc, dpe.dpe_derimp desc", vg_db, adOpenStatic
Set RS1 = vg_db.Execute("sgpadm_Sel_ValidarUsuarioAcceso '" & vg_NUsr & "', " & Formu.HelpContextID & " ")
ValidaPerfil = "00000"
DoEvents
If Not RS1.EOF Then
    
    acceder = Trim(RS1!dpe_deracc)
    incluir = Trim(RS1!dpe_deragr)
    alterar = Trim(RS1!dpe_dermod)
    eliminar = Trim(RS1!dpe_dereli)
    imprimir = Trim(RS1!dpe_derimp)
    ValidaPerfil = acceder & incluir & alterar & eliminar & imprimir

End If
RS1.Close
Set RS1 = Nothing

'--------------------------------------------------------------
End Function

'----> Incluido por Andrea M <----
Function BuscarCodReceta(NombreReceta As String) As Integer

Dim RS1 As New ADODB.Recordset
RS1.Open "Select rec_codigo from b_receta where rec_nombre = '" & Trim(NombreReceta) & "'", vg_db, adOpenStatic
DoEvents
If Not RS1.EOF Then
  BuscarCodReceta = RS1!rec_codigo
Else
  BuscarCodReceta = 0
End If
RS1.Close: Set RS1 = Nothing

End Function

Function BuscarCodRecetaSitRem(NombreReceta As String) As Integer
Dim RS1 As New ADODB.Recordset

    Call RS1.Open("Select rec_codigo from b_receta where rec_nombre = '" & Trim(NombreReceta) & "'", vg_db, adOpenStatic)
    DoEvents
    If Not RS1.EOF Then
      Let BuscarCodRecetaSitRem = RS1!rec_codigo
    Else
      Let BuscarCodRecetaSitRem = 0
    End If
    Call RS1.Close
    Set RS1 = Nothing
End Function

Function fg_CalCtoRecInv(CodRec As Long, tiprec As Long, ctacon As String) As Double
Dim RS1 As New ADODB.Recordset
RS1.Open "SELECT a.rec_codigo, a.rec_nombre, " & _
         "SUM(b.red_canpro*c.ing_precos) AS cosrec " & _
         "FROM b_receta a, b_recetadet b, b_ingrediente c, b_productos d " & _
         "WHERE b.red_codigo=a.rec_codigo " & _
         "AND   b.red_codpro=c.ing_codigo " & _
         "AND  (c.ing_codcom=d.pro_codigo OR c.ing_codped=d.pro_codigo) " & _
         "AND   d.pro_ctacon='" & ctacon & "' " & _
         "AND   b.red_codigo=" & CodRec & " " & _
         "GROUP BY a.rec_codigo, a.rec_nombre order by a.rec_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then fg_CalCtoRecInv = RS1!cosrec
RS1.Close: Set RS1 = Nothing
End Function
'MVA - MVI RECALCULO DE LA RECETA MINUTA
Function fg_CalCtoRecInvSitRem_MVI(CodRec As Long, tiprec As Long, ctacon As String, cencos As String, periodo As Long, codReg As Long) As Double
Dim RS1 As New ADODB.Recordset
fg_CalCtoRecInvSitRem_MVI = 0

Set RS1 = vg_db.Execute("sgpadm_s_calcostorecsitioremoto_MVI '" & cencos & "', " & CodRec & ", " & tiprec & ", '" & ctacon & "', " & periodo & "," & codReg)

If Not RS1.EOF Then fg_CalCtoRecInvSitRem_MVI = RS1!cosrec
RS1.Close: Set RS1 = Nothing
End Function
'FIN MVA - MVI RECALCULO DE LA RECETA MINUTA


Function fg_CalCtoRecInvSitRem(CodRec As Long, tiprec As Long, ctacon As String, cencos As String) As Double

Dim RS1 As New ADODB.Recordset
fg_CalCtoRecInvSitRem = 0
Set RS1 = vg_db.Execute("sgpadm_s_calcostorecsitioremoto '" & cencos & "', " & CodRec & ", " & tiprec & ", '" & ctacon & "'")
If Not RS1.EOF Then fg_CalCtoRecInvSitRem = RS1!cosrec
RS1.Close: Set RS1 = Nothing

End Function

Function fg_CalCtoRecListaPrecio(CodRec As Long, tiprec As Long, codlpr As Long, ctacon As String, anomes As Long) As Double

Dim RS1 As New ADODB.Recordset
Set RS1 = vg_db.Execute("sgpadm_s_CalCostoReceta " & vg_codsubseg & "," & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & ",'" & ctacon & "'," & CodRec & ", '" & anomes & "'")
'sgpadm_s_CalCostoReceta 2,10075,1,'410001',2168,200905
If Not RS1.EOF Then fg_CalCtoRecListaPrecio = RS1!cosrec
RS1.Close: Set RS1 = Nothing

End Function

Function Redondear(variable As Variant, numdec As Integer) As Variant

Dim cientos As Long, i As Integer
cientos = 1
For i = 1 To numdec
    cientos = cientos * 10
Next i
If IsNumeric(variable) Then
    If (variable * cientos) Mod 2 <> 0 Then
        Redondear = Round(variable, numdec)
    Else
        If (variable * cientos) - Int((variable * cientos)) >= 0.5 Then
            Redondear = Round((variable + 0.5), numdec)
'            Redondear = Round((variable * cientos + 0.5) / cientos, numdec)
        Else
            Redondear = Round(variable, numdec)
        End If
    End If
Else
    Redondear = variable
End If

End Function

Function fg_CalCtoRecPlan(Fecha As Long, TipMin As String, CodRec As Long) As Double

Dim RS1 As New ADODB.Recordset
RS1.Open "select b_receta.rec_codigo, b_receta.rec_nombre, sum(b_recetadet.red_canpro*b_minutacosto.mic_cospro) as cosrec from b_receta, b_recetadet, b_minutacosto " & _
         "where b_recetadet.red_codigo=b_receta.rec_codigo and b_recetadet.red_codpro=b_minutacosto.mic_codpro and b_recetadet.red_codigo=" & CodRec & " " & _
         "and b_minutacosto.mic_fecval=" & Fecha & " and b_minutacosto.mic_tipmin='" & TipMin & "' group by b_receta.rec_codigo, b_receta.rec_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then fg_CalCtoRecPlan = RS1!cosrec
RS1.Close: Set RS1 = Nothing

End Function

Function TipoDato(variable, valor)

If VarType(variable) = vbNull Then
    
    TipoDato = IIf(VarType(valor) <> vbNull, valor, " ")

Else
    
    TipoDato = variable

End If

End Function
'vbEmpty 0   Empty (no inicializado).
'vbNull  1   Null (sin datos válidos).
'vbInteger   2   Entero.
'vbLong  3   Entero largo.
'vbSingle    4   Un número de punto flotante de precisión simple.
'vbDouble    5   Un número de punto flotante de precisión doble.
'vbCurrency  6   Moneda.
'vbDate  7   Fecha.
'vbString    8   Cadena.
'vbObject    9   Objeto de Automatización OLE.
'vbError 10  Error.
'vbBoolean   11  Boolean.
'vbVariant   12  Variante (utilizada sólo con matrices de Variantes).
'vbDataObject    13  Objeto sin Automatización OLE.
'vbByte  17  Byte
'vbArray 8192    Matriz.

Function fg_BuscaArr(ByVal Arre As Variant, vBus As Variant, ByVal nDim As Long, ByVal nCol As Long) As Long
Dim i As Long
If nDim = 1 Then
    For i = 1 To UBound(Arre)
        If Arre(nCol) = vBus Then Exit For
    Next i
    If i > UBound(Arre) Then i = 0
Else
    For i = 1 To UBound(Arre)
        If Arre(i, nCol) = vBus Then Exit For
    Next i
    If i > UBound(Arre) Then i = 0
End If
fg_BuscaArr = i
End Function

Function fg_Desencripta(psw_encriptada As String) As String

Dim result As String, count As Integer
result = ""
psw_encriptada = Trim$(psw_encriptada)

For count = 1 To Len(psw_encriptada)
    
    result = result & Chr$(Asc(Mid$(psw_encriptada, count, 1)) - 73 - count)

Next

fg_Desencripta = result

End Function

Function fg_Encripta(password As String) As String

Dim encrip As String, count As Integer
encrip = ""
password = Trim$(password)

For count = 1 To Len(password)
    
    encrip = encrip & Chr$(Asc(Mid$(password, count, 1)) + 73 + count)

Next

fg_Encripta = encrip

End Function

Function ValidaBod(ByVal cBod As Long, ByVal cPro As String)

Dim RS As New ADODB.Recordset
RS.Open "select * from b_bodegas where bod_codbod=" & cBod & " and bod_codpro='" & cPro & "'", vg_db, adOpenStatic
If RS.EOF Then
    vg_db.BeginTrans
    vg_db.Execute "insert into b_bodegas values (" & cBod & ", '" & cPro & "', 0)"
    vg_db.CommitTrans
End If
RS.Close: Set RS = Nothing

End Function

Function dBoM(ByVal Fecha As Date) As Date

Dim mes As Integer
mes = Month(Fecha)
Do While mes = Month(Fecha)
    Fecha = Fecha - 1
Loop
dBoM = Fecha + 1

End Function

Function dEoM(ByVal Fecha As Date) As Date

Dim mes As Integer
mes = Month(Fecha)

Do While mes = Month(Fecha)
    
    Fecha = Fecha + 1

Loop

dEoM = Fecha - 1
End Function

Function Bom(Fecha As Date) As String

Dim mes As Integer
mes = Month(Fecha)
Do While mes = Month(Fecha)
    Bom = fg_pone_cero(Year(Fecha), 4) & fg_pone_cero(Month(Fecha), 2) & fg_pone_cero(Day(Fecha), 2)
    Fecha = Fecha - 1
Loop
Bom = Fecha

End Function

Function EoM(ByVal Fecha As Date) As Date

Dim mes As Integer
mes = Month(Fecha)
Do While mes = Month(Fecha)
    Fecha = Fecha + 1
Loop
EoM = Fecha

End Function

Function GetParametro(ByVal cpar As String) As Variant

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_Parametros_V01 '" & cpar & "'")
If Not RS1.EOF Then
    
    Select Case RS1!par_tipo
    
        Case "N"
            
            GetParametro = Val(TipoDato(RS1!par_valor, ""))
        
        Case "C"
            
            GetParametro = Trim(TipoDato(RS1!par_valor, ""))
    
    End Select

Else
    
    GetParametro = Null

End If
RS1.Close
Set RS1 = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

Function fg_WeekNumber(dDate As Date) As Integer
fg_WeekNumber = (DateDiff("ww", CDate("01/01/" & Year(dDate)), dDate, vbMonday, vbFirstFourDays) Mod 52) + 1
End Function

Sub fg_CheckTmp(cTabla As String)
Dim RS1 As New ADODB.Recordset
On Error GoTo ManError
RS1.Open "select * from " & cTabla, vg_db, adOpenStatic
RS1.Close: Set RS1 = Nothing
vg_db.Execute "drop table " & cTabla
Exit Sub
ManError:
    If Err = -2147217865 Then Exit Sub
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error"
End Sub

'Sub fg_CheckTmp(cTabla As String)
'Dim tdfBucle As TableDef
'vg_db.TableDefs.Refresh
'For Each tdfBucle In vg_db.TableDefs
'    If tdfBucle.Name = cTabla Then
'        vg_db.Execute "Drop Table " & cTabla
'    End If
'Next tdfBucle
'End Sub

Function ModCasino() As Boolean
ModCasino = True
RS1.Open "SELECT par_valor FROM a_param WHERE par_codigo='casinomod'", vg_db, adOpenStatic
If Not RS1.EOF Then ModCasino = IIf(RS1!par_valor = 0, False, True)
RS1.Close: Set RS1 = Nothing
End Function

Function MuestraCasino(op As Integer) As String
Dim RS1 As New ADODB.Recordset
MuestraCasino = ""
Select Case op
Case 1
    MuestraCasino = GetParametro("casino")
Case 2
    RS1.Open "select cli_nombre from b_clientes where cli_codigo='" & GetParametro("casino") & "' and cli_tipo=0", vg_db, adOpenStatic
    If Not RS1.EOF Then MuestraCasino = RS1!Cli_nombre
    RS1.Close: Set RS1 = Nothing
End Select
End Function

Function BubbleSort(varArray As Variant, bAscending As Boolean)

'Option Base 0 is assumed
Dim HoldEntry As Long
Dim SwapOccurred As Boolean
Dim iItteration As Integer
Dim i As Integer
SwapOccurred = True
iItteration = 1

Do Until Not SwapOccurred
   
   SwapOccurred = False
   
   For i = LBound(varArray) To UBound(varArray) - iItteration Step 1
       
       If bAscending Then
          
          If varArray(i) > varArray(i + 1) Then
             HoldEntry = varArray(i)
             varArray(i) = varArray(i + 1)
             varArray(i + 1) = HoldEntry
             SwapOccurred = True
          End If

       Else
          
          If varArray(i + 1) > varArray(i) Then
             HoldEntry = varArray(i)
             varArray(i) = varArray(i + 1)
             varArray(i + 1) = HoldEntry
             SwapOccurred = True
          End If
       
       End If
   
   Next i
   
   iItteration = iItteration + 1   'reduce iteration each time as greatest/lowest
                                   'item already at end/start of array
Loop
End Function

Function TeclasNoPermitidas(tecla As Integer) As Boolean
If tecla = 16 Or tecla = 17 Or tecla = 18 Or tecla = 19 Or tecla = 20 Or tecla = 32 Or tecla = 33 Or tecla = 34 Or tecla = 35 Or tecla = 36 Or tecla = 37 Or tecla = 38 Or tecla = 39 Or tecla = 40 Or tecla = 44 Or tecla = 45 Or tecla = 46 Or tecla = 91 Or tecla = 93 Or tecla = 145 Or tecla = 144 Then TeclasNoPermitidas = False: Exit Function
TeclasNoPermitidas = True
End Function

Function SendMail(cObj As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String, Cmailaviso As Integer)
'Dim cNombU As String, cMailU As String
Dim CHost As String, Caddr As String, Cuser As String, Cpass As String
On Error GoTo Man_Error
'cNombU = "Jorge Paz": CmailU = "jpaz@sodexho.cl"
If Trim(CmailU) <> "" Then
    '-------> Traer parametro de cuenta correo
    Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("cor")) & "%'")
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: MsgBox "No existe Parametrización Correo, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Function
    Do While Not RS1.EOF
       If RS1!par_codigo = "corser" Then CHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corcum" Then Caddr = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corusu" Then Cuser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corpas" Then Cpass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    cObj.UnlockComponent "1mundoedwardsMAIL_3e5SOpaZmRkg"
    cObj.SmtpHost = CHost '"10.0.100.21"
    cObj.SmtpUsername = Cuser '"adminbdcasinos"
    cObj.SmtpPassword = Cpass '"admincasinos"
    cObj.ConnectTimeout = 30
'    Dim email As ChilkatEmail, Success As Long
'    Set email = New ChilkatEmail
    Dim email As New ChilkatEmail2, Success As Long
    email.AddTo cNombU, CmailU
    email.Subject = cSubject
    email.Body = cBody
    If Cmailaviso = 1 Then email.AddFileAttachment cArchivo
    If UCase(Mid(cArchivo, Len(cArchivo) - 23, 2)) = "MP" Then
       email.FromName = IIf(Cmailaviso = 1, "Administrador Maestro de Productos", "")
    ElseIf UCase(Mid(cArchivo, Len(cArchivo) - 23, 2)) = "MR" Then
       email.FromName = IIf(Cmailaviso = 1, "Administrador Maestro de Recetas", "Administrador Maestro de Recetas")
    ElseIf UCase(Mid(cArchivo, Len(cArchivo) - 23, 2)) = "MM" Then
       email.FromName = IIf(Cmailaviso = 1, "Administrador Maestro de Planificación", "Administrador Maestro de Planificación")
    End If
    email.FromAddress = Caddr '"adminbdcasinos@sodexho.cl"
    cObj.LogMailSentFilename = "mailSent.log"
    Success = cObj.SendEmail(email)
    If (Success = 0) Then
        MsgBox cObj.LastErrorText
    End If
End If
Exit Function
Man_Error:
    cObj.SaveXmlLog "log.xml"
    If Err = -2147467259 Then MsgBox "Cuenta no válida" & VgLinea & Trim(cNombU), vbCritical, "Envio Archivos a contratos": Exit Function
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error al enviar mail."
End Function

Sub LeerArbProducto(codigo As String, codtec As String, nivel As String, posin1 As Integer, posfi1 As Integer, posin2 As Integer, posfi2 As Integer, op As Integer)
Dim RS1 As New ADODB.Recordset
Dim indice As String
indice = ""
If nivel = "1" Then
   RS1.Open "select cdarvprod, cdprodinte from produto where cdprodinte='" & codigo & "' and cdnvproduto='" & nivel & "'", vg_dbtec, adOpenStatic
Else
   RS1.Open "select cdarvprod, cdprodinte from produto where cdprodinte='" & codigo & "' and cdnvproduto='" & nivel & "' and (substr(cdarvprod," & posin2 & "," & posfi2 & ")='" & codtec & "')", vg_dbtec, adOpenStatic
End If
If RS1.EOF Then
   RS1.Close: Set RS1 = Nothing
   If op = 1 Then
      RS1.Open "select cdarvprod pro from produto where cdnvproduto='" & nivel & "' order by pro desc", vg_dbtec, adOpenStatic
   Else
      RS1.Open "select substr(cdarvprod," & posin1 & "," & posfi1 & ") pro from produto where cdnvproduto='" & nivel & "' and substr(cdarvprod," & posin2 & "," & posfi2 & ")='" & IIf(nivel = "5", CStr(Mid(codtec, 1, 7)), CStr(codtec)) & "' order by pro desc", vg_dbtec, adOpenStatic
   End If
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If IsNumeric(RS1!pro) Then indice = Val(RS1!pro + 1): Exit Do
         RS1.MoveNext
      Loop
   Else
      indice = "1"
   End If
   RS1.Close: Set RS1 = Nothing
   vg_codigo = indice: vg_opcion = 1
Else
   vg_codigo = Mid(RS1!cdarvprod, posin1, posfi1)
   vg_opcion = 2
   RS1.Close: Set RS1 = Nothing
End If
End Sub

Sub LeerArbProductoAnexo(CodIng As String, codigo As String, codtec As String, nivel As String, posin1 As Integer, posfi1 As Integer, posin2 As Integer, posfi2 As Integer, op As Integer)
Dim RS1 As New ADODB.Recordset
Dim indice As String
indice = ""
If nivel = "1" Then
   RS1.Open "select cdarvprod, cdprodinte from produto where cdprodinte='" & codigo & "' and cdnvproduto='" & nivel & "'", vg_dbtec, adOpenStatic
Else
   RS1.Open "select cdproduto, cdarvprod, cdprodinte from produto where cdprodinte='" & codigo & "' and cdnvproduto='" & nivel & "' and (substr(cdarvprod," & posin2 & "," & posfi2 & ")='" & codtec & "')", vg_dbtec, adOpenStatic
End If
If RS1.EOF Then
   RS1.Close: Set RS1 = Nothing
   If op = 1 Then
      RS1.Open "select cdarvprod pro from produto where cdnvproduto='" & nivel & "' order by pro desc", vg_dbtec, adOpenStatic
   Else
      RS1.Open "select substr(cdarvprod," & posin1 & "," & posfi1 & ") pro from produto where cdnvproduto='" & nivel & "' and substr(cdarvprod," & posin2 & "," & posfi2 & ")='" & IIf(nivel = "5", CStr(Mid(codtec, 1, 7)), CStr(codtec)) & "' order by pro desc", vg_dbtec, adOpenStatic
   End If
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If IsNumeric(RS1!pro) Then indice = Val(RS1!pro + 1): Exit Do
         RS1.MoveNext
      Loop
   Else
      indice = IIf((CodIng = "720" Or CodIng = "762" Or CodIng = "742"), 0, 1)
   End If
   RS1.Close: Set RS1 = Nothing
   vg_codigo = indice: vg_opcion = 1
Else
  vg_codigo = RS1!cdproduto
   vg_codigo = Mid(RS1!cdarvprod, posin1, posfi1)
   vg_opcion = 2
   RS1.Close: Set RS1 = Nothing
End If
End Sub

Sub LeerArbReceta(codigo As String, codtec As String, nivel As String, posin1 As Integer, posfi1 As Integer, posin2 As Integer, posfi2 As Integer, op As Integer)
Dim RS1 As New ADODB.Recordset
Dim indice As String

indice = ""
If nivel = "1" Then
   RS1.Open "select cdprato, usr_cdpratinte from prato where usr_cdpratinte='" & codigo & "' and cdnvprato='" & nivel & "'", vg_dbtec, adOpenStatic
Else
'   RS1.Open "select cdprato, usr_cdpratinte from prato where usr_cdpratinte='" & codigo & "' and cdnvprato='" & nivel & "' and (substr(cdprato," & posin2 & "," & posfi2 & ")='" & codtec & "' or " & Val(codtec) & ">0)", vg_dbtec, adOpenStatic
   RS1.Open "select cdprato, usr_cdpratinte from prato where usr_cdpratinte='" & codigo & "' and cdnvprato='" & nivel & "' and (substr(cdprato," & posin2 & "," & posfi2 & ")='" & codtec & "')", vg_dbtec, adOpenStatic
End If
If RS1.EOF Then
   RS1.Close: Set RS1 = Nothing
   If op = 1 Then
      RS1.Open "select cdprato pra from prato where cdnvprato='" & nivel & "' order by pra desc", vg_dbtec, adOpenStatic
   Else
      RS1.Open "select substr(cdprato," & posin1 & "," & posfi1 & ") pra from prato where cdnvprato='" & nivel & "' and substr(cdprato," & posin2 & "," & posfi2 & ")='" & IIf(nivel = "6", CStr(Mid(codtec, 1, 7)), CStr(codtec)) & "' order by pra desc", vg_dbtec, adOpenStatic
   End If
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         If IsNumeric(RS1!pra) Then indice = Val(RS1!pra + 1): Exit Do
         RS1.MoveNext
      Loop
   Else
      indice = IIf(nivel = 6, "0", "1")
   End If
   RS1.Close: Set RS1 = Nothing
   vg_codigo = indice: vg_opcion = 1
Else
   vg_codigo = Mid(RS1!cdprato, posin1, posfi1)
   vg_opcion = 2
   RS1.Close: Set RS1 = Nothing
End If
End Sub
Public Function BuscarNivel(TablaGen As String, SufGen As String, codigo As Long) As String
Dim RS1 As New ADODB.Recordset
Dim codpadre As Long
codpadre = 0: BuscarNivel = "1"
RS1.Open "select " & SufGen & "codigo, " & SufGen & "previo from " & TablaGen & " where " & SufGen & "codigo=" & codigo & "", vg_db, adOpenKeyset
If Not RS1.EOF Then codpadre = RS1(1)
RS1.Close: Set RS1 = Nothing
RS1.Open "select " & SufGen & "codigo, " & SufGen & "previo from " & TablaGen & " where " & SufGen & "codigo=" & codpadre & "", vg_db, adOpenStatic
If Not RS1.EOF Then BuscarNivel = IIf(RS1(1) = 0, "2", "3")
RS1.Close: Set RS1 = Nothing
End Function

Function GrabarProductoTecfood(codfa1 As String, codfa2 As String, codfa3 As String, CodIng As String, noming As String, unimed As Long, codume As Long, codpro As String, nompro As String, faconv As Double, prodact As String) As Boolean
Dim RS1 As New ADODB.Recordset
Dim indtec As String, auxtec As String
Dim i As Long
GrabarProductoTecfood = False

'----- Arbol de producto nivel 1
indtec = "": auxtec = ""
vg_codigo = "": vg_opcion = 0
LeerArbProducto codfa1, 0, 1, 1, 1, 1, 1, 1
indtec = vg_codigo
If Val(indtec) > 9 Then GrabarProductoTecfood = True: MsgBox "Tecfood considera un maximo de 9 niveles, proceso cancelado ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = indtec
If vg_opcion = 1 Then
   RS1.Open "select tip_nombre from a_tipopro where tip_codigo=" & codfa1 & "", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarProductoTecfood = True: MsgBox "No existe familia producto...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
   vg_dbtec.Execute "insert into produto (cdproduto, cdnvproduto, nmproduto, vrpesounid, vrpefatocorr, vrfatoconv, cdprodinte, idimpproduto, idpesaprod, cdarvprod, idprodprod, idprodativpr, idtppreco, idcobtxserv, idsollocimp, iddiasvalida) " & _
                    "values ('" & auxtec & "', '1', '" & Trim(Mid(RS1!tip_nombre, 1, 50)) & "', 1, 1, 1, '" & codfa1 & "', '1', 'N', '" & auxtec & "', 'N', 'S', '1', 'S', 'N', 'M')"
   RS1.Close: Set RS1 = Nothing
End If
'----- Fin arbol de producto nivel 1

'----- Arbol de producto nivel 2
vg_codigo = "": vg_opcion = 0
LeerArbProducto codfa2, auxtec, 2, 2, 2, 1, 1, 2
indtec = vg_codigo
If Val(indtec) > 99 Then GrabarProductoTecfood = True: MsgBox "Tecfood considera un maximo de 99 niveles, proceso cancelado ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & fg_pone_cero(indtec, 2)
If vg_opcion = 1 Then
   RS1.Open "select tip_nombre from a_tipopro where tip_codigo=" & codfa2 & "", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarProductoTecfood = True: MsgBox "No existe familia producto...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
   vg_dbtec.Execute "insert into produto (cdproduto, cdnvproduto, nmproduto, vrpesounid, vrpefatocorr, vrfatoconv, cdprodinte, idimpproduto, idpesaprod, cdarvprod, idprodprod, idprodativpr, idtppreco, idcobtxserv, idsollocimp, iddiasvalida) " & _
                    "values ('" & auxtec & "', '2', '" & Trim(Mid(RS1!tip_nombre, 1, 50)) & "', 1, 1, 1, '" & codfa2 & "', '1', 'N', '" & auxtec & "', 'N', 'S', '1', 'S', 'N', 'M')"
   RS1.Close: Set RS1 = Nothing
End If
'----- Fin arbol de producto nivel 2

'----- Arbol de producto nivel 3
vg_codigo = "": vg_opcion = 0
LeerArbProducto codfa3, auxtec, 3, 4, 2, 1, 3, 2
indtec = vg_codigo
If Val(indtec) > 99 Then GrabarProductoTecfood = True: MsgBox "Tecfood considera un maximo de 99 niveles, proceso cancelado ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & fg_pone_cero(indtec, 2)
If vg_opcion = 1 Then
   RS1.Open "select tip_nombre from a_tipopro where tip_codigo=" & codfa3 & "", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarProductoTecfood = True: MsgBox "No existe familia producto...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
   vg_dbtec.Execute "insert into produto (cdproduto, cdnvproduto, nmproduto, vrpesounid, vrpefatocorr, vrfatoconv, cdprodinte, idimpproduto, idpesaprod, cdarvprod, idprodprod, idprodativpr, idtppreco, idcobtxserv, idsollocimp, iddiasvalida) " & _
                    "values ('" & auxtec & "', '3', '" & Trim(Mid(RS1!tip_nombre, 1, 50)) & "', 1, 1, 1, '" & codfa3 & "', '1', 'N', '" & auxtec & "', 'N', 'S', '1', 'S', 'N', 'M')"
   RS1.Close: Set RS1 = Nothing
End If
'----- Fin arbol de producto nivel 3

'----- Agregar Ingredientes nivel 4
vg_codigo = "": vg_opcion = 0
LeerArbProducto CodIng, auxtec, 4, 6, 2, 1, 5, 2
indtec = vg_codigo
If Val(indtec) > 99 Then GrabarProductoTecfood = True: MsgBox "Tecfood considera un maximo de 99 niveles, proceso cancelado ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & fg_pone_cero(indtec, 2)
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into produto (cdproduto, cdnvproduto, nmproduto, vrpesounid, vrpefatocorr, vrfatoconv, cdprodinte, idimpproduto, idpesaprod, cdarvprod, idprodprod, idprodativpr, idtppreco, idcobtxserv, idsollocimp, iddiasvalida) " & _
                    "values ('" & auxtec & "', '4', '" & Trim(Mid(noming, 1, 50)) & "', 1, 1, 1, '" & CodIng & "', '1', 'N', '" & auxtec & "', 'N', 'S', '1', 'S', 'N', 'M')"
Else
   vg_dbtec.Execute "update produto set nmproduto='" & Trim(Mid(noming, 1, 50)) & "' where cdproduto='" & auxtec & "' and cdarvprod='" & auxtec & "' and cdnvproduto='4' and cdprodinte='" & CodIng & "'"
End If
'----- Fin agregar ingredientes nivel 4

'----- Agregar Ingredientes nivel 5
 If CodIng <> "720" And CodIng <> "762" And CodIng <> "742" Then
   vg_codigo = "": vg_opcion = 0
   LeerArbProducto "I" & CodIng, auxtec, 5, 8, 3, 1, 7, 2
   indtec = vg_codigo
   auxtec = auxtec & "000"
   '------- Traer codigo unidad medida ingrediente
   Dim uninco As String
'   RS1.Open "select uni_nomcor from a_unidad  where uni_codunm=" & IIf(unimed <> codume, codume, unimed) & "", vg_db, adOpenStatic
   RS1.Open "select uni_nomcor from a_unidad where uni_codigo=" & codume & "", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarProductoTecfood = True: MsgBox "No existe unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
   uninco = UCase(Trim(RS1!uni_nomcor))
   RS1.Close: Set RS1 = Nothing
   If vg_opcion = 1 Then
      vg_dbtec.Execute "insert into produto (cdproduto, cdnvproduto, sgunidade, nmproduto, vrpesounid, vrpefatocorr, cdprodesto, vrfatoconv, cdprodinte, idimpproduto, idpesaprod, cdarvprod, idprodprod, idprodativpr, idtppreco, idcobtxserv, idsollocimp, iddiasvalida) " & _
                       "values ('" & auxtec & "', '5', '" & UCase(Trim(uninco)) & "', '" & Trim(Mid(noming, 1, 50)) & "', 1, 0, '" & auxtec & "', 1, '" & "I" & CodIng & "', '1', 'N', '" & auxtec & "', 'N', 'S', '1', 'S', 'N', 'M')"
      '------- grabar aportes nutricionales
      RS1.Open "select * from b_productonut where pnu_codpro='" & CodIng & "'", vg_db, adOpenStatic
      If Not RS1.EOF Then
         vg_dbtec.Execute "delete from nutrprod where cdproduto='" & auxtec & "'"
         Do While Not RS1.EOF
            vg_dbtec.Execute "insert into nutrprod (cdproduto, cdnutriente, qtnutrprod) values ('" & auxtec & "', '" & fg_pone_cero(RS1!pnu_codapo, 3) & "', " & RS1!pnu_canapo & ")"
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
      '------- Fin grabar aportes nutricionales
   Else
'      vg_dbtec.Execute "update produto set nmproduto='" & Trim(Mid(noming, 1, 50)) & "', sgunidade='" & UCase(Trim(uninco)) & "' where cdproduto='" & auxtec & "' and cdarvprod='" & auxtec & "' and cdnvproduto='5' and cdprodinte='" & "I" & coding & "'"
      vg_dbtec.Execute "update produto set nmproduto='" & Trim(Mid(noming, 1, 50)) & "'  where cdproduto='" & auxtec & "' and cdarvprod='" & auxtec & "' and cdnvproduto='5' and cdprodinte='" & "I" & CodIng & "'"
   End If
End If
'----- Fin agregar ingredientes nivel 5

'----- Agregar ó modificar Producto nivel 5
If Trim(codpro) = "" Then Exit Function
vg_codigo = "": vg_opcion = 0
LeerArbProductoAnexo CodIng, codpro, Mid(auxtec, 1, 7), 5, 8, 3, 1, 7, 2
indtec = vg_codigo
If Val(indtec) > 999 Then RS1.Close: Set RS1 = Nothing: vg_db.RollbackTrans: MsgBox "Tecfood considera un maximo de 999 productos ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = Mid(auxtec, 1, 7) & fg_pone_cero(indtec, 3)
RS1.Open "select uni_nomcor from a_unidad where uni_codigo=" & codume & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarProductoTecfood = True:  MsgBox "No existe unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into produto (cdproduto, cdnvproduto, sgunidade, nmproduto, nmprodingles, vrpesounid, vrpefatocorr, cdprodesto, vrfatoconv, cdprodinte, idimpproduto, idpesaprod, cdarvprod, idprodprod, idprodativpr, idtppreco, idcobtxserv, idsollocimp, iddiasvalida) " & _
                    "values ('" & auxtec & "', '5', '" & UCase(Trim(RS1!uni_nomcor)) & "', '" & Trim(Mid(nompro, 1, 50)) & "', 'PRODUCTOS DE COMPRA', 1, 0, '" & IIf(CodIng = "720" Or CodIng = "762" Or CodIng = "742", auxtec, Mid(auxtec, 1, 7) & "000") & "', " & faconv & ", '" & codpro & "', '1', 'N', '" & auxtec & "', '" & prodact & "', 'S', '1', 'S', 'N', 'M')"
Else
'   vg_dbtec.Execute "update produto set nmproduto='" & Trim(Mid(NomPro, 1, 50)) & "', sgunidade='" & UCase(Trim(RS1!uni_nomcor)) & "', vrfatoconv=" & faconv & ", idprodprod='" & prodact & "' where cdproduto='" & auxtec & "' and cdarvprod='" & auxtec & "' and cdnvproduto='5' and cdprodinte='" & codpro & "'"
   vg_dbtec.Execute "update produto set nmproduto='" & Trim(Mid(nompro, 1, 50)) & "', idprodprod='" & prodact & "' where cdproduto='" & auxtec & "' and cdarvprod='" & auxtec & "' and cdnvproduto='5' and cdprodinte='" & codpro & "'"
End If
RS1.Close: Set RS1 = Nothing
'----- Fin agregar ó modificar producto nivel 5
End Function

Function GrabarRecetaTecfood(CodRec As String, nomrec As String, nomfan As String, coddi1 As String, coddi2 As String, codti1 As String, codti2 As String, codti3 As String, Form As Object, op As String) As Boolean

Dim RS1     As New ADODB.Recordset
Dim RS2     As New ADODB.Recordset
Dim RS3     As New ADODB.Recordset
Dim indtec  As String, auxtec As String
Dim i       As Long

Let GrabarRecetaTecfood = False

'----- Arbol de receta nivel 1
indtec = "": auxtec = ""
vg_codigo = "": vg_opcion = 0
LeerArbReceta CStr(coddi1), 0, 1, 1, 1, 1, 1, 1
indtec = vg_codigo
If Val(indtec) > 9 Then GrabarRecetaTecfood = True: MsgBox "Tecfood considera un maximo de 9 niveles ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = indtec
RS1.Open "select car_nombre from a_recetacatdie where car_codigo=" & Val(coddi1) & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe categoria dietetica ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into prato (cdfilauxprat, cdprato, nmprato, cdnvprato, usr_cdpratinte) " & _
                    "values ('P', '" & auxtec & "', '" & UCase(Trim(Mid(RS1!car_nombre, 1, 40))) & "', 1, '" & coddi1 & "')"
Else
   vg_dbtec.Execute "update prato set nmprato='" & UCase(Trim(Mid(RS1!car_nombre, 1, 40))) & "' where cdprato='" & auxtec & "' and cdnvprato='1' and usr_cdpratinte='" & coddi1 & "'"
End If
RS1.Close: Set RS1 = Nothing
'----- Fin arbol de receta nivel 1

'----- Arbol de receta nivel 2
vg_codigo = "": vg_opcion = 0
LeerArbReceta CStr(coddi2), auxtec, 2, 2, 1, 1, 1, 2
indtec = vg_codigo
If Val(indtec) > 9 Then GrabarRecetaTecfood = True: MsgBox "Tecfood considera un maximo de 9 niveles ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & indtec
RS1.Open "select car_nombre from a_recetacatdie where car_codigo=" & Val(coddi2) & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe subcategoria dietetica ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
If vg_codigo = "A" Then
   vg_dbtec.Execute "insert into prato (cdfilauxprat, cdprato, nmprato, cdnvprato, usr_cdpratinte) " & _
                    "values ('P', '" & auxtec & "', '" & UCase(Trim(Mid(RS1!car_nombre, 1, 40))) & "', '2', '" & coddi2 & "')"
Else
   vg_dbtec.Execute "update prato set nmprato='" & UCase(Trim(Mid(RS1!car_nombre, 1, 40))) & "' where cdprato='" & auxtec & "' and cdnvprato='1' and usr_cdpratinte='" & coddi2 & "'"
End If
RS1.Close: Set RS1 = Nothing
'----- Fin arbol de receta nivel 2

'----- Arbol de receta nivel 3
vg_codigo = "": vg_opcion = 0
LeerArbReceta CStr(codti1), auxtec, 3, 3, 2, 1, 2, 2
indtec = vg_codigo
If Val(indtec) > 99 Then GrabarRecetaTecfood = True: MsgBox "Tecfood considera un maximo de 99 niveles ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & fg_pone_cero(indtec, 2)
RS1.Open "select tip_nombre from a_recetatippla where tip_codigo=" & Val(codti1) & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe categoria plato ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into prato (cdfilauxprat, cdprato, nmprato, cdnvprato, usr_cdpratinte) " & _
                    "values ('P', '" & auxtec & "', '" & UCase(Trim(Mid(RS1!tip_nombre, 1, 40))) & "', '3', '" & Trim(codti1) & "')"
Else
   vg_dbtec.Execute "update prato set nmprato='" & UCase(Trim(Mid(RS1!tip_nombre, 1, 40))) & "' where cdprato='" & auxtec & "' and cdnvprato='3' and usr_cdpratinte='" & codti1 & "'"
End If
RS1.Close: Set RS1 = Nothing
'------- Fin arbol de receta nivel 3

'------- Arbol de receta nivel 4
vg_codigo = "": vg_opcion = 0
LeerArbReceta codti2, auxtec, 4, 5, 2, 1, 4, 2
indtec = vg_codigo
If Val(indtec) > 99 Then GrabarRecetaTecfood = True: MsgBox "Tecfood considera un maximo de 99 niveles ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & fg_pone_cero(indtec, 2)
RS1.Open "select tip_nombre from a_recetatippla where tip_codigo=" & Val(codti2) & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe categoria plato ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into prato (cdfilauxprat, cdprato, nmprato, cdnvprato, usr_cdpratinte) " & _
                    "values ('P', '" & auxtec & "', '" & UCase(Trim(Mid(RS1!tip_nombre, 1, 40))) & "', '4', '" & codti2 & "')"
Else
   vg_dbtec.Execute "update prato set nmprato='" & UCase(Trim(Mid(RS1!tip_nombre, 1, 40))) & "' where cdprato='" & auxtec & "' and cdnvprato='4' and usr_cdpratinte='" & codti2 & "'"
End If
RS1.Close: Set RS1 = Nothing
'------- Fin arbol de receta nivel 4

'------- Arbol de receta nivel 5
vg_codigo = "": vg_opcion = 0
LeerArbReceta IIf(codti3 > 0, codti3, codti2), auxtec, 5, 7, 1, 1, 6, 2
indtec = vg_codigo
If Val(indtec) > 9 Then GrabarRecetaTecfood = True: MsgBox "Tecfood considera un maximo de 9 niveles ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & indtec
RS1.Open "select tip_nombre from a_recetatippla where tip_codigo=" & IIf(Val(codti3) > 0, Val(codti3), Val(codti2)) & "", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe categoria plato ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into prato (cdfilauxprat, cdprato, nmprato, cdnvprato, usr_cdpratinte) " & _
                    "values ('P', '" & auxtec & "', '" & UCase(Trim(Mid(RS1!tip_nombre, 1, 40))) & "', '5', '" & IIf(Val(codti3) > 0, codti3, codti2) & "')"
Else
   vg_dbtec.Execute "update prato set nmprato='" & UCase(Trim(Mid(RS1!tip_nombre, 1, 40))) & "' where cdprato='" & auxtec & "' and cdnvprato='5' and usr_cdpratinte='" & IIf(Val(codti3) > 0, codti3, codti2) & "'"
End If
RS1.Close: Set RS1 = Nothing
'------- Fin arbol de receta nivel 5

'------- Arbol de receta nivel 6
vg_codigo = "": vg_opcion = 0
Call LeerArbReceta(CodRec, auxtec, 6, 8, 3, 1, 7, 2)
indtec = vg_codigo
If Val(indtec) > 999 Then GrabarRecetaTecfood = True: MsgBox "Tecfood considera un maximo de 999 recetas ...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
auxtec = auxtec & fg_pone_cero(indtec, 3)
If vg_opcion = 1 Then
   vg_dbtec.Execute "insert into prato (cdfilauxprat, cdprato, nmprato, cdnvprato, sgunidade, usr_cdpratinte) " & _
                    "values ('P', '" & auxtec & "', '" & UCase(Trim(Mid(nomrec, 1, 40))) & "', '6', 'KG', '" & CodRec & "')"
   '------- Grabar nombre fantasia
   vg_dbtec.Execute "insert into pratpadr (cdfilauxprat, cdprato, nmfantasiapr) values ('P', '" & auxtec & "', '" & Trim(Mid(nomfan, 1, 80)) & "')"
Else
   vg_dbtec.Execute "update prato set nmprato='" & UCase(Trim(Mid(nomrec, 1, 40))) & "' where cdprato='" & auxtec & "' and cdnvprato='6' and usr_cdpratinte='" & CodRec & "'"
   '------- Grabar nombre fantasia
   RS1.Open "select cdprato from pratpadr where cdprato='" & auxtec & "'", vg_dbtec, adOpenStatic
   If RS1.EOF Then
      vg_dbtec.Execute "insert into pratpadr (cdfilauxprat, cdprato, nmfantasiapr) values ('P', '" & auxtec & "', '" & Trim(Mid(nomfan, 1, 80)) & "')"
   Else
      vg_dbtec.Execute "update pratpadr set nmfantasiapr='" & Trim(Mid(nomfan, 1, 80)) & "' where cdprato='" & auxtec & "'"
   End If
   RS1.Close: Set RS1 = Nothing
End If
'------- Fin arbol de receta nivel 6

Dim CodIng As String, codpro As String, codume As String
Dim caning  As Double, pctapr As Double, pctcoc As Double, pctnut As Double, canser As Double, cannet As Double, valuni As Double
'------- Detalle receta tecfood

If op = "0" Then
   vg_dbtec.Execute "delete from recepadr where cdprato='" & auxtec & "' and cdfilauxprat='P'"
   For i = 1 To Form.vaSpread1(1).MaxRows
       Form.vaSpread1(1).Row = i
       Form.vaSpread1(1).Col = 1
       If Form.vaSpread1(1).text <> "" Then
          CodIng = 0: caning = 0: pctapr = 0: pctcoc = 0: pctnut = 0: canser = 0: cannet = 0: valuni = 0
          Form.vaSpread1(1).Col = 1: CodIng = Form.vaSpread1(1).text
          Form.vaSpread1(1).Col = 3: caning = Form.vaSpread1(1).text
          Form.vaSpread1(1).Col = 5: pctapr = Form.vaSpread1(1).text
          Form.vaSpread1(1).Col = 6: pctcoc = Form.vaSpread1(1).text
          Form.vaSpread1(1).Col = 7: canser = Form.vaSpread1(1).text
          Form.vaSpread1(1).Col = 8: pctnut = Form.vaSpread1(1).text
          Form.vaSpread1(1).Col = 9: cannet = Form.vaSpread1(1).text
          
          RS1.Open "select cdproduto, cdarvprod, sgunidade from produto where cdprodinte='" & "I" & CodIng & "' and cdnvproduto='5'", vg_dbtec, adOpenStatic
          If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe producto generico en tecfood, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
          codpro = "": codpro = RS1!cdarvprod: codume = "": codume = Trim(RS1!sgunidade)
          RS1.Close: Set RS1 = Nothing
          If codume <> "KG" And codume <> "LT" Then
             RS1.Open "select distinct a_unidad.uni_codunm, a_unidad.uni_nomcor, a_unidad.uni_valuni, a_unidadmed.unm_codigo, b_ingrediente.ing_codigo, b_ingrediente.ing_unimed, b_productos.pro_facing " & _
                      "from (a_unidad inner join b_productos on a_unidad.uni_codigo = b_productos.pro_coduni) inner join ((b_ingrediente inner join a_unidadmed on b_ingrediente.ing_unimed = a_unidadmed.unm_codigo) inner join b_productosing on b_ingrediente.ing_codigo = b_productosing.pri_coding) on b_productos.pro_codigo = b_productosing.pri_codpro " & _
                      "where b_ingrediente.ing_codigo='" & CodIng & "'", vg_db, adOpenStatic
             If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe relación codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
             valuni = RS1!uni_valuni
             codume = UCase(Trim(RS1!uni_nomcor))
             If RS1!uni_codunm <> RS1!unm_codigo Then
                RS2.Open "select * from a_unidad where uni_codunm=" & RS1!ing_unimed & "", vg_db, adOpenStatic
                If RS2.EOF Then RS2.Close: Set RS2 = Nothing: RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
                valuni = RS2!uni_valuni
                codume = UCase(Trim(RS2!uni_nomcor))
                RS2.Close: Set RS2 = Nothing
                RS2.Open "select * from convproduto where cdproduto='" & codpro & "' and sgunidade='" & codume & "'", vg_dbtec, adOpenStatic
                If RS2.EOF Then
                   vg_dbtec.Execute "insert into convproduto (sgunidade, cdproduto, vrfatoconves) values ('" & codume & "', '" & codpro & "', " & (RS1!uni_valuni / RS1!pro_facing) & ")"
                Else
                   codume = Trim(RS2!sgunidade)
                End If
                RS2.Close: Set RS2 = Nothing
             End If
             RS1.Close: Set RS1 = Nothing
          Else
             RS1.Open "select a.* from a_unidad a, b_ingrediente b where a.uni_codunm=b.ing_unimed and b.ing_codigo='" & CodIng & "'", vg_db, adOpenStatic
             If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
             valuni = RS1!uni_valuni
             codume = UCase(Trim(RS1!uni_nomcor))
             RS1.Close: Set RS1 = Nothing
          End If
          RS1.Open "select b.* from produto a, recepadr b where b.cdprato='" & auxtec & "' and b.cdproduto=a.cdarvprod and a.cdprodinte='" & "I" & CodIng & "' and b.cdfilauxprat='P'", vg_dbtec, adOpenStatic
          If RS1.EOF Then
             vg_dbtec.Execute "insert into recepadr (cdfilauxprat, cdprato, cdproduto, qtprodprat, vrpefatcocpr, vrpefatcorpr, sgunidade, qtunidestpr, qtbrutapr, idprodbasepr, vrpefatnutpr, qtcalcnutrpr, qtbrutaprunr) " & _
                              "values ('P', '" & auxtec & "', '" & codpro & "', " & (canser / valuni) & ", " & (pctcoc - 100) & ", " & (100 - pctapr) & ", '" & Trim(codume) & "', " & (caning / valuni) & ", " & (caning / valuni) & ", 'N', " & pctnut & ", " & (cannet / valuni) & ", " & (caning / valuni) & ")"
          Else
             caning = ((RS1!qtbrutapr * valuni) + caning)
             canser = (((pctapr / 100) * ((RS1!qtbrutapr * valuni) + caning)) * (pctcoc / 100))
             cannet = ((pctnut / 100) * ((RS1!qtbrutapr * valuni) + caning))
             vg_dbtec.Execute "update recepadr set qtprodprat=" & (canser / valuni) & ", qtunidestpr=" & (caning / valuni) & ", qtbrutapr=" & (caning / valuni) & ", qtcalcnutrpr=" & (cannet / valuni) & ", qtbrutaprunr=" & (caning / valuni) & " where cdprato='" & auxtec & "' and cdproduto='" & RS1!cdproduto & "' and cdfilauxprat='P'"
          End If
          RS1.Close: Set RS1 = Nothing
       End If
   Next i
ElseIf op = "1" Then
   vg_dbtec.Execute "delete from recepadr where cdprato='" & auxtec & "' and cdfilauxprat='P'"
   RS1.Open "select * from b_recetadet where red_codigo=" & vg_codreceta & "", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         CodIng = 0: caning = 0: pctapr = 0: pctcoc = 0: pctnut = 0: canser = 0: cannet = 0: valuni = 0
         CodIng = RS1!red_codpro
         caning = RS1!red_canpro
         pctapr = RS1!red_pctapr
         pctcoc = RS1!red_pctcoc
         canser = (((RS1!red_pctapr / 100) * RS1!red_canpro) * (RS1!red_pctcoc / 100))
         cannet = ((RS1!red_pctnut / 100) * (RS1!red_canpro))
         cannet = RS1!red_pctnut
         RS2.Open "select cdproduto, sgunidade, cdarvprod from produto where cdprodinte='" & "I" & CodIng & "' and cdnvproduto='5'", vg_dbtec, adOpenStatic
         If RS2.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe producto generico en tecfood...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
         codpro = "": codpro = RS2!cdarvprod: codume = "": codume = Trim(RS2!sgunidade)
         RS2.Close: Set RS2 = Nothing
         If codume <> "KG" And codume <> "LT" Then
            RS2.Open "select distinct a_unidad.uni_codunm, a_unidad.uni_nomcor, a_unidad.uni_valuni, a_unidadmed.unm_codigo, b_ingrediente.ing_codigo, b_ingrediente.ing_unimed, b_productos.pro_facing " & _
                     "from (a_unidad inner join b_productos on a_unidad.uni_codigo = b_productos.pro_coduni) inner join ((b_ingrediente inner join a_unidadmed on b_ingrediente.ing_unimed = a_unidadmed.unm_codigo) inner join b_productosing on b_ingrediente.ing_codigo = b_productosing.pri_coding) on b_productos.pro_codigo = b_productosing.pri_codpro " & _
                     "where b_ingrediente.ing_codigo='" & CodIng & "'", vg_db, adOpenStatic
            If RS2.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe relación codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
            valuni = RS2!uni_valuni
            codume = UCase(Trim(RS2!uni_nomcor))
            If RS2!uni_codunm <> RS2!unm_codigo Then
               RS3.Open "select * from a_unidad where uni_codunm=" & RS2!ing_unimed & "", vg_db, adOpenStatic
               If RS3.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: RS3.Close: Set RS3 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
               valuni = RS3!uni_valuni
               codume = UCase(Trim(RS3!uni_nomcor))
               RS3.Close: Set RS3 = Nothing
               RS3.Open "select * from convproduto where cdproduto='" & codpro & "'  and sgunidade='" & codume & "'", vg_dbtec, adOpenStatic
               If RS3.EOF Then
                  vg_dbtec.Execute "insert into convproduto (sgunidade, cdproduto, vrfatoconves) values ('" & codume & "', '" & codpro & "', " & (valuni / RS2!pro_facing) & ")"
               Else
                  codume = Trim(RS3!sgunidade)
               End If
               RS3.Close: Set RS3 = Nothing
            End If
            RS2.Close: Set RS2 = Nothing
         Else
            RS2.Open "select a.* from a_unidad a, b_ingrediente b where a.uni_codunm=b.ing_unimed and b.ing_codigo='" & CodIng & "'", vg_db, adOpenStatic
            If RS2.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
            valuni = RS2!uni_valuni
            codume = UCase(Trim(RS2!uni_nomcor))
            RS2.Close: Set RS2 = Nothing
         End If
         RS2.Open "select b.* from produto a, recepadr b where b.cdprato='" & auxtec & "' and b.cdproduto=a.cdarvprod and a.cdprodinte='" & "I" & CodIng & "' and b.cdfilauxprat='P'", vg_dbtec, adOpenStatic
         If RS2.EOF Then
            vg_dbtec.Execute "insert into recepadr (cdfilauxprat, cdprato, cdproduto, qtprodprat, vrpefatcocpr, vrpefatcorpr, sgunidade, qtunidestpr, qtbrutapr, idprodbasepr, vrpefatnutpr, qtcalcnutrpr, qtbrutaprunr) " & _
                             "values ('P', '" & auxtec & "', '" & codpro & "', " & (canser / valuni) & ", " & (pctcoc - 100) & ", " & (100 - pctapr) & ", '" & Trim(codume) & "', " & (caning / valuni) & ", " & (caning / valuni) & ", 'N', " & pctnut & ", " & (cannet / valuni) & ", " & (caning / valuni) & ")"
         Else
            caning = ((RS2!qtbrutapr * valuni) + caning)
            canser = (((pctapr / 100) * ((RS2!qtbrutapr * valuni) + caning)) * (pctcoc / 100))
            cannet = ((pctnut / 100) * ((RS2!qtbrutapr * valuni) + caning))
            vg_dbtec.Execute "update recepadr set qtprodprat=" & (canser / valuni) & ", qtunidestpr=" & (caning / valuni) & ", qtbrutapr=" & (caning / valuni) & ", qtcalcnutrpr=" & (cannet / valuni) & ", qtbrutaprunr=" & (caning / valuni) & " where cdprato='" & auxtec & "' and cdproduto='" & RS2!cdproduto & "' and cdfilauxprat='P'"
         End If
         RS2.Close: Set RS2 = Nothing
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
ElseIf op = "2" Then
    For i = 1 To Form.vaSpread1.MaxRows
        Form.vaSpread1.Row = i
        Form.vaSpread1.Col = 1
        CodIng = 0: caning = 0: pctapr = 0: pctcoc = 0: pctnut = 0: canser = 0: cannet = 0
        If Form.vaSpread1.text = "1" Then
           Form.vaSpread1.Col = 2: CodIng = Form.vaSpread1.text
           Form.vaSpread1.Col = 4: caning = Form.vaSpread1.text
           Form.vaSpread1.Col = 5: pctapr = Form.vaSpread1.text
           Form.vaSpread1.Col = 6: pctcoc = Form.vaSpread1.text
           Form.vaSpread1.Col = 7: canser = Form.vaSpread1.text
           Form.vaSpread1.Col = 8: pctnut = Form.vaSpread1.text
           Form.vaSpread1.Col = 9: cannet = Form.vaSpread1.text
           RS1.Open "select cdproduto, sgunidade, cdarvprod from produto where cdprodinte='" & "I" & CodIng & "' and cdnvproduto='5'", vg_dbtec, adOpenStatic
           If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe producto generico en tecfood...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
           codpro = "": codpro = RS1!cdarvprod: codume = "": codume = Trim(RS1!sgunidade)
           RS1.Close: Set RS1 = Nothing
           If codume <> "KG" And codume <> "LT" Then
              RS1.Open "select distinct a_unidad.uni_codunm, a_unidad.uni_nomcor, a_unidad.uni_valuni, a_unidadmed.unm_codigo, b_ingrediente.ing_codigo, b_ingrediente.ing_unimed, b_productos.pro_facing " & _
                       "from (a_unidad inner join b_productos on a_unidad.uni_codigo = b_productos.pro_coduni) inner join ((b_ingrediente inner join a_unidadmed on b_ingrediente.ing_unimed = a_unidadmed.unm_codigo) inner join b_productosing on b_ingrediente.ing_codigo = b_productosing.pri_coding) on b_productos.pro_codigo = b_productosing.pri_codpro " & _
                       "where b_ingrediente.ing_codigo='" & CodIng & "'", vg_db, adOpenStatic
              If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe relación codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
              valuni = RS1!uni_valuni
              codume = UCase(Trim(RS1!uni_nomcor))
              If RS1!uni_codunm <> RS1!unm_codigo Then
                 RS2.Open "select * from a_unidad where uni_codunm=" & RS1!ing_unimed & "", vg_db, adOpenStatic
                 If RS2.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
                 valuni = RS2!uni_valuni
                 codume = UCase(Trim(RS2!uni_nomcor))
                 RS2.Close: Set RS2 = Nothing
                 RS2.Open "select * from convproduto where cdproduto='" & codpro & "'  and sgunidade='" & codume & "'", vg_dbtec, adOpenStatic
                 If RS2.EOF Then
                    vg_dbtec.Execute "insert into convproduto (sgunidade, cdproduto, vrfatoconves) values ('" & codume & "', '" & codpro & "', " & (valuni / RS1!pro_facing) & ")"
                 Else
                    codume = Trim(RS2!sgunidade)
                 End If
                 RS2.Close: Set RS2 = Nothing
              End If
              RS1.Close: Set RS1 = Nothing
           Else
              RS1.Open "select a.* from a_unidad a, b_ingrediente b where a.uni_codunm=b.ing_unimed and b.ing_codigo='" & CodIng & "'", vg_db, adOpenStatic
              If RS1.EOF Then RS1.Close: Set RS1 = Nothing: GrabarRecetaTecfood = True: MsgBox "No existe codigo unidad medida, proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo: Exit Function
              valuni = RS1!uni_valuni
              codume = UCase(Trim(RS1!uni_nomcor))
              RS1.Close: Set RS1 = Nothing
           End If
           RS1.Open "select b.* from produto a, recepadr b where b.cdprato='" & auxtec & "' and b.cdproduto=a.cdarvprod and a.cdprodinte='" & "I" & CodIng & "' and b.cdfilauxprat='P'", vg_dbtec, adOpenStatic
           If RS1.EOF Then
              vg_dbtec.Execute "insert into recepadr (cdfilauxprat, cdprato, cdproduto, qtprodprat, vrpefatcocpr, vrpefatcorpr, sgunidade, qtunidestpr, qtbrutapr, idprodbasepr, vrpefatnutpr, qtcalcnutrpr, qtbrutaprunr) " & _
                               "values ('P', '" & auxtec & "', '" & codpro & "', " & (canser / valuni) & ", " & (pctcoc - 100) & ", " & (100 - pctapr) & ", '" & Trim(codume) & "', " & (caning / valuni) & ", " & (caning / valuni) & ", 'N', " & pctnut & ", " & (cannet / valuni) & ", " & (caning / valuni) & ")"
           Else
              caning = ((RS1!qtbrutapr * valuni) + caning)
              canser = (((pctapr / 100) * ((RS1!qtbrutapr * valuni) + caning)) * (pctcoc / 100))
              cannet = ((pctnut / 100) * ((RS1!qtbrutapr * valuni) + caning))
              vg_dbtec.Execute "update recepadr set qtprodprat=" & (canser / valuni) & ", qtunidestpr=" & (caning / valuni) & ", qtbrutapr=" & (caning / valuni) & ", qtcalcnutrpr=" & (cannet / valuni) & ", qtbrutaprunr=" & (caning / valuni) & " where cdprato='" & auxtec & "' and cdproduto='" & RS1!cdproduto & "' and cdfilauxprat='P'"
           End If
           RS1.Close: Set RS1 = Nothing
        End If
    Next i
End If
'------- Fin detalle receta tecfood
End Function

Function fg_CanServir(codigo As Integer) As Double
Dim RS1 As New ADODB.Recordset
    Set RS1 = vg_db.Execute(" select sum(((red_pctapr/100) * red_canpro * (red_pctcoc/100))) as Canservir from b_recetadet where red_codigo= " & codigo & "") ', vg_db, adOpenForwardOnly ', adOpenStatic
    If Not RS1.EOF Then
      fg_CanServir = RS1!Canservir
    Else
      fg_CanServir = 0
    End If
RS1.Close: Set RS1 = Nothing
End Function


Function EspFecha(Fecha As fpDateTime)
'Fecha.DateTimeFormat = 5
'If Bandera = 1 Then
'   Fecha.UserDefinedFormat = "dd/mm/yyyy"
'Else
'   Fecha.UserDefinedFormat = "mmmm"
'End If
'Fecha.CalFirstDay (1)

'Fecha.ShortDayName(1) = "Dom"
'Fecha.ShortDayName(2) = "Lun"
'Fecha.ShortDayName(3) = "Mar"
'Fecha.ShortDayName(4) = "Mie"
'Fecha.ShortDayName(5) = "Jue"
'Fecha.ShortDayName(6) = "Vie"
'Fecha.ShortDayName(7) = "Sab"
'Fecha.LongDayName(1) = "Domingo"
'Fecha.LongDayName(2) = "Lunes"
'Fecha.LongDayName(3) = "Martes"
'Fecha.LongDayName(4) = "Miercoles"
'Fecha.LongDayName(5) = "Jueves"
'Fecha.LongDayName(6) = "Viernes"
'Fecha.LongDayName(7) = "Sabado"
'Fecha.LongMonthName(1) = "Enero"
'Fecha.LongMonthName(2) = "Febrero"
'Fecha.LongMonthName(3) = "Marzo"
'Fecha.LongMonthName(4) = "Abril"
'Fecha.LongMonthName(5) = "Mayo"
'Fecha.LongMonthName(6) = "Junio"
'Fecha.LongMonthName(7) = "Julio"
'Fecha.LongMonthName(8) = "Agosto"
'Fecha.LongMonthName(9) = "Septiembre"
'Fecha.LongMonthName(10) = "Octubre"
'Fecha.LongMonthName(11) = "Noviembre"
'Fecha.LongMonthName(12) = "Diciembre"
End Function

Function GenerarBaseEnviado(cDBI As String, aAp1 As String, aAp2 As String, dBo As String, op As Integer, Fecha As Long, subseg As String, codReg As String, CCeco As String, ProdXML As String, RecetaXML As String, SubSegXML As String, RegXML As String, CecoXML As String)

Dim RS As New ADODB.Recordset
Dim nArch As String
Dim pArch As String
Dim db7 As Database

nArch = "TRFSGP" & fg_Dtos(Date) & Trim(fg_Quitachar(Time, ":")) & ".TXT"
pArch = dir_trabajo & nArch
Close #1
Open pArch For Output As #1

Print #1, "CC"

'-------> cuenta contable optimun
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbCuentas_Sap_AX")
Print #1, "CREATE TABLE Cuentas_Sap_AX (Cuentas_Sap char(50), Cuentas_AX char(50))"
Print #1, "DELETE FROM Cuentas_Sap_AX"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO Cuentas_Sap_AX (Cuentas_Sap, Cuentas_AX) " & _
             "VALUES ('" & RS!Cuentas_Sap & "', '" & RS!cuentas_AX & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Ceco SAP AX
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbCeco_Sap_AX")
Print #1, "CREATE TABLE Cecos_Sap_AX (Cecos_Sap char(10), Cecos_AX char(10), Sociedad_Sap char(10))"
Print #1, "DELETE FROM Cecos_Sap_AX"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO Cecos_Sap_AX (Cecos_Sap, Cecos_AX, Sociedad_Sap) " & _
             "VALUES ('" & RS!Cecos_Sap & "', '" & RS!Cecos_AX & "', '" & RS!Sociedad_Sap; "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Sociedad SAP AX
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbSociedad_Sap_AX")
Print #1, "CREATE TABLE Sociedad_Sap_AX (Sociedad_Sap char(50), Sociedad_AX char(50))"
Print #1, "DELETE FROM Sociedad_Sap_AX"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO Sociedad_Sap_AX (Sociedad_Sap, Sociedad_AX) " & _
             "VALUES ('" & RS!Sociedad_Sap & "', '" & RS!Sociedad_AX & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> b_sac_listaprecio
Print #1, "CREATE TABLE b_sac_listaprecio (lps_cencos char(10), lps_periodo char(6), lps_codsac char(20), lps_precio double)"
Print #1, "DELETE FROM b_sac_listaprecio"

'-------> b_formatocompras
Print #1, "CREATE TABLE b_formatocompras (foc_codsac char(20), foc_codcat int, foc_nomsac char(100), foc_unisac char(20), foc_vigini datetime, foc_flexec int, foc_vigfin datetime, foc_faccon double)"
Print #1, "DELETE b_formatocompras FROM b_formatocompras"

'-------> b_formatocomprassgp
Print #1, "CREATE TABLE b_formatocomprassgp (fcs_codsac char(20), fcs_codsgp char(20), fcs_sgppre int)"
Print #1, "DELETE FROM b_formatocomprassgp"

'-------> a_tipodocumento
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbTipoDocumento")
'Print #1, "CREATE TABLE a_tipodocumento (tdo_codigo char(2), tdo_nombre char(50), tdo_cladoc char(2), tdo_orden int, tdo_IdCodigo char(2), tdo_VisualizaDoc char(1))"
Print #1, "CREATE TABLE a_tipodocumento (tdo_codigo char(2), tdo_nombre char(50), tdo_cladoc char(2), tdo_orden int)"
Print #1, "DELETE FROM a_tipodocumento"
Do While Not RS.EOF
   
   DoEvents
'   Print #1, "INSERT INTO a_tipodocumento (tdo_codigo, tdo_nombre, tdo_cladoc, tdo_orden, tdo_IdCodigo, tdo_VisualizaDoc) " & _
'             "VALUES ('" & RS!tdo_codigo & "', '" & RS!tdo_nombre & "', '" & RS!tdo_cladoc & "', " & RS!tdo_orden & ", " & RS!tdo_IdCodigo & ", '" & IIf(IsNull(RS!tdo_VisualizaDoc) Or Not RS!tdo_VisualizaDoc, "0", "1") & "')"
   
   Print #1, "INSERT INTO a_tipodocumento (tdo_codigo, tdo_nombre, tdo_cladoc, tdo_orden) " & _
             "VALUES ('" & RS!tdo_codigo & "', '" & RS!tdo_nombre & "', '" & RS!tdo_cladoc & "', " & RS!tdo_orden & ")"
   
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> a_clasedocsap
Print #1, "CREATE TABLE a_clasedocsap (cds_coddoc char(2), cds_codreg int, cds_cdosap char(2))"
Print #1, "DELETE FROM a_clasedocsap"

'-------> a_cencos
Print #1, "CREATE TABLE a_cencos (cen_codigo char(10), cen_socsap char(4), cen_envsap char(1), cen_sobrec char(1), cen_codmun int, cen_ccisac int, cen_cecsac char(4), cen_codreg int, cen_codopt char(10))"
Print #1, "DELETE FROM a_cencos"

'-------> Generar familia productos (a_tipopro)
Print #1, "CREATE TABLE a_tipopro (tip_codigo int, tip_nombre char(35), tip_previo int)"
Print #1, "DELETE FROM a_tipopro"

'-------> Unidad Producto (a_unidad)
Print #1, "CREATE TABLE a_unidad (uni_codigo int, uni_nombre char(10), uni_nomcor char(5))"
Print #1, "DELETE FROM a_unidad"

'-------> Generar embalaje productos
Print #1, "CREATE TABLE a_embalaje (emb_codigo int, emb_nombre char(20), emb_nomcor char(5))"
Print #1, "DELETE FROM a_embalaje"

'-------> Generar cuentas contables productos
Print #1, "CREATE TABLE a_ctacontable (cta_codigo char(10), cta_nombre char(40))"
Print #1, "DELETE FROM a_ctacontable"

'-------> Generar parametro sistema
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbParamI")
Print #1, "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255))"
Print #1, "DELETE FROM a_param"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) " & _
             "VALUES ('" & RS!par_codigo & "', '" & RS!par_nombre & "', '" & RS!par_tipo & "', '" & RS!par_valor & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbParamII")
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) " & _
             "VALUES ('" & RS!par_codigo & "', '" & RS!par_nombre & "', '" & RS!par_tipo & "', '" & RS!par_valor & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Generar impuesto productos
Print #1, "CREATE TABLE a_impuesto (imp_codigo int, imp_nombre char(15), imp_pctimp double, imp_inccos int, imp_codsap char(20), imp_indmod char(1), imp_adicional char(1), imp_cimsap1 char(2), imp_cimsap2 char(2), imp_cimsap3 char(2), imp_cimsap4 char(4))"
Print #1, "DELETE FROM a_impuesto"

'-------> Generar unidad medida ingrediente
Print #1, "CREATE TABLE a_unidadmed (unm_codigo int, unm_nombre char(10), unm_nomcor char(5))"
Print #1, "DELETE FROM a_unidadmed"

'-------> Generar nutriente aporte
Print #1, "CREATE TABLE a_nutriente (nut_codigo int, nut_nombre char(30), nut_nomuni char(5), nut_indpri int, nut_secnro int)"
Print #1, "DELETE FROM a_nutriente"

'-------> Generar productos
Print #1, "CREATE TABLE b_productos (pro_codigo char(20), pro_codbar char(20), pro_codcom char(20), pro_codtip int, pro_nombre char(50), pro_coduni int, pro_facing double, pro_facsto double, pro_codemb int, pro_uniemb double, pro_upreco double, pro_fecuco datetime, pro_propon double, pro_ctacon char(10), pro_fecven int, pro_ctrsto int, pro_maepro int, pro_codref int, pro_codrei int, pro_cuohor char(1), pro_tipord char(1), pro_tippro char(1))"
Print #1, "DELETE FROM b_productos"

'-------> Generar productos impuestos
Print #1, "CREATE TABLE b_productosimp (ipr_codpro char(20), ipr_codimp int)"
Print #1, "DELETE FROM b_productosimp"

'-------> Generar productos ingredientes
Print #1, "CREATE TABLE b_productosing (pri_codpro char(20), pri_coding char(20))"
Print #1, "DELETE FROM b_productosing"

'-------> Generar ingredientes
Print #1, "CREATE TABLE b_ingrediente (ing_codigo char(20), ing_nombre char(50), ing_nomfan char(50), ing_unimed int, ing_pctapr double, ing_pctcoc double, ing_pctnut double, ing_facnut double, ing_indpav int, ing_indgrv int, ing_precos double, ing_feccos int, ing_codcom char(20), ing_codped char(20))"
Print #1, "DELETE FROM b_ingrediente"

'-------> Generar nutriente del ingrediente
Print #1, "CREATE TABLE b_productonut (pnu_codpro char(20), pnu_codapo int, pnu_canapo double)"
Print #1, "DELETE FROM b_productonut"

'-------> Generar proveedores
Print #1, "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo date, prv_origen char(1), prv_regimp char(1), prv_autret char(1), prv_cuohor char(1), prv_codmun char(1), prv_docele char(1))"
Print #1, "DELETE FROM b_proveedor"

'-------> Generar tipo servicio
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbTipoServicio")
Print #1, "CREATE TABLE a_tiposervicio (tis_codigo int, tis_nombre char(50))"
Print #1, "DELETE FROM a_tiposervicio"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_tiposervicio (tis_codigo, tis_nombre) " & _
             "VALUES (" & RS!tis_codigo & ", '" & RS!tis_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Generar segmento
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbSgemento")
Print #1, "CREATE TABLE a_segmento (seg_codigo int, seg_nombre char(50))"
Print #1, "DELETE FROM a_segmento"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_segmento (seg_codigo, seg_nombre) " & _
             "VALUES (" & RS!seg_codigo & ", '" & RS!seg_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Generar envío sap
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbTipoInterfaz")
Print #1, "CREATE TABLE a_tipointerfaz (tii_codigo int, tii_nombre char(50))"
Print #1, "DELETE FROM a_tipointerfaz"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_tipointerfaz (tii_codigo, tii_nombre) " & _
             "VALUES (" & RS!tii_codigo & ", '" & RS!tii_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Generar casino envío sap
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbCasinoInterfaz")
Print #1, "CREATE TABLE b_casinointerfaz (cai_cencos char(10), cai_codtii int)"
Print #1, "DELETE FROM b_casinointerfaz"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO b_casinointerfaz (cai_cencos, cai_codtii) " & _
             "VALUES ('" & RS!cai_cencos & "', " & RS!cai_codtii & ")"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Generar envío tipo actividad
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbTipoActividad")
Print #1, "CREATE TABLE a_tipoactividad (tia_codigo int, tia_nombre char(50))"
Print #1, "DELETE FROM a_tipoactividad"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_tipoactividad (tia_codigo, tia_nombre) " & _
             "VALUES (" & RS!tia_codigo & ", '" & RS!tia_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Generar envío casino tipo actividad
Print #1, "CREATE TABLE b_casinotipoactividades (cta_cencos char(10), cta_tipact int)"
Print #1, "DELETE FROM b_casinotipoactividades"

'-------> Generar envío parametro despacho
Print #1, "CREATE TABLE b_paramdesp (pad_cencos char(10), pad_codtip int, pad_tipo char(3), pad_diaseg int, pad_diario char(7))"
Print #1, "DELETE FROM b_paramdesp"

'-------> Generar envío días inhabiles
Print #1, "CREATE TABLE b_Fecha_Inhabiles (CFI_CeCo char(10), CFI_Fecha datetime, CFI_Glosa char(100))"
Print #1, "DELETE FROM b_Fecha_Inhabiles"

'-------> Generar envío casino paramero stock
Print #1, "CREATE TABLE b_casinoparametrostock (cps_cencos char(10), cps_invsto char(1), cps_reqmen char(1), cps_porinv double, cps_liscri char(1), cps_diario char(1), cps_ajuimp char(1))"
Print #1, "DELETE FROM b_casinoparametrostock"

'-------> Insert envío pais
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbPais")
Print #1, "CREATE TABLE a_pais (pai_codigo char(10), pai_nombre char(100))"
Print #1, "DELETE FROM a_pais"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_pais (pai_codigo, pai_nombre) " & _
             "VALUES ('" & RS!pai_codigo & "', '" & RS!pai_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Insert envío región
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbRegion")
Print #1, "CREATE TABLE a_region (reg_codigo int, reg_nombre char(50))"
Print #1, "DELETE FROM a_region"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_region (reg_codigo, reg_nombre) " & _
             "VALUES (" & RS!Reg_Codigo & ", '" & RS!reg_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Insert envío municipio
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbMunicipio")
Print #1, "CREATE TABLE a_municipio (mun_codigo int, mun_nombre char(50), mun_retobl char(1))"
Print #1, "DELETE FROM a_municipio"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_municipio (mun_codigo, mun_nombre, mun_retobl) " & _
             "VALUES (" & RS!mun_codigo & ", '" & RS!mun_nombre & "', '" & RS!mun_retobl & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Insert envío retención fuentes
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbRetencionFuente")
Print #1, "CREATE TABLE b_retencionfuente (ref_codigo int, ref_nombre char(100), ref_portar double, ref_codcta char(10), ref_tipret char(10), ref_indret char(10))"
Print #1, "DELETE FROM b_retencionfuente"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO b_retencionfuente (ref_codigo, ref_nombre, ref_portar, ref_codcta, ref_tipret, ref_indret) " & _
             "VALUES (" & RS!ref_codigo & ", '" & RS!ref_nombre & "', " & RS!ref_portar & ", '" & RS!ref_codcta & "', '" & RS!ref_tipret & "', '" & RS!ref_indret & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Insert envío retención ica
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbRetencionIca")
Print #1, "CREATE TABLE b_retencionica (rei_codigo int, rei_nombre char(100))"
Print #1, "DELETE FROM b_retencionica"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO b_retencionica (rei_codigo, rei_nombre) " & _
             "VALUES (" & RS!rei_codigo & ", '" & RS!rei_nombre & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Insert envío retención ica
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbDetRetencionIca")
Print #1, "CREATE TABLE b_detretencionica (dri_codigo int, dri_codmun int, dri_portar double, dri_codcta char(10), dri_tipret char(10), dri_indret char(10))"
Print #1, "DELETE FROM b_detretencionica"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO b_detretencionica (dri_codigo, dri_codmun, dri_portar, dri_codcta, dri_tipret, dri_indret) " & _
             "VALUES (" & RS!dri_codigo & ", " & RS!dri_codmun & ", " & RS!dri_portar & ", '" & RS!dri_codcta & "', '" & RS!dri_tipret & "', '" & RS!dri_indret & "')"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Insert envío parametro código barra
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbParCodigoBarra")
Print #1, "CREATE TABLE a_par_codigo_barra (a_par_id_codigo int, atr_codigo_barra int, cli_codigo char(10), cbar_posinicial int, cbar_largo int)"
Print #1, "DELETE FROM a_par_codigo_barra"
Do While Not RS.EOF
   
   DoEvents
   Print #1, "INSERT INTO a_par_codigo_barra (a_par_id_codigo, atr_codigo_barra, cli_codigo, cbar_posinicial, cbar_largo) " & _
             "VALUES (" & RS!a_par_id_codigo & ", " & RS!atr_codigo_barra & ", '" & RS!Cli_codigo & "', " & RS!cbar_posinicial & ", " & RS!cbar_largo & ")"
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

'-------> Creación tabla categoria dietetica
Print #1, "CREATE TABLE a_recetacatdie (car_codigo int, car_nombre char(50), car_previo int)"
Print #1, "DELETE FROM a_recetacatdie"

'-------> Creación tabla
Print #1, "CREATE TABLE a_recetatippla (tip_codigo int, tip_nombre char(50), tip_previo int)"
Print #1, "DELETE FROM a_recetatippla"

'-------> creación tabla receta
Print #1, "CREATE TABLE b_receta (rec_codigo int, rec_catdie int, rec_tippla int, rec_nombre char(80), rec_nomfan char(80), rec_metpre longtext, rec_conche longtext, rec_sugere longtext, rec_basrac int, rec_tiprec char(1), rec_fecvig int, rec_gruvul longtext, Constraint b_receta_pk Primary Key (rec_codigo))"
Print #1, "DELETE FROM b_receta"
'-------> creación tabla receta temporal
If (op = 1 Or op = 2 Or op = 3) Then
    
    Print #1, "CREATE TABLE b_recetaaux (rec_codigo int, rec_catdie int, rec_tippla int, rec_nombre char(80), rec_nomfan char(80), rec_metpre longtext, rec_conche longtext, rec_sugere longtext, rec_basrac int, rec_tiprec char(1), rec_fecvig int, rec_gruvul longtext)"
    Print #1, "DELETE FROM b_recetaaux"

End If

'-------> creación tabla receta detalle
Print #1, "CREATE TABLE b_recetadet (red_codigo int, red_nroite int, red_codpro char(20), red_canpro double, red_cospro double, red_pctapr double, red_pctcoc double, red_pctnut double, red_tiprec int)"
Print #1, "DELETE FROM b_recetadet"
'-------> creación tabla receta detalle aux
If (op = 1 Or op = 2 Or op = 3) Then
   
   Print #1, "CREATE TABLE b_recetadetaux (red_codigo int, red_nroite int, red_codpro char(20), red_canpro double, red_cospro double, red_pctapr double, red_pctcoc double, red_pctnut double, red_tiprec int)"
   Print #1, "DELETE FROM b_recetadetaux"

End If

'-------> creación tabla costo patron
Print #1, "CREATE TABLE b_costopatron (cpa_cencos char(10), cpa_codreg int, cpa_codser int, cpa_anomes int, cpa_descripcion char(10), cpa_valor double)"
Print #1, "DELETE FROM b_costopatron"
   
'-------> creación tabla gramo familia producto
Print #1, "CREATE TABLE b_gramofamproducto (gfp_cencos char(10), gfp_codreg int, gfp_catdie int, gfp_tiprec int, gfp_fampro int, gfp_graini double, gfp_grafin double)"
Print #1, "DELETE FROM b_costopatron"

'-------> creación tabla regimen
Print #1, "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
Print #1, "DELETE FROM a_regimen"

'-------> creación tabla servicio
Print #1, "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(50), ser_orden int, ser_codsap char(20), ser_facturable char(1))"
Print #1, "DELETE FROM a_servicio"

'-------> creación tabla estructura servicio
Print #1, "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double)"
Print #1, "DELETE FROM a_estservicio"

'-------> creación tabla minuta
Print #1, "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int, Constraint b_minuta_pk Primary Key (min_codigo))"
Print #1, "DELETE FROM b_minuta"

'-------> creación tabla minuta detalle
Print #1, "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac int, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer int, mid_rec5eta char(1), mid_cosdes double, Constraint b_minutadet_pk Primary Key (mid_codigo, mid_numlin))"
Print #1, "DELETE FROM b_minutadet"

'-------> creación tabla gramaje
Print #1, "CREATE TABLE b_tablagramaje (tgr_codreg int, tgr_codrec int, tgr_coding char(20), tgr_codzon int, tgr_codins char(20), tgr_cantgr double)"
Print #1, "DELETE FROM b_tablagramaje"

'-------> creación tabla gramaje auxiliares y recetas
If (op = 1 Or op = 2 Or op = 3) Then
    
    Print #1, "CREATE TABLE b_tablagramajeaux (tgr_subseg int, tgr_codreg int, tgr_codrec int, tgr_coding char(20), tgr_codzon int, tgr_codins char(20), tgr_cantgr double, Constraint b_tablagramajeaux Primary Key (tgr_subseg, tgr_codreg, tgr_codrec, tgr_coding, tgr_codzon))"
    Print #1, "DELETE FROM b_tablagramajeaux"
    
    Print #1, "CREATE TABLE b_tablagramajeauxceco (tgc_ceco char(10), tgc_codreg int, tgc_codrec int, tgc_coding char(20), tgc_codins char(20), tgc_cantgr double, Constraint b_tablagramajeauxceco Primary Key (tgc_ceco, tgc_codreg, tgc_codrec, tgc_coding))"
    Print #1, "DELETE FROM b_tablagramajeauxceco"
    
    Print #1, "CREATE TABLE gra_receta (rec_codigo int)"
    Print #1, "DELETE FROM gra_receta"

    Print #1, "CREATE TABLE tmp_receta (rec_codigo int)"
    Print #1, "DELETE FROM tmp_receta"

End If

If (op = 1 Or op = 2 Or op = 3) Then
   
   Print #1, "CREATE TABLE a_subsegmentoaux (sub_codigo int, sub_nombre char(50))"
   Print #1, "DELETE FROM a_subsegmentoaux"

End If


If op <> 3 Then
   
   '-------> Generar familia productos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbFamiliaProducto")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_tipopro (tip_codigo, tip_nombre, tip_previo) " & _
                "VALUES (" & RS!tip_codigo & ", '" & RS!tip_nombre & "', '" & RS!tip_previo & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar unidad medida productos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbUnidadProducto '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_unidad (uni_codigo, uni_nombre, uni_nomcor) " & _
                "VALUES( " & RS!uni_codigo & ", '" & RS!uni_nombre & "', '" & RS!uni_nomcor & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar embalaje productos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbEmbalajeProducto '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_embalaje (emb_codigo, emb_nombre, emb_nomcor) " & _
               "VALUES (" & RS!emb_codigo & ", '" & RS!emb_nombre & "', '" & RS!emb_nomcor & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar cuentas contables productos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbCtaConProducto '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_ctacontable (cta_codigo, cta_nombre) " & _
               "VALUES ('" & RS!cta_codigo & "', '" & RS!cta_nombre & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar parametros cuentas contables
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbParamCtaCon")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) " & _
                "VALUES ('" & RS!par_codigo & "', '" & RS!par_nombre & "', '" & RS!par_tipo & "', '" & RS!par_valor & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar impuesto productos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbImpuesto")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_impuesto (imp_codigo, imp_nombre, imp_pctimp, imp_inccos, imp_codsap, imp_indmod, imp_adicional, imp_cimsap1, imp_cimsap2, imp_cimsap3, imp_cimsap4) " & _
                "VALUES (" & RS!imp_codigo & ", '" & RS!imp_nombre & "', " & RS!imp_pctimp & ", " & RS!imp_inccos & ", '" & RS!imp_codsap & "', '" & RS!imp_indmod & "', '" & RS!imp_adicional & "', '" & RS!imp_cimsap1 & "', '" & RS!imp_cimsap2 & "', '" & RS!imp_cimsap3 & "', '" & RS!imp_cimsap4 & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar unidad medida ingrediente
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
  
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbUnidadMedidaIngrediente '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_unidadmed (unm_codigo, unm_nombre, unm_nomcor) " & _
                "VALUES ('" & RS!unm_codigo & "', '" & RS!unm_nombre & "', '" & RS!unm_nomcor & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar nutriente aporte
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbNutriente '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_nutriente (nut_codigo, nut_nombre, nut_nomuni, nut_indpri, nut_secnro) " & _
                "VALUES (" & RS!nut_codigo & ", '" & RS!nut_nombre & "', '" & RS!nut_nomuni & "', " & RS!nut_indpri & ", " & RS!nut_secnro & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar productos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbProducto '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_productos (pro_codigo, pro_codbar, pro_codcom, pro_codtip, pro_nombre, pro_coduni, pro_facing, pro_facsto, pro_codemb, pro_uniemb, pro_upreco, pro_fecuco, pro_propon, pro_ctacon, pro_fecven, pro_ctrsto, pro_maepro, pro_codref, pro_codrei, pro_cuohor, pro_tipord, pro_tippro) " & _
                "VALUES ('" & RS!pro_codigo & "', '" & RS!pro_codbar & "', '" & RS!pro_codcom & "', " & RS!pro_codtip & ", '" & RS!pro_nombre & "', " & RS!pro_coduni & ", " & RS!pro_facing & ", " & RS!pro_facsto & ", " & RS!pro_codemb & ", " & RS!pro_uniemb & ", " & RS!pro_upreco & ", '" & RS!pro_fecuco & "', " & RS!pro_propon & ", '" & RS!pro_ctacon & "', " & RS!pro_fecven & ", " & RS!pro_ctrsto & ", " & RS!pro_maepro & ", " & RS!pro_codref & ", " & RS!pro_codrei & ", '" & RS!pro_cuohor & "', '" & RS!pro_tipord & "', '" & RS!pro_tippro & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar productos impuestos
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbProductoImp '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_productosimp (ipr_codpro, ipr_codimp) " & _
                "VALUES ('" & RS!ipr_codpro & "', " & RS!ipr_codimp & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar productos ingredientes
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbProductoIng '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_productosing (pri_codpro, pri_coding) " & _
                "VALUES ('" & RS!pri_codpro & "', '" & RS!pri_coding & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar ingredientes
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbIngrediente '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_ingrediente (ing_codigo, ing_nombre, ing_nomfan, ing_unimed, ing_pctapr, ing_pctcoc, ing_pctnut, ing_facnut, ing_indpav, ing_indgrv, ing_precos, ing_feccos, ing_codcom, ing_codped) " & _
                "VALUES ('" & RS!ing_codigo & "', '" & RS!ing_nombre & "', '" & RS!ing_nomfan & "', " & RS!ing_unimed & ", " & RS!ing_pctapr & ", " & RS!ing_pctcoc & ", " & RS!ing_pctnut & ", " & RS!ing_facnut & ", " & RS!ing_indpav & ", " & RS!ing_indgrv & ", " & RS!ing_precos & ", " & RS!ing_feccos & ", '" & RS!ing_codcom & "', '" & RS!ing_codped & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing

   '-------> Generar nutriente del ingrediente
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbIngNut '" & ProdXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_productonut (pnu_codpro, pnu_codapo, pnu_canapo) " & _
                "VALUES ('" & RS!pnu_codpro & "', " & RS!pnu_codapo & ", " & RS!pnu_canapo & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing

   '-------> Generar proveedores
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbProveedor")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_proveedor (prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen, prv_regimp, prv_autret, prv_cuohor, prv_codmun, prv_docele) " & _
                "VALUES ('" & RS!prv_codigo & "', '" & LimpiaDato(RS!prv_nombre) & "', '" & LimpiaDato(RS!prv_direccion) & "', '" & LimpiaDato(RS!prv_comuna) & "', '" & LimpiaDato(RS!prv_ciudad) & "', '" & LimpiaDato(RS!prv_fono1) & "', '" & LimpiaDato(RS!prv_fono2) & "', '" & LimpiaDato(RS!prv_fax) & "', '" & LimpiaDato(RS!prv_percon) & "', '" & LimpiaDato(RS!prv_giro) & "', '" & LimpiaDato(RS!prv_emapro) & "', '" & RS!prv_activo & "', '" & RS!prv_fecumo & "', '" & RS!prv_origen & "', '" & RS!prv_regimp & "', '" & RS!prv_autret & "', '" & RS!prv_cuohor & "', '" & RS!prv_codmun & "', '" & RS!prv_docele & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing


End If


If op = 1 Or op = 2 Then
   
   '-------> Generar subsegmento
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbSubSegmento")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_subsegmentoaux (sub_codigo, sub_nombre) " & _
                "VALUES (" & RS!sub_codigo & ", '" & RS!sub_nombre & "')"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
'   db1.Execute "INSERT INTO a_subsegmentoaux SELECT DISTINCT sub_codigo, sub_nombre FROM a_subsegmento IN " & dBo & ""
   
   '-------> Generar envió categoria dietetica receta
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
  
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbParCatDietetica")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_recetacatdie (car_codigo, car_nombre, car_previo) " & _
                "VALUES (" & RS!car_codigo & ", '" & RS!car_nombre & "', " & RS!car_previo & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar tipo plato
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbTipoPlato")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO a_recetatippla (tip_codigo, tip_nombre, tip_previo) " & _
                "VALUES (" & RS!tip_codigo & ", '" & RS!tip_nombre & "', " & RS!tip_previo & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar encabezado receta
   Dim MetPre As String
   Dim ConChef As String
   Dim Sugere As String
   Dim GruVul As String
   Dim cSpi As Long
   '-------> Buscar spid
   Set RS = vg_db.Execute("SELECT @@spid spid")
   If Not RS.EOF Then cSpi = RS!spid
   RS.Close: Set RS = Nothing

   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbReceta '" & RecetaXML & "', " & cSpi & ", '" & vg_NUsr & "'")
'   glosa = RS.GetString(adClipString, -1, ",", vbCrLf, "")
'   Print #1, glosa
'   Do While Not RS.EOF
'      DoEvents
'      '-------> Mover metodo preparación
'      M_GenRec.RichTextBox1.text = ""
'      M_GenRec.RichTextBox1.SelRTF = RS!rec_metpre
'      MetPre = ""
'      MetPre = M_GenRec.RichTextBox1.text
'      '-------> Mover consejo chef
'      M_GenRec.RichTextBox1.text = ""
'      M_GenRec.RichTextBox1.SelRTF = RS!rec_conche
'      ConChef = ""
'      ConChef = M_GenRec.RichTextBox1.text
'      '-------> Mover Sugerencia
'      M_GenRec.RichTextBox1.text = ""
'      M_GenRec.RichTextBox1.SelRTF = RS!rec_sugere
'      Sugere = ""
'      Sugere = M_GenRec.RichTextBox1.text
'      '-------> Mover Grupo Vulnerable
'      M_GenRec.RichTextBox1.text = ""
'      M_GenRec.RichTextBox1.SelRTF = RS!rec_gruvul
'      GruVul = ""
'      GruVul = M_GenRec.RichTextBox1.text
'
'      Print #1, "INSERT INTO b_recetaaux (rec_codigo, rec_catdie, rec_tippla, rec_nombre, rec_nomfan, rec_metpre, rec_conche, rec_sugere, rec_basrac, rec_tiprec, rec_fecvig, rec_gruvul)  " & _
'                "VALUES (" & RS!rec_codigo & ", " & RS!rec_catdie & ", " & RS!rec_tippla & ", '" & LimpiaDato(RS!rec_nombre) & "', '" & LimpiaDato(RS!rec_nomfan) & "', '" & MetPre & "', '" & ConChef & "', '" & Sugere & "', " & RS!rec_basrac & ", '" & RS!rec_tiprec & "', " & RS!rec_fecvig & ", '" & GruVul & "')"
'      Print #1, "INSERT INTO b_receta (rec_codigo, rec_catdie, rec_tippla, rec_nombre, rec_nomfan, rec_metpre, rec_conche, rec_sugere, rec_basrac, rec_tiprec, rec_fecvig, rec_gruvul) " & _
'                "VALUES (" & RS!rec_codigo & ", " & RS!rec_catdie & ", " & RS!rec_tippla & ", '" & LimpiaDato(RS!rec_nombre) & "', '" & LimpiaDato(RS!rec_nomfan) & "', '" & MetPre & "', '" & ConChef & "', '" & Sugere & "', " & RS!rec_basrac & ", '" & RS!rec_tiprec & "', " & RS!rec_fecvig & ", '" & GruVul & "')"
'      RS.MoveNext
'   Loop
   RS.Close: Set RS = Nothing
   
   '-------> Generar encabezado receta
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbRecetaDet '" & RecetaXML & "'")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_recetadetaux (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) " & _
                "VALUES (" & RS!red_codigo & ", " & RS!red_nroite & ", '" & RS!red_codpro & "', " & RS!red_canpro & ", " & RS!red_cospro & ", " & RS!red_pctapr & ", " & RS!red_pctcoc & ", " & RS!red_pctnut & ", " & RS!red_tiprec & ")"
      Print #1, "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) " & _
                "VALUES (" & RS!red_codigo & ", " & RS!red_nroite & ", '" & RS!red_codpro & "', " & RS!red_canpro & ", " & RS!red_cospro & ", " & RS!red_pctapr & ", " & RS!red_pctcoc & ", " & RS!red_pctnut & ", " & RS!red_tiprec & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing
   
'   '-------> Generar gramaje aux
'   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbTablaGramajeAux '" & RecetaXML & "', '" & SubSegXML & "', '" & RegXML & "'")
'   Do While Not RS.EOF
'      DoEvents
'      Print #1, "INSERT INTO b_tablagramajeaux (tgr_subseg, tgr_codreg, tgr_codrec, tgr_coding, tgr_codzon, tgr_codins, tgr_cantgr)  " & _
'                "VALUES (" & RS!tgr_subseg & ", " & RS!tgr_codreg & ", " & RS!tgr_codrec & ", '" & RS!tgr_coding & "', " & RS!tgr_codzon & ", '" & RS!tgr_codins & "', " & RS!tgr_cantgr & ")"
'      RS.MoveNext
'   Loop
'   RS.Close: Set RS = Nothing
'
'   '-------> Generar gramaje aux
'   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbTablaGramajeAuxCeco '" & RecetaXML & "', '" & CecoXML & "', '" & RegXML & "'")
'   Do While Not RS.EOF
'      DoEvents
'      Print #1, "INSERT INTO b_tablagramajeauxceco (tgc_ceco, tgc_codreg, tgc_codrec, tgc_coding, tgc_codins, tgc_cantgr)  " & _
'                "VALUES ('" & RS!tgc_ceco & "', " & RS!tgc_codreg & ", " & RS!tgc_codrec & ", '" & RS!tgc_coding & "', '" & RS!tgc_codins & "', " & RS!tgc_cantgr & ")"
'      RS.MoveNext
'   Loop
'   RS.Close: Set RS = Nothing
   
   
End If

If op = 2 Or op = 3 Then
   
   If op <> 3 Then
      '-------> Generar regimen
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbRegimen ")
      Do While Not RS.EOF
         
         DoEvents
         Print #1, "INSERT INTO a_regimen (reg_codigo, reg_nombre) " & _
                   "VALUES (" & RS!Reg_Codigo & ", '" & RS!reg_nombre & "')"
         RS.MoveNext
      
      Loop
      RS.Close: Set RS = Nothing
   
      '-------> Generar servicio
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbServicio ")
      Do While Not RS.EOF
         
         DoEvents
         Print #1, "INSERT INTO a_servicio (ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable) " & _
                   "VALUES (" & RS!Ser_codigo & ", '" & RS!ser_nombre & "', " & RS!ser_orden & ", '" & RS!ser_codsap & "', '" & RS!ser_facturable & "')"
         RS.MoveNext
      
      Loop
      RS.Close: Set RS = Nothing
      
      '-------> Generar estructura servicio
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
     
      Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbEstServicio ")
      Do While Not RS.EOF
         
         DoEvents
         Print #1, "INSERT INTO a_estservicio (ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin) " & _
                   "VALUES (" & RS!ess_codser & ", " & RS!ess_codigo & ", '" & RS!ess_nombre & "', " & RS!ess_orden & ", " & RS!ess_codsec & ", " & RS!ess_racmin & ")"
         RS.MoveNext
      
      Loop
      RS.Close: Set RS = Nothing
   
   End If
   
   If op = 2 Then
      '-------> Generar encabezado planificación minutas
'      Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbMinuta '" & SubSegXML & "', " & Fecha & "")
'      Do While Not RS.EOF
'         DoEvents
'         Print #1, "INSERT INTO b_minuta (min_codigo, min_cencos, min_codreg, min_codser, min_fecmin, min_indblo, min_racteo, min_racrea) " & _
'                   "VALUES (" & RS!min_codigo & ", '" & RS!min_subseg & "', " & RS!min_codreg & ", " & RS!min_codser & ", " & RS!min_fecmin & ", " & RS!MIN_INDBLO & ", " & RS!min_racteo & ", " & RS!min_racrea & ")"
'         RS.MoveNext
'      Loop
'      RS.Close: Set RS = Nothing

   
      '-------> Generar detalle planificación minutas
'      Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbMinutaDetalle '" & SubSegXML & "', " & Fecha & "")
'      Do While Not RS.EOF
'         DoEvents
'         Print #1, "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_fecval, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes) " & _
'                   "VALUES (" & RS!mid_codigo & ", '" & RS!mid_tipmin & "', " & RS!mid_numlin & ", " & RS!mid_estser & ", " & RS!mid_codrec & ", " & RS!mid_numrac & ", '" & RS!mid_descri & "', " & RS!mid_cosrec & ", " & RS!mid_fecval & ", " & RS!mid_tiprec & ", " & RS!mid_nummer & ", '" & RS!mid_rec5eta & "', " & RS!mid_cosdes & ")"
'         RS.MoveNext
'      Loop
'      RS.Close: Set RS = Nothing

  
   End If
   
   '-------> Costo patron
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbParCostoPatron " & Fecha & "")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_costopatron (cpa_cencos, cpa_codreg, cpa_codser, cpa_anomes, cpa_descripcion, cpa_valor) " & _
                 "VALUES ('" & RS!pcp_cencos & "', " & RS!pcp_codreg & ", " & RS!pcp_codser & ", " & RS!pcp_anomes & ", '" & RS!pcp_descripcion & "', " & RS!pcp_valor & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing

   '-------> Gramo familia producto
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbGramoFamProducto ")
   Do While Not RS.EOF
      
      DoEvents
      Print #1, "INSERT INTO b_gramofamproducto (gfp_cencos, gfp_codreg, gfp_catdie, gfp_tiprec, gfp_fampro, gfp_graini, gfp_grafin) " & _
                "VALUES (" & RS!gfp_subseg & ", " & RS!gfp_codreg & ", " & RS!gfp_catdie & ", " & RS!gfp_tiprec & ", " & RS!gfp_fampro & ", " & RS!gfp_graini & ", " & RS!gfp_grafin & ")"
      RS.MoveNext
   
   Loop
   RS.Close: Set RS = Nothing

End If

Close #1

Dim strLocalFileName As String
Dim strLineReg As String
Dim strFileNameMDB As String

strFileNameMDB = ""
strLocalFileName = pArch

Open strLocalFileName For Input As #1
If Not EOF(1) Then
   Line Input #1, strLineReg
   Do While Not EOF(1)

'      lblStatus.Caption = "Procesando registros, " & Trim(Str(lngRow)) & "/" & Trim(Str(prbStatus.max))
      DoEvents
      If Mid(strLineReg, 1, 2) = "CC" Then
         
         If Trim(strFileNameMDB) <> "" Then db7.Close: Set db7 = Nothing
         strFileNameMDB = cDBI 'dir_trabajo & Trim(strLineReg) & ".mdb"
         If Dir(strFileNameMDB) <> "" Then Kill (strFileNameMDB)
         Set db7 = DBEngine(0).CreateDatabase(strFileNameMDB, dbLangGeneral)
      
      Else
         
         db7.Execute Trim(strLineReg)
      
      End If
      strLineReg = ""
      Line Input #1, strLineReg
'      lngRow = lngRow + 1
'      prbStatus.Value = lngRow
   
   Loop
   If Trim(strLineReg) <> "" Then db7.Execute Trim(strLineReg)
   
   If Trim(strFileNameMDB) <> "" Then
      
      If op = 1 Or op = 2 Then
         
         '-------> Buscar spid
         cSpi = 0
         Set RS = vg_db.Execute("SELECT @@spid spid")
         If Not RS.EOF Then cSpi = RS!spid
         RS.Close: Set RS = Nothing
         '-------> Genera tabla receta en tabla paso_receta
         Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbReceta '" & RecetaXML & "', " & cSpi & ", '" & vg_NUsr & "'")
         RS.Close: Set RS = Nothing
         
         db7.Execute "INSERT INTO b_receta (rec_codigo, rec_catdie, rec_tippla, rec_nombre, rec_nomfan, rec_metpre, rec_conche, rec_sugere, rec_basrac, rec_tiprec, rec_fecvig, rec_gruvul) SELECT a.rec_codigo, a.rec_catdie, a.rec_tippla, a.rec_nombre, a.rec_nomfan, a.rec_metpre, a.rec_conche, a.rec_sugere, a.rec_basrac, a.rec_tiprec, a.rec_fecvig, a.rec_gruvul FROM b_receta a, paso_receta b IN " & dBo & " WHERE a.rec_codigo = b.rec_codigo AND a.rec_indppr = 1 and b.rec_spid = " & cSpi & " and rec_usr = '" & vg_NUsr & "'"
      
         db7.Execute "INSERT INTO b_recetaaux (rec_codigo, rec_catdie, rec_tippla, rec_nombre, rec_nomfan, rec_metpre, rec_conche, rec_sugere, rec_basrac, rec_tiprec, rec_fecvig, rec_gruvul) SELECT a.rec_codigo, a.rec_catdie, a.rec_tippla, a.rec_nombre, a.rec_nomfan, a.rec_metpre, a.rec_conche, a.rec_sugere, a.rec_basrac, a.rec_tiprec, a.rec_fecvig, a.rec_gruvul FROM b_receta a "
      
      End If
      '-------> borrar index la tabla receta
      db7.Execute "drop index b_receta_pk on b_receta"
      db7.Close: Set db7 = Nothing
   
   End If

End If
'lblStatus.Visible = False: prbStatus.Visible = False
'-------> Cerrar tabla temporal y borrar
Close #1
If Dir(pArch) <> "" Then Kill pArch

End Function

Function SendMail1(cObj As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String, Cmailaviso As Integer, logenv As String)
Dim CHost As String, Caddr As String, Cuser As String, Cpass As String
On Error GoTo Man_Error

vg_GlosaEnvioCorreo = ""

If Trim(CmailU) <> "" Then
    DoEvents
    '-------> Traer parametro de cuenta correo
    Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("cor")) & "%'")
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: MsgBox "No existe Parametrización Correo, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Function
    Do While Not RS1.EOF
       If RS1!par_codigo = "corser" Then CHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corcum" Then Caddr = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corusu" Then Cuser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corpas" Then Cpass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    cObj.UnlockComponent "1mundoedwardsMAIL_3e5SOpaZmRkg"
    cObj.SmtpHost = CHost '"10.0.100.21"
    cObj.SmtpUsername = Cuser '"adminbdcasinos"
    cObj.SmtpPassword = Cpass '"admincasinos"
    cObj.ConnectTimeout = 30
'    Dim email As ChilkatEmail, Success As Long
'    Set email = New ChilkatEmail
    Dim email As New ChilkatEmail2, Success As Long
    email.AddTo cNombU, CmailU
    email.Subject = cSubject
    email.Body = cBody
    If Cmailaviso = 1 Then email.AddFileAttachment cArchivo
    email.FromName = IIf(Cmailaviso = 1, "Administrador SGP", "Administrador SGP")
    email.FromAddress = Caddr '"adminbdcasinos@sodexho.cl"
    cObj.LogMailSentFilename = logenv '"mailSent.log"
    Success = cObj.SendEmail(email)
'    MsgBox CHost & " " & Cuser & " " & Cpass & " " & Caddr & " "
    If (Success = 0) Then
'        MsgBox cObj.LastErrorText
        vg_GlosaEnvioCorreo = cObj.LastErrorText
    End If
End If
Exit Function
Man_Error:
    cObj.SaveXmlLog "log.xml"
    If Err = -2147467259 Then MsgBox "Cuenta no válida" & VgLinea & Trim(cNombU), vbCritical, "Envio Archivos a contratos": Exit Function
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error al enviar mail."
End Function

Function SendMail2(cObj As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String, Cmailaviso As Integer, logenv As String)
Dim CHost As String, Caddr As String, Cuser As String, Cpass As String
On Error GoTo Man_Error

vg_GlosaEnvioCorreo = ""

If Trim(CmailU) <> "" Then
    DoEvents
    '-------> Traer parametro de cuenta correo
    Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("cor")) & "%'")
    If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: MsgBox "No existe Parametrización Correo, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Function
    
    Do While Not RS1.EOF
       
       If RS1!par_codigo = "corser" Then CHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corcum" Then Caddr = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corusu" Then Cuser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       If RS1!par_codigo = "corpas" Then Cpass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    
    cObj.UnlockComponent "1mundoedwardsMAIL_3e5SOpaZmRkg"
    cObj.SmtpHost = CHost '"10.0.100.21"
    cObj.SmtpUsername = Cuser '"adminbdcasinos"
    cObj.SmtpPassword = Cpass '"admincasinos"
    cObj.ConnectTimeout = 30
'    Dim email As ChilkatEmail, Success As Long
'    Set email = New ChilkatEmail
    Dim email As New ChilkatEmail2, Success As Long
    email.AddTo cNombU, CmailU
    email.Subject = cSubject
    email.Body = cBody
    If Cmailaviso = 1 Then email.AddFileAttachment cArchivo
    email.FromName = IIf(Cmailaviso = 1, "Administrador SGP", "Administrador SGP")
    email.FromAddress = Caddr '"adminbdcasinos@sodexho.cl"
    cObj.LogMailSentFilename = logenv '"mailSent.log"
    Success = cObj.SendEmail(email)
'    MsgBox CHost & " " & Cuser & " " & Cpass & " " & Caddr & " "
    If (Success = 0) Then

'        MsgBox cObj.LastErrorText
        vg_GlosaEnvioCorreo = cObj.LastErrorText
    
    End If
End If

Exit Function
Man_Error:
    
    cObj.SaveXmlLog "log.xml"
    
    If Err = -2147467259 Then
       
       MsgBox "Cuenta no válida" & VgLinea & Trim(cNombU), vbCritical, "Envio Archivos a contratos"
       Exit Function
    
    End If
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error al enviar mail."

End Function

Function ExportHeaderFooter(vp As VSPrinter)
    
    ' no RTF export file? no work!
    If Len(vp.ExportFile) = 0 Then Exit Function
    If vp.ExportFormat < vpxRTF Then Exit Function
    
    ' build rtf style string for headers and foooters
    Dim rtfStyle$
'    rtfStyle = "\rtf1\ansi\ansicpg1252\deff0\deflang1033 " & _
'               "{\fonttbl{\f999 {{fname}};}}\li0\tqc\tx{{center}}\tqr\tx{{right}}\f999\fs{{fsize}}"
    rtfStyle = ""
    vp.GetMargins
    rtfStyle = Replace(rtfStyle, "{{center}}", (vp.x1 + vp.X2) / 2 - vp.x1)
    rtfStyle = Replace(rtfStyle, "{{right}}", vp.X2 - vp.x1)
    rtfStyle = Replace(rtfStyle, "{{fname}}", vp.HdrFontName)
    rtfStyle = Replace(rtfStyle, "{{fsize}}", CInt(2 * vp.HdrFontSize))
    If vp.HdrFontBold Then rtfStyle = rtfStyle & "\b"
    If vp.HdrFontItalic Then rtfStyle = rtfStyle & "\i"
    If vp.HdrFontUnderline Then rtfStyle = rtfStyle & "\ul"
    
    ' output header field
    Dim RTF$, s$, v
    s = vp.Header
    If Len(s) Then
        s = Replace(s, "\", "\\")
        s = Replace(s, "%d", "{\field{\*\fldinst PAGE}}")
        v = Split(s, "|")
        RTF = "{\header{" & rtfStyle & " {{left}} \tab {{center}} \tab {{right}}\par }}"
        RTF = Replace(RTF, "{{left}}", v(0))
        If UBound(v) >= 1 Then RTF = Replace(RTF, "{{center}}", v(1))
        If UBound(v) >= 2 Then RTF = Replace(RTF, "{{right}}", v(2))
        RTF = Replace(RTF, "{{center}}", "")
        RTF = Replace(RTF, "{{right}}", "")
        vp.ExportRaw = RTF
    End If

    ' output footer field
    s = vp.Footer
    If Len(s) Then
        s = Replace(s, "\", "\\")
        s = Replace(s, "%d", "{\field{\*\fldinst PAGE}}")
        v = Split(s, "|")
        RTF = "{\footer{" & rtfStyle & " {{left}} \tab {{center}} \tab {{right}}\par }}"
        RTF = Replace(RTF, "{{left}}", v(0))
        If UBound(v) >= 1 Then RTF = Replace(RTF, "{{center}}", v(1))
        If UBound(v) >= 2 Then RTF = Replace(RTF, "{{right}}", v(2))
        RTF = Replace(RTF, "{{center}}", "")
        RTF = Replace(RTF, "{{right}}", "")
        vp.ExportRaw = RTF
    End If

End Function

Function fg_ponepiepagina() As String
fg_ponepiepagina = "SGPADM v " & Trim(Str(App.Major)) & "." & Trim(Str(App.Minor)) & "." & Trim(Str(App.Revision) & "  " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm"))
End Function

Function formAbierto(op As String) As Boolean
Dim intIndexForm As Long
formAbierto = False
For intIndexForm = 0 To Forms.count - 1
    If (Forms(intIndexForm).Tag = Trim(op)) Then
        Forms(intIndexForm).WindowState = 0
        Forms(intIndexForm).SetFocus
        formAbierto = True
        Exit Function
    End If
Next
End Function

Function cFecha(Nform As Form, indice As Integer, Bandera As Integer)
Dim X As Long
For X = 0 To indice
    Nform.Fecha(X).DateTimeFormat = 5
    If Bandera = 1 Then
       Nform.Fecha(X).UserDefinedFormat = "dd/mm/yyyy"
    Else
       Nform.Fecha(X).UserDefinedFormat = "mmmm"
    End If
    Nform.Fecha(X).CalFirstDay (1)
    Nform.Fecha(X).ShortDayName(1) = "Dom"
    Nform.Fecha(X).ShortDayName(2) = "Lun"
    Nform.Fecha(X).ShortDayName(3) = "Mar"
    Nform.Fecha(X).ShortDayName(4) = "Mie"
    Nform.Fecha(X).ShortDayName(5) = "Jue"
    Nform.Fecha(X).ShortDayName(6) = "Vie"
    Nform.Fecha(X).ShortDayName(7) = "Sab"
    Nform.Fecha(X).LongDayName(1) = "Domingo"
    Nform.Fecha(X).LongDayName(2) = "Lunes"
    Nform.Fecha(X).LongDayName(3) = "Martes"
    Nform.Fecha(X).LongDayName(4) = "Miercoles"
    Nform.Fecha(X).LongDayName(5) = "Jueves"
    Nform.Fecha(X).LongDayName(6) = "Viernes"
    Nform.Fecha(X).LongDayName(7) = "Sabado"
    Nform.Fecha(X).LongMonthName(1) = "Enero"
    Nform.Fecha(X).LongMonthName(2) = "Febrero"
    Nform.Fecha(X).LongMonthName(3) = "Marzo"
    Nform.Fecha(X).LongMonthName(4) = "Abril"
    Nform.Fecha(X).LongMonthName(5) = "Mayo"
    Nform.Fecha(X).LongMonthName(6) = "Junio"
    Nform.Fecha(X).LongMonthName(7) = "Julio"
    Nform.Fecha(X).LongMonthName(8) = "Agosto"
    Nform.Fecha(X).LongMonthName(9) = "Septiembre"
    Nform.Fecha(X).LongMonthName(10) = "Octubre"
    Nform.Fecha(X).LongMonthName(11) = "Noviembre"
    Nform.Fecha(X).LongMonthName(12) = "Diciembre"
Next X
End Function

Public Function ExraeCodCombo(text As String) As Integer
'Creado Samuel Melendez 15/09/09
    If text = "" Then
        ExraeCodCombo = 0
    Else
        ExraeCodCombo = Val(Mid(text, InStr(1, text, "(") + 1, (Len(text) - InStr(1, text, "(")) - 1))
    End If
End Function

Public Function ExtraeFecha(Fec As fpDateTime) As Long
'Creado Samuel Melendez 15/09/09
ExtraeFecha = 0
If Not IsNull(Fec.text) And Trim(Fec.text) <> "" Then
    ExtraeFecha = Mid(Fec.text, 4, 4) & Mid(Fec.text, 1, 2)
End If
End Function

Function fg_Dtos(Fecha As String) As String

Dim cFec As Date

If Not IsDate(Fecha) Then fg_Dtos = "": Exit Function

If Trim(Fecha) = "" Then
    
    fg_Dtos = ""

Else
    
    cFec = CDate(Fecha)
    fg_Dtos = fg_pone_cero(Str(Year(Fecha)), 4) & fg_pone_cero(Str(Month(Fecha)), 2) & fg_pone_cero(Str(Day(Fecha)), 2)

End If

End Function

Function MoverDatosExcel(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
   sheet.Range(col1 & row1).Value = datos

End Function

Function MoverDatosExcelDosDecimales(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
    sheet.Range(ColFec & row1 & ":" & ColPor & row2).Select
    excel.Selection.NumberFormat = "0.00"
    
End Function

Function MoverDatosExcelCombo(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)

  'Range("A1:A9").Select
     sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    With excel.Selection.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=hoja1!$a$1:$a$8523"
        .IgnoreBlank = False 'True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
       sheet.Range(col1 & row1).Value = datos
End Function

Function MoverDatosExcelFormula(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
   
   If sheet.Range(ColFec & row1).Value <> "" Then
      
      sheet.Range(ColPor & row1).Value = "=(" & col1 & row1 & " / " & col2 & row2 & ")"
      sheet.Range(ColPor & row1 & ":" & ColPor & row2).Select
      excel.Selection.NumberFormat = "0%"
   
   End If

End Function

'Function MoverDatosExcelFormulaII(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
'
'   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
'
'   If sheet.Range(ColFec & row1).Value <> "" Then
'
'      sheet.Range(ColPor & row1).Value = "=(" & col1 & row1 & " * " & col2 & row1 & ")/" & col1 & row2 & ""
'      sheet.Range(ColPor & row1 & ":" & ColPor & row2).Select
'      excel.Selection.NumberFormat = "0.00"
'
'   End If
'
'End Function

Function MoverDatosExcelCostoBandeja(excel As Object, sheet As Object, Col As String, Row As String, datos As String)
   
   sheet.Range(Col & Row & ":" & Col & Row).Select
   
   sheet.Range(Col & Row).Value = "=(" & datos & ")" '"=(" & col1 & row1 & " * " & col2 & row1 & ")/" & col1 & row2 & ""
   sheet.Range(Col & Row & ":" & Col & Row).Select
   excel.Selection.NumberFormat = "0.00"

End Function

Function MoverDatosExcelFormulaEntero(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
   
   If sheet.Range(ColFec & row1).Value <> "" Then
      
      sheet.Range(col1 & row1).Value = "=(" & ColPor & row1 & " * " & col2 & row2 & ")"
'      sheet.Range(ColPor & row1 & ":" & ColPor & row2).Select
'      excel.Selection.NumberFormat = "0%"
   
   End If

End Function

Function MoverDatosExcelFormulaSum(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
'sheet.Range(col1 & 6).Select
'excel.ActiveCell.FormulaR1C1 = "=SUM(R[" & row1 - 7 & "]" & "C" & ":R[" & row2 - 7 & "]" & "C" & ")"
'excel.Selection.NumberFormat = "0.00"
'sheet.Range(col1 & 6).Select

sheet.Range(col1 & 7).Select
excel.ActiveCell.FormulaR1C1 = "=iferror(SUM(R[" & row1 - 7 & "]" & "C" & ":R[" & row2 - 7 & "]" & "C" & "),0)"
excel.Selection.NumberFormat = "0.00"
sheet.Range(col1 & 7).Select

End Function

Function MoverDatosExcelFormulaPegadoEspecial(excel As Object, sheet As Object, col1 As String, row1 As Long, row2)
   
sheet.Range(col1 & row2).Select
excel.Selection.Copy
sheet.Range(col1 & row1).Select
excel.Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                            :=False, Transpose:=True
                             excel.Selection.NumberFormat = "0.00"
                            sheet.Range(col1 & row1).Select
End Function

'Function MoverDatosExcelBuscarV(excel As Object, sheet As Object, ColRec As String, col1 As String, row1 As Long, row2 As Long)
'
'On Error GoTo Man_Error
'
'sheet.Range(col1 & row1).Select
'excel.ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],Recetas!R2C2:R[" & row2 & "]C[-2],2,FALSE),0)" '"=VLOOKUP(RC[-3],Hoja1!RC:R[-3]C[-1],2,FALSE)" '"=VLOOKUP(RC[-3],Hoja1!R[-7]C[-117]:R[9019]C[-116],2,FALSE)" '"=BUSCARV(" & ColRec & row1 & ";recetas!B2:C45028;2;FALSO)"
''excel.ActiveCell.Value = "=BUSCARV(" & ColRec & row1 & ";Recetas!B2:C45028;2;FALSO)"
'
''ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Hoja1!RC:R[1]C[1],2,FALSE)"
''"=VLOOKUP(RC[-3],Recetas!R[-7]C[-117]:R[9019]C[-116],2,FALSE)"
''=BUSCARV(B9;Recetas!B2:C9028;2;FALSO)
''"=SUM(R[" & row1 - 7 & "]" & "C" & ":R[" & row2 - 7 & "]" & "C" & ")"
'
''excel.Selection.NumberFormat = "0.00"
'sheet.Range(col1 & row1).Select
'
'Exit Function
'Man_Error:
'Resume Next
'
'End Function

Function MoverDatosExcelBuscarV(excel As Object, sheet As Object, ColRec As String, col1 As String, row1 As Long, row2 As Long, EstVacio)
   
On Error GoTo Man_Error

Dim RowPla As Long
RowPla = 0

If EstVacio Then
      
   RowPla = 9 'row1 + 1
      
   Do While RowPla <= vg_RowEnd + 8
   
      sheet.Range(col1 & RowPla).Select
      excel.ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],Recetas!R2C2:R[" & row2 & "]C[-2],2,FALSE),0)"
   
      RowPla = RowPla + 1
        
   Loop

End If

sheet.Range(col1 & row1).Select
excel.ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],Recetas!R2C2:R[" & row2 & "]C[-2],2,FALSE),0)" '"=VLOOKUP(RC[-3],Hoja1!RC:R[-3]C[-1],2,FALSE)" '"=VLOOKUP(RC[-3],Hoja1!R[-7]C[-117]:R[9019]C[-116],2,FALSE)" '"=BUSCARV(" & ColRec & row1 & ";recetas!B2:C45028;2;FALSO)"
'excel.ActiveCell.Value = "=BUSCARV(" & ColRec & row1 & ";Recetas!B2:C45028;2;FALSO)"

'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Hoja1!RC:R[1]C[1],2,FALSE)"
'"=VLOOKUP(RC[-3],Recetas!R[-7]C[-117]:R[9019]C[-116],2,FALSE)"
'=BUSCARV(B9;Recetas!B2:C9028;2;FALSO)
'"=SUM(R[" & row1 - 7 & "]" & "C" & ":R[" & row2 - 7 & "]" & "C" & ")"

'excel.Selection.NumberFormat = "0.00"

sheet.Range(col1 & row1).Select

Exit Function
Man_Error:
Resume Next

End Function

Function ValidarExisteDato(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String) As Boolean

ValidarExisteDato = False
If Trim(sheet.Range(col1 & row1).Value) <> "" Then
   ValidarExisteDato = True
End If

End Function

Function PonerFontBold(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Font.Bold = True

End Function

Function PonerCombinarCentrar(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & Trim(col2) & row2).Select
    
    With excel.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    excel.Selection.Merge

End Function

Function PonerTipoLetraTamaño(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, tamaño As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    
    With excel.Selection.Font
        .Name = "Arial"
        .Size = tamaño '12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
'        .ColorIndex = 2
    End With

End Function

Function DibujarLineas(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With excel.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With excel.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With excel.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
    With excel.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = xlAutomatic
    End With
End Function

Function OcultarColumna(excel As Object, sheet As Object, col1 As String, col2 As String)
    sheet.Columns(col1 & ":" & col2).Select
'    sheet.Range(col1 & ":" & col2).Select
    With excel.Selection

        .EntireColumn.Hidden = True
    
    End With
    
End Function

Function ExtraerDatos(excel As Object, sheet As Object, Col As String, Row As Long) As String

ExtraerDatos = ""
ExtraerDatos = sheet.Range(Col & Row).Value

End Function

Function PonerColorInterior(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    excel.Selection.Interior.ColorIndex = 5

End Function

Function PonerColorInteriorN(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, Color As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    excel.Selection.Interior.ColorIndex = Color

End Function

Function PonerColorFont(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    excel.Selection.Font.ColorIndex = 2

End Function

Function PonerNegrilla(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
   
   sheet.Range(col1 & row1 & ":" & col2 & row2).Select
   excel.Selection.Font.Bold = True

End Function

Function PonerCentrado(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    With excel.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With

End Function

Function PonerTodosLosBorde(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)

'    Range("F3:H11").Select
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select

    excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With excel.Selection.Borders(xlEdgeLeft)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    
    End With
    
    With excel.Selection.Borders(xlEdgeTop)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    
    End With
    
    With excel.Selection.Borders(xlEdgeBottom)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    
    End With
    
    With excel.Selection.Borders(xlEdgeRight)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    
    End With
    
    With excel.Selection.Borders(xlInsideVertical)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    
    End With
    
    With excel.Selection.Borders(xlInsideHorizontal)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    
    End With

End Function

Function PonerAnchoColumna(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, Ancho As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    excel.Selection.ColumnWidth = Ancho

End Function

Function PonerCombinarLeft(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
    
    With excel.Selection
        
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    
    End With

End Function

Function PonerBordeColor(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long)
    
    sheet.Range(col1 & row1 & ":" & col2 & row2).Select
'    Range("E2:G2").Select
    excel.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    excel.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    With excel.Selection.Borders(xlEdgeLeft)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    
    End With
    
    With excel.Selection.Borders(xlEdgeTop)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    
    End With
    
    With excel.Selection.Borders(xlEdgeBottom)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    
    End With
    
    With excel.Selection.Borders(xlEdgeRight)
        
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = 5
    
    End With
 '   excel.Selection.Borders(xlInsideVertical).LineStyle = xlNone

End Function

Function VistaPreliminarExcel(excel As Object, sheet As Object, OpOrientacion As Boolean)
    
    '-------> Vista Preliminar
    With sheet.PageSetup
        
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = 0
        .RightMargin = 0
        .TopMargin = 0
        .BottomMargin = 0
        .HeaderMargin = 0
        .FooterMargin = 0
        .PrintHeadings = False 'True
        .PrintGridlines = False 'True
        .PrintComments = xlPrintInPlace
'        .PrintQuality = 600
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = IIf(OpOrientacion, xlLandscape, xlPortrait) ' 1º=Horizontal;2º=Vertical
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False 'True
    
    End With

'    Cells.Select
    sheet.DisplayAutomaticPageBreaks = False

End Function

Function ValidarNombreHoja(excel As Object, sheet As Object, nombrehoja As String) As Boolean

Dim numeroHojas
Dim i As Long
ValidarNombreHoja = False
numeroHojas = excel.Sheets.count

For i = 1 To numeroHojas
   
   If nombrehoja = excel.Sheets(i).Name Then
       
       ValidarNombreHoja = True
   
   End If

Next

End Function

Function BloquearColumnaExcel(excel As Object, sheet As Object, Col As String, row1, row2)
    
    sheet.Range(Col & row1 & ":" & Col & row2).Select
    excel.Selection.Locked = False
    excel.Selection.FormulaHidden = False

End Function
    
Function PonerColorGrisExcel(excel As Object, sheet As Object, Col As String, row1, row2)

    sheet.Range(Col & row1 & ":" & Col & row2).Select
    With excel.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Function

Function FormatearColumnaNumericoExcel(excel As Object, sheet As Object, Col As String)
    
    sheet.Range(Col & ":" & Col).Select
    excel.Selection.NumberFormat = "0"

End Function
    
Function FormatearColumnaPorcentajeExcel(excel As Object, sheet As Object, Col As String)
    
    sheet.Range(Col & ":" & Col).Select
    excel.Selection.NumberFormat = "0%"

End Function
    
Function SendMailOutlook(cObj As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String, Cmailaviso As Integer, logenv As String)

On Error GoTo Man_Error

Dim Out As New Outlook.Application
Dim msg As Outlook.MailItem

   ' Dim Archivo As String
   ' Casilla = Trim(RS!cli_email)   '"marco.maturana@sodexo.com"
   ' Sujeto = "Adjunto archivo traspaso " & Format(Date, "dd/mm/yyyy")
   vg_GlosaEnvioCorreo = ""
   Set msg = Out.Session.GetDefaultFolder(olFolderOutbox).Items.Add
   
   With msg

'       If ConCopia.text <> "" Then .CC = ConCopia.text
       If cArchivo <> "" Then .Attachments.Add (cArchivo)
       .BodyFormat = olFormatRichText
       .Subject = cSubject  'Asunto.text
       .Body = cBody 'Cuerpo.text
   
   End With
   
   With msg.Recipients.Add(CmailU) 'cuenta)
       
       .Type = Outlook.olTo
       .Resolve
   
   End With
   'jpazSendMail_Status "Dirección: " & CmailU 'cuenta
   msg.send
   'jpazSendMail_Status "Enviado..."
'   Grilla.Col = 3: Grilla.text = "Enviado Ok"
   Set msg = Nothing


Exit Function
Man_Error:
    'oMail.SaveXmlLog "log.xml"
'    If Err = -2147467259 Then MsgBox "Cuenta no válida" & VgLinea & Trim(cNombU), vbCritical, "Envio Archivos a contratos": Exit Function
    If Err = 424 Or Err = 429 Or Err = 91 Then vg_GlosaEnvioCorreo = "Error envio": Resume Next
    vg_GlosaEnvioCorreo = "Error envio"
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error al enviar mail."

End Function

Function SendMailGmail(oMail As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String)

On Error GoTo Man_Error

If Trim(CmailU) <> "" Then
    
    oMail.UnlockComponent "1mundoedwardsMAIL_3e5SOpaZmRkg"
    oMail.SmtpAuthMethod = "NONE"
    oMail.SmtpHost = "smtp.gmail.com" '"svrmail2003"
    oMail.SmtpPort = 465
    oMail.SmtpSsl = 1
    oMail.SmtpUsername = "SGP.Pedidos@gmail.com" '"j.gonzalez.prb@gmail.com" '"adminbdcasinos"
    oMail.SmtpPassword = "Sgpsodexo" '"jgprb2013" '"admincasinos"
    oMail.ConnectTimeout = 30
    Dim email As ChilkatEmail2, Success As Long
    Set email = New ChilkatEmail2
    email.AddTo cNombU, CmailU
    email.Subject = cSubject
    email.Charset = "windows-1252"
    email.Body = cBody
    email.AddFileAttachment cArchivo
    email.FromName = "Administrador SGP Pedidos"
    email.FromAddress = "SGP.Pedidos@gmail.com" '"j.gonzalez.prb@gmail.com" '"adminbdcasinos@sodexho.cl"
    oMail.LogMailSentFilename = dir_trabajo & "mailSent.log"
    Success = oMail.SendEmail(email)
    vg_codigo = ""
    
    If (Success = 0) Then
        
        vg_codigo = "1"
        MsgBox oMail.LastErrorText
    
    End If

End If

Exit Function
Man_Error:

    oMail.SaveXmlLog "log.xml"
'    If Err = -2147467259 Then MsgBox "Cuenta no válida" & VgLinea & Trim(cNombU), vbCritical, "Envio Archivos a contratos": Exit Function
    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error al enviar mail."

End Function

Sub valida_Pedidos(Form As Object, cecos As Long, fecha_desde As String, fecha_hasta As String, estado As Long)

Dim strSQL As String
   pedidos = 0
   fg_carga ""
   
   strSQL = " sgpadm_val_EstadoPedido "
   strSQL = strSQL & " '" & cecos & "'"
   strSQL = strSQL & " , " & fecha_desde & " "
   strSQL = strSQL & " , " & fecha_hasta & ""
   strSQL = strSQL & " , " & estado & ""
   Set RS1 = vg_db.Execute(strSQL)
    
   '-------> Inicio LLenar grilla
   
   M_GenPed.vaSpread3.MaxRows = 0
    
   
   Do While Not RS1.EOF
        
        M_GenPed.vaSpread3.MaxRows = M_GenPed.vaSpread3.MaxRows + 1
        M_GenPed.vaSpread3.Row = M_GenPed.vaSpread3.MaxRows
        M_GenPed.vaSpread3.Col = 1 ' IdCompra
        M_GenPed.vaSpread3.text = Val(RS1(0))
        M_GenPed.vaSpread3.Col = 2 ' Cod. Ingrediente
        M_GenPed.vaSpread3.text = RS1(1)
        M_GenPed.vaSpread3.Col = 3 ' Des. Ingrediente
        M_GenPed.vaSpread3.text = RS1(2)
        M_GenPed.vaSpread3.Col = 4 ' Proveedor
        M_GenPed.vaSpread3.text = RS1(3)
        M_GenPed.vaSpread3.Col = 5 ' Familia Producto
        pedidos = 1
 
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    
    fg_descarga

End Sub

Sub Valida_Generacion_pedido(Form As Object)

Dim i As Integer
Dim IdRuta As Integer
Dim CodIngrediente As String
Dim CodProveedor As String
Dim codproducto As String
Dim FechaDespacho As String
Dim total As Integer
Dim CodProductoSGP As String
Dim CantidadIngrediente As Integer
Dim CantidadProducto As Integer
Dim descripcion As String
valida = True
 
For i = 1 To M_GenPed.vaSpread1.MaxRows
M_GenPed.vaSpread1.Row = i
M_GenPed.vaSpread1.Col = 1 'Id Ruta de Compras
IdRuta = IIf(M_GenPed.vaSpread1.text = "", 0, M_GenPed.vaSpread1.text)
M_GenPed.vaSpread1.Col = 2 'Código Ingrediente
CodIngrediente = M_GenPed.vaSpread1.text
M_GenPed.vaSpread1.Col = 3 'Descripción Ingrediente
descripcion = M_GenPed.vaSpread1.text
M_GenPed.vaSpread1.Col = 4 'Código Proveedor SAP
CodProveedor = M_GenPed.vaSpread1.text
M_GenPed.vaSpread1.Col = 5 'Código Familia Producto
M_GenPed.vaSpread1.Col = 6 'Centro costo
M_GenPed.vaSpread1.Col = 7 'Código Producto SAP
codproducto = M_GenPed.vaSpread1.text
If codproducto = "" Then
MsgBox " El Ingrediente " & CodIngrediente & " " & descripcion & " No Tiene Fromato Sap.... ", 16
valida = False
Exit Sub
End If

M_GenPed.vaSpread1.Col = 8 'Descripción Producto
M_GenPed.vaSpread1.Col = 9 'Unidad
M_GenPed.vaSpread1.Col = 10 'Fecha Despacho
FechaDespacho = IIf(M_GenPed.vaSpread1.text = "", 0, Format(M_GenPed.vaSpread1.text, "yyyymmdd"))
M_GenPed.vaSpread1.Col = 11 'Total
total = M_GenPed.vaSpread1.text
'M_GenPed.vaSpread1.Col = 12 'Còdigo Producto SGP
'CodProductoSGP = M_GenPed.vaSpread1.text
'M_GenPed.vaSpread1.Col = 13 'Cantidad Ingrediente SGP
'CantidadIngrediente = M_GenPed.vaSpread1.text
'M_GenPed.vaSpread1.Col = 14 'Cantidad Producto SGP
'CantidadProducto = M_GenPed.vaSpread1.text
Next i

End Sub

Function fg_codigolistaNuevo(List As String, Index As Integer, Largo As Integer, cDefa As Variant)

If Len(Trim(List)) > Largo Then
    fg_codigolistaNuevo = Mid(Trim(List), Len(Trim(List)) + 1 - Largo, Largo)
Else
    fg_codigolistaNuevo = cDefa
End If

End Function

Function fg_buscacboNuevo(Combo As Object, Index As Integer, Largo As Integer, cBusca As String)

Dim i As Integer
fg_buscacboNuevo = -1
For i = 0 To Combo(Index).ListCount - 1
    Combo(Index).ListIndex = i
    If Mid(Trim(Combo(Index).text), Len(Trim(Combo(Index).text)) + 1 - Largo, Largo) = Val(cBusca) Then
        fg_buscacboNuevo = i
        Exit For
    End If
Next

End Function

Function Traerfechadia(Fecha As Date, numdia As Integer) As String

Traerfechadia = Format(Fecha, "dd/mm/yyyy")
Do While DatePart("w", Fecha, 2) <> numdia
     Fecha = Fecha + 1
     Traerfechadia = Format(Fecha, "dd/mm/yyyy")
Loop

End Function

Function SacarNumeroRight(valor As String) As Long

Dim i         As Long
Dim X         As Long

SacarNumeroRight = 0

If Trim(valor) = "" Then

    Exit Function
    
End If

For i = Len(valor) To 1 Step -1
    
    If Mid(valor, i, 1) = " " Then
       
       X = IIf(i + 1 < 1, 1, i + 1)
       
       Exit For
        
    End If
    
Next i

If X = 0 Then X = 1
If IsNumeric(Mid(valor, X, Len(valor))) Then
   SacarNumeroRight = Mid(valor, X, Len(valor))
End If

End Function

Function encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

Function encabezado_V02(ByRef RS As ADODB.Recordset, ByRef xlWs As Object, ByRef Row As Long, ByRef titulo1 As String, ByRef titulo2 As String)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long
Dim I_Row As Long
I_Row = Row

xlWs.Cells(1, 4).Value = titulo1
xlWs.Cells(2, 4).Value = titulo2

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(I_Row, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

Function fg_GrabaLogSistema(cNUsuario As String, cOpcion As Long, cOpcionSistema As String, cDatoNuevo As String, cDatoAnterior As String, CDetalleOperacion As String) As String

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Ins_logsistema_V01 '" & cNUsuario & "', " & cOpcion & ", '" & IIf(cOpcionSistema <> "MINSAN", Fg_Ponerpunto(cOpcionSistema), cOpcionSistema) & "', '" & cDatoAnterior & "', '" & cDatoNuevo & "', '" & CDetalleOperacion & "'")

If Not RS.EOF Then
   
   If RS(0) > 0 Then
                       
      MsgBox RS(0) & " " & RS(1)
   
    End If
    
End If

RS.Close: Set RS = Nothing

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

Function fg_TraeLogConcepto(referencia As String) As Long

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_TraeLogConcepto = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_TraeLogConcepto '" & referencia & "'")

If Not RS.EOF Then
   
   fg_TraeLogConcepto = TipoDato(RS!loc_codigo, "")
   
End If

RS.Close
Set RS = Nothing

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

Function Fg_Ponerpunto(ByVal Punto As String) As String

On Error GoTo Man_Error

Dim X%, j%
Dim ValLcntH$
ValLcntH = ""
j = 1

For X = 1 To Len(Punto)
    
    If Asc(Mid(Punto, X, 1)) <> 46 Then
       
       ValLcntH = IIf(j = 4, ValLcntH + "." + Mid(Punto, X, 1), ValLcntH + Mid(Punto, X, 1))
       j = IIf(j = 4, 2, j + 1)
    
    End If

Next X

Fg_Ponerpunto = ValLcntH

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
    
End Function

Function fg_ArchivoRtf()

Dim i As Long

i = 1

For i = 1 To 9999999
    
    If Dir(dir_trabajo & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "Reporte.rtf") = "" Then
        
        fg_ArchivoRtf = dir_trabajo & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "Reporte.rtf": Exit Function
    
    End If

Next i

End Function

Function fg_ArchivoExcel()

Dim i As Long

i = 1

For i = 1 To 9999999
    
    If Dir(dir_trabajo & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "ReporteExcel.rtf") = "" Then
        
        fg_ArchivoExcel = dir_trabajo & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "ReporteExcel.rtf": Exit Function
    
    End If

Next i

End Function

Function fg_ArchivoPDF()

Dim i As Long

i = 1

For i = 1 To 9999999
    
    If Dir(dir_trabajo & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "Reporte.pdf") = "" Then
        
        fg_ArchivoPDF = dir_trabajo & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "Reporte.pdf": Exit Function
    
    End If

Next i

End Function

Function fg_ArchivoTXT_1(ruta As String)

Dim i As Long

i = 1

For i = 1 To 9999999
    
    If Dir(ruta & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "Reporte.txt") = "" Then
        
        fg_ArchivoTXT_1 = ruta & vg_NUsr & fg_pone_cero(Trim(Str(i)), 7) & "Reporte.txt": Exit Function
    
    End If

Next i

End Function

Function fg_ValidaPassword(cUsr As String, Cpass As String, cMsg As String) As Boolean

Dim RS_Dato2    As New ADODB.Recordset
Dim cPassLong   As Long
Dim cPassAnt    As Long
Dim cContPass   As Long
Dim i           As Long
Dim cCar        As String
Dim cCantNum    As Integer
Dim cCantCar    As Integer
Dim cCantCarMay As Integer
Dim cCantEsp    As Integer
Dim cCantExep   As Integer
Dim cCarEsp     As String

fg_ValidaPassword = True

'Revisa caracteres
cCarEsp = GetParametro("pscara")
cCantNum = 0
cCantCar = 0
cCantEsp = 0
cCantExep = 0
cCantCarMay = 0

For i = 1 To Len(Cpass)
    
    cCar = Mid(Cpass, i, 1)
    
    If (Asc(cCar) >= 48 And Asc(cCar) <= 57) Then
        
        cCantNum = cCantNum + 1
    
    End If
    
    If (Asc(cCar) >= 65 And Asc(cCar) <= 90) Then
    
       cCantCarMay = cCantCarMay + 1
    
    End If
    
    If (Asc(cCar) >= 97 And Asc(cCar) <= 122) Or Asc(cCar) = 241 Or Asc(cCar) = 209 Then
        
        cCantCar = cCantCar + 1
    
    End If
    
    If InStr(cCarEsp, cCar) <> 0 Then
        
        cCantEsp = cCantEsp + 1
    
    End If
    
    If Not ((Asc(cCar) >= 48 And Asc(cCar) <= 57)) And Not ((Asc(cCar) >= 97 And Asc(cCar) <= 122) Or (Asc(cCar) >= 65 And Asc(cCar) <= 90) Or Asc(cCar) = 241 Or Asc(cCar) = 209) And Not (InStr(cCarEsp, cCar) <> 0) Then
        
        cCantExep = cCantExep + 1
    
    End If

Next i

If cCantCarMay = 0 Then
    
    MsgBox "La password debe contener por lo menos " & VgLinea & "una letra mayuscula. ", vbCritical + vbOKOnly, cMsg
    
    fg_ValidaPassword = False
    Exit Function

End If

If cCantCar = 0 Then

    MsgBox "La password debe contener por lo menos " & VgLinea & "una letra minuscula. ", vbCritical + vbOKOnly, cMsg
    
    fg_ValidaPassword = False
    Exit Function

End If

If cCantExep > 0 Then
    
    MsgBox "La password no puede contener caracteres que no sean " & VgLinea & _
           "una letra, un número o un caracter especial (" & cCarEsp & "). ", vbCritical + vbOKOnly, cMsg
    
    fg_ValidaPassword = False
    Exit Function

End If

If cCantNum = 0 Or cCantCar = 0 Or cCantEsp = 0 Then
    
    MsgBox "La password debe contener por lo menos " & VgLinea & _
           "una letra, un número y un caracter especial (" & cCarEsp & "). ", vbCritical + vbOKOnly, cMsg
    
    fg_ValidaPassword = False
    Exit Function

End If

'Revisa largo de la Password
cPassLong = GetParametro("pslong")
If Len(Cpass) < cPassLong Then
    
    MsgBox "La password debe tener una longitud mínima de " & cPassLong & " caracteres.", vbCritical + vbOKOnly, cMsg
    fg_ValidaPassword = False
    Exit Function

End If

If RS_Dato2.State = 1 Then RS_Dato2.Close
RS_Dato2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'Revisa password anteriores
cPassAnt = GetParametro("psante")
cContPass = 0
Set RS_Dato2 = vg_db.Execute("sgpadm_Sel_Log_CambiaPass 1, '" & cUsr & "', " & fg_TraeLogConcepto("vg_logsis_CambiaPass") & "")
If Not RS_Dato2.EOF Then
    
    Do While Not RS_Dato2.EOF
        
        If fg_Encripta(Cpass) = RS_Dato2!datonuevo Then
            
            MsgBox "La password no puede ser igual a las " & cPassAnt & " password anteriores.", vbCritical + vbOKOnly, cMsg
            RS_Dato2.Close: Set RS_Dato2 = Nothing
            fg_ValidaPassword = False
            Exit Function
        
        End If
        cContPass = cContPass + 1
        If cContPass = cPassAnt Then
            
            Exit Do
        End If
        RS_Dato2.MoveNext
    
    Loop

End If
RS_Dato2.Close: Set RS_Dato2 = Nothing

End Function

Function fg_pone_rchar(ByVal cadena As String, ByVal cuanto As Integer, ByVal char As String) As String

'pone caracteres a la derecha
fg_pone_rchar = Trim(cadena) & String(cuanto - Len(Trim(cadena)), char)

End Function

Function GetSerialNumber(strDrive As String) As Long

    Dim SerialNum As Long
    Dim res As Long
    Dim Temp1 As String
    Dim Temp2 As String

    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))

    res = GetVolumeInformation(strDrive, Temp1, _
    Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))

    GetSerialNumber = SerialNum

End Function

Function fg_ValidarDirectorio(Directorio As String) As Boolean
Dim fso As Object

'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
Set fso = CreateObject("Scripting.FileSystemObject")

fg_ValidarDirectorio = False

'validar si existe carpeta
If fso.FolderExists(Directorio) Then
   
   fg_ValidarDirectorio = True
       
End If

End Function

Function fg_ValidarUnidadDisco(Driver As String) As Boolean

Dim i   As Long
Dim ret As Long
'
ret = GetLogicalDrives()
fg_ValidarUnidadDisco = False

If ret Then
    
    For i = 0 To 25
        ' Si el bit es cero, es que no existe la unidad o no está mapeada
        If (ret And 2 ^ i) <> 0 Then
           
           ' Mostrar el nombre de la unidad ocupada
            If Driver = Chr$(i + 65) & ":\" Then
            
               fg_ValidarUnidadDisco = True
            
            End If
            
        End If
    
    Next
    
End If

End Function

Function Unix2Dos(file) As Boolean

On Error GoTo Man_Error

Unix2Dos = True

    Dim fs As Object, txt As String
    Set fs = CreateObject("Scripting.FileSystemObject")

    txt = fs.OpenTextFile(file, 1).ReadAll
    txt = Replace(txt, vbLf, vbCrLf)
    fs.OpenTextFile(file, 2).Write txt

Exit Function
Man_Error:
    
    Unix2Dos = False
    MsgBox Err.Description, vbCritical, MsgTitulo

End Function


' Funcción que abre el cuadro de dialogo y retorna la ruta
'******************************************************************
Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String

On Local Error GoTo errFunction

    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object

    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")

    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder(0, Titulo, 0, Path_Inicial)

    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self

    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path

Exit Function
'Error
errFunction:
    
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString

End Function

Function MoverDatosExcelFormulaRaciones(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String, EstVacio As Boolean)
   
   Dim RowPla As Long
   RowPla = 0
   
   Dim RowPla1 As Long
   RowPla1 = 0
   RowPla1 = row2 - 9
   
   If EstVacio Then
        
      RowPla = 9 'row1
        
      Do While RowPla <= vg_RowEnd + 8
        
         sheet.Range(ColPor & RowPla).Value = "=(" & col1 & RowPla & " / " & col2 & RowPla & ")"
         sheet.Range(ColPor & RowPla & ":" & ColPor & RowPla).Select
         excel.ActiveCell.FormulaR1C1 = "=IFERROR((RC[-1] / R[" & RowPla1 & "]C[-1]),0)"
         
         excel.Selection.NumberFormat = "0%"
        
         RowPla = RowPla + 1
         RowPla1 = RowPla1 - 1
              
      Loop
    
   End If
   
   RowPla1 = row2 - row1
   
   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
   
   If sheet.Range(ColFec & row1).Value <> "" Then
      
      sheet.Range(ColPor & row1).Value = "=(" & col1 & row1 & " / " & col2 & row2 & ")"
      sheet.Range(ColPor & row1 & ":" & ColPor & row2).Select
      excel.ActiveCell.FormulaR1C1 = "=IFERROR((RC[-1] / R[" & RowPla1 & "]C[-1]),0)"
      
      excel.Selection.NumberFormat = "0%"
   
   End If

End Function

Function MoverDatosExcelFormulaPonderacion(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String, EstVacio As Boolean)
   
   Dim RowPla As Long
   RowPla = 0
   
   Dim RowPla1 As Long
   RowPla1 = 0
   RowPla1 = row2 - 9
   
   If EstVacio Then

      RowPla = 9

      Do While RowPla <= vg_RowEnd + 8

         sheet.Range(col1 & RowPla).Value = "=(" & ColPor & RowPla & " * " & col2 & RowPla & ")"
         sheet.Range(col1 & RowPla & ":" & col1 & RowPla).Select
         excel.ActiveCell.FormulaR1C1 = "=IFERROR((RC[1] * R[" & RowPla1 & "]C),0)"
         excel.Selection.NumberFormat = "0"

         RowPla = RowPla + 1
         RowPla1 = RowPla1 - 1

      Loop

   End If
   
   RowPla1 = row2 - row1
   
   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
   
   If sheet.Range(ColFec & row1).Value <> "" Then
      
      sheet.Range(col1 & row1).Value = "=(" & ColPor & row1 & " * " & col2 & row2 & ")"
      sheet.Range(col1 & row1 & ":" & col1 & row2).Select
      excel.ActiveCell.FormulaR1C1 = "=IFERROR((RC[1] * R[" & RowPla1 & "]C),0)"
      
      excel.Selection.NumberFormat = "0"
   
   End If

End Function

Function MoverDatosExcelFormulaII(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String, EstVacio As Boolean)
   
   Dim RowPla As Long
   RowPla = 0
   Dim Rowpla2 As Long
   Rowpla2 = row2 - 9
   
   If EstVacio Then
      
      RowPla = 9 'row1
      
      Do While RowPla <= vg_RowEnd + 8
   
         sheet.Range(ColPor & RowPla).Value = "=(" & col1 & RowPla & " * " & col2 & RowPla & ")/" & col1 & RowPla & ""
         sheet.Range(ColPor & RowPla & ":" & ColPor & RowPla).Select
         excel.Selection.NumberFormat = "0.00"
         excel.ActiveCell.FormulaR1C1 = "=IFERROR(((RC[-3] * RC[-1])/R[" & Rowpla2 & "]C[-3]),0)"
   
         RowPla = RowPla + 1
         Rowpla2 = Rowpla2 - 1
         
      Loop

   End If
   
   Rowpla2 = row2 - row1
   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
   
   If sheet.Range(ColFec & row1).Value <> "" Then
      
      sheet.Range(ColPor & row1).Value = "=(" & col1 & row1 & " * " & col2 & row1 & ")/" & col1 & row2 & ""
      sheet.Range(ColPor & row1 & ":" & ColPor & row2).Select
      excel.Selection.NumberFormat = "0.00"
      excel.ActiveCell.FormulaR1C1 = "=IFERROR(((RC[-3] * RC[-1])/R[" & Rowpla2 & "]C[-3]),0)"
   
   End If
  
End Function

Function MoverDatosExcelClave(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
   Dim ArregloClave() As String
   Dim i              As Long
   Dim RowPla         As Long
   RowPla = 0
   RowPla = 9 'row1 + 1
   
   ArregloClave = Split(datos, ";")
   i = 1 'ArregloClave(5)
   Do While RowPla <= vg_RowEnd + 8
   
      sheet.Range(col1 & RowPla).Value = ArregloClave(0) & ";" & ArregloClave(1) & ";" & ArregloClave(2) & ";" & ArregloClave(3) & ";" & ArregloClave(4) & ";" & i

      RowPla = RowPla + 1
      i = i + 1
         
   Loop

End Function

Function MoverDatosExcelValorNumerico(excel As Object, sheet As Object, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String, EstVacio As Boolean)
   
   Dim RowPla As Long
   RowPla = 0
   RowPla = 9 'row1 + 1
   
   If EstVacio Then
   
      Do While RowPla <= vg_RowEnd + 8
   
         sheet.Range(col1 & RowPla).Value = 0
   
         RowPla = RowPla + 1
         
      Loop

   End If
   
   sheet.Range(col1 & row1).Value = datos
   
End Function

Function MoverDatosExcelFormulaIII(excel As Object, sheet As Object, ColFec As String, ColPor As String, col1 As String, col2 As String, row1 As Long, row2 As Long, datos As String)
   
   sheet.Range(ColFec & row1 & ":" & ColFec & row2).Select
   
   If sheet.Range(ColFec & row1).Value <> "" Then
      
      sheet.Range(ColPor & row1).Value = "=(" & col1 & row1 & " * " & col2 & row1 & ")/" & col1 & row2 & ""
      sheet.Range(ColPor & row1 & ":" & ColPor & row2).Select
      excel.Selection.NumberFormat = "0.00"
   
   End If

End Function

Function MoverDatosExcelBuscarVI(excel As Object, sheet As Object, ColRec As String, col1 As String, row1 As Long, row2 As Long)
   
On Error GoTo Man_Error

sheet.Range(col1 & row1).Select
excel.ActiveCell.FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-3],Recetas!R2C2:R[" & row2 & "]C[-2],2,FALSE),0)" '"=VLOOKUP(RC[-3],Hoja1!RC:R[-3]C[-1],2,FALSE)" '"=VLOOKUP(RC[-3],Hoja1!R[-7]C[-117]:R[9019]C[-116],2,FALSE)" '"=BUSCARV(" & ColRec & row1 & ";recetas!B2:C45028;2;FALSO)"
'excel.ActiveCell.Value = "=BUSCARV(" & ColRec & row1 & ";Recetas!B2:C45028;2;FALSO)"

'ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],Hoja1!RC:R[1]C[1],2,FALSE)"
'"=VLOOKUP(RC[-3],Recetas!R[-7]C[-117]:R[9019]C[-116],2,FALSE)"
'=BUSCARV(B9;Recetas!B2:C9028;2;FALSO)
'"=SUM(R[" & row1 - 7 & "]" & "C" & ":R[" & row2 - 7 & "]" & "C" & ")"

'excel.Selection.NumberFormat = "0.00"
sheet.Range(col1 & row1).Select

Exit Function
Man_Error:
Resume Next

End Function

Function ValidarCampo(op As String) As Boolean

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

ValidarCampo = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_Parametros '" & op & "'")

If Not RS.EOF Then
    
   If UCase(RS(0)) = "S" Then
   
      ValidarCampo = True
   
   ElseIf UCase(RS(0)) = "N" Then
   
      ValidarCampo = False
   
   End If

End If
RS.Close
Set RS = Nothing

Exit Function
Man_Error:
Resume Next

End Function
