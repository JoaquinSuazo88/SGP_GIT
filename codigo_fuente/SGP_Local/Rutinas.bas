Attribute VB_Name = "Rutinas"
Option Explicit
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Private Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wmsg As Long, ByVal wparam As Long, lparam As Any) As Long
Const WM_PASTE = &H302

Public Function fg_BuscaenArbol(codigo As Long, Tabla As String, CampoBus As String) As String
Dim Nombre As String
Dim i As Long
Dim RS1 As New ADODB.Recordset
'-------> Buscar raiz en un TreeView
Nombre = ""
For i = 1 To 5
    If codigo = 0 Then Exit For
    RS1.Open "SELECT * FROM " & Tabla & " WHERE " & CampoBus & " = " & codigo & "", vg_db, adOpenForwardOnly
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit For
    Nombre = RS1(1) & "\" & Nombre
    codigo = IIf(IsNull(RS1(2)), 0, RS1(2))
    If RS1(0) = 0 Then RS1.Close: Set RS1 = Nothing: Exit For
    RS1.Close: Set RS1 = Nothing
Next
If Trim(Nombre) <> "" Then fg_BuscaenArbol = Mid(Nombre, 1, Len(Nombre) - 1) Else fg_BuscaenArbol = ""
End Function

Public Function fg_BuscaenArbolNivel2(codigo As Long, Tabla As String, CampoBus As String) As Long
Dim i As Long
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
'-------> Buscar raiz en un TreeView
fg_BuscaenArbolNivel2 = 0
For i = 1 To 5
    If codigo = 0 Then Exit For
    RS1.Open "SELECT * FROM " & Tabla & " WHERE " & CampoBus & " = " & codigo & "", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit For
    codigo = RS1(2)
    RS2.Open "SELECT * FROM b_paramdesp WHERE pad_codigo=" & codigo & "", vg_db, adOpenStatic
    If Not RS2.EOF Then fg_BuscaenArbolNivel2 = codigo: RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: Exit For
    If RS1(0) = 0 Then RS2.Close: Set RS2 = Nothing: RS1.Close: Set RS1 = Nothing: Exit For
    RS1.Close: Set RS1 = Nothing
    RS2.Close: Set RS2 = Nothing
Next
End Function

Sub fg_centra(frm As Form)
frm.Top = ((Screen.Height - IIf(frm.MDIChild, 1100, 0)) \ 2 - frm.Height \ 2)
frm.Left = Screen.Width \ 2 - frm.Width \ 2
'    frm.Top = Screen.Height \ 2 - frm.Height \ 2
'    frm.Left = Screen.Width \ 2 - frm.Width \ 2
End Sub

Sub validar_respuesta(Title As String)
Dim msg As String, Style, help, Ctxt
msg = "               Esta Seguro ?"
Style = vbYesNoCancel + vbQuestion + vbDefaultButton2
help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, Title, help, Ctxt)
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

If Not IsNull(numero) Then

Do While InStr(numero, Caracter) > 0
    
    numero = Mid(numero, 1, InStr(numero, Caracter) - 1) & Nuevo & Mid(numero, InStr(numero, Caracter) + 1, Len(numero))

Loop

Else
   
   numero = ""

End If

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

Function fg_CreaRS(Base As Connection, Sql As String, Tipo As Integer, modo As Integer) As Recordset

If modo = dbSQLPassThrough Then
    
    Clipboard.SetText Sql
    Set fg_CreaRS = Base.OpenRecordset(Sql, Tipo) ', Modo)

Else
    
    Set fg_CreaRS = Base.OpenRecordset(Sql, Tipo)

End If

End Function

Function AbrirBase()

On Error GoTo Man_Error

Dim cDrv As String
Set vg_db = New ADODB.Connection

If vg_tipbase = "1" Then
   
   vg_db.ConnectionString = "Provider='" + LTrim(RTrim(Provider)) + "';Data Source= '" + dir_trabajo + BaseDeDato + "' ;Persist Security Info=False" ';Jet OLEDB:Database Password='jpaz'"

Else
   
   vg_db.ConnectionString = "PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";OLE DB Services = -2;uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";"
 
End If
vg_db.ConnectionTimeout = 15 '3600
vg_db.CommandTimeout = 600 '3600
vg_db.Open

Exit Function
Man_Error:
    MsgBox Err.Number & ":  " & Err.Description, vbCritical, "Mantención sistema SGP"
    End

'If Err.Number = -2147467259 Then cDrv = "{Microsoft ODBC para SqlServer}": Resume
'If Err.Number = -2147467259 Then MsgBox Err & ":  " & "El sistema esta respaldando información por otro usuario. " & Chr(13) & Chr(13) & "Inténtelo en unos minutos mas tarde.", vbExclamation + vbOKOnly, "Mantención sistema SGP": End
'If Err.Number = -2147217843 Or Err.Number = -2147467259 Then MsgBox Err & ":  " & Error$(Err), vbCritical, "Mantención sistema SGP": End
End Function

Function fg_Fecha_Escrita(Fecha As String) As String
'Escribe la fecha en el siguiente formato:
'Jueves 16 de Junio de 1994
    Dim MiFecha$, Dia%, mes%, Año%, mes1$, FechaP$, diasem%
    
    If Fecha = "" Then Exit Function

    MiFecha = Format$(Fecha, "mmm dd yyyy")
    Dia = Day(DateValue(MiFecha))
    diasem = Weekday(DateValue(MiFecha))
    mes = Month(DateValue(MiFecha))
    Año = Year(DateValue(MiFecha))
    FechaP = ""
    Select Case diasem
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

    FechaP = FechaP + " " + Format$(Dia) + " de "
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
VerConfReg
vg_CSep = ","
vg_CDec = "."
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
'Escribe la fecha en el siguiente formato:
'Jueves 16 de Junio de 1994
    Dim MiFecha$, Dia%, FechaP$, diasem%
    If Fecha = "" Then Exit Function
    MiFecha = Format$(fg_Ctod1(Fecha), "mmm dd yyyy")
    Dia = Day(DateValue(MiFecha))
    diasem = Weekday(DateValue(MiFecha))
    FechaP = ""
    Select Case opcion
      Case 1
        Select Case diasem
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
        Select Case diasem
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
    FechaP = FechaP + " " + Format$(Dia)
    fg_Fecha_Dia = FechaP
End Function

Function fg_Fecha_Dia1(Fecha As String, opcion As Integer) As String
'Escribe la fecha en el siguiente formato:
'Jueves 16 de Junio de 1994
    Dim MiFecha$, Dia%, FechaP$, diasem%
    If Fecha = "" Then Exit Function
    MiFecha = Format$(fg_Ctod1(Fecha), "mmm dd yyyy")
    Dia = Day(DateValue(MiFecha))
    diasem = Weekday(DateValue(MiFecha))
    FechaP = ""
    Select Case opcion
      Case 1
        Select Case diasem
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
        Select Case diasem
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
    FechaP = FechaP + " " + fg_pone_cero(Format$(Dia), 2)
    fg_Fecha_Dia1 = FechaP
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
    If Dir(dir_trabajo_Inf & "txt" & fg_pone_cero(Trim(Str(i)), 5) & ".txt") = "" Then
        fg_ArchivoTxt = dir_trabajo_Inf & "txt" & fg_pone_cero(Trim(Str(i)), 5) & ".txt": Exit Function
    End If
Next i
End Function

Function fg_ArchivoRtf()
Dim i As Long
i = 1
For i = 1 To 99999
    If Dir(dir_trabajo_Inf & "Reporte" & fg_pone_cero(Trim(Str(i)), 5) & ".rtf") = "" Then
        fg_ArchivoRtf = dir_trabajo_Inf & "Reporte" & fg_pone_cero(Trim(Str(i)), 5) & ".rtf": Exit Function
    End If
Next i
End Function

Function fg_ArchivoXls(NombreArchivo As String) As String

Dim i As Long

i = 1

'-------> Crear directorio ExcelSGP
If Dir(dir_trabajo_Inf & "\" & "ExcelSGP", vbDirectory) = "" Then MkDir dir_trabajo_Inf & "\" & "ExcelSGP"
'-------> Fin crear directorio Excel Versión

For i = 1 To 99999
    
    If Dir(dir_trabajo_Inf & "\" & "ExcelSGP\" & NombreArchivo & fg_pone_cero(Trim(Str(i)), 5) & ".Xls") = "" Then
        
        fg_ArchivoXls = dir_trabajo_Inf & "ExcelSGP\" & NombreArchivo & fg_pone_cero(Trim(Str(i)), 5) & ".Xls": Exit Function
    
    End If

Next i

End Function

Function fg_pone_cero(ByVal cadena As String, ByVal cuanto As Integer) As String
'pone ceros a la izquierda
fg_pone_cero = ""
If cadena <> "" Then
   
   Do While Len(Trim(cadena)) < cuanto
      
      cadena = "0" + Trim(cadena)
   
   Loop
   fg_pone_cero = Trim(cadena)

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
Dim msg As String, Style, help, Ctxt
msg = "Cancelar Operación ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
help = "DEMO.HLP"
Ctxt = 1000
'respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
respuesta = MsgBox(msg, Style, Title) ', Help, Ctxt)
End Function

Function Resp_Delete(Title As String)
Dim msg As String, Style, help, Ctxt
msg = "Confirma Eliminar ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
help = "DEMO.HLP"
Ctxt = 1000
respuesta = MsgBox(msg, Style, Title, help, Ctxt)
End Function

Function Resp_Habilitar(Title As String)
Dim msg As String, Style, help, Ctxt
msg = "Todos las opciones subordinadas al selecionar serán habilidata. ¿ Cancelar Operación ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
help = "DEMO.HLP"
Ctxt = 1000
respuesta = MsgBox(msg, Style, Title)
End Function

Function Resp_Deshabilitar(Title As String)
Dim msg As String, Style, help, Ctxt
msg = "Todos las opciones subordinadas serán deshabilidata. ¿ Cancelar Operación ? "
Style = vbYesNo + vbQuestion + vbDefaultButton2
help = "DEMO.HLP"
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
For i = 0 To Combo(Index).listcount - 1
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

For i = 0 To Combo(Index).listcount - 1
    
    Combo(Index).ListIndex = i
    
    If Mid(Trim(Combo(Index).text), Len(Trim(Combo(Index).text)) - Largo, Largo) = cBusca Then
        
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

Function fg_codigolistaNuevo(List As String, Index As Integer, Largo As Integer, cDefa As Variant)

If Len(Trim(List)) > Largo Then
    fg_codigolistaNuevo = Mid(Trim(List), Len(Trim(List)) + 1 - Largo, Largo)
Else
    fg_codigolistaNuevo = cDefa
End If

End Function

Function fg_Busca_En_Lista(List As Control, dato As String, Largo As Integer) As Integer
Dim i%
fg_Busca_En_Lista = 0
For i = 0 To List.listcount - 1
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
    Dim MiFecha$, Dia%, FechaP$, diasem%
    If Fecha = "" Then Exit Function
    MiFecha = Format$(fg_Ctod1(Fecha), "mmm dd yyyy")
    Dia = Day(DateValue(MiFecha))
    diasem = Weekday(DateValue(MiFecha))
    fg_Dia = diasem
End Function

Function fg_NomDia(Dia As Long) As String
Select Case Dia
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

Function fg_NumDia(Dia As String) As Integer
Select Case Dia
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
    .X1 = 500: .X2 = Preview.P1.Width + .X1: .Y1 = 500: .Y2 = Preview.P1.Height + .Y1
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
        Dim ML!
        ML = vp.IndentLeft / 1440
        vp.ExportRaw = vbCrLf & _
            "<img style='margin-left:" & ML & "in;' src=" & sPicFile & ">" & vbCrLf
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
    sendmessage Preview.rtfPic.hWnd, WM_PASTE, 0, 0
    Preview.rtfPic.SelIndent = vp.IndentLeft
    vp.ExportRaw = vbCrLf & Preview.rtfPic.TextRTF & vbCrLf
    Clipboard.Clear
End Sub

Sub Gl_Mo_Botones(Form As Object, op As Integer)

Dim BtnX As Object

Select Case op

Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 2 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR DATO-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Datos"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 3 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-COPIAR DATO-FILTRAR-BUSQUEDA-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
'    Set btnX = Form.Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDropdown, "A_CopiarD"): btnX.Visible = True: btnX.ToolTipText = "Copiar Receta Patrón" ': btnX.ButtonMenus.Add Text:="Copiar Recetas" ': btnX.ButtonMenus.Add Text:="Pegar Recetas": btnX.ButtonMenus.Add Text:="Mover Recetas"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Receta Patrón" ': btnX.ButtonMenus.Add Text:="Copiar Recetas" ': btnX.ButtonMenus.Add Text:="Pegar Recetas": btnX.ButtonMenus.Add Text:="Mover Recetas"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_BuscarPro", , tbrDefault, "A_BuscarPro"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar Productos"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 4 'INCLUIR-GRABAR-CANCELAR(ANULAR)-HISTORICO-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Anular "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "

Case 5 'GRABAR-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = "Grabar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "

Case 6 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-IMPRIMIR-EXPORTAR-IMPORTAR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Candado", , tbrDefault, "I_Candado"): BtnX.Visible = True: BtnX.ToolTipText = "Autorizado Ajuste"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.ToolTipText = "Enviar Inventario "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_NoEnviar", , tbrDefault, "A_NoEnviar"): BtnX.Visible = True: BtnX.ToolTipText = "Anular Envio Inventario "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Ajuste", , tbrDefault, "A_Ajuste"): BtnX.Visible = True: BtnX.ToolTipText = "Ajustar Inventario "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_AnulaAjuste", , tbrDefault, "A_AnulaAjuste"): BtnX.Visible = True: BtnX.ToolTipText = "Anular Ajuste Inventario "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_ExportarInventario", , tbrDefault, "A_ExportarInventario"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Inventario "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_ImportarInventario", , tbrDefault, "A_ImportarInventario"): BtnX.Visible = True: BtnX.ToolTipText = "Importar Inventario "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_GenerarArchivo", , tbrDefault, "A_GenerarArchivo"): BtnX.Visible = True: BtnX.ToolTipText = "Generar Inventario OPTIMUM "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_ExploradorWindows", , tbrDefault, "A_ExploradorWindows"): BtnX.Visible = True: BtnX.ToolTipText = "Ver Carpeta"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 7 'INCLUIR-BORRAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
'    Set btnX = Form.Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): btnX.Visible = True: btnX.ToolTipText = "Grabar "
'    Set btnX = Form.Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): btnX.Visible = False: btnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "

Case 8 'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
    
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
    Set BtnX = Form.Toolbar2.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir Ingrediente"
    Set BtnX = Form.Toolbar2.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""

Case 9 'INCLUIR(1)-ALTERAR(3)-BORRAR(5)-ACTUALIZAR(7)-CANCELAR(10)-CONFIRMAR(12)-SUBIR(15)-BAJAR(16)-IMPRIMIR(18)-SALIR(21)
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Gasto A13"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_SubirF", , tbrDefault, "A_SubirF"): BtnX.Visible = True: BtnX.ToolTipText = "Subir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): BtnX.Visible = True: BtnX.ToolTipText = "Bajar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 10 'CANDADO-LLAVE-ACTUALIZAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Candado", , tbrDefault, "I_Candado"): BtnX.Visible = True: BtnX.ToolTipText = "Cerrar Mes"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Llave", , tbrDefault, "I_Llave"): BtnX.Visible = False: BtnX.ToolTipText = "Abrir Mes"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 11
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Candado", , tbrDefault, "I_Candado"): BtnX.Visible = True: BtnX.ToolTipText = "Cerrar Venta del Día"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Llave", , tbrDefault, "I_Llave"): BtnX.Visible = False: BtnX.ToolTipText = "Reabrir Venta del Día"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 12 'INCLUIR-GRABAR-CANCELAR(ANULAR)-HISTORICO-IMPRIMIR-CERRAR SALIDA-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Anular "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Candado", , tbrDefault, "I_Candado"): BtnX.Visible = True: BtnX.ToolTipText = "Cerrar Salida"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "

Case 13 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-CANCELAR-CONFIRMAR-AYUDA-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Toma Pedido Pacientes "
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 14 'INCLUIR-GRABAR-CANCELAR(ANULAR)-HISTORICO-IMPRIMIR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Anular "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.ToolTipText = "Enviar Guia Venta SAP"
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir "

Case 19 '-------> INCLUIR-ALTERAR-BORRAR-SALIR
    
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Modificar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Case 21
        
    Form.Toolbar1.ImageList = Partida.IL1
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Incluir"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Alterar", , tbrDefault, "A_Alterar"): BtnX.Visible = True: BtnX.ToolTipText = "Alterar"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Alterar", , tbrDefault, "I_Alterar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "Borrar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Borrar ", , tbrDefault, "I_Borrar "): BtnX.Visible = False: BtnX.ToolTipText = ""
    
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Actualizar", , tbrDefault, "A_Actualizar"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Lista   "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Actualizar", , tbrDefault, "I_Actualizar"): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Cancelar ", , tbrDefault, "A_Cancelar "): BtnX.Visible = False: BtnX.ToolTipText = "Cancelar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Cancelar ", , tbrDefault, "I_Cancelar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDropdown, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir ": BtnX.ButtonMenus.Add text:="Imprimir Usuario Perfil": BtnX.ButtonMenus.Add text:="Transacciones Usuarios"
    Set BtnX = Form.Toolbar1.Buttons.Add(, "I_Imprimir ", , tbrDefault, "I_Imprimir "): BtnX.Visible = False: BtnX.ToolTipText = ""
    Set BtnX = Form.Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Form.Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

End Select

End Sub

Function Gl_Ac_Botones(Form As Form, op1 As Integer, op2 As Integer, modo As String)

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir As String, vtcest As Boolean, sql1 As String
vtcest = False
'-----------------------------VALIDAR USUARIO-----------------
'RS1.Open "SELECT dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
'         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
'         "INNER JOIN a_usuarios usu ON per.per_codigo = usu.usu_perfil " & _
'         "WHERE usu.usu_codigo = '" & vg_NUsr & "' AND dpe.dpe_codopc = " & Form.HelpContextID & "", vg_db, adOpenStatic

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgp_Sel_DerechosPerfil '" & vg_NUsr & "', " & Form.HelpContextID & "")

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
Select Case op1

Case 1
    
    Select Case op2
    
    Case 0 '-------> CANCELAR-CONFIRMAR
        
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
        
        End If
    
    Case 1 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        
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
    
    Case 2 '-------> INCLUIR
        
        If incluir = "1" Then
            
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        
        End If
        
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
    
    Case 3 '-------> Habilitar Solamente Salir
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
    
    Case 4 '-------> ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
    
    Case 5 '-------> ALTERAR-ACTUALIZAR-IMPRIMIR
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = False
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
    
    Case 6 '-------> IMPRIMIR
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
    
    Case 7 '-------> ALTERAR-IMPRIMIR-SALIR
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        If alterar = "1" Then Form.Toolbar1.Buttons(3).Visible = True: Form.Toolbar1.Buttons(4).Visible = False
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
    
    Case 8 '-------> ALTERAR-Actualizar-IMPRIMIR-SALIR
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
   
   Case 16 'INCLUIR-ALTERAR-BORRAR
        
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = False
        Form.Toolbar1.Buttons(6).Visible = True
'        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
'        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
    
    End Select

Case 2
    
    Select Case op2
    
    Case 0 '-------> CANCELAR-CONFIRMAR
        
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Enabled = False
            Form.Toolbar1.Buttons(17).Visible = False: Form.Toolbar1.Buttons(18).Visible = True
        
        End If
    
    Case 1 '-------> INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    
    Case 2 '-------> INCLUIR
        
        If incluir = "1" Then
            
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        
        End If
        
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False
        Form.Toolbar1.Buttons(17).Visible = False: Form.Toolbar1.Buttons(18).Visible = True
    
    Case 3 '-------> Habilitar Solamente Salir
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False
        Form.Toolbar1.Buttons(17).Visible = False: Form.Toolbar1.Buttons(18).Visible = True
    
    End Select

Case 3
    
    Select Case op2
    
    Case 0 'CANCELAR-CONFIRMAR
        
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Enabled = False
            Form.Toolbar1.Buttons(17).Enabled = False
            Form.Toolbar1.Buttons(19).Enabled = False
            Form.Toolbar1.Buttons(21).Visible = False: Form.Toolbar1.Buttons(22).Visible = True
        
        End If
    
    Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(17).Enabled = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(19).Enabled = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(21).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(22).Visible = IIf(imprimir = "1", False, True)
    
    Case 2 'INCLUIR
        
        If incluir = "1" Then
            
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        
        End If
        
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False
        Form.Toolbar1.Buttons(17).Enabled = True
        Form.Toolbar1.Buttons(19).Enabled = False
        Form.Toolbar1.Buttons(21).Visible = False: Form.Toolbar1.Buttons(22).Visible = True
    
    Case 3 'Habilitar Solamente Salir
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False
        Form.Toolbar1.Buttons(17).Enabled = False
        Form.Toolbar1.Buttons(19).Enabled = False
        Form.Toolbar1.Buttons(21).Visible = False: Form.Toolbar1.Buttons(22).Visible = True
     
     Case 4 'ALTERAR-ACTUALIZAR-IMPRIMIR
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1" And ("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Or vg_5etapas = False), True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1" And ("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Or vg_5etapas = False), False, True)
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
'        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(7).Visible = IIf(vg_newcodrec = 0, True, False): Form.Toolbar1.Buttons(8).Visible = IIf(vg_newcodrec = 0, False, True)
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        
'        Form.Toolbar1.Buttons(15).Enabled = IIf(("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Or Not vg_5etapas Or vg_tiprec = 0), True, False)
        Form.Toolbar1.Buttons(15).Enabled = IIf(("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Or Not vg_5etapas Or vg_tiprec = 0 Or vg_newcodrec > 0), True, False)
        Form.Toolbar1.Buttons(17).Enabled = IIf(vg_newcodrec = 0, True, False)
        Form.Toolbar1.Buttons(19).Enabled = False
        Form.Toolbar1.Buttons(21).Visible = IIf(imprimir = "1" And vg_newcodrec = 0, True, False)
        Form.Toolbar1.Buttons(22).Visible = IIf(imprimir = "1", False, True)
    
    End Select

Case 4
    
    Form.Toolbar1.Refresh
    
    Select Case op2
    
    Case 1 'Ninguno
        
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(9).Visible = False: Form.Toolbar1.Buttons(10).Visible = True
        Form.Toolbar1.Buttons(11).Visible = True: Form.Toolbar1.Buttons(12).Visible = True
    
    Case 2, 5 'Grabar
        
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Enabled = True: Form.Toolbar1.Buttons(11).ToolTipText = "Buscar "
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
    
    Case 3, 4 'Anular - Imprimir
        
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = IIf(op2 = 3, IIf(eliminar = 1, True, False), False)
        Form.Toolbar1.Buttons(4).Visible = IIf(op2 = 3, IIf(eliminar = 0, True, False), True)
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Enabled = True: Form.Toolbar1.Buttons(11).ToolTipText = "Buscar "
        Form.Toolbar1.Buttons(12).Visible = IIf(imprimir = 1, True, False)
        Form.Toolbar1.Buttons(13).Visible = IIf(imprimir = 0, True, False)
    
    Case 6
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = True: Form.Toolbar1.Buttons(7).Visible = False
        Form.Toolbar1.Buttons(8).Visible = True: Form.Toolbar1.Buttons(9).Visible = False
        Form.Toolbar1.Buttons(11).Enabled = False: Form.Toolbar1.Buttons(11).ToolTipText = ""
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
    
    End Select

Case 5
    
    Form.Toolbar1.Refresh
    Select Case op2
    
    Case 1 'Grabar
        
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
    
    Case 2 'Imprimir
        
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = True: Form.Toolbar1.Buttons(4).Visible = False
    
    End Select

Case 6
    
    Dim Fecha As String
    sql1 = IIf(vg_tipbase = "1", " Cdate('" & IIf(Form.Date1(0).text = "", 0, Form.Date1(0).text) & "') ", " '" & IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")) & "' ")
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT COUNT(tov_fecemi) AS suma FROM b_totventas " & _
             "WHERE tov_fecemi = " & sql1 & " " & _
             "AND tov_codbod = " & Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)) & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS2.Open "SELECT MAX(tin_fectom) AS fecha FROM b_tomainv WHERE tin_codbod = " & Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)), vg_db, adOpenStatic
    If Not RS2.EOF Then Fecha = fg_Ctod1(RS2!Fecha) Else Fecha = Form.Date1(0).text
    Dim auxhelp As Long, envio As Boolean, autaju As Boolean
    '-------> Traer autorización ajuste
    autaju = True
    
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS3.Open "SELECT DISTINCT tin_autaju FROM b_tomainv WHERE tin_fectom = " & IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")) & " AND tin_codbod = " & vg_codbod & "", vg_db, adOpenStatic
    If Not RS3.EOF Then autaju = IIf(IsNull(RS3!tin_autaju) Or RS3!tin_autaju = "0" Or Trim(RS3!tin_autaju) = "", False, True)
    RS3.Close: Set RS3 = Nothing
    '-------> Validar si toma fue enviada
    envio = True
    
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS3.Open "SELECT DISTINCT cencos FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso = '2' AND num_documento = '" & IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")) & "' AND estado = '1' AND (anulado) IS NULL", vg_db, adOpenStatic
    If Not RS3.EOF Then envio = False
    RS3.Close: Set RS3 = Nothing
    '-------> Validar si toma fue anulada
    
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS3.Open "SELECT DISTINCT cencos FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso = '3' AND num_documento = '" & IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")) & "' AND estado = '1' AND (anulado) IS NULL", vg_db, adOpenStatic
    If Not RS3.EOF Then envio = True
    RS3.Close: Set RS3 = Nothing
    auxhelp = Form.HelpContextID
    Form.Toolbar1.Buttons(21).Enabled = False
    '-------> Validar si no hay ingresado stock fisico
    
    If RS3.State = 1 Then RS3.Close
    RS3.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS3.Open "SELECT DISTINCT tin_fectom FROM b_tomainv " & _
              "WHERE  tin_codbod = " & vg_codbod & " " & _
              "AND    tin_fectom = " & IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")) & " " & _
              "AND    round(tin_stofis,2) <> 0 AND round(tin_propon,2) <> 0", vg_db, adOpenStatic
    If Not RS3.EOF And Not CierrePeriodo(IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")), vg_codbod, 0) Then
       '-------> Activar opción autorización
       Form.HelpContextID = 2079010
       Form.Toolbar1.Buttons(21).Enabled = IIf(autaju Or modo = "A" Or modo = "M" Or RS1!Suma = 0 Or Fecha <> IIf(Form.Date1(0).text = "", 0, Form.Date1(0).text) Or Mid(ValidarUsuarioAcceso(M_TomInv), 1, 1) = "0" Or Not envio, False, True)
    End If
    RS3.Close: Set RS3 = Nothing
    
    '-------> Activar opciones envio
    Form.HelpContextID = 2079100
    Form.Toolbar1.Buttons(23).Enabled = False
    Form.Toolbar1.Buttons(24).Enabled = False
    If (ValidarOpEnvio(MuestraCasino(1), 2) Or ValidarOpEnvio(MuestraCasino(1), 5)) And Not CierrePeriodo(IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")), vg_codbod, 16) _
       And Not CierrePeriodo(IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")), vg_codbod, 7) And Not CierrePeriodo(IIf(Form.Date1(0).text = "", 0, Format(Form.Date1(0).text, "yyyymmdd")), vg_codbod, 0) Then
       
       Form.Toolbar1.Buttons(23).Enabled = IIf(Not autaju Or modo = "A" Or modo = "M" Or Mid(ValidarUsuarioAcceso(M_TomInv), 1, 1) = "0" Or Not envio, False, True)
       Form.Toolbar1.Buttons(24).Enabled = IIf((Not autaju And Not envio) Or autaju And (modo <> "A" And modo <> "M" And (Form.Toolbar1.Buttons(23).Enabled = False)), True, False)
    
    End If
    
    Form.HelpContextID = auxhelp
    Form.Toolbar1.Buttons(27).Enabled = IIf(modo = "A" Or modo = "M" Or RS1!Suma = 0 Or Fecha <> Form.Date1(0).text Or (Not envio And ValidarOpEnvio(MuestraCasino(1), 2)), False, True)
    Form.Toolbar1.Buttons(26).Enabled = IIf(modo = "A" Or modo = "M", False, True)
    Form.Toolbar1.Buttons(19).Enabled = False
    Form.Toolbar1.Buttons(18).Enabled = True
    RS1.Close: Set RS1 = Nothing
    RS2.Close: Set RS2 = Nothing
    Select Case op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
            Form.Toolbar1.Buttons(18).Enabled = IIf(modo = "A", False, True)
            Form.Toolbar1.Buttons(19).Enabled = IIf(modo = "A", True, False)
            Form.Toolbar1.Buttons(29).Enabled = False
            Form.Toolbar1.Buttons(30).Enabled = False
            Form.Toolbar1.Buttons(32).Enabled = False
            Form.Toolbar1.Buttons(34).Enabled = False
        End If
    Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1" And (Not ValidarInventarioRotativo(MuestraCasino(1)) Or CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29)), True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1" And (Not ValidarInventarioRotativo(MuestraCasino(1)) Or CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29)), False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1" And Form.Toolbar1.Buttons(27).Enabled = False And Form.Toolbar1.Buttons(24).Enabled = False, True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1" And Form.Toolbar1.Buttons(27).Enabled = False And Form.Toolbar1.Buttons(24).Enabled = False, False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1" And Form.Toolbar1.Buttons(24).Enabled = False, True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1" And Form.Toolbar1.Buttons(24).Enabled = False, False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
        Form.Toolbar1.Buttons(29).Enabled = True
        Form.Toolbar1.Buttons(30).Enabled = IIf(Form.Date1(0).text = (CDate(vg_ciedia) - 1) And Form.Toolbar1.Buttons(27).Enabled = False, True, False)
        Form.Toolbar1.Buttons(32).Enabled = True
        Form.Toolbar1.Buttons(34).Enabled = True
    Case 2 'INCLUIR
        If incluir = "1" Then
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        End If
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
    Case 3 'Habilitar Solamente Salir
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
    Case 4 'ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
    Case 5 'DESACTIVAR TODO
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = False
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = False
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = False
            Form.Toolbar1.Buttons(9).Visible = False: Form.Toolbar1.Buttons(10).Visible = False
            Form.Toolbar1.Buttons(11).Visible = False: Form.Toolbar1.Buttons(12).Visible = False
            Form.Toolbar1.Buttons(13).Visible = False: Form.Toolbar1.Buttons(14).Visible = False
            Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Enabled = False
            Form.Toolbar1.Buttons(17).Enabled = False: Form.Toolbar1.Buttons(18).Enabled = False
            Form.Toolbar1.Buttons(19).Enabled = False: Form.Toolbar1.Buttons(20).Enabled = False
            Form.Toolbar1.Buttons(21).Enabled = False: Form.Toolbar1.Buttons(22).Enabled = False
            Form.Toolbar1.Buttons(23).Enabled = False: Form.Toolbar1.Buttons(24).Enabled = False
            Form.Toolbar1.Buttons(25).Enabled = False: Form.Toolbar1.Buttons(26).Enabled = False
            Form.Toolbar1.Buttons(27).Enabled = False: Form.Toolbar1.Buttons(28).Enabled = False
            Form.Toolbar1.Buttons(29).Enabled = False
            Form.Toolbar1.Buttons(31).Enabled = False
            Form.Toolbar1.Buttons(32).Enabled = False
            Form.Toolbar1.Buttons(34).Enabled = False
        End If
    Case 6 'CANCELAR-CONFIRMAR-SALIR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = False
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = False
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = False
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = False
            Form.Toolbar1.Buttons(9).Visible = False: Form.Toolbar1.Buttons(10).Visible = True
            Form.Toolbar1.Buttons(11).Visible = False: Form.Toolbar1.Buttons(12).Visible = True
            Form.Toolbar1.Buttons(13).Visible = False: Form.Toolbar1.Buttons(14).Visible = False
            Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Enabled = False
            Form.Toolbar1.Buttons(17).Enabled = False: Form.Toolbar1.Buttons(18).Enabled = False
            Form.Toolbar1.Buttons(19).Enabled = False: Form.Toolbar1.Buttons(20).Enabled = False
            Form.Toolbar1.Buttons(21).Enabled = False: Form.Toolbar1.Buttons(22).Enabled = False
            Form.Toolbar1.Buttons(23).Enabled = False: Form.Toolbar1.Buttons(24).Enabled = False
            Form.Toolbar1.Buttons(25).Enabled = False: Form.Toolbar1.Buttons(26).Enabled = False
            Form.Toolbar1.Buttons(27).Enabled = False: Form.Toolbar1.Buttons(28).Enabled = False
            Form.Toolbar1.Buttons(29).Enabled = False
            Form.Toolbar1.Buttons(31).Enabled = False
            Form.Toolbar1.Buttons(32).Enabled = True
            Form.Toolbar1.Buttons(34).Enabled = True
        End If
    Case 7 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR-SALIR
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1" And Form.Toolbar1.Buttons(27).Enabled = False And Form.Toolbar1.Buttons(24).Enabled = False, True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1" And Form.Toolbar1.Buttons(27).Enabled = False And Form.Toolbar1.Buttons(24).Enabled = False, False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1" And Form.Toolbar1.Buttons(24).Enabled = False, True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1" And Form.Toolbar1.Buttons(24).Enabled = False, False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
        Form.Toolbar1.Buttons(29).Enabled = True
        Form.Toolbar1.Buttons(30).Enabled = IIf(Form.Date1(0).text = (CDate(vg_ciedia) - 1) And Form.Toolbar1.Buttons(27).Enabled = False, True, False)
        Form.Toolbar1.Buttons(32).Enabled = True
        Form.Toolbar1.Buttons(34).Enabled = True
    End Select
Case 7 'INCLUIR-BORRAR-CANCELAR-CONFIRMAR-IMPRIMIR-SALIR
    Form.Toolbar1.Refresh
    Select Case op2
    Case 1 'Incluir
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Visible = False: Form.Toolbar1.Buttons(12).Visible = True
        Form.Toolbar1.Buttons(14).Visible = True
    Case 2 'ELIMINAR - IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = IIf(eliminar = 1, True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(eliminar = 0, True, False)
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Visible = IIf(imprimir = 1, True, False)
        Form.Toolbar1.Buttons(12).Visible = IIf(imprimir = 0, True, False)
        Form.Toolbar1.Buttons(11).Visible = True
    Case 3 'CANCELAR - CONFIRMAR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = True: Form.Toolbar1.Buttons(7).Visible = False
        Form.Toolbar1.Buttons(8).Visible = True: Form.Toolbar1.Buttons(9).Visible = False
        Form.Toolbar1.Buttons(11).Visible = False: Form.Toolbar1.Buttons(12).Visible = True
        Form.Toolbar1.Buttons(14).Visible = True
    End Select
Case 8 'INCLUIR(1)-BORRAR(3)-CANCELAR(6)-CONFIRMAR(8)-BUSCAR(11)-IMPRIMIR(12)
    Form.Toolbar2.Refresh
    Select Case op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar2.Buttons(1).Visible = False: Form.Toolbar2.Buttons(2).Visible = True
            Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
            Form.Toolbar2.Buttons(6).Visible = True: Form.Toolbar2.Buttons(7).Visible = False
            Form.Toolbar2.Buttons(8).Visible = True: Form.Toolbar2.Buttons(9).Visible = False
            Form.Toolbar2.Buttons(11).Enabled = False
            Form.Toolbar2.Buttons(12).Visible = False: Form.Toolbar2.Buttons(13).Visible = True
        End If
    Case 1 'INCLUIR-BORRAR-BUSCAR-IMPRIMIR
        Form.Toolbar2.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar2.Buttons(2).Visible = IIf(incluir = "0", True, False)
        Form.Toolbar2.Buttons(3).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar2.Buttons(4).Visible = IIf(eliminar = "0", True, False)
        Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
        Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
        Form.Toolbar2.Buttons(11).Enabled = True
        Form.Toolbar2.Buttons(12).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar2.Buttons(13).Visible = IIf(imprimir = "0", True, False)
    Case 2 'INCLUIR
        If incluir = "1" Then
            Form.Toolbar2.Buttons(1).Visible = True: Form.Toolbar2.Buttons(2).Visible = False
        End If
        Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
        Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
        Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
        Form.Toolbar2.Buttons(11).Enabled = False
        Form.Toolbar2.Buttons(12).Visible = False: Form.Toolbar2.Buttons(13).Visible = True
    Case 3 'IMPRIMIR
        Form.Toolbar2.Buttons(1).Visible = False: Form.Toolbar2.Buttons(2).Visible = True
        Form.Toolbar2.Buttons(3).Visible = False: Form.Toolbar2.Buttons(4).Visible = True
        Form.Toolbar2.Buttons(6).Visible = False: Form.Toolbar2.Buttons(7).Visible = True
        Form.Toolbar2.Buttons(8).Visible = False: Form.Toolbar2.Buttons(9).Visible = True
        Form.Toolbar2.Buttons(11).Enabled = False
        Form.Toolbar2.Buttons(12).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar2.Buttons(13).Visible = IIf(imprimir = "0", True, False)
    End Select
Case 9
    'INCLUIR(1)-ALTERAR(3)-BORRAR(5)-ACTUALIZAR(7)-CANCELAR(10)-CONFIRMAR(12)-SUBIR(15)-BAJAR(16)-IMPRIMIR(18)-SALIR(21)
    Select Case op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Enabled = False:: Form.Toolbar1.Buttons(15).ToolTipText = ""
            Form.Toolbar1.Buttons(17).Enabled = False: Form.Toolbar1.Buttons(18).Enabled = False
            Form.Toolbar1.Buttons(20).Visible = False: Form.Toolbar1.Buttons(21).Visible = True
        End If
    Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Copiar Gasto A13"
        Form.Toolbar1.Buttons(17).Enabled = True: Form.Toolbar1.Buttons(18).Enabled = True
        Form.Toolbar1.Buttons(20).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(21).Visible = IIf(imprimir = "1", False, True)
    Case 2 'INCLUIR
        If incluir = "1" Then
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        End If
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(16).Enabled = False
        Form.Toolbar1.Buttons(18).Visible = False: Form.Toolbar1.Buttons(19).Visible = True
    End Select
Case 10
    Select Case op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Enabled = False: Form.Toolbar1.Buttons(2).Enabled = False
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(6).Visible = True: Form.Toolbar1.Buttons(7).Visible = False
            Form.Toolbar1.Buttons(8).Visible = True: Form.Toolbar1.Buttons(9).Visible = False
            Form.Toolbar1.Buttons(11).Visible = False: Form.Toolbar1.Buttons(12).Visible = True
        End If
    Case 5 'ALTERAR-ACTUALIZAR-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(1).Enabled = True: Form.Toolbar1.Buttons(2).Enabled = True
        Form.Toolbar1.Buttons(3).Visible = True: Form.Toolbar1.Buttons(4).Visible = False
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(12).Visible = IIf(imprimir = "1", False, True)
    Case 6 'CERRAR-REABRIR-ACTUALIZAR-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(1).Enabled = True: Form.Toolbar1.Buttons(2).Enabled = True
        Form.Toolbar1.Buttons(3).Visible = True: Form.Toolbar1.Buttons(4).Visible = False
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Visible = False: Form.Toolbar1.Buttons(12).Visible = True
    End Select
Case 11
    If Format(M_VenCaf.fpDateTime1(0).text, "dd/mm/yyyy") <> "" Then
       sql1 = IIf(vg_tipbase = "1", " cdate('" & Format(M_VenCaf.fpDateTime1(0).text, "dd/mm/yyyy") & "') ", " '" & Format(M_VenCaf.fpDateTime1(0).text, "yyyymmdd") & "' ")
       RS1.Open "SELECT * FROM b_totventascaf WHERE tvc_cencos = '" & Trim(LimpiaDato(M_VenCaf.fpText1(0).text)) & "' AND tvc_fecing = " & sql1 & " AND tvc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
       If Not RS1.EOF Then op2 = IIf(RS1!tvc_estado = "C", 4, op2): vtcest = True
       RS1.Close: Set RS1 = Nothing
    End If
    Select Case op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
            Form.Toolbar1.Buttons(18).Visible = True: Form.Toolbar1.Buttons(18).Enabled = False: Form.Toolbar1.Buttons(19).Visible = False
        End If
    Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
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
        Form.Toolbar1.Buttons(18).Visible = True: Form.Toolbar1.Buttons(18).Enabled = IIf(incluir = "1", True, False): Form.Toolbar1.Buttons(19).Visible = False
    Case 2 'INCLUIR
        If incluir = "1" Then
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        End If
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
        Form.Toolbar1.Buttons(18).Visible = True: Form.Toolbar1.Buttons(18).Enabled = False: Form.Toolbar1.Buttons(19).Visible = False
    Case 3 'Habilitar Solamente Salir
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = False: Form.Toolbar1.Buttons(16).Visible = True
        Form.Toolbar1.Buttons(18).Visible = True: Form.Toolbar1.Buttons(18).Enabled = False: Form.Toolbar1.Buttons(19).Visible = False
'        Form.Toolbar1.Buttons(18).Enabled = False
    Case 4 'Habilitar Imprimir y Salir
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(16).Visible = IIf(imprimir = "1", False, True)
        If CDate(M_VenCaf.fpDateTime1(0).text) < CDate(vg_ciedia) And Format(CDate(vg_ciedia), "mm/yyyy") = Format(CDate(M_VenCaf.fpDateTime1(0).text), "mm/yyyy") Then
           Form.Toolbar1.Buttons(18).Visible = True: Form.Toolbar1.Buttons(18).Enabled = False: Form.Toolbar1.Buttons(19).Visible = False
        Else
           Form.Toolbar1.Buttons(18).Visible = IIf(vtcest = True, False, True): Form.Toolbar1.Buttons(18).Enabled = IIf(vtcest = True, False, True): Form.Toolbar1.Buttons(19).Visible = IIf(vtcest = True, True, False)
        End If
    End Select
Case 12
    Form.Toolbar1.Refresh
    Select Case op2
    Case 1 'Ninguno
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(9).Visible = False: Form.Toolbar1.Buttons(10).Visible = True
        Form.Toolbar1.Buttons(11).Visible = True: Form.Toolbar1.Buttons(12).Visible = True
    Case 2, 5 'Grabar
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Enabled = True: Form.Toolbar1.Buttons(11).ToolTipText = "Buscar "
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
    Case 3, 4 'Anular - Imprimir
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = IIf(op2 = 3, IIf(eliminar = 1, True, False), False)
        Form.Toolbar1.Buttons(4).Visible = IIf(op2 = 3, IIf(eliminar = 0, True, False), True)
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Enabled = True: Form.Toolbar1.Buttons(11).ToolTipText = "Buscar "
        Form.Toolbar1.Buttons(12).Visible = IIf(imprimir = 1, True, False)
        Form.Toolbar1.Buttons(13).Visible = IIf(imprimir = 0, True, False)
        RS1.Open "SELECT dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
                 "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
                 "INNER JOIN a_usuarios usu ON per.per_codigo = usu.usu_perfil " & _
                 "WHERE usu.usu_codigo='" & vg_NUsr & "' AND dpe.dpe_codopc=2032000", vg_db, adOpenStatic
        If Not RS1.EOF Then
           Form.Toolbar1.Buttons(15).Enabled = IIf(op2 = 3, IIf(RS1!dpe_deragr = 1, True, False), False): Form.Toolbar1.Buttons(15).ToolTipText = "Cerrar Salida"
        End If
        RS1.Close: Set RS1 = Nothing
    Case 6
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = True: Form.Toolbar1.Buttons(7).Visible = False
        Form.Toolbar1.Buttons(8).Visible = True: Form.Toolbar1.Buttons(9).Visible = False
        Form.Toolbar1.Buttons(11).Enabled = False: Form.Toolbar1.Buttons(11).ToolTipText = ""
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
    End Select
Case 13
    Select Case op2
    Case 0 'CANCELAR-CONFIRMAR
        If (incluir = "1" And modo = "A") Or (alterar = "1" And modo = "M") Then
            Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
            Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
            Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
            Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
            Form.Toolbar1.Buttons(10).Visible = True: Form.Toolbar1.Buttons(11).Visible = False
            Form.Toolbar1.Buttons(12).Visible = True: Form.Toolbar1.Buttons(13).Visible = False
            Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
            Form.Toolbar1.Buttons(17).Visible = False: Form.Toolbar1.Buttons(18).Visible = True
        End If
    Case 1 'INCLUIR-ALTERAR-BORRAR-ACTUALIZAR-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", False, True)
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", False, True)
    Case 2 'INCLUIR
        If incluir = "1" Then
            Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        End If
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
        Form.Toolbar1.Buttons(17).Visible = False: Form.Toolbar1.Buttons(18).Visible = True
    Case 3 'Habilitar Solamente Salir
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(17).Visible = False: Form.Toolbar1.Buttons(18).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
    Case 4 'ALTERAR-BORRAR-ACTUALIZAR-HISTORICO-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = IIf(eliminar = "1", True, False)
        Form.Toolbar1.Buttons(6).Visible = IIf(eliminar = "1", False, True)
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    Case 5 'ALTERAR-ACTUALIZAR-HISTORICO-IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = IIf(alterar = "1", True, False)
        Form.Toolbar1.Buttons(4).Visible = IIf(alterar = "1", False, True)
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = False
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    Case 6 'IMPRIMIR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    Case 7 'ALTERAR-HISTORICO-IMPRIMIR-SALIR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        If alterar = "1" Then Form.Toolbar1.Buttons(3).Visible = True: Form.Toolbar1.Buttons(4).Visible = False
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    Case 8 'INCLUIR-ACTUALIZAR-HISTORICO-IMPRIMIR-SALIR
        Form.Toolbar1.Buttons(1).Visible = IIf(incluir = "1", True, False): Form.Toolbar1.Buttons(2).Visible = IIf(incluir = "1", True, False)
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = False
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    Case 9 'HISTORICO-IMPRIMIR-SALIR
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = False: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(10).Visible = False: Form.Toolbar1.Buttons(11).Visible = True
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = True: Form.Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
        Form.Toolbar1.Buttons(17).Visible = IIf(imprimir = "1", True, False)
        Form.Toolbar1.Buttons(18).Visible = IIf(imprimir = "1", False, True)
    End Select
Case 14
    Form.Toolbar1.Refresh
    Select Case op2
    Case 1 '-------> Ninguno
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(5).Visible = False: Form.Toolbar1.Buttons(6).Visible = True
        Form.Toolbar1.Buttons(7).Visible = True: Form.Toolbar1.Buttons(8).Visible = True
        Form.Toolbar1.Buttons(9).Visible = False: Form.Toolbar1.Buttons(10).Visible = True
        Form.Toolbar1.Buttons(11).Visible = True: Form.Toolbar1.Buttons(12).Visible = True
    Case 2, 5 '-------> Grabar
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Enabled = True: Form.Toolbar1.Buttons(11).ToolTipText = "Buscar "
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
        Form.Toolbar1.Buttons(15).Enabled = False: Form.Toolbar1.Buttons(15).ToolTipText = ""
    Case 3, 4 '-------> Anular - Imprimir
        Form.Toolbar1.Buttons(1).Visible = True: Form.Toolbar1.Buttons(2).Visible = False
        Form.Toolbar1.Buttons(3).Visible = IIf(op2 = 3, IIf(eliminar = 1, True, False), False)
        Form.Toolbar1.Buttons(4).Visible = IIf(op2 = 3, IIf(eliminar = 0, True, False), True)
        Form.Toolbar1.Buttons(6).Visible = False: Form.Toolbar1.Buttons(7).Visible = True
        Form.Toolbar1.Buttons(8).Visible = False: Form.Toolbar1.Buttons(9).Visible = True
        Form.Toolbar1.Buttons(11).Enabled = True: Form.Toolbar1.Buttons(11).ToolTipText = "Buscar "
        Form.Toolbar1.Buttons(12).Visible = IIf(imprimir = 1, True, False)
        Form.Toolbar1.Buttons(13).Visible = IIf(imprimir = 0, True, False)
    Case 6 '------->
        Form.Toolbar1.Buttons(1).Visible = False: Form.Toolbar1.Buttons(2).Visible = True
        Form.Toolbar1.Buttons(3).Visible = False: Form.Toolbar1.Buttons(4).Visible = True
        Form.Toolbar1.Buttons(6).Visible = True: Form.Toolbar1.Buttons(7).Visible = False
        Form.Toolbar1.Buttons(8).Visible = True: Form.Toolbar1.Buttons(9).Visible = False
        Form.Toolbar1.Buttons(11).Enabled = False: Form.Toolbar1.Buttons(11).ToolTipText = ""
        Form.Toolbar1.Buttons(12).Visible = False: Form.Toolbar1.Buttons(13).Visible = True
    End Select
End Select
End Function

Function ValidarUsuario(Formu As Form) As String
Dim RS1 As New ADODB.Recordset
Dim incluir As String, alterar As String, eliminar As String, imprimir
'-----------------------------VALIDAR USUARIO-----------------

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
         "INNER JOIN a_usuarios usu ON per.per_codigo = usu.usu_perfil " & _
         "WHERE usu.usu_codigo='" & vg_NUsr & "' and dpe.dpe_codopc=" & Formu.HelpContextID, vg_db, adOpenStatic
ValidarUsuario = "0000"
If Not RS1.EOF Then
    incluir = Trim(RS1!dpe_deragr)
    alterar = Trim(RS1!dpe_dermod)
    eliminar = Trim(RS1!dpe_dereli)
    imprimir = Trim(RS1!dpe_derimp)
    ValidarUsuario = incluir & alterar & eliminar & imprimir
End If
RS1.Close: Set RS1 = Nothing
'--------------------------------------------------------------
End Function

Function ValidarUsuarioAcceso(Formu As Form) As String
Dim RS1 As New ADODB.Recordset
Dim acceso As String, incluir As String, alterar As String, eliminar As String, imprimir
'-----------------------------VALIDAR USUARIO-----------------

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT dpe.dpe_deracc, dpe.dpe_deragr, dpe.dpe_dermod, dpe.dpe_dereli, dpe.dpe_derimp " & _
         "FROM (a_perfil per INNER JOIN a_derechosperfil dpe ON per.per_codigo = dpe.dpe_codper) " & _
         "INNER JOIN a_usuarios usu ON per.per_codigo = usu.usu_perfil " & _
         "WHERE usu.usu_codigo='" & vg_NUsr & "' and dpe.dpe_codopc=" & Formu.HelpContextID, vg_db, adOpenStatic
ValidarUsuarioAcceso = "00000"
If Not RS1.EOF Then
    acceso = Trim(RS1!dpe_deracc)
    incluir = Trim(RS1!dpe_deragr)
    alterar = Trim(RS1!dpe_dermod)
    eliminar = Trim(RS1!dpe_dereli)
    imprimir = Trim(RS1!dpe_derimp)
    ValidarUsuarioAcceso = acceso & incluir & alterar & eliminar & imprimir
End If
RS1.Close: Set RS1 = Nothing
'--------------------------------------------------------------
End Function

Function ValidarOpEnvio(cencos As String, op As Integer) As Boolean

Dim RS1 As New ADODB.Recordset
'Dim sql As String
ValidarOpEnvio = False
'sql = ""
'Select Case op
'Case 1
'    sql = "AND cai_codtii=1"
'Case 2
'    sql = "AND cai_codtii=2"
'Case 3
'    sql = "AND cai_codtii=3"
'Case 5
'    sql = "AND cai_codtii=5"
'End Select
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT DISTINCT cai_cencos FROM b_casinointerfaz WHERE cai_cencos='" & cencos & "' and cai_codtii = " & op & "")
If Not RS1.EOF Then
   
   ValidarOpEnvio = True

End If
RS1.Close
Set RS1 = Nothing

End Function

Function fg_CalCtoRecInv(CodRec As Long, tiprec As Long, ctacon As String) As Double
Dim sql1 As String
Dim RS1 As New ADODB.Recordset
sql1 = IIf(vg_tipbase = "1", " SUM(b.red_canpro*iif(ISNULL(c.cpi_precos),0,c.cpi_precos)) AS cosrec ", " SUM(b.red_canpro*CASE WHEN (c.cpi_precos) IS NULL THEN 0 ELSE c.cpi_precos END) AS cosrec ")
RS1.Open "SELECT a.rec_codigo, a.rec_nombre, " & sql1 & " " & _
         "FROM b_receta a, b_recetadet b, b_contlistpreing c, b_productos d " & _
         "WHERE b.red_codigo = a.rec_codigo " & _
         "AND   b.red_codpro = c.cpi_coding AND c.cpi_cencos = '" & MuestraCasino(1) & "' " & _
         "AND  (c.cpi_codcom = d.pro_codigo) " & _
         "AND   d.pro_ctacon = '" & ctacon & "' " & _
         "AND   b.red_codigo = " & CodRec & " " & _
         "AND   b.red_tiprec = " & tiprec & " " & _
         "AND   b.red_cencos = '" & IIf(tiprec = 0, 0, MuestraCasino(1)) & "' " & _
         "GROUP BY a.rec_codigo, a.rec_nombre ORDER BY a.rec_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then fg_CalCtoRecInv = RS1!cosrec
RS1.Close: Set RS1 = Nothing
End Function

Function Redondear(Variable As Variant, numdec As Integer) As Variant
Dim cientos As Long, i As Integer
cientos = 1
For i = 1 To numdec
    cientos = cientos * 10
Next i
If IsNumeric(Variable) Then
    If (Variable * cientos) Mod 2 <> 0 Then
        Redondear = Round(Variable, numdec)
    Else
        If (Variable * cientos) - Int((Variable * cientos)) >= 0.5 Then
            Redondear = Round((Variable + 0.5), numdec)
'            Redondear = Round((variable * cientos + 0.5) / cientos, numdec)
        Else
            Redondear = Round(Variable, numdec)
        End If
    End If
Else
    Redondear = Variable
End If
End Function

Function fg_CalCtoRecPlan(Fecha As Long, tipmin As String, CodRec As Long, tiprec As Long, ctacon As String) As Double
Dim RS1 As New ADODB.Recordset
RS1.Open "SELECT a.rec_codigo, a.rec_nombre, SUM(b.red_canpro*c.mic_cospro) AS cosrec " & _
         "FROM b_receta a, b_recetadet b, b_minutacosto c, b_contlistpreing d, b_productos e " & _
         "WHERE b.red_codigo=a.rec_codigo AND b.red_codpro=c.mic_codpro " & _
         "AND   b.red_codpro=d.cpi_coding AND d.cpi_cencos='" & MuestraCasino(1) & "' " & _
         "AND  (d.cpi_codcom=e.pro_codigo OR d.cpi_codped=e.pro_codigo) AND e.pro_ctacon='" & ctacon & "' " & _
         "AND   b.red_codigo=" & CodRec & " AND b.red_tiprec=" & tiprec & " " & _
         "AND   b.red_cencos='" & IIf(tiprec = 0, 0, MuestraCasino(1)) & "' " & _
         "AND   c.mic_cencos='" & MuestraCasino(1) & "' AND c.mic_fecval=" & Fecha & " " & _
         "AND   c.mic_tipmin='" & tipmin & "' group by a.rec_codigo, a.rec_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then fg_CalCtoRecPlan = RS1!cosrec
RS1.Close: Set RS1 = Nothing
End Function

Function TipoDato(Variable, valor)

If VarType(Variable) = vbNull Then
    
    TipoDato = IIf(VarType(valor) <> vbNull, valor, " ")

Else
    
    TipoDato = Variable

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

Function fg_Encripta(Password As String) As String
Dim encrip As String, count As Integer
encrip = ""
Password = Trim$(Password)
For count = 1 To Len(Password)
    encrip = encrip & Chr$(Asc(Mid$(Password, count, 1)) + 73 + count)
Next
fg_Encripta = encrip
End Function

Function ValidaBod(ByVal cBod As Long, ByVal cPro As String)
Dim RS As New ADODB.Recordset
RS.Open "SELECT * FROM b_productos WHERE pro_codigo = '" & cPro & "' AND pro_ctrsto = 1", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: Exit Function
RS.Close: Set RS = Nothing
RS.Open "SELECT * FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = '" & cPro & "'", vg_db, adOpenStatic
If RS.EOF Then
    If vg_tipbase = "1" Then vg_db.BeginTrans
    vg_db.Execute "INSERT INTO b_bodegas VALUES (" & vg_codbod & ", '" & cPro & "', 0)"
    If vg_tipbase = "1" Then vg_db.CommitTrans
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

Function BoM(Fecha As Date) As String
Dim mes As Integer
mes = Month(Fecha)
Do While mes = Month(Fecha)
    BoM = fg_pone_cero(Year(Fecha), 4) & fg_pone_cero(Month(Fecha), 2) & fg_pone_cero(Day(Fecha), 2)
    Fecha = Fecha - 1
Loop
BoM = Fecha
End Function

Function BEoM(Fecha As Date) As String
Dim mes As Integer
mes = Month(Fecha)
Do While mes = Month(Fecha)
    Fecha = Fecha + 1
Loop
BEoM = Fecha
End Function

Function GetParametro(ByVal cpar As String) As Variant

Dim RS1 As New ADODB.Recordset

If Trim(vg_contra) = "" Then
   
   RS1.Open "SELECT * FROM a_param WHERE par_codigo = '" & cpar & "'", vg_db, adOpenStatic

Else

   Set RS1 = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', '" & cpar & "'")

End If

If Not RS1.EOF Then
    
    Select Case RS1!par_tipo
    
    Case "N"
        
        GetParametro = Val(TipoDato(RS1!par_valor, ""))
    
    Case "C"
        
        GetParametro = Trim(TipoDato(RS1!par_valor, ""))
    
    End Select

Else
    
    GetParametro = "" 'Null

End If
RS1.Close: Set RS1 = Nothing

End Function

Function GetParametro_Seguridad(ByVal cpar As String) As Variant

Dim RS1 As New ADODB.Recordset

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Trim(vg_contra) = "" Then
   
   RS1.Open "SELECT distinct * FROM a_param as a with (nolock) " & _
            "inner join b_clientes as b on b.cli_codigo = a.par_cencos " & _
            "and b.cli_tipo = 0 and b.cli_codbod > 0 and b.cli_activo = '1' " & _
            "WHERE a.par_codigo = '" & cpar & "'", vg_db, adOpenStatic

Else

   Set RS1 = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', '" & cpar & "'")

End If

If Not RS1.EOF Then
    
    Select Case RS1!par_tipo
    
    Case "N"
        
        GetParametro_Seguridad = Val(TipoDato(RS1!par_valor, ""))
    
    Case "C"
        
        GetParametro_Seguridad = Trim(TipoDato(RS1!par_valor, ""))
    
    End Select

Else
    
    GetParametro_Seguridad = "" 'Null

End If
RS1.Close
Set RS1 = Nothing

End Function

Function GetParametroCamIng(ByVal cpar As String) As Variant
Dim RS1 As New ADODB.Recordset
If Trim(vg_contra) = "" Then
   RS1.Open "SELECT * FROM a_param WHERE Mid(par_codigo,1,6) = '" & cpar & "'", vg_db, adOpenStatic
Else
   RS1.Open "SELECT * FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND mid(par_codigo,1,6) = '" & cpar & "'", vg_db, adOpenStatic
End If
If Not RS1.EOF Then
    Select Case RS1!par_tipo
    Case "N"
        GetParametroCamIng = Val(TipoDato(RS1!par_valor, ""))
    Case "C"
        GetParametroCamIng = Trim(TipoDato(RS1!par_valor, ""))
    End Select
Else
    GetParametroCamIng = Null
End If
RS1.Close: Set RS1 = Nothing
End Function

Function fg_WeekNumber(dDate As Date) As Integer
fg_WeekNumber = (DateDiff("ww", CDate("01/01/" & Year(dDate)), dDate, vbMonday, vbFirstFourDays) Mod 52) + 1
End Function

Sub fg_CheckTmp(ByVal cTabla As String)
Dim RS1 As New ADODB.Recordset
On Error GoTo ManError
RS1.Open "SELECT * FROM " & cTabla, vg_db, adOpenStatic
RS1.Close: Set RS1 = Nothing
vg_db.Execute "drop table " & cTabla
Exit Sub
ManError:
    If Err = -2147217865 Then Exit Sub
    MsgBox Err & ":  " & error$(Err), vbCritical, "Error"
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
ModCasino = False
'20070906 skipper ModCasino = True
'20070906 skipper RS1.Open "SELECT par_valor FROM a_param WHERE  par_cencos='" & MuestraCasino(1) & "' AND par_codigo='casinomod'", vg_db, adOpenStatic
'20070906 skipper if Not RS1.EOF Then ModCasino = IIf(RS1!par_valor = 0, False, True)
'20070906 skipper RS1.Close: Set RS1 = Nothing
End Function

Function MuestraCasino(op As Integer) As String

Dim RS1 As New ADODB.Recordset
MuestraCasino = ""
Select Case op

Case 1
    
    MuestraCasino = Trim(vg_contra) ' 20070906 skipper GetParametro("casino")

Case 2
    
    RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & vg_contra & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If Not RS1.EOF Then MuestraCasino = Trim(RS1!cli_nombre)
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
If tecla = 16 Or tecla = 17 Or tecla = 18 Or tecla = 19 Or tecla = 20 Or tecla = 32 Or tecla = 33 Or tecla = 34 Or tecla = 35 Or tecla = 36 Or tecla = 37 Or tecla = 38 Or tecla = 39 Or tecla = 40 Or tecla = 44 Or tecla = 45 Or tecla = 46 Or tecla = 91 Or tecla = 93 Or tecla = 120 Or tecla = 145 Or tecla = 144 Then TeclasNoPermitidas = False: Exit Function
TeclasNoPermitidas = True
End Function

Function EspFecha(Fecha As fpDateTime)
'Fecha.DateTimeFormat = 5
'If Bandera = 1 Then
'   Fecha.UserDefinedFormat = "dd/mm/yyyy"
'Else
'   Fecha.UserDefinedFormat = "mmmm"
'End If
Fecha.CalFirstDay (1)

Fecha.ShortDayName(1) = "Dom"
Fecha.ShortDayName(2) = "Lun"
Fecha.ShortDayName(3) = "Mar"
Fecha.ShortDayName(4) = "Mie"
Fecha.ShortDayName(5) = "Jue"
Fecha.ShortDayName(6) = "Vie"
Fecha.ShortDayName(7) = "Sab"
Fecha.LongDayName(1) = "Domingo"
Fecha.LongDayName(2) = "Lunes"
Fecha.LongDayName(3) = "Martes"
Fecha.LongDayName(4) = "Miercoles"
Fecha.LongDayName(5) = "Jueves"
Fecha.LongDayName(6) = "Viernes"
Fecha.LongDayName(7) = "Sabado"
Fecha.LongMonthName(1) = "Enero"
Fecha.LongMonthName(2) = "Febrero"
Fecha.LongMonthName(3) = "Marzo"
Fecha.LongMonthName(4) = "Abril"
Fecha.LongMonthName(5) = "Mayo"
Fecha.LongMonthName(6) = "Junio"
Fecha.LongMonthName(7) = "Julio"
Fecha.LongMonthName(8) = "Agosto"
Fecha.LongMonthName(9) = "Septiembre"
Fecha.LongMonthName(10) = "Octubre"
Fecha.LongMonthName(11) = "Noviembre"
Fecha.LongMonthName(12) = "Diciembre"
End Function

Function UltCierre(Bodega As Long) As Date
Dim RS1 As New ADODB.Recordset
UltCierre = CDate("01/01/1899")
RS1.Open "SELECT Max(tin_fectom) AS fectom FROM b_tomainv WHERE left(tin_fectom,6) = tin_ciemes AND tin_codbod = " & Bodega, vg_db, adOpenStatic
If Not RS1.EOF And Not IsNull(RS1!fectom) Then UltCierre = fg_Ctod1(RS1!fectom)
RS1.Close: Set RS1 = Nothing
End Function

Function CierreAjuste() As Boolean

Dim fechaCA As Date, rutcli As String, Bodega As Long, bodegas As String, sql1 As String
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
rutcli = MuestraCasino(1)
CierreAjuste = False

RS3.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE a.bod_codigo = b.cli_codbod AND b.cli_codigo = '" & vg_contra & "' ORDER BY bod_nombre", vg_db, adOpenForwardOnly
bodegas = ""

Do While Not RS3.EOF
    
    Bodega = RS3!bod_codigo
    RS1.Open "SELECT tin_fectom, tin_stofis, tin_stosis FROM b_tomainv WHERE left(tin_fectom,6) = tin_ciemes AND tin_codbod = " & Bodega & " " & _
             "AND tin_fectom = (SELECT MAX(tin_fectom) FROM b_tomainv WHERE left(tin_fectom,6) = tin_ciemes AND tin_codbod = " & Bodega & ")", vg_db, adOpenForwardOnly
    If Not RS1.EOF Then
       sql1 = IIf(vg_tipbase = "1", " CDate('" & fg_Ctod1(RS1!tin_fectom) & "') ", " '" & Format(fg_Ctod1(RS1!tin_fectom), "yyyymmdd") & "' ")
       
       RS2.Open "SELECT COUNT(*) AS nroreg FROM b_totventas WHERE tov_rutcli = '" & rutcli & "' AND tov_tipdoc = 'AI' " & _
                "AND tov_codbod = " & Bodega & " AND tov_fecemi = " & sql1 & " AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenForwardOnly
        
        If RS2!nroreg = 0 Then
             
             Do While Not RS1.EOF
                
                If Round(RS1!tin_stofis, vg_DCa) <> Round(RS1!tin_stosis, vg_DCa) Then
                    CierreAjuste = True
                    bodegas = bodegas & vbCrLf & "   * " & RS3!bod_codigo & " - " & RS3!bod_nombre & " (" & fg_Ctod1(RS1!tin_fectom) & ")"
                    Exit Do
                
                End If
                RS1.MoveNext
             
             Loop
        
        End If
        
        RS2.Close: Set RS2 = Nothing
    
    End If
    RS1.Close: Set RS1 = Nothing
    RS3.MoveNext
Loop
RS3.Close: Set RS3 = Nothing
If CierreAjuste Then MsgBox "Las sig. bodegas tienen cierre con diferencias y no ha hecho ajuste:" & vbCrLf & bodegas & vbCrLf & vbCrLf & "No puede mover inventario hasta que realice el ajuste...", vbCritical

End Function

Function DiferenciaInventario(FecCie As Long) As Boolean
Dim fechaCA As Date, rutcli As String, Bodega As Long, bodegas As String
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim MsgTitulo As String
MsgTitulo = "Cierre Diario"
rutcli = MuestraCasino(1)
DiferenciaInventario = False
RS3.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE a.bod_codigo = b.cli_codbod AND b.cli_codigo = '" & vg_contra & "' ORDER BY bod_nombre", vg_db, adOpenStatic
bodegas = ""
Do While Not RS3.EOF
    Bodega = RS3!bod_codigo
    RS1.Open "SELECT tin_fectom, tin_stofis, tin_stosis FROM b_tomainv WHERE tin_fectom = " & FecCie & " AND tin_codbod = " & Bodega & " " & _
             "AND Round(tin_stofis, " & vg_DCa & ") <> Round(tin_stosis, " & vg_DCa & ")", vg_db, adOpenStatic
    If Not RS1.EOF Then
       DiferenciaInventario = True
       Exit Do
    End If
    RS1.Close: Set RS1 = Nothing
    RS3.MoveNext
Loop
RS3.Close: Set RS3 = Nothing
End Function

Function CierrePeriodo(Fecha, codbod As Long, opcion As Integer) As Boolean

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset

Dim fecini          As Long
Dim fecfin          As Long
Dim sql1            As String
Dim sql2            As String
Dim periodo         As Long

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim NomArchivoExcel As String

CierrePeriodo = False

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If RS3.State = 1 Then RS3.Close
RS3.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case opcion

Case 0
    RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_fecini <= " & Fecha & " AND cie_fecter >= " & Fecha & " AND cie_estado = 1", vg_db, adOpenStatic
    If RS1.EOF Then CierrePeriodo = True
    RS1.Close: Set RS1 = Nothing
Case 1
    RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_fecter <= " & Val(Format(Date, "yyyymmdd")) & " AND cie_estado = 1", vg_db, adOpenStatic
    If Not RS1.EOF Then MsgBox "Se recuerda que tiene que realizar su cierre de mes", vbExclamation
    RS1.Close: Set RS1 = Nothing
Case 2
   '-------> Verificar si existen datos en esa fecha en ventas
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventas WHERE tov_codbod = " & codbod & " AND ((tov_fecemi > " & sql1 & " AND tov_tipdoc NOT IN ('DP', 'SP')) OR (tov_fecpro > " & sql1 & " AND tov_tipdoc IN ('DP', 'SP') AND (tov_fecpro) IS NOT NULL )) AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en ventas
   
   '-------> Verificar si existen datos en esa fecha en documentos
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_fecrem > " & sql1 & "", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en documentos

   '-------> Verificar si existen datos en esa fecha en ventas servicios especiales
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventaserviciosespeciales WHERE tos_IdBodega = " & codbod & " AND tos_fecha_produccion > " & sql1 & " AND tos_tipo_documento IN ('DE', 'SE') AND tos_estado_documento <> 'A' AND tos_estado_documento <> 'P'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en ventas servicios espaciales

   '-------> Verificar si existen datos en esa fecha en ventas cafeteria
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventascaf WHERE tvc_codbod = " & codbod & " AND tvc_fecing > " & sql1 & " AND tvc_estado <> 'A' AND tvc_estado <> 'P'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en ventas cafeteria
   
Case 3
   '-------> Verificar si existe toma inventario
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom > " & Fecha & "", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Din verificar si existe toma inventario
   
   '-------> Verificar si existen datos en esa fecha en ventas
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventas WHERE tov_codbod = " & codbod & " AND tov_fecemi > " & sql1 & " AND tov_fecpro = " & sql1 & " AND (tov_fecpro) IS NOT NULL AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en ventas
   
   '-------> Verificar si existen datos en esa fecha en documentos
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_fecrem > " & sql1 & "", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en documentos
Case 4
   '-------> Verificar existe toma inventario de fin de mes
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom = " & Fecha & "", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg = 0 Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar existe toma inventario de fin de mes
Case 5
    '-------> Verificar si existe un cierre de mes entre el periodo
    RS1.Open "SELECT * FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_ciemes = " & Val(Mid(Fecha, 1, 6)) & "", vg_db, adOpenStatic
    If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
    RS1.Close: Set RS1 = Nothing
    '-------> Fin verificar si existe un cierre de mes entre el periodo
Case 6
    sql1 = IIf(vg_tipbase = "1", " val(format(b.tov_fecemi, 'yyyymmdd')) ", " convert(varchar(10),b.tov_fecemi, 112) ")
    RS1.Open "SELECT DISTINCT a.tin_fectom FROM b_tomainv a, b_totventas b WHERE a.tin_codbod = " & codbod & " AND a.tin_fectom = " & sql1 & " AND a.tin_codbod = b.tov_codbod AND b.tov_tipdoc = 'AI' AND b.tov_estdoc <> 'A' AND b.tov_estdoc <> 'P' AND a.tin_fectom >= " & Fecha & "", vg_db, adOpenStatic
    If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
    RS1.Close: Set RS1 = Nothing
Case 7
   '-------> Verificar si si existe diferencia toma inventario
   sql1 = IIf(vg_tipbase = "1", " val(format(b.tov_fecemi, 'yyyymmdd')) ", " convert(varchar(10),b.tov_fecemi, 112) ")
   RS1.Open "SELECT DISTINCT a.tin_fectom FROM b_tomainv a, b_totventas b WHERE a.tin_codbod = " & codbod & " AND a.tin_fectom =  " & sql1 & " AND a.tin_codbod = b.tov_codbod AND b.tov_tipdoc = 'AI' AND b.tov_estdoc <> 'A' AND b.tov_estdoc <> 'P' AND a.tin_fectom >= " & Fecha & "", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = False: Exit Function
    RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si si existe diferencia toma inventario
   
   '-------> Verificar si existen datos en esa fecha en ventas
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventas WHERE tov_codbod = " & codbod & " AND ((tov_fecemi > " & sql1 & " AND tov_tipdoc NOT IN ('DP', 'SP')) OR (tov_fecpro > " & sql1 & " AND tov_tipdoc IN ('DP', 'SP') AND (tov_fecpro) IS NOT NULL)) AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en ventas
   
   '-------> Verificar si existen datos en esa fecha en documentos
   sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_fecrem > " & sql1 & "", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos en esa fecha en documentos
Case 8
   RS3.Open "SELECT max(tin_fectom) AS tin_fectom FROM b_tomainv WHERE tin_codbod = " & codbod & "", vg_db, adOpenStatic
   If Not RS3.EOF And Not IsNull(RS3!tin_fectom) Then
      RS1.Open "SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom >= " & RS3!tin_fectom & " AND Round(tin_stosis," & vg_DCa & ") <> Round(tin_stofis," & vg_DCa & ") ORDER BY tin_fectom", vg_db, adOpenStatic
      If Not RS1.EOF Then
         sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(RS3!tin_fectom) & "') ", " '" & Format(fg_Ctod1(RS3!tin_fectom), "yyyymmdd") & "' ")
         RS2.Open "SELECT DISTINCT tov_tipdoc FROM b_totventas WHERE tov_codbod = " & codbod & " AND tov_fecemi = " & sql1 & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
         If Not RS2.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: RS3.Close: Set RS3 = Nothing: CierrePeriodo = False: Exit Function
         RS1.Close: Set RS1 = Nothing
         RS2.Close: Set RS2 = Nothing
         RS3.Close: Set RS3 = Nothing
         CierrePeriodo = True: Exit Function
       End If
       RS1.Close: Set RS1 = Nothing
   End If
   RS3.Close: Set RS3 = Nothing
   CierrePeriodo = False: Exit Function
Case 9
   '-------> Fecha del periodo
   fecini = 0: fecfin = 0
   RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
   If Not RS1.EOF Then fecini = RS1!cie_fecini: fecfin = RS1!cie_fecter
   RS1.Close: Set RS1 = Nothing
   '-------> Verificar si existen datos pendientes salida producción
   If fecini > 0 And fecfin > 0 Then
      sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
      sql2 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
      RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventas WHERE tov_codbod = " & codbod & " AND tov_fecpro >= " & sql1 & " AND tov_fecpro <= " & sql2 & " AND (tov_fecpro)  IS NOT NULL AND tov_estdoc = 'P' And tov_tipdoc = 'SP'", vg_db, adOpenStatic
      If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
      RS1.Close: Set RS1 = Nothing
   End If
   '-------> Fin verificar si existen datos pendientes salida producción
Case 10
   '-------> Verificar si existen datos pendientes salida producción
   sql1 = IIf(vg_tipbase = "1", " FORMAT(tov_fecpro,'yyyymm') ", " substring(CONVERT(varchar(10), tov_fecpro,112),1,6) ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventas WHERE tov_codbod = " & codbod & " AND " & sql1 & " >= '" & Fecha & "' AND " & sql1 & " <= '" & Fecha & "' AND NOT (tov_fecpro) IS NULL AND (tov_estdoc = 'P' OR tov_estdoc <> 'A') AND tov_tipdoc = 'SP'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos pendientes salida producción
Case 11
    RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If RS1.EOF Then CierrePeriodo = True
    RS1.Close: Set RS1 = Nothing
Case 12
    RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_fecter >= " & Fecha & " AND cie_estado = 1", vg_db, adOpenStatic
    If RS1.EOF Then CierrePeriodo = True
    RS1.Close: Set RS1 = Nothing
Case 13
    RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_periodo = " & Fecha & " AND cie_estado = 1", vg_db, adOpenStatic
    If RS1.EOF Then CierrePeriodo = True
    RS1.Close: Set RS1 = Nothing
Case 14
   '-------> Verificar si existen datos pendientes salida producción
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "'  ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventas WHERE tov_codbod = " & codbod & " AND tov_fecpro = " & sql1 & " AND (tov_fecpro) IS NOT NULL AND tov_estdoc = 'P' And tov_tipdoc = 'SP'", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos pendientes salida producción
Case 15
   '-------> Verificar si existen datos pendientes salida producción
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   sql2 = IIf(vg_tipbase = "1", " trim(tvc_estado) ", " ltrim(tvc_estado) ")
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_totventascaf WHERE tvc_cencos = '" & MuestraCasino(1) & "' AND tvc_codbod = " & codbod & " AND tvc_fecing = " & sql1 & " AND ((tvc_estado) IS NULL OR " & sql2 & " = '')", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar si existen datos pendientes salida producción
Case 16
   '-------> Verificar existe toma inventario de fin de mes
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom = " & Fecha & " AND tin_ciemes <> 0 AND (tin_ciemes) IS NOT NULL", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg = 0 Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing
   '-------> Fin verificar existe toma inventario de fin de mes
Case 17
   '-------> Verificar si se han ingresado documentos de proveedores
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
'   RS1.Open "SELECT DISTINCT toc_fecemi  FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND toc_fecemi = " & sql1 & " " & _
'            "AND toc_tipdoc IN ('FA','FE','CE', 'DE', 'GD', 'NC', 'ND')", vg_db, adOpenStatic
'20200129   RS1.Open "SELECT DISTINCT toc_fecrem  FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = " & sql1 & " " & _
'            "AND toc_tipdoc IN ('FA','FE','CE', 'DE', 'GD', 'NC', 'ND')", vg_db, adOpenStatic

   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   RS1.Open "SELECT DISTINCT toc_fecrem  FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = " & sql1 & " " & _
            "AND toc_tipdoc IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('FA','FE','CE', 'DE', 'GD', 'NC', 'ND'))", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close
   Set RS1 = Nothing
   '-------> Fin Verificar si se han ingresado documentos de proveedores
Case 18
   '-------> Verificar salidas a producción
    Dim XLS As New Excel.Application 'Crea el objeto excel
    
    '-------> Ini : Validar servicio Principales 20201001
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & ", "
    sql1 = sql1 & " '2' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarServiciosSalidaProduccion " & sql1 & "")
    If Not RS1.EOF Then
    
        MsgBox "No se han realizado salidas producción. A continuación se generará archivo excel con el detalle de los datos faltantes." & VgLinea & VgLinea & Space(20) & " Cierre Cancelado ", vbCritical + vbOKOnly, MsgTitulo
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteErrorServciosSinSalida" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteErrorServciosSinSalida")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XLS.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLS.Visible = True
        XLS.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLS = Nothing
        
        RS1.Close
        Set RS1 = Nothing
        CierrePeriodo = True
        Exit Function
    
    End If
    
    RS1.Close: Set RS1 = Nothing
   
   '-------> Fin Verificar salidas a producción
Case 19
    '-------> Verificar devoluciones a bodega
    sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
    RS1.Open "SELECT DISTINCT tov_rutcli FROM b_totventas WHERE tov_rutcli = '" & MuestraCasino(1) & "' AND tov_tipdoc = 'DP' " & _
        "AND tov_codbod = " & vg_codbod & " AND tov_fecpro = " & sql1 & " " & _
        "AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
    RS1.Close: Set RS1 = Nothing
    '-------> Fin Verificar devoluciones a bodega
Case 20
    '-------> Verificar Mermas
    sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
    RS1.Open "SELECT DISTINCT tov_rutcli FROM b_totventas WHERE tov_rutcli = '" & MuestraCasino(1) & "' AND tov_tipdoc = 'ME' " & _
        "AND tov_codbod = " & vg_codbod & " AND tov_fecemi = " & sql1 & " " & _
        "AND tov_estdoc <> 'A'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
    RS1.Close: Set RS1 = Nothing
    '-------> Fin Verificar Mermas
Case 21
    '-------> Verificar raciones no vendidas
    sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
    RS1.Open "SELECT DISTINCT reg_codigo, reg_nombre, ser_codigo, ser_nombre FROM b_minuta a, b_minutadet b, a_regimen c, a_servicio d " & _
             "WHERE a.min_codigo = b.mid_codigo AND a.min_codreg = c.reg_codigo AND a.min_codser = d.ser_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' " & _
             "AND a.min_fecmin = " & Fecha & " AND b.mid_tipmin = '2'", vg_db, adOpenStatic
    Do While Not RS1.EOF
        RS2.Open "SELECT DISTINCT b.min_cencos FROM b_minuta b, b_minutadet c WHERE b.min_codigo = c.mid_codigo " & _
                 "AND b.min_cencos = '" & MuestraCasino(1) & "' AND b.min_codreg = " & RS1!reg_codigo & " AND b.min_codser = " & RS1!ser_codigo & "" & _
                 "AND b.min_fecmin = " & sql1 & " AND c.mid_nummer > 0", vg_db, adOpenStatic
        If RS2.EOF Then
           MsgBox "No Existen raciones no vendidas " & VgLinea & VgLinea & "del regimen " & RS1!reg_codigo & " - " & Trim(RS1!reg_nombre) & VgLinea & "del servicio " & RS1!ser_codigo & " - " & RS1!ser_nombre & VgLinea & VgLinea & "          Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo
           RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS1 = Nothing
           CierrePeriodo = True: Exit Function
        End If
        RS2.Close: Set RS2 = Nothing
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    '-------> Fin Verificar raciones no vendidas

Case 22 '-------> Verificar control de raciones
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT DISTINCT reg_codigo, reg_nombre, ser_codigo, ser_nombre FROM b_minuta a, b_minutadet b, a_regimen c, a_servicio d " & _
            "WHERE a.min_codigo = b.mid_codigo AND a.min_codreg = c.reg_codigo AND a.min_codser = d.ser_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' " & _
            "AND a.min_fecmin = " & Fecha & " AND b.mid_tipmin = '2' AND b.mid_numrac > 0", vg_db, adOpenStatic
   Do While Not RS1.EOF
      RS2.Open "SELECT DISTINCT mir_cencos FROM b_minutaraciones " & _
               "WHERE mir_cencos = '" & MuestraCasino(1) & "' AND mir_fecmin = " & sql1 & " " & _
               "AND mir_codreg = " & RS1!reg_codigo & " AND mir_codser = " & RS1!ser_codigo & " " & _
               "AND mir_rutcli NOT IN ('PERSONAL', 'PRODUCIDAS')", vg_db, adOpenStatic
      If RS2.EOF Then
         MsgBox "No existe control de raciones " & VgLinea & VgLinea & "del regimen " & RS1!reg_codigo & " - " & Trim(RS1!reg_nombre) & VgLinea & "del servicio " & RS1!ser_codigo & " - " & RS1!ser_nombre & VgLinea & VgLinea & "              Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo
         RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS1 = Nothing
         CierrePeriodo = True: Exit Function
      End If
      RS2.Close: Set RS2 = Nothing
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
   '-------> Fin Verificar control de raciones
   
Case 23 '-------> Verificar Registro de Venta Cafetería
    sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
    RS1.Open "SELECT DISTINCT tvc_cencos FROM b_totventascaf WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
             "AND tvc_codbod = " & vg_codbod & " AND tvc_fecing = " & sql1 & " AND tvc_estado = 'C'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
    RS1.Close: Set RS1 = Nothing
    '-------> Fin Verificar Registro de Venta Cafetería
    
Case 24 '-------> Verificar Venta Servicio Contado
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS2.Open "SELECT DISTINCT * FROM b_ventacontado " & _
            "WHERE a.vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   a.vtc_fecvta = " & Fecha & "", vg_db, adOpenStatic
   If RS2.EOF Then
      MsgBox "No existen Ventas Servicio Contado " & VgLinea & VgLinea & Space(3) & "Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo
      RS2.Close: Set RS2 = Nothing
      CierrePeriodo = True: Exit Function
   End If
   RS2.Close: Set RS2 = Nothing
   '-------> Fin Verificar Venta Servicio Contado
   
Case 25 '-------> Verificar Venta Directa
    sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
    RS1.Open "SELECT DISTINCT tov_rutcli FROM b_totventas " & _
             "WHERE tov_rutcli = '" & MuestraCasino(1) & "' " & _
             "AND tov_tipdoc IN ('FA', 'GD') " & _
             "AND tov_codbod  = " & vg_codbod & " " & _
             "AND tov_fecemi  = " & sql1 & " " & _
             "AND tov_estdoc <> 'A'", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
    RS1.Close: Set RS1 = Nothing
   '-------> Fin Verificar Venta Directa
   
Case 26 '-------> Verificar Inventario Rotativo
   If ValidarInventarioRotativo(MuestraCasino(1)) Then
        RS2.Open "SELECT tin_fectom, tin_codbod FROM b_tomainv " & _
                 "WHERE tin_fectom = " & Format(vg_ciedia, "yyyymmdd") & " " & _
                 "AND tin_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        If RS2.EOF Then
            RS2.Close: Set RS2 = Nothing: CierrePeriodo = True: Exit Function
        End If
        If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 8) Then RS2.Close: Set RS2 = Nothing: CierrePeriodo = True: Exit Function
        RS2.Close: Set RS2 = Nothing
    End If
   '-------> Fin Verificar Inventario Rotativo

Case 27
   RS1.Open "SELECT * FROM b_casinotipoactividades WHERE cta_cencos = '" & MuestraCasino(1) & "' AND cta_tipact = 10 ", vg_db, adOpenStatic
   If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 28 '-------> Validar si existe día feriado
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(Fecha) & "') ", " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' ")
   RS1.Open "SELECT DISTINCT CFI_Fecha FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & MuestraCasino(1) & "' AND CFI_Fecha = " & sql1 & " ORDER BY CFI_Fecha", vg_db, adOpenStatic
   If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 29 '-------> Validar si es fin del periodo
   RS1.Open "SELECT DISTINCT cie_cencos FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1 AND cie_fecter = " & Fecha & "", vg_db, adOpenStatic
   If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 30 '-------> Validar si existe solo toma inventario
   RS1.Open "SELECT COUNT(*) AS nreg FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom = " & Fecha & "", vg_db, adOpenStatic
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 31 '-------> Validar si día feriado cambio de periodo
   RS1.Open "SELECT DISTINCT cie_cencos FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1 AND cie_periodo = " & Format(fg_Ctod1(Fecha), "yyyymm") & "", vg_db, adOpenStatic
   If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 32 '-------> Validar si existe Ajuste Inventario creado
   RS3.Open "SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom = " & Fecha & "", vg_db, adOpenStatic
   If Not RS3.EOF Then
      RS1.Open "SELECT DISTINCT tin_fectom FROM b_tomainv WHERE tin_codbod = " & codbod & " AND tin_fectom >= " & RS3!tin_fectom & " AND Round(tin_stosis," & vg_DCa & ") <> Round(tin_stofis," & vg_DCa & ") ORDER BY tin_fectom", vg_db, adOpenStatic
      If Not RS1.EOF Then
         sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(RS3!tin_fectom) & "') ", " '" & Format(fg_Ctod1(RS3!tin_fectom), "yyyymmdd") & "' ")
         RS2.Open "SELECT DISTINCT tov_tipdoc FROM b_totventas WHERE tov_codbod = " & codbod & " AND tov_fecemi = " & sql1 & " AND tov_tipdoc = 'AI' AND tov_estdoc <> 'A' AND tov_estdoc <> 'P'", vg_db, adOpenStatic
         If RS2.EOF Then RS1.Close: Set RS1 = Nothing: RS2.Close: Set RS2 = Nothing: RS3.Close: Set RS3 = Nothing: CierrePeriodo = False: Exit Function
         RS1.Close: Set RS1 = Nothing
         RS2.Close: Set RS2 = Nothing
         RS3.Close: Set RS3 = Nothing
         CierrePeriodo = True: Exit Function
       End If
       RS1.Close: Set RS1 = Nothing
   End If
   RS3.Close: Set RS3 = Nothing
   CierrePeriodo = False: Exit Function

Case 33 '-------> validar si estan los días cerrado.
   '-------> Fecha del periodo
   RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
   If Not RS1.EOF Then periodo = RS1!cie_periodo
   RS1.Close: Set RS1 = Nothing
   sql1 = IIf(vg_tipbase = "1", " val(format(fecha, 'yyyymm')) ", " substring(CONVERT(varchar(10), fecha,112),1,6) ")
   RS1.Open "SELECT DISTINCT fecha, estenv, fecsub FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = '" & periodo & "' AND estenv = '0'", vg_db, adOpenStatic
   If Not RS1.EOF Then
      RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   End If
   RS1.Close: Set RS1 = Nothing
   CierrePeriodo = False: Exit Function

Case 34 '-------> Validar si esta cerrado periodo
   RS1.Open "SELECT DISTINCT cie_cencos FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 0 and cie_periodo = " & Fecha & "", vg_db, adOpenStatic
   If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 35 '-------> validar si día de calendario mes corresponde cierre diario
   RS1.Open "SELECT DISTINCT cie_cencos FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1 and cie_fecter = " & Fecha & "", vg_db, adOpenStatic
   If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: CierrePeriodo = True: Exit Function
   RS1.Close: Set RS1 = Nothing

Case 36 '-------> Verificar control de raciones
      
    Dim XL As New Excel.Application 'Crea el objeto excel
    
    '-------> Validar raciones produccidas
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & " "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarRacionesProducidas " & sql1 & "")
    If Not RS1.EOF Then
    
        If MsgBox("Esta seguro cerrar día, sin registrar raciones producidas ... ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
           
           RS1.Close: Set RS1 = Nothing
           CierrePeriodo = True
           Exit Function
           
        End If
    
        MsgBox "No se han asignado raciones producidas - (Planificación Real), acontinuación se muestra el detalle de los datos faltante" & VgLinea & VgLinea & " ", vbCritical + vbOKOnly, MsgTitulo
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteError" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteError")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XL.Visible = True
        XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XL = Nothing
        
        RS1.Close: Set RS1 = Nothing
    
    Else
        
        RS1.Close: Set RS1 = Nothing
    
    End If

    '-------> Fin validar raciones produccidas

Case 37 '-------> Validar precio venta
    
    Dim XLPVta As New Excel.Application 'Crea el objeto excel
    
    '-------> Ini : Validar precio venta 20201001
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & " "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarPrecioVenta " & sql1 & "")
    If Not RS1.EOF Then
    
        MsgBox "No se han ajustado los precios ventas de sus clientes - (Mantenedor Precio Venta), acontinuación se mostrar el detalle de los datos faltantes" & VgLinea & VgLinea & Space(20) & " ", vbCritical + vbOKOnly, MsgTitulo
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteErrorAjustePrecioCliente" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteErrorAjustePrecioCliente")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XLPVta.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLPVta.Visible = True
        XLPVta.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLPVta = Nothing
        
        RS1.Close
        Set RS1 = Nothing
        CierrePeriodo = True
        Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    '-------> Fin : Validar Precio venta 20201001

Case 38 'Validar si inventario esta en proceso
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("SELECT isnull(par_valor,'') as par_valor FROM a_param WHERE par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "' AND isnull(par_valor,'') = '1'")
   If Not RS1.EOF Then
   
      RS1.Close
      Set RS1 = Nothing
      CierrePeriodo = True
      Exit Function
   
   End If
   
   RS1.Close
   Set RS1 = Nothing

Case 39 'Validar inventario calendarizado
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgp_Sel_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(fg_Ctod1(Fecha), "yyyymmdd") & "")
   If Not RS1.EOF Then
   
      RS1.Close
      Set RS1 = Nothing
      CierrePeriodo = True
      Exit Function
   
   End If
   
   RS1.Close
   Set RS1 = Nothing

Case 40 'Validar ingreso documento inventario calendarizado
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgp_Sel_ValidarIngresoDocumentoInvCalendarizado '" & MuestraCasino(1) & "', " & Format(fg_Ctod1(Fecha), "yyyymmdd") & "")
   If Not RS1.EOF Then
   
      If Not IsNull(RS1(0)) Then
         
         RS1.Close
         Set RS1 = Nothing
         CierrePeriodo = True
         Exit Function
      
      End If
      
   End If
   
   RS1.Close
   Set RS1 = Nothing

Case 41 'Validar fecha fin de mes
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgp_Sel_ValidarFechaFinMes '" & MuestraCasino(1) & "', " & Format(fg_Ctod1(Fecha), "yyyymmdd") & "")
   If Not RS1.EOF Then
   
      If Not IsNull(RS1(0)) Then
         
         RS1.Close
         Set RS1 = Nothing
         CierrePeriodo = True
         Exit Function
      
      End If
      
   End If
   
   RS1.Close
   Set RS1 = Nothing

Case 42 'Validar que estado parametro toma inventario esta activado y que haya toma inventario - inventario calendarizado
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgp_Sel_ValidarParametroTomInvCalenToma '" & MuestraCasino(1) & "', " & Format(fg_Ctod1(Fecha), "yyyymmdd") & "")
   If Not RS1.EOF Then
   
      If Not IsNull(RS1(0)) Then
         
         RS1.Close
         Set RS1 = Nothing
         CierrePeriodo = True
         Exit Function
      
      End If
      
   End If
   
   RS1.Close
   Set RS1 = Nothing

Case 43

    Dim XLSp As New Excel.Application 'Crea el objeto excel
    
    MsgTitulo = "Cierre Diario"

    '-------> Ini : Validar servicio Principales 20201001
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & ", "
    sql1 = sql1 & " '2' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarServiciosPreferidos " & sql1 & "")
    If Not RS1.EOF Then
    
       If RS2.State = 1 Then RS2.Close
       RS2.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

       Set RS2 = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'parbserpri'")
       If Not RS2.EOF Then
   
           If RS2(0) = "0" Then
            
               MsgBox "No hay ingreso de comensales totales, se generará archivo excel con información faltante - (Planificación Real)  o No podrá visualizar requisiciones por falta de comensales totales, Se generará archivo con información pendientes..." & VgLinea & VgLinea & Space(20) & " ", vbInformation + vbOKOnly, MsgTitulo
               CierrePeriodo = False
            
           Else
              
               MsgBox "Falta ingresar raciones producidas (Planificación Real). Debe ingresar toda la información antes de cerrar día. A continuación se generará archivo con detalle de los datos faltates..." & VgLinea & VgLinea & Space(20) & " Cierre Cancelado ", vbCritical + vbOKOnly, MsgTitulo
               CierrePeriodo = True
            
           End If

        End If
        RS2.Close
        Set RS2 = Nothing
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteErrorServciosSinRaciones" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteErrorServciosSinRaciones")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XLSp.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLSp.Visible = True
        XLSp.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLSp = Nothing
        
        RS1.Close
        Set RS1 = Nothing
        
        Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '-------> Fin : Validar servicio Principales 20201001

Case 44

    Dim XLRv As New Excel.Application 'Crea el objeto excel

    MsgTitulo = "Cierre Diario"

    '-------> Validar raciones vendidas
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & " "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarClienteSinraciones " & sql1 & "")
    If Not RS1.EOF Then
        
        MsgBox "Falta ingresar raciones vendidas. (Control de raciones). Debe ingresar toda la información antes de cerrar día. A continuación se generará archivo excel con el detalle de los datos faltantes." & VgLinea & VgLinea & Space(20) & " Cierre cancelado ", vbCritical + vbOKOnly, MsgTitulo
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteError" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteError")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
 '       xlApp.Columns("A:A").Select
 '       xlApp.Selection.Delete Shift:=xlToLeft
  
        xlWb.Close True, NomArchivoExcel

        XLRv.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLRv.Visible = True
        XLRv.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLRv = Nothing
        
        RS1.Close: Set RS1 = Nothing
        CierrePeriodo = True
        Exit Function
    
    End If
    RS1.Close: Set RS1 = Nothing
   '-------> Fin Validar raciones vendida

Case 45 'Validar inventario hay ajuste
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgp_Sel_ValidarInventarioAjuste '" & MuestraCasino(1) & "', " & Format(fg_Ctod1(Fecha), "yyyymmdd") & "")
   If Not RS1.EOF Then
   
      If Not IsNull(RS1(0)) Then
         
         RS1.Close
         Set RS1 = Nothing
         CierrePeriodo = True
         Exit Function
      
      End If
      
   End If
   
   RS1.Close
   Set RS1 = Nothing

Case 46
   '-------> Verificar si existen datos pendientes ventas servicios espaciales
   sql1 = " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "'  "
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("sgp_Sel_ValidarServiciosEspecialesPendientes " & codbod & ", " & sql1 & "")
   If Not RS1.EOF And RS1!nreg > 0 And Not IsNull(RS1!nreg) Then
   
      RS1.Close
      Set RS1 = Nothing
      CierrePeriodo = True
      Exit Function
   
   End If
   RS1.Close
   Set RS1 = Nothing
   '-------> Fin verificar si existen datos ventas servicios especiales

Case 47

    Dim XLSm As New Excel.Application 'Crea el objeto excel
    
    MsgTitulo = "Cierre Diario"

    '-------> Ini : Validar raciones no vendidas 20235015
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & ", "
    sql1 = sql1 & " '2' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarRacionesNoVendidasPreferidos " & sql1 & "")
    If Not RS1.EOF Then
    
        MsgBox "Falta ingresar raciones no vendidas (Raciones no Vendidas). Debe ingresar toda la información antes de cerrar día. A continuación se generará archivo con detalle de los datos faltates..." & VgLinea & VgLinea & Space(20) & " Cierre Cancelado ", vbCritical + vbOKOnly, MsgTitulo
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteErrorServciosSinRacionesNoVendidas" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteErrorServciosSinRacionesNoVendidas")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XLSm.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLSm.Visible = True
        XLSm.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLSm = Nothing
        
        RS1.Close
        Set RS1 = Nothing
        CierrePeriodo = True
        Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '-------> Fin : Validar raciones no vendidas 20235015

Case 48

    '-------> ini : Validar ajuste inventario
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & vg_codbod & ", "
    sql1 = sql1 & " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarAjusteInventario " & sql1 & "")
    If Not RS1.EOF Then
    
      RS1.Close
      Set RS1 = Nothing
      CierrePeriodo = True
      Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '-------> fin : Validar ajuste inventario

Case 49

    '-------> ini : Validar envio de inventario
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " '" & Format(fg_Ctod1(Fecha), "yyyymmdd") & "', "
    sql1 = sql1 & " '2', "
    sql1 = sql1 & " '1' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarEnvioInventario " & sql1 & "")
    If Not RS1.EOF Then
    
      RS1.Close
      Set RS1 = Nothing
      CierrePeriodo = True
      Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '-------> fin : Validar envio inventario

Case 50

    Dim XLSmDes As New Excel.Application 'Crea el objeto excel
    
    MsgTitulo = "Cierre Diario"
    
    '-------> Ini : Validar raciones no vendidas desconche
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & ", "
    sql1 = sql1 & " '2' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarRacionesNoVendidasDesconchePreferidos " & sql1 & "")
    If Not RS1.EOF Then
    
        MsgBox "Falta ingresar desconche (Raciones no Vendidas). Debe ingresar toda la información antes de cerrar día. A continuación se generará archivo con detalle de los datos faltates..." & VgLinea & VgLinea & Space(20) & " Cierre Cancelado ", vbCritical + vbOKOnly, MsgTitulo
        
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteErrorServciosSinRacionesNoVendidasDesconche" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteErrorServciosSinRacionesNoVendidasDesconche")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XLSmDes.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLSmDes.Visible = True
        XLSmDes.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLSmDes = Nothing
        
        RS1.Close
        Set RS1 = Nothing
        CierrePeriodo = True
        Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '-------> Fin : Validar raciones no vendidas desconche

Case 51

    Dim XLSpSP As New Excel.Application 'Crea el objeto excel
    
    MsgTitulo = "Cierre Diario"
    
    '-------> Ini : Validar servicio Principales salida producción
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = ""
    sql1 = sql1 & " '" & MuestraCasino(1) & "', "
    sql1 = sql1 & " " & Format(fg_Ctod1(Fecha), "yyyymmdd") & ", "
    sql1 = sql1 & " '2' "
    Set RS1 = vg_db.Execute("sgp_Sel_ValidarServiciosPreferidosSalidaProduccion " & sql1 & "")
    If Not RS1.EOF Then
    
              
        MsgBox "Falta ingresar salida produción - (Salida Producción). Debe ingresar toda la información antes de cerrar día. A continuación se generará archivo con detalle de los datos faltates..." & VgLinea & VgLinea & Space(20) & " Cierre Cancelado ", vbCritical + vbOKOnly, MsgTitulo
        CierrePeriodo = True
                   
        M_BorrarArchivos.InicioBorrado dir_trabajo_Inf & "ExcelSGP\", "ReporteErrorServciosSinSalidaProduccion" & "*.Xls", 4
        
        NomArchivoExcel = fg_ArchivoXls("ReporteErrorServciosSinSalidaProduccion")
        '-------> Create an instance of Excel and add a workbook
        Set xlApp = CreateObject("Excel.Application")
        Set xlWb = xlApp.Workbooks.Add
        Set xlWs = xlWb.Worksheets("Hoja1")
  
        '-------> Display Excel and give user control of Excel's lifetime
        xlApp.UserControl = True
    
        '-------> Check version of Excel
        Call encabezado(RS1, xlWs)
          
        xlWs.Cells(2, 1).CopyFromRecordset RS1

        '-------> Auto-fit the column widths and row heights
        xlApp.Selection.CurrentRegion.Columns.AutoFit
        xlApp.Selection.CurrentRegion.Rows.AutoFit
    
        xlWb.Close True, NomArchivoExcel

        XLSpSP.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
        XLSpSP.Visible = True
        XLSpSP.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
        ' -- Cerrar Excel
        xlApp.Quit
        '-------> Release Excel references
        Set xlWs = Nothing
        Set xlWb = Nothing
        Set xlApp = Nothing
        Set XLSpSP = Nothing
        
        RS1.Close
        Set RS1 = Nothing
        
        Exit Function
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '-------> Fin : Validar servicio Principales Salida Producción

End Select

End Function

Function CierreFecha() As String
Dim RS1 As New ADODB.Recordset
CierreFecha = ""
RS1.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
If Not RS1.EOF Then CierreFecha = "Del   " & Mid(RS1!cie_fecini, 7, 2) & "/" & Mid(RS1!cie_fecini, 5, 2) & "/" & Mid(RS1!cie_fecini, 1, 4) & "  al  " & Mid(RS1!cie_fecter, 7, 2) & "/" & Mid(RS1!cie_fecter, 5, 2) & "/" & Mid(RS1!cie_fecter, 1, 4)
RS1.Close: Set RS1 = Nothing
End Function

Function RecalPpp(fecper As Long, fecini As Long, fecfin As Long, codpro As String)
On Local Error GoTo Error_Mover
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String

'00 Inventario inicial
'10 Ajuste de Entrada
'20 Proveedores Entrada
'30 Traspaso Entrada
'40 Prodccion Salida
'50 Produccion Entrada
'60 Traspaso Salida
'70 Mermas Salida
'80 Venta Directa Salida
'90 Ajuste Salida

If fecper = 0 Then Exit Function
vg_db.BeginTrans
'------- Mover maestro de productos a un temporal
aAp = Trim(vg_NUsr) & "_tmp_RecalPppproducto"
fg_CheckTmp aAp
RS1.Open "SELECT a.pro_codigo, a.pro_nombre INTO " & aAp & " " & _
         "FROM b_productos a, b_productospmpdia b " & _
         "WHERE a.pro_codigo = b.ppd_codpro " & _
         "AND   b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   b.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
         "AND   b.ppd_propon > 0 AND a.pro_ctrsto = 1", vg_db, adOpenStatic
Set RS1 = Nothing
'------- Fin mover maestro de productos a un temporal

'------- Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_RecalPpp"
fg_CheckTmp aAp
RS1.Open "SELECT DISTINCT tin_fectom AS fecpro, tin_codpro AS codpro, tin_stofis AS cansto, tin_propon AS propon, " & _
         "'E' AS tipmov , 0 AS numdoc, 'E' AS tipdoc, 'E' AS rutcli, '00' AS orden INTO " & aAp & " FROM b_tomainv " & _
         "WHERE tin_ciemes = " & fecper & " " & _
         "AND   tin_stofis > 0 " & _
         "AND   tin_propon > 0 " & _
         "AND  (tin_codpro = '" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "'='') ORDER BY tin_fectom, tin_codpro", vg_db, adOpenStatic
Set RS1 = Nothing
DoEvents
'------- Fin traer Inventario primer inventario

'------- Traer ajuste inventario
vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmin AS cansto, dev.dev_precos AS propon, iif(aju.aju_codigo=3,'E','S') AS tipmov, " & _
              "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, IIf(aju.aju_tipo = 'A', '10', '90') AS orden " & _
              "FROM b_totventas tov, b_detventas dev, b_productos pro, a_tipoajuste aju " & _
              "WHERE tov.tov_rutcli = dev.dev_rutcli " & _
              "AND   tov.tov_tipdoc = dev.dev_tipdoc " & _
              "AND   tov.tov_numdoc = dev.dev_numdoc " & _
              "AND   dev.dev_codmer = pro.pro_codigo " & _
              "AND   tov.tov_codser = aju.aju_codigo " & _
              "AND   format(tov.tov_fecemi, 'yyyymmdd')>=" & fecini & " " & _
              "AND   format(tov.tov_fecemi, 'yyyymmdd')<=" & fecfin & " " & _
              "AND   tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc='AI' AND tov.tov_estdoc<>'A' AND pro.pro_ctrsto=1 AND (pro.pro_codigo='" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "'='') ORDER BY tov.tov_fecemi, pro.pro_codigo"
'------- Fin traer ajuste inventario
    
'------- Traer salida y devolución produción
vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecpro, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, " & _
              "tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, iif(tov.tov_tipdoc='SP','40','50') AS orden " & _
              "FROM b_totventas tov, b_detventas dev, b_productos pro " & _
              "WHERE tov.tov_rutcli = dev.dev_rutcli " & _
              "AND   tov.tov_tipdoc = dev.dev_tipdoc " & _
              "AND   tov.tov_numdoc = dev.dev_numdoc " & _
              "AND   dev.dev_codmer = pro.pro_codigo " & _
              "AND   pro.pro_ctrsto = 1 " & _
              "AND  (tov.tov_tipdoc = 'SP' OR tov.tov_tipdoc = 'DP') " & _
              "AND   dev.dev_canmer <> 0 AND tov.tov_codbod = " & vg_codbod & " AND tov.tov_estdoc <> 'A' " & _
              "AND   tov.tov_fecpro >= cdate('" & fg_Ctod1(fecini) & "')" & _
              "AND   tov.tov_fecpro <= cdate('" & fg_Ctod1(fecfin) & "')" & _
              "AND  (pro.pro_codigo = '" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "'='')"
'------- Fin traer salida y devolución produción
    
'------- Traer mermas
vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, " & _
              "tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '70' AS orden " & _
              "FROM  b_totventas tov, b_detventas dev, b_productos pro " & _
              "WHERE tov.tov_rutcli = dev.dev_rutcli " & _
              "AND   tov.tov_tipdoc = dev.dev_tipdoc " & _
              "AND   tov.tov_numdoc = dev.dev_numdoc " & _
              "AND   dev.dev_codmer = pro.pro_codigo " & _
              "AND   tov.tov_tipdoc = 'ME' AND tov.tov_codbod = " & vg_codbod & " AND tov.tov_estdoc <> 'A' " & _
              "AND   format(tov.tov_fecemi,'yyyymmdd') >= " & fecini & "" & _
              "AND   format(tov.tov_fecemi,'yyyymmdd') >= " & fecfin & "" & _
              "AND   (pro.pro_codigo = '" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "' = '')"
'------- Fin traer mermas
    
'------- Traer documento traspaso entrada
vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, iif(tov.tov_codreg=0,'E','S') AS tipmov, " & _
              "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, iif(tov.tov_codreg=0,'30','60') AS orden " & _
              "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
              "WHERE  tov.tov_rutcli = dev.dev_rutcli " & _
              "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
              "AND    tov.tov_numdoc = dev.dev_numdoc " & _
              "AND    dev.dev_codmer = pro.pro_codigo " & _
              "AND    pro.pro_ctrsto = 1 " & _
              "AND    tov.tov_codbod = " & vg_codbod & " AND tov.tov_tipdoc = 'TR' AND tov.tov_estdoc <> 'A' " & _
              "AND    format(tov.tov_fecemi,'yyyymmdd') >= " & fecini & " AND format(tov.tov_fecemi,'yyyymmdd') <= " & fecfin & " AND dev.dev_canmer > 0 AND (pro.pro_codigo = '" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "' = '') ORDER BY tov.tov_fecemi, pro.pro_codigo"
'------- Fin traer documento traspaso entrada
  
'------- Traer Documento Proveedor
Dim pctimp As Double, pctdes  As Double, Precio As Double

RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, de.dec_numdoc, " & _
         "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
         "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
         "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
         "WHERE toc.toc_rutpro = de.dec_rutpro " & _
         "AND   toc.toc_tipdoc = de.dec_tipdoc " & _
         "AND   toc.toc_numdoc = de.dec_numdoc " & _
         "AND   de.dec_codmer = pro.pro_codigo " & _
         "AND   de.dec_mueinv = 'S' and toc.toc_tipdoc not in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN' ) " & _
         "AND   de.dec_canrec > 0 " & _
         "AND  format(toc.toc_fecrem, 'yyyymmdd') >= '" & fecini & "' " & _
         "AND  format(toc.toc_fecrem, 'yyyymmdd') <= '" & fecfin & "' " & _
         "AND  toc.toc_codbod = " & vg_codbod & " AND (pro.pro_codigo = '" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "' = '') " & _
         "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      pctimp = 0: Precio = 0
      RS2.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
               "FROM  b_detcomprasimp a, a_impuesto b " & _
               "WHERE a.imd_rutdoc = '" & RS1!dec_rutpro & "' " & _
               "AND   a.imd_tipdoc = '" & RS1!dec_tipdoc & "' " & _
               "AND   a.imd_numdoc = " & RS1!dec_numdoc & " " & _
               "AND   a.imd_numlin = " & RS1!dec_numlin & " " & _
               "AND   a.imd_codpro = '" & RS1!pro_codigo & "' " & _
               "AND   a.imd_codimp = b.imp_codigo " & _
               "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
      If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
      pctdes = 0
      If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
      If RS1!dec_prefle > 0 Then
         Precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec)) + (RS1!dec_prefle / RS1!dec_canrec)
      Else
         Precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec))
      End If
      vg_db.Execute "insert into " & aAp & " values (" & Val(Format(RS1!toc_fecrem, "yyyymmdd")) & ", " & _
                    "'" & Trim(RS1!pro_codigo) & "', " & RS1!dec_canrec & ", " & Precio & ", '" & "E+" & "', " & RS1!dec_numdoc & ", '" & Trim(RS1!dec_tipdoc) & "', '" & Trim(RS1!dec_rutpro) & "', '20')"
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
'------- Fin traer Documento Proveedor
   
'------- Procesar Información Precio Promedio Ponderado
Dim auxCanmer As Double, auxPropon As Long, propon As Long, auxfec As Long
Dim auxcodpro As String, auxtipdoc As String
RS1.Open "SELECT * FROM " & aAp & " WHERE (codpro = '" & LimpiaDato(Trim(codpro)) & "' OR '" & Trim(codpro) & "' = '') ORDER BY codpro, fecpro, orden, tipmov", vg_db, adOpenStatic
If Not RS1.EOF Then
   auxCanmer = 0: auxPropon = 0: propon = 0: auxfec = 0: auxcodpro = "": auxtipdoc = ""
   Do While Not RS1.EOF
      If RS1!codpro <> auxcodpro Then
         If Trim(auxcodpro) <> "" And propon > 0 Then
            '------- Actualizar maestro producto y ingrediente
            '------- Fin actualizar maestro producto y ingrediente
         End If
         auxcodpro = RS1!codpro:  auxCanmer = 0: auxPropon = 0: propon = 0
      End If
      If RS1!tipmov = "S" Then
         '------- Actualizar ajuste-salida y devolución producción-mermas
         If propon > 0 Then
            '------- Actualizar encabezado y detalle ventas
            vg_db.Execute "UPDATE b_totventas INNER JOIN b_detventas ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli) SET b_detventas.dev_precos = " & propon & ", b_detventas.dev_predoc = " & propon & ", b_detventas.dev_ptotal=(" & propon & " * " & RS1!cansto & ") " & _
                          "WHERE b_totventas.tov_numdoc=" & RS1!NumDoc & " AND b_totventas.tov_tipdoc='" & RS1!tipdoc & "' AND b_totventas.tov_rutcli='" & RS1!rutcli & "' AND b_detventas.dev_codmer='" & RS1!codpro & "' AND b_totventas.tov_estdoc<>'A' AND b_totventas.tov_codbod=" & vg_codbod & ""
            RS2.Open "SELECT SUM(dev_ptotal) AS ptotal FROM b_detventas WHERE dev_rutcli='" & RS1!rutcli & "' AND dev_tipdoc='" & RS1!tipdoc & "' AND dev_numdoc=" & RS1!NumDoc & " GROUP BY dev_rutcli", vg_db, adOpenStatic
            If Not RS2.EOF Then
               vg_db.Execute "UPDATE b_totventas SET b_totventas.tov_totdoc=" & RS2!ptotal & " " & _
                             "WHERE b_totventas.tov_estdoc<>'A' AND b_totventas.tov_rutcli='" & RS1!rutcli & "' AND b_totventas.tov_tipdoc='" & RS1!tipdoc & "' AND b_totventas.tov_numdoc=" & RS1!NumDoc & " AND b_totventas.tov_codbod=" & vg_codbod & " AND " & RS2!ptotal & " > 0 AND NOT ISNULL(" & RS2!ptotal & ")"
            End If
            RS2.Close: Set RS2 = Nothing
            '------- Fin actualizar encabezado y detalle ventas
            If RS1!tipdoc = "AI" Then
            '------- Actualizar toma inventario
               vg_db.Execute "UPDATE b_tomainv SET tin_propon = " & propon & " WHERE tin_fectom=" & RS1!fecpro & " AND tin_codpro='" & RS1!codpro & "'"
            '------- Fin actualizar toma inventario
            End If
         End If
         '------- Fin actualizar ajuste-salida y devolución producción-mermas
      Else
         If RS1!cansto > 0 Then
            propon = ((auxPropon * IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer)) + (RS1!propon * RS1!cansto)) / (IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer) + RS1!cansto)
         End If
         auxPropon = propon: auxCanmer = (auxCanmer + RS1!cansto)
      End If
      RS1.MoveNext
   Loop
   If propon > 0 Then
   End If
End If
RS1.Close: Set RS1 = Nothing
vg_db.CommitTrans

Exit Function
Error_Mover:
    fg_descarga
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
    vg_db.RollbackTrans
    Resume Next
End Function

Function CalcularProvisiones(cencos As String, Fecha As Long, fecini As Long, fecter As Long, est As Boolean)

Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim pctimp As Double, pctdes As Double
Dim tgdali As Double, tgddes As Double, tgdgrl As Double
Dim tscali As Double, tscdes As Double, tscgrl As Double
Dim tncali As Double, tncdes As Double, tncgrl As Double
Dim tgdantali As Double, tgdantdes As Double, tgdantgrl As Double
Dim tsnantali As Double, tsnantdes As Double, tsnantgrl As Double
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, sql6 As String, sql7 As String

pctdes = 0
tgdali = 0
tgddes = 0
tgdgrl = 0
pctimp = 0
tscali = 0
tscdes = 0
tscgrl = 0
tncali = 0
tncdes = 0
tncgrl = 0
tgdantali = 0
tgdantdes = 0
tgdantgrl = 0
tsnantali = 0
tsnantdes = 0
tsnantgrl = 0

'-------> Traer guías pendientes del mes
sql1 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
sql2 = IIf(vg_tipbase = "1", " CDATE('" & fg_Ctod1(fecter) & "') ", " '" & Format(fg_Ctod1(fecter), "yyyymmdd") & "' ")
sql3 = IIf(vg_tipbase = "1", " trim(a.toc_docaso) ", " ltrim(a.toc_docaso) ")

'         "AND   a.toc_fecemi >= " & sql1 & " " & _
'         "AND   a.toc_fecemi <= " & sql2 & " " & _

RS2.Open "SELECT b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
         "b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
         "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
         "WHERE a.toc_rutpro = b.dec_rutpro " & _
         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
         "AND   a.toc_numdoc = b.dec_numdoc " & _
         "AND   b.dec_codmer = c.pro_codigo " & _
         "AND   a.toc_tipinf = 'C' AND a.toc_codbod = " & vg_codbod & " " & _
         "AND   a.toc_tipdoc in (select tdo_Codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
         "AND   a.toc_fecrem >= " & sql1 & " " & _
         "AND   a.toc_fecrem <= " & sql2 & " " & _
         "AND  (" & sql3 & " = '' OR (a.toc_docaso) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic
If Not RS2.EOF Then
   Do While Not RS2.EOF
      DoEvents
      pctdes = 0
      If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
      If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
         tgdali = Round(tgdali + (RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100)), vg_DPr) 'vg_DCa)
      ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         tgddes = Round(tgddes + ((RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100))), vg_DPr) 'vg_DCa)
      ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         tgdgrl = Round(tgdgrl + ((RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100))), vg_DPr) 'vg_DCa)
      End If
      RS2.MoveNext
   Loop
End If
RS2.Close: Set RS2 = Nothing

If est = True Then
   Dim aAp As String
   aAp = Trim(vg_NUsr) & "_tmp_GuiadelMes"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   sql3 = "": sql4 = "": sql5 = "": sql7 = ""
   sql3 = IIf(vg_tipbase = "1", " IIF((b_totcompras.toc_docaso) IS NULL,0, VAL(b_totcompras.toc_docaso)) ", " CASE WHEN (b_totcompras.toc_docaso) IS NULL THEN 0 ELSE  convert(int,b_totcompras.toc_docaso) END ")
   sql4 = IIf(vg_tipbase = "1", "  ", " OR lTrim(b_totcompras.toc_docaso) <> '' ")
   sql5 = IIf(vg_tipbase = "1", " cstr(d.toc_numdoc) ", " convert(varchar(20),d.toc_numdoc) ")

'   sql7 = IIf(vg_tipbase = "1", " FORMAT(b_totcompras.toc_fecemi,'yyyymm') ", " convert(int,convert(varchar(06),b_totcompras.toc_fecemi,112)) ")
   
   sql7 = IIf(vg_tipbase = "1", " FORMAT(b_totcompras.toc_fecrem,'yyyymm') ", " convert(int,convert(varchar(06),b_totcompras.toc_fecrem,112)) ")
   vg_db.Execute ("SELECT toc_numdoc INTO " & aAp & " FROM b_totcompras " & _
            "WHERE b_totcompras.toc_numdoc IN (SELECT DISTINCT " & sql3 & " FROM b_totcompras WHERE b_totcompras.toc_codbod = " & vg_codbod & " AND b_totcompras.toc_tipdoc = 'GD' AND ((b_totcompras.toc_docaso) IS NOT NULL " & sql4 & " )) " & _
            "AND  (b_totcompras.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('FA','FE') ) ) AND b_totcompras.toc_codbod = " & vg_codbod & " " & _
            "AND  ((b_totcompras.toc_docaso) IS NOT NULL " & sql4 & " ) " & _
            "AND   " & sql7 & " > " & Fecha & "")
   Set RS2 = Nothing

'            "AND   a.toc_fecemi >= " & sql1 & " " & _
'            "AND   a.toc_fecemi <= " & sql2 & " ORDER BY c.pro_ctacon", vg_db, adOpenStatic
   
   If RS2.State = 1 Then RS2.Close
   RS2.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS2.Open "SELECT b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
            "b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
            "FROM  b_totcompras a, b_detcompras b, b_productos c, " & aAp & " d " & _
            "WHERE a.toc_rutpro = b.dec_rutpro " & _
            "AND   a.toc_tipdoc = b.dec_tipdoc " & _
            "AND   a.toc_numdoc = b.dec_numdoc " & _
            "AND   b.dec_codmer = c.pro_codigo " & _
            "AND   a.toc_docaso = " & sql5 & " " & _
            "AND   a.toc_tipinf = 'C' AND a.toc_codbod = " & vg_codbod & " " & _
            "AND   a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND   a.toc_fecrem >= " & sql1 & " " & _
            "AND   a.toc_fecrem <= " & sql2 & " ORDER BY c.pro_ctacon", vg_db, adOpenStatic
   
   If Not RS2.EOF Then
      
      Do While Not RS2.EOF
         
         DoEvents
         pctdes = 0
         
         If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
         
         If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
            
            tgdali = Round(tgdali + (RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100)), vg_DPr) 'vg_DCa)
         
         ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
            
            tgddes = Round(tgddes + ((RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100))), vg_DPr) 'vg_DCa)
         
         ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
            
            tgdgrl = Round(tgdgrl + ((RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100))), vg_DPr) 'vg_DCa)
         
         End If
         RS2.MoveNext
      
      Loop
   
   End If
   RS2.Close
   Set RS2 = Nothing

End If
'-------> Fin traer guías pendientes del mes
    
'-------> Traer guías pendientes del meses anteriores
sql3 = IIf(vg_tipbase = "1", " TRIM(a.toc_docaso) ", " lTRIM(a.toc_docaso) ")

'         "AND   a.toc_fecemi < " & sql1 & " " & _

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS2.Open "SELECT b.dec_canmer, b.dec_precom, b.dec_pctdes, " & _
         "b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
         "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
         "WHERE a.toc_rutpro = b.dec_rutpro " & _
         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
         "AND   a.toc_numdoc = b.dec_numdoc " & _
         "AND   b.dec_codmer = c.pro_codigo " & _
         "AND   a.toc_tipinf = 'C' AND a.toc_codbod=" & vg_codbod & " " & _
         "AND   a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
         "AND   a.toc_fecrem < " & sql1 & " " & _
         "AND  (" & sql3 & " = '' OR (a.toc_docaso) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic
If Not RS2.EOF Then
   
   Do While Not RS2.EOF
      
      DoEvents
      pctdes = 0
      
      If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
      
      If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
         
         tgdantali = Round(tgdantali + (RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100)), vg_DPr) 'vg_DCa)
      
      ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         tgdantdes = Round(tgdantdes + ((RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100))), vg_DPr) 'vg_DCa)
      
      ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         tgdantgrl = Round(tgdantgrl + ((RS2!dec_ptotrec) - ((RS2!dec_ptotrec) * (pctdes / 100))), vg_DPr) 'vg_DCa)
      
      End If
      RS2.MoveNext
   
   Loop

End If
RS2.Close
Set RS2 = Nothing
'-------> Fin traer guías pendientes del meses anteriores
   
'-------> Traer solicitud nota creditos pendientes del mes
sql3 = IIf(vg_tipbase = "1", " TRIM(a.toc_docsnc) ", " lTRIM(a.toc_docsnc) ")

'         "AND   a.toc_fecemi >= " & sql1 & " " & _
'         "AND   a.toc_fecemi <= " & sql2 & " " & _

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
'         "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
'         "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
'         "WHERE a.toc_rutpro = b.dec_rutpro " & _
'         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'         "AND   a.toc_numdoc = b.dec_numdoc " & _
'         "AND   b.dec_codmer = c.pro_codigo " & _
'         "AND   a.toc_fecrem >= " & sql1 & " " & _
'         "AND   a.toc_fecrem <= " & sql2 & " " & _
'         "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc = 'SN' AND a.toc_codbod = " & vg_codbod & " " & _
'         "AND  (" & sql3 & " = '' OR (a.toc_docsnc) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic

RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
         "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
         "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
         "WHERE a.toc_rutpro = b.dec_rutpro " & _
         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
         "AND   a.toc_numdoc = b.dec_numdoc " & _
         "AND   b.dec_codmer = c.pro_codigo " & _
         "AND   a.toc_fecrem >= " & sql1 & " " & _
         "AND   a.toc_fecrem <= " & sql2 & " " & _
         "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND a.toc_codbod = " & vg_codbod & " " & _
         "AND  (" & sql3 & " = '' OR (a.toc_docsnc) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic
         
If Not RS2.EOF Then
   
   Do While Not RS2.EOF
      
      '-------> Traer Impuesto adicionales
      DoEvents
      pctimp = 0
      
      If RS3.State = 1 Then RS3.Close
      RS3.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      RS3.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
               "FROM b_detcomprasimp a, a_impuesto b " & _
               "WHERE a.imd_rutdoc = '" & RS2!dec_rutpro & "' " & _
               "AND   a.imd_tipdoc = '" & RS2!dec_tipdoc & "' " & _
               "AND   a.imd_numdoc = " & RS2!dec_numdoc & " " & _
               "AND   a.imd_numlin = " & RS2!dec_numlin & " " & _
               "AND   a.imd_codpro = '" & RS2!dec_codmer & "' " & _
               "AND   a.imd_codimp = b.imp_codigo " & _
               "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
      
      If RS3.EOF Then RS3.Close: Set RS3 = Nothing Else pctimp = RS3!imd_monimp: RS3.Close: Set RS3 = Nothing
      '-------> Fin traer Impuesto adicionales
      pctdes = 0
      If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
      
      If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
         
         tscali = Round(tscali + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
      
      ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         tscdes = Round(tscdes + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
      
      ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         tscgrl = Round(tscgrl + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
      
      End If
      RS2.MoveNext
   
   Loop

End If
RS2.Close
Set RS2 = Nothing

If est Then
   aAp = Trim(vg_NUsr) & "_tmp_SolicituddelMes"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   sql3 = IIf(vg_tipbase = "1", " IIF((b_totcompras.toc_docsnc) IS NULL OR trim(b_totcompras.toc_docsnc) = '', 0, VAL(b_totcompras.toc_docsnc)) ", " CASE WHEN (b_totcompras.toc_docsnc) IS NULL THEN 0 WHEN LTRIM(b_totcompras.toc_docsnc) = '' THEN 0 ELSE convert(int,b_totcompras.toc_docsnc) END ")
   sql4 = IIf(vg_tipbase = "1", " TRIM(b_totcompras.toc_docsnc) ", " lTRIM(b_totcompras.toc_docsnc) ")
   vg_db.Execute ("SELECT toc_numdoc INTO " & aAp & " FROM b_totcompras " & _
            "WHERE b_totcompras.toc_numdoc IN (SELECT DISTINCT " & sql3 & " FROM b_totcompras WHERE b_totcompras.toc_codbod=" & vg_codbod & " AND b_totcompras.toc_tipdoc = 'SN' AND ((b_totcompras.toc_docsnc) IS NULL OR " & sql4 & " <> '')) " & _
            "AND  (b_totcompras.toc_tipdoc in (select tdo_Codigo from a_tipodocumento where tdo_IdCodigo in ('NC','CE'))) AND b_totcompras.toc_codbod = " & vg_codbod & " " & _
            "AND  ((b_totcompras.toc_docsnc) IS NULL OR " & sql4 & " <> '') " & _
            " " & _
            "AND   " & sql7 & " > " & Fecha & "")
   
   If RS2.State = 1 Then RS2.Close
   RS2.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS2.Open "SELECT * FROM " & aAp & "", vg_db, adOpenStatic
   If Not RS2.EOF Then
      
      RS2.Close
      Set RS2 = Nothing
      
'               "AND   a.toc_fecemi >= " & sql1 & " " & _
'               "AND   a.toc_fecemi <= " & sql2 & " " & _

      sql3 = IIf(vg_tipbase = "1", " cstr(d.toc_numdoc) ", " convert(varchar(20),d.toc_numdoc) ")
      
      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
'      RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
'               "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
'               "FROM  b_totcompras a, b_detcompras b, b_productos c, " & aAp & " d " & _
'               "WHERE a.toc_rutpro = b.dec_rutpro " & _
'               "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'               "AND   a.toc_numdoc = b.dec_numdoc " & _
'               "AND   b.dec_codmer = c.pro_codigo " & _
'               "AND   a.toc_docsnc = " & sql3 & " AND a.toc_codbod = " & vg_codbod & " " & _
'               "AND   d.toc_numdoc > 0 " & _
'               "AND   a.toc_fecrem >= " & sql1 & " " & _
'               "AND   a.toc_fecrem <= " & sql2 & " " & _
'               "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc = 'SN' ORDER BY c.pro_ctacon", vg_db, adOpenStatic
      
      RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
               "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
               "FROM  b_totcompras a, b_detcompras b, b_productos c, " & aAp & " d " & _
               "WHERE a.toc_rutpro = b.dec_rutpro " & _
               "AND   a.toc_tipdoc = b.dec_tipdoc " & _
               "AND   a.toc_numdoc = b.dec_numdoc " & _
               "AND   b.dec_codmer = c.pro_codigo " & _
               "AND   a.toc_docsnc = " & sql3 & " AND a.toc_codbod = " & vg_codbod & " " & _
               "AND   d.toc_numdoc > 0 " & _
               "AND   a.toc_fecrem >= " & sql1 & " " & _
               "AND   a.toc_fecrem <= " & sql2 & " " & _
               "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') ORDER BY c.pro_ctacon", vg_db, adOpenStatic
      
      If Not RS2.EOF Then
         
         Do While Not RS2.EOF
            
            DoEvents
            '-------> Traer Impuesto adicionales
            pctimp = 0
            
            If RS3.State = 1 Then RS3.Close
            RS3.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            RS3.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
                     "FROM b_detcomprasimp a, a_impuesto b " & _
                     "WHERE a.imd_rutdoc = '" & RS2!dec_rutpro & "' " & _
                     "AND   a.imd_tipdoc = '" & RS2!dec_tipdoc & "' " & _
                     "AND   a.imd_numdoc = " & RS2!dec_numdoc & " " & _
                     "AND   a.imd_numlin = " & RS2!dec_numlin & " " & _
                     "AND   a.imd_codpro = '" & RS2!dec_codmer & "' " & _
                     "AND   a.imd_codimp = b.imp_codigo " & _
                     "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
            If RS3.EOF Then RS3.Close: Set RS3 = Nothing Else pctimp = RS3!imd_monimp: RS3.Close: Set RS3 = Nothing
            '-------> Fin traer Impuesto adicionales
            pctdes = 0
            If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
            
            If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
               
               tscali = Round(tscali + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
            
            ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
               
               tscdes = Round(tscdes + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
            
            ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
               
               tscgrl = Round(tscgrl + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
            
            End If
            RS2.MoveNext
         
         Loop
      
      End If
      RS2.Close
      Set RS2 = Nothing
   
   Else
      
      RS2.Close
      Set RS2 = Nothing
   
   End If

End If
'-------> Fin traer solicitud nota creditos pendientes del mes

'-------> Traer solicitud nota creditos pendientes del meses anteriores
sql3 = IIf(vg_tipbase = "1", " TRIM(a.toc_docsnc) ", " lTRIM(a.toc_docsnc) ")

'         "AND   a.toc_fecemi < " & sql1 & " " & _

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
'         "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
'         "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
'         "WHERE a.toc_rutpro = b.dec_rutpro " & _
'         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'         "AND   a.toc_numdoc = b.dec_numdoc " & _
'         "AND   b.dec_codmer = c.pro_codigo " & _
'         "AND   a.toc_fecrem < " & sql1 & " " & _
'         "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc = 'SN' AND a.toc_codbod = " & vg_codbod & " " & _
'         "AND  (" & sql3 & " = '' OR (a.toc_docsnc) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic

RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
         "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
         "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
         "WHERE a.toc_rutpro = b.dec_rutpro " & _
         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
         "AND   a.toc_numdoc = b.dec_numdoc " & _
         "AND   b.dec_codmer = c.pro_codigo " & _
         "AND   a.toc_fecrem < " & sql1 & " " & _
         "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND a.toc_codbod = " & vg_codbod & " " & _
         "AND  (" & sql3 & " = '' OR (a.toc_docsnc) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic

If Not RS2.EOF Then
   
   Do While Not RS2.EOF
      
      DoEvents
      pctimp = 0
      '-------> Traer Impuesto adicionales
      
      If RS3.State = 1 Then RS3.Close
      RS3.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      RS3.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
               "FROM b_detcomprasimp a, a_impuesto b " & _
               "WHERE a.imd_rutdoc = '" & RS2!dec_rutpro & "' " & _
               "AND   a.imd_tipdoc = '" & RS2!dec_tipdoc & "' " & _
               "AND   a.imd_numdoc = " & RS2!dec_numdoc & " " & _
               "AND   a.imd_numlin = " & RS2!dec_numlin & " " & _
               "AND   a.imd_codpro = '" & RS2!dec_codmer & "' " & _
               "AND   a.imd_codimp = b.imp_codigo " & _
               "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
      If RS3.EOF Then RS3.Close: Set RS3 = Nothing Else pctimp = RS3!imd_monimp: RS3.Close: Set RS3 = Nothing
      '-------> Fin traer Impuesto adicionales
      pctdes = 0
      
      If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
      
      If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
         
         tsnantali = Round(tsnantali + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
      
      ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         tsnantdes = Round(tsnantdes + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
      
      ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
         
         tsnantgrl = Round(tsnantgrl + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
      
      End If
      RS2.MoveNext
      
   Loop
End If
RS2.Close
Set RS2 = Nothing

If est Then
   
   aAp = Trim(vg_NUsr) & "_tmp_SolicituddelMes"
   '-------> Creo tabla temporal y chequeo si existe antes
   fg_CheckTmp aAp
   sql3 = IIf(vg_tipbase = "1", " IIF(ISNULL(b_totcompras.toc_docsnc) OR trim(b_totcompras.toc_docsnc) = '',0,VAL(b_totcompras.toc_docsnc)) ", " CASE WHEN (b_totcompras.toc_docsnc) IS NULL THEN 0 WHEN LTRIM(b_totcompras.toc_docsnc) = '' THEN 0 ELSE convert(int,b_totcompras.toc_docsnc) END ")
   sql4 = IIf(vg_tipbase = "1", " TRIM(b_totcompras.toc_docsnc) ", " lTRIM(b_totcompras.toc_docsnc) ")
   sql5 = IIf(vg_tipbase = "1", " TRIM(toc_docsnc) ", " lTRIM(toc_docsnc) ")
   vg_db.Execute ("SELECT toc_numdoc INTO " & aAp & " FROM b_totcompras " & _
            "WHERE toc_numdoc IN (SELECT " & sql3 & " FROM b_totcompras WHERE b_totcompras.toc_codbod = " & vg_codbod & " AND b_totcompras.toc_tipdoc = 'SN' AND ((b_totcompras.toc_docsnc) IS NULL OR " & sql4 & " <> '')) " & _
            "AND   (toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('NC','CE'))) AND toc_codbod = " & vg_codbod & " " & _
            "AND  ((toc_docsnc) IS NULL OR " & sql5 & " <> '') " & _
            "AND   " & sql7 & " > " & Fecha & "")
   Set RS2 = Nothing
   RS2.Open "SELECT * FROM " & aAp & "", vg_db, adOpenStatic
   If Not RS2.EOF Then
      
      RS2.Close: Set RS2 = Nothing
      sql3 = IIf(vg_tipbase = "1", " cstr(d.toc_numdoc) ", " convert(varchar(20),d.toc_numdoc) ")
      
'               "AND   a.toc_fecemi < " & sql1 & " " & _

      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

'      RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
'               "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
'               "FROM  b_totcompras a, b_detcompras b, b_productos c, " & aAp & " d " & _
'               "WHERE a.toc_rutpro = b.dec_rutpro " & _
'               "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'               "AND   a.toc_numdoc = b.dec_numdoc " & _
'               "AND   b.dec_codmer = c.pro_codigo " & _
'               "AND   a.toc_docsnc = " & sql3 & " AND a.toc_codbod=" & vg_codbod & " " & _
'               "AND   d.toc_numdoc > 0 " & _
'               "AND   a.toc_fecrem < " & sql1 & " " & _
'               "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc = 'SN' ORDER BY c.pro_ctacon", vg_db, adOpenStatic
      RS2.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
               "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
               "FROM  b_totcompras a, b_detcompras b, b_productos c, " & aAp & " d " & _
               "WHERE a.toc_rutpro = b.dec_rutpro " & _
               "AND   a.toc_tipdoc = b.dec_tipdoc " & _
               "AND   a.toc_numdoc = b.dec_numdoc " & _
               "AND   b.dec_codmer = c.pro_codigo " & _
               "AND   a.toc_docsnc = " & sql3 & " AND a.toc_codbod=" & vg_codbod & " " & _
               "AND   d.toc_numdoc > 0 " & _
               "AND   a.toc_fecrem < " & sql1 & " " & _
               "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') ORDER BY c.pro_ctacon", vg_db, adOpenStatic
      
      If Not RS2.EOF Then
         
         Do While Not RS2.EOF
            
            DoEvents
            
            '-------> Traer Impuesto adicionales
            pctimp = 0
            
            If RS3.State = 1 Then RS3.Close
            RS3.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            RS3.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
                     "FROM b_detcomprasimp a, a_impuesto b " & _
                     "WHERE a.imd_rutdoc = '" & RS2!dec_rutpro & "' " & _
                     "AND   a.imd_tipdoc = '" & RS2!dec_tipdoc & "' " & _
                     "AND   a.imd_numdoc = " & RS2!dec_numdoc & " " & _
                     "AND   a.imd_numlin = " & RS2!dec_numlin & " " & _
                     "AND   a.imd_codpro = '" & RS2!dec_codmer & "' " & _
                     "AND   a.imd_codimp = b.imp_codigo " & _
                     "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
            If RS3.EOF Then RS3.Close: Set RS3 = Nothing Else pctimp = RS3!imd_monimp: RS3.Close: Set RS3 = Nothing
            '-------> Fin traer Impuesto adicionales
            pctdes = 0
            If RS2!dec_pctdes > 0 Then pctdes = RS2!dec_pctdes
            
            If RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
               
               tsnantali = Round(tsnantali + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
            
            ElseIf RS2!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
               
               tsnantdes = Round(tsnantdes + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
            
            ElseIf RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS2!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
               
               tsnantgrl = Round(tsnantgrl + ((RS2!dec_ptotal - RS2!dec_ptotrec)) - (((RS2!dec_ptotal - RS2!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
            
            End If
            RS2.MoveNext
         
         Loop
         
      End If
      RS2.Close
      Set RS2 = Nothing
      
   Else
      
      RS2.Close
      Set RS2 = Nothing
      
   End If
   
End If
'-------> Fin traer solicitud nota creditos pendientes del mes

'-------> Traer solicitud nota creditos pedientes desde parametros
If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS2.Open "SELECT * FROM a_param WHERE par_codigo = 'parsn' AND par_cencos = '" & cencos & "'", vg_db, adOpenStatic
If Not RS2.EOF Then
   
   Dim ParFec As Date, mes As Integer, imes As Integer
'   mes = IIf(IsNull(RS2!par_valor), 6, RS2!par_valor)
   imes = 0: mes = 0
   ParFec = fg_Ctod1(fecini)
'   Do While mes <> Month(parfec)
   Do While imes <> IIf(IsNull(RS2!par_valor), 6, RS2!par_valor)
      
      ParFec = ParFec - 1
      If mes <> Month(ParFec) Then mes = Month(ParFec): imes = imes + 1
   
   Loop
   
'               "AND   a.toc_fecemi <= " & sql3 & " " & _

   sql3 = IIf(vg_tipbase = "1", " CDATE('" & ParFec & "') ", " '" & Format(ParFec, "yyyymmdd") & "' ")
   sql4 = IIf(vg_tipbase = "1", " TRIM(a.toc_docsnc) ", " lTRIM(a.toc_docsnc) ")
   
   If RS3.State = 1 Then RS3.Close
   RS3.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

'   RS3.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
'            "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
'            "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
'            "WHERE a.toc_rutpro = b.dec_rutpro " & _
'            "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'            "AND   a.toc_numdoc = b.dec_numdoc " & _
'            "AND   b.dec_codmer = c.pro_codigo " & _
'            "AND   a.toc_fecrem <= " & sql3 & " " & _
'            "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc = 'SN' AND a.toc_codbod = " & vg_codbod & " " & _
'            "AND  (" & sql4 & " = '' OR (a.toc_docsnc) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic
   
   RS3.Open "SELECT b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, b.dec_numlin, b.dec_canmer, b.dec_precom, b.dec_ptotal, b.dec_pctdes, " & _
            "b.dec_codmer, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, c.pro_ctacon, c.pro_ctrsto " & _
            "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
            "WHERE a.toc_rutpro = b.dec_rutpro " & _
            "AND   a.toc_tipdoc = b.dec_tipdoc " & _
            "AND   a.toc_numdoc = b.dec_numdoc " & _
            "AND   b.dec_codmer = c.pro_codigo " & _
            "AND   a.toc_fecrem <= " & sql3 & " " & _
            "AND   b.dec_prerec > 0 AND a.toc_tipinf = 'C' AND a.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND a.toc_codbod = " & vg_codbod & " " & _
            "AND  (" & sql4 & " = '' OR (a.toc_docsnc) IS NULL) ORDER BY c.pro_ctacon", vg_db, adOpenStatic
   
   If Not RS3.EOF Then
      
      Do While Not RS3.EOF
         DoEvents
         pctimp = 0
         '-------> Traer Impuesto adicionales
         
         If RS4.State = 1 Then RS4.Close
         RS4.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient

         RS4.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
                  "FROM b_detcomprasimp a, a_impuesto b " & _
                  "WHERE a.imd_rutdoc = '" & RS3!dec_rutpro & "' " & _
                  "AND   a.imd_tipdoc = '" & RS3!dec_tipdoc & "' " & _
                  "AND   a.imd_numdoc = " & RS3!dec_numdoc & " " & _
                  "AND   a.imd_numlin = " & RS3!dec_numlin & " " & _
                  "AND   a.imd_codpro = '" & RS3!dec_codmer & "' " & _
                  "AND   a.imd_codimp = b.imp_codigo " & _
                  "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
         If RS4.EOF Then RS4.Close: Set RS4 = Nothing Else pctimp = RS4!imd_monimp: RS4.Close: Set RS4 = Nothing
         '-------> Fin traer Impuesto adicionales
         pctdes = 0
         If RS3!dec_pctdes > 0 Then pctdes = RS3!dec_pctdes
         
         If RS3!pro_ctacon = (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) Then
            
            tsnantali = Round(tsnantali + ((RS3!dec_ptotal - RS3!dec_ptotrec)) - (((RS3!dec_ptotal - RS3!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
         
         ElseIf RS3!pro_ctacon = (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
            
            tsnantdes = Round(tsnantdes + ((RS3!dec_ptotal - RS3!dec_ptotrec)) - (((RS3!dec_ptotal - RS3!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
         
         ElseIf RS3!pro_ctacon <> (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) And RS3!pro_ctacon <> (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) Then
            
            tsnantgrl = Round(tsnantgrl + ((RS3!dec_ptotal - RS3!dec_ptotrec)) - (((RS3!dec_ptotal - RS3!dec_ptotrec)) * (pctdes / 100)) + pctimp, vg_DPr) 'vg_DCa)
         
         End If
         RS3.MoveNext
      
      Loop
      
      '-------> Actualizar solicitud nota de credito
      sql3 = IIf(vg_tipbase = "1", " CDATE('" & ParFec & "') ", " '" & Format(ParFec, "yyyymmdd") & "' ")
      sql4 = IIf(vg_tipbase = "1", " TRIM(toc_docsnc) ", " lTRIM(toc_docsnc) ")
      sql5 = IIf(vg_tipbase = "1", " 'XXX' & " & Fecha & " ", " 'XXX' + '" & Fecha & "' ")
      
'            vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = " & sql5 & " WHERE toc_fecemi <= " & sql3 & " " & _

'      vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = " & sql5 & " WHERE toc_fecrem <= " & sql3 & " " & _
'                    "AND toc_tipinf = 'C' AND toc_tipdoc = 'SN' AND toc_codbod = " & vg_codbod & " " & _
'                    "AND (" & sql4 & " = '' OR (toc_docsnc) IS NULL)"
   
      vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = " & sql5 & " WHERE toc_fecrem <= " & sql3 & " " & _
                    "AND toc_tipinf = 'C' AND toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND toc_codbod = " & vg_codbod & " " & _
                    "AND (" & sql4 & " = '' OR (toc_docsnc) IS NULL)"
   End If
   RS3.Close
   Set RS3 = Nothing

End If
RS2.Close
Set RS2 = Nothing
'-------> Fin traer solicitud nota creditos pedientes desde parametros

'-------> Actualizar periodo cierre
vg_db.Execute "UPDATE b_cierreperiodo SET cie_proantali = " & ((-tgdantali) + (tsnantali)) + (-(tgdali) + (tscali)) & ", cie_gdpenmesali = " & tgdali & ", cie_gdpenmesantali = " & (tgdantali) & ", cie_sncpenmesali = " & tscali & ", cie_sncpenmesantali = " & (tsnantali) * -1 & ", " & _
                                         "cie_proantgrl = " & ((-tgdantgrl) + (tsnantgrl)) + (-(tgdgrl) + (tscgrl)) & ", cie_gdpenmesgrl = " & tgdgrl & ", cie_gdpenmesantgrl = " & (tgdantgrl) & ", cie_sncpenmesgrl = " & tscgrl & ", cie_sncpenmesantgrl = " & (tsnantgrl) * -1 & ", " & _
                                         "cie_proantdes = " & ((-tgdantdes) + (tsnantdes)) + (-(tgddes) + (tscdes)) & ", cie_gdpenmesdes = " & tgddes & ", cie_gdpenmesantdes = " & (tgdantdes) & ", cie_sncpenmesdes = " & tscdes & ", cie_sncpenmesantdes = " & (tsnantdes) * -1 & " WHERE cie_cencos = '" & cencos & "' AND cie_periodo = " & Fecha & ""
'-------> Fin actualizar periodo cierre
End Function

Function SendMail(oMail As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String)
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
    oMail.LogMailSentFilename = dir_trabajo_Inf & "mailSent.log"
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
    MsgBox Err & ":  " & error$(Err), vbCritical, "Error al enviar mail."
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

Function fg_poneencpagina() As String

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim nomemp As String
nomemp = "Sodexo Chile S.A."


If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT DISTINCT a.par_nombre, a.par_valor FROM a_param as a  with (nolock ) WHERE a.par_codigo = 'nomempresa' AND a.par_cencos = '" & MuestraCasino(1) & "'")
If Not RS1.EOF Then nomemp = Trim(RS1!par_valor) & "Versión " & TipoDato(GetParametro("version"), 0)
RS1.Close
Set RS1 = Nothing

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("SELECT a.cli_nombre, isnull(b.Cecos_AX,'') as Cecos_AX FROM b_clientes as a with (nolock) left join Cecos_Sap_AX as b on b.Cecos_Sap = a.cli_codigo WHERE a.cli_codigo = '" & MuestraCasino(1) & "' AND a.cli_tipo = 0")
If Not RS1.EOF Then fg_poneencpagina = nomemp & VgLinea & "Centro de Costo : " & vg_contra & " - " & Trim(RS1!cli_nombre) & Space(15) & "Centro de Costo OPTIMUM : " & Trim(RS1!Cecos_AX)
RS1.Close
Set RS1 = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Function fg_ponepiepagina() As String
RS1.Open "SELECT DISTINCT par_nombre, par_valor FROM a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
If Not RS1.EOF Then fg_ponepiepagina = "SGP v " & Trim(Str(App.Major)) & "." & Trim(Str(App.Minor)) & "." & Trim(Str(App.Revision) & "  " & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm")) & VgLinea & Trim(RS1!par_nombre) & " : " & CDate(fg_Desencripta(TipoDato(RS1!par_valor, ""))) - 1
RS1.Close: Set RS1 = Nothing
End Function

Function SendMail1(cObj As Object, cSubject As String, cBody As String, cArchivo As String, cNombU As String, CmailU As String, Cmailaviso As Integer) As Boolean
On Error GoTo Man_Error
If Trim(CmailU) <> "" Then
    cObj.UnlockComponent "1mundoedwardsMAIL_3e5SOpaZmRkg"
    cObj.SmtpAuthMethod = "NONE"
    cObj.SmtpHost = "mailserver.sodexhochile.cl"
    cObj.SmtpUsername = ""
    cObj.SmtpPassword = ""
    cObj.ConnectTimeout = 30
    Dim email As ChilkatEmail, Success As Long
    Set email = New ChilkatEmail
'chilkatmail 7.0.1    Dim email As New ChilkatEmail2, Success As Long
    email.AddTo cNombU, CmailU
    email.Subject = cSubject
    email.Body = cBody
    'If Cmailaviso = 1 Then
    email.AddFileAttachment cArchivo
    email.FromName = "Administrador Control Factura Compras" 'IIf(Cmailaviso = 1, "Administrador Control Factura Compras", "")
    email.FromAddress = "adminbdcasinos@sodexho.cl"
    cObj.LogMailSentFilename = "mailSent.log"
    Success = cObj.SendEmail(email)
End If
SendMail1 = True
Exit Function
Man_Error:
    'SendMail1 = False
    SendMail1 = True
    cObj.SaveXmlLog "log.xml"
    On Error Resume Next
'    If Err = -2147467259 Then MsgBox "Cuenta no válida ó no esta conectado a internet" & VgLinea & Trim(cNombU), vbCritical, "Envio Archivos a contratos": Exit Function
'    MsgBox Err & ":  " & Error$(Err), vbCritical, "Error al enviar mail."
End Function

Function ExportHeaderFooter(vp As VSPrinter)
    
    ' no RTF export file? no work!
    If Len(vp.ExportFile) = 0 Then Exit Function
    If vp.ExportFormat < vpxRTF Then Exit Function
    
    ' build rtf style string for headers and foooters
    Dim rtfStyle$
    rtfStyle = ""
'    rtfStyle = "\rtf1\ansi\ansicpg1252\deff0\deflang1033 " & _
'               "{\fonttbl{\f999 {{fname}};}}\li0\tqc\tx{{center}}\tqr\tx{{right}}\f999\fs{{fsize}}"
    vp.GetMargins
    rtfStyle = Replace(rtfStyle, "{{center}}", (vp.X1 + vp.X2) / 2 - vp.X1)
    rtfStyle = Replace(rtfStyle, "{{right}}", vp.X2 - vp.X1)
    rtfStyle = Replace(rtfStyle, "{{fname}}", vp.HdrFontName)
    rtfStyle = Replace(rtfStyle, "{{fsize}}", CInt(2 * vp.HdrFontSize))
    If vp.HdrFontBold Then rtfStyle = rtfStyle & "\b"
    If vp.HdrFontItalic Then rtfStyle = rtfStyle & "\i"
    If vp.HdrFontUnderline Then rtfStyle = rtfStyle & "\ul"
    
    ' output header field
    Dim rtf$, s$, v
    s = vp.Header
    If Len(s) Then
        s = Replace(s, "\", "\\")
        s = Replace(s, "%d", "{\field{\*\fldinst PAGE}}")
        v = Split(s, "|")
        rtf = "{\header{" & rtfStyle & " {{left}} \tab {{center}} \tab {{right}}\par }}"
        rtf = Replace(rtf, "{{left}}", v(0))
        If UBound(v) >= 1 Then rtf = Replace(rtf, "{{center}}", v(1))
        If UBound(v) >= 2 Then rtf = Replace(rtf, "{{right}}", v(2))
        rtf = Replace(rtf, "{{center}}", "")
        rtf = Replace(rtf, "{{right}}", "")
        vp.ExportRaw = rtf
    End If

    ' output footer field
    s = vp.Footer
    If Len(s) Then
        s = Replace(s, "\", "\\")
        s = Replace(s, "%d", "{\field{\*\fldinst PAGE}}")
        v = Split(s, "|")
        rtf = "{\footer{" & rtfStyle & " {{left}} \tab {{center}} \tab {{right}}\par }}"
        rtf = Replace(rtf, "{{left}}", v(0))
        If UBound(v) >= 1 Then rtf = Replace(rtf, "{{center}}", v(1))
        If UBound(v) >= 2 Then rtf = Replace(rtf, "{{right}}", v(2))
        rtf = Replace(rtf, "{{center}}", "")
        rtf = Replace(rtf, "{{right}}", "")
        vp.ExportRaw = rtf
    End If

End Function

Function SetCursorBaseline(vp As VSPrinter8LibCtl.VSPrinter, X!, y!)
'
' This function positions a VSPrinter cursor at a
' position that produces the same results as the
' ta*Baseline settings for VSPrinter3.
'
' It does this by moving the cursor up by one FontSize.
' The FontSize is converted from points to twips.
'
    With vp
        .CurrentX = X
        .CurrentY = y - .FontSize * 1440 / 72
    End With
End Function

Function fg_aletras(ByVal Num As Double, ByVal cMonS As String, cMonP As String, cDecS As String, cDecP As String) As String
Dim Centenas As Variant, Decenas As Variant, Numeros As Variant, cNum As String
Dim Fin As String, Enlet As String
Dim Nu1 As String, Nu2 As String, Nu3 As String, nu4 As String, nu5 As String
Dim nu6 As String, nu7 As String, nu8 As String, nu9 As String
Dim Num1 As String
Num = Round(Num, 0)
If Num > 999999999 Then
   fg_aletras = " ": Exit Function
End If
Centenas = Array("", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
Decenas = Array("", "DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
Numeros = Array("", "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
cNum = String(9 - Len(Trim(Str(Num))), "0") + Trim(Str(Num))
If Val(cNum) = 0 Then
    fg_aletras = " ": Exit Function
End If
Nu1 = Mid(cNum, 1, 1)
Nu2 = Mid(cNum, 2, 1)
Nu3 = Mid(cNum, 3, 1)
nu4 = Mid(cNum, 4, 1)
nu5 = Mid(cNum, 5, 1)
nu6 = Mid(cNum, 6, 1)
nu7 = Mid(cNum, 7, 1)
nu8 = Mid(cNum, 8, 1)
nu9 = Mid(cNum, 9, 1)
Fin = " "
Enlet = " "
'Centenas de Millon
If Nu1 = "1" And Val(Nu2) = 0 And Val(Nu3) = 0 Then
    Enlet = Enlet + " CIEN MILLONES"
Else
    If Val(Nu1) <> 0 Then
        Enlet = Enlet + " " + Centenas(Nu1)
    End If
End If
'Decenas de Millon
If Val(Nu2) <> 0 Then
    Num1 = Mid(cNum, 2, 2)
    If Val(Num1) < 30 Then
        Enlet = Enlet + " " + Numeros(Num1) + " MILLONES"
    Else
        Enlet = Enlet + " " + Decenas(Nu2)
        If Val(Nu3) <> 0 Then
            Enlet = Enlet + " Y " + Numeros(Nu3)
        End If
        Enlet = Enlet + " MILLONES"
    End If
End If
'unidades de millon
If Val(Nu3) <> 0 And Val(Nu2) = 0 Then
    If Nu3 = "1" And Val(Nu1) = 0 Then
        Enlet = Enlet + " UN MILLON"
    Else
        Enlet = Enlet + " " + Numeros(Nu3) + " MILLONES"
    End If
End If
'centenas de mil
If nu4 = "1" And Val(nu5) = 0 And Val(nu6) = 0 Then
    Enlet = Enlet + " CIEN"
Else
    If Val(nu4) <> 0 Then
        Enlet = Enlet + " " + Centenas(nu4)
    End If
End If
'decenas de mil
If Val(nu5) <> 0 Then
    Num1 = Mid(cNum, 5, 2)
    If Val(Num1) < 30 Then
        Enlet = Enlet + " " + Numeros(Num1)
    Else
        Enlet = Enlet + " " + Decenas(nu5)
        If Val(nu6) <> 0 Then
            Enlet = Enlet + " Y " + Numeros(nu6)
        End If
    End If
End If
'unidades de mil
If Val(nu5) <> 0 Or Val(nu4) <> 0 And Val(nu6) = 0 Then
    Enlet = Enlet + " MIL"
Else
    If Val(nu5) = 0 And Val(nu6) <> 0 Then
        If Val(nu4) = 0 And nu6 = "1" Then
            Enlet = Enlet + " UN MIL"
        Else
            Enlet = Enlet + " " + Numeros(nu6) + " MIL"
        End If
    End If
End If
'centenas
If nu7 = "1" And Val(nu8) = 0 And Val(nu9) = 0 Then
    Enlet = Enlet + " CIEN"
Else
    If Val(nu7) <> 0 Then
        Enlet = Enlet + " " + Centenas(nu7)
    End If
End If
'decenas
If Val(nu8) <> 0 Then
    Num1 = Mid(cNum, 8, 2)
    If Val(Num1) < 30 Then
        Enlet = Enlet + " " + Numeros(Num1)
    Else
        Enlet = Enlet + " " + Decenas(nu8)
        If Val(nu9) <> 0 Then
            Enlet = Enlet + " Y " + Numeros(nu9)
        Else
            Enlet = Enlet
        End If
    End If
End If
'unidades
If Val(nu9) <> 0 And Val(nu8) = 0 Then
    If Val(nu7) = 0 And nu9 = "1" And Val(nu8) = 0 Then
        Enlet = Enlet & " UN " & cMonS
    Else
        Enlet = Enlet & " " & Numeros(nu9) & " " & cMonP
    End If
Else
    Enlet = Enlet & " " & cMonP
End If
If Trim(Enlet) = "" Then Enlet = "CERO PESOS"
fg_aletras = Enlet
End Function

Function RecalPrecioDoc(Fecha As Long, codbod As Long, codpro As String)
On Local Error GoTo Error_RecalPrecioDoc
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String
Dim fecini As Long, fecfin As Long, i As Long
Dim vecpro() As Variant
'00 Inventario inicial
'10 Ajuste de Entrada
'20 Proveedores Entrada
'30 Traspaso Entrada
'40 Produccion Salida
'50 Produccion Entrada
'60 Traspaso Salida
'70 Mermas Salida
'80 Venta Directa Salida
'90 Ajuste Salida
'100 Venta Cafeteria

'-------> guardar productos en vector
If Trim(codpro) = "" Then Exit Function
RS1.Open "SELECT COUNT(pro_codigo) AS nreg FROM b_productos WHERE pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") AND pro_ctrsto=1", vg_db, adOpenStatic
If RS1.EOF Or RS1!nreg < 1 Then RS1.Close: Set RS1 = Nothing: Exit Function
ReDim vecpro(RS1!nreg, 3)
RS1.Close: Set RS1 = Nothing
RS1.Open "SELECT a.pro_codigo, b.ppd_propon " & _
         "FROM b_productos a, b_productospmpdia b " & _
         "WHERE a.pro_codigo = b.ppd_codpro " & _
         "AND   b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   b.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
         "AND   a.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") " & _
         "AND   a.pro_ctrsto = 1", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Function
i = 1
Do While Not RS1.EOF
   vecpro(i, 1) = RS1!pro_codigo
   vecpro(i, 2) = RS1!ppd_propon
   vecpro(i, 3) = 0
   RS1.MoveNext: i = i + 1
Loop
RS1.Close: Set RS1 = Nothing
'-------> buscar ultimo inventario
RS1.Open "SELECT max(tin_fectom) AS tin_fectom FROM b_tomainv WHERE tin_fectom < " & Fecha & " AND tin_codbod=" & codbod & "", vg_db, adOpenStatic
If RS1.EOF Or IsNull(RS1!tin_fectom) Then
   RS1.Close: Set RS1 = Nothing
   DoEvents
   '-------> actualizar precio que no estan en documento de entrada
   vg_db.BeginTrans
   For i = 1 To UBound(vecpro)
       If vecpro(i, 3) = 0 Then
'*          vg_db.Execute "UPDATE b_contlistprepro SET cpp_propon=0, cpp_upreco=0, cpp_fecuco=NULL WHERE cpp_cencos='" & MuestraCasino(1) & "' AND cpp_codigo='" & vecpro(i, 1) & "'"
'*          vg_db.Execute "UPDATE b_contlistpreing INNER JOIN (b_productos INNER JOIN b_productosing ON b_productos.pro_codigo = b_productosing.pri_codpro) ON b_contlistpreing.cpi_coding = b_productosing.pri_coding SET b_contlistpreing.cpi_precos=0, b_contlistpreing.cpi_feccos=" & Format(Date, "yyyymmdd") & " " & _
'*                        "WHERE b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "' AND b_productosing.pri_codpro='" & vecpro(i, 1) & "'"
       End If
   Next i
   vg_db.CommitTrans
   Exit Function
End If
fecini = RS1!tin_fectom
RS1.Close: Set RS1 = Nothing
'-------> buscar fecha termino
RS1.Open "SELECT cie_fecter FROM b_cierreperiodo WHERE cie_cencos='" & MuestraCasino(1) & "' AND cie_periodo=" & Val(Mid(Fecha, 1, 6)) & " AND cie_estado=1", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Function
fecfin = RS1!cie_fecter
RS1.Close: Set RS1 = Nothing
vg_db.BeginTrans
'-------> Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_RecalPrecioDoc"
fg_CheckTmp aAp
'         "AND   tin_stofis>0 " &
fg_carga ""
RS1.Open "SELECT DISTINCT tin_fectom AS fecpro, tin_codpro AS codpro, tin_stofis AS cansto, tin_propon AS propon, " & _
         "'E' AS tipmov , 0 AS numdoc, 'E' AS tipdoc, 'E' AS rutcli, '000' AS orden INTO " & aAp & " FROM b_tomainv " & _
         "WHERE tin_fectom=" & fecini & " " & _
         "AND   tin_propon>0 " & _
         "AND   tin_codbod=" & codbod & " " & _
         "AND   tin_codpro IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") ORDER BY tin_fectom, tin_codpro", vg_db, adOpenStatic
Set RS1 = Nothing
DoEvents
'-------> Fin traer Inventario primer inventario

'-------> Traer ajuste inventario
'vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
'              "dev.dev_canmin AS cansto, dev.dev_precos AS propon, IIF(aju.aju_codigo=3,'E','S') AS tipmov, " & _
'              "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, IIF(aju.aju_tipo = 'A', '10', '90') AS orden " & _
'              "FROM b_totventas tov, b_detventas dev, b_productos pro, a_tipoajuste aju " & _
'              "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
'              "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
'              "AND   tov.tov_numdoc=dev.dev_numdoc " & _
'              "AND   dev.dev_codmer=pro.pro_codigo " & _
'              "AND   tov.tov_codser=aju.aju_codigo " & _
'              "AND   tov.tov_codbod=" & CodBod & " " & _
'              "AND   format(tov.tov_fecemi, 'yyyymmdd')>=" & fecini & " " & _
'              "AND   format(tov.tov_fecemi, 'yyyymmdd')<=" & fecfin & " " & _
'              "AND   tov.tov_tipdoc='AI' AND tov.tov_estdoc<>'A' AND pro.pro_ctrsto=1 AND pro.pro_codigo IN (" & Mid(CodPro, 1, Len(CodPro) - 1) & ") ORDER BY tov.tov_fecemi, pro.pro_codigo"
'-------> Fin traer ajuste inventario
    
'-------> Traer salida y devolución produción
vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecpro, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, " & _
              "tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, IIF(tov.tov_tipdoc='SP','40','50') AS orden " & _
              "FROM b_totventas tov, b_detventas dev, b_productos pro " & _
              "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
              "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
              "AND   tov.tov_numdoc=dev.dev_numdoc " & _
              "AND   dev.dev_codmer=pro.pro_codigo " & _
              "AND   pro.pro_ctrsto=1 " & _
              "AND  (tov.tov_tipdoc='SP' or tov.tov_tipdoc='DP') " & _
              "AND   dev.dev_canmer<>0 and tov.tov_estdoc<>'A' " & _
              "AND   tov.tov_codbod=" & codbod & " " & _
              "AND   tov.tov_fecpro>cdate('" & fg_Ctod1(fecini) & "')" & _
              "AND   tov.tov_fecpro<=cdate('" & fg_Ctod1(fecfin) & "')" & _
              "AND   pro.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
'-------> Fin traer salida y devolución produción
    
'-------> Traer mermas
vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, " & _
              "tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '70' AS orden " & _
              "FROM  b_totventas tov, b_detventas dev, b_productos pro " & _
              "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
              "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
              "AND   tov.tov_numdoc=dev.dev_numdoc " & _
              "AND   dev.dev_codmer=pro.pro_codigo " & _
              "AND   tov.tov_tipdoc='ME' and tov.tov_estdoc<>'A' " & _
              "AND   tov.tov_codbod=" & codbod & " " & _
              "AND   format(tov.tov_fecemi,'yyyymmdd')>" & fecini & "" & _
              "AND   format(tov.tov_fecemi,'yyyymmdd')<=" & fecfin & "" & _
              "AND   pro.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
'-------> Fin traer mermas
    
'-------> Traer documento traspaso entrada
'              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, IIF(tov.tov_codreg=0,'E','S') AS tipmov, " & _
'              "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, IIF(tov.tov_codreg=0,'30','60') AS orden " & _

vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
              "dev.dev_canmer AS cansto, dev.dev_precos AS propon, 'E' AS tipmov, " & _
              "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '30' AS orden " & _
              "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
              "WHERE  tov.tov_rutcli=dev.dev_rutcli " & _
              "AND    tov.tov_tipdoc=dev.dev_tipdoc " & _
              "AND    tov.tov_numdoc=dev.dev_numdoc " & _
              "AND    tov.tov_codbod=" & codbod & " " & _
              "AND    tov.tov_codreg=0 " & _
              "AND    dev.dev_codmer=pro.pro_codigo " & _
              "AND    pro.pro_ctrsto=1 " & _
              "AND    tov.tov_tipdoc='TR' AND tov.tov_estdoc<>'A' " & _
              "AND    format(tov.tov_fecemi,'yyyymmdd')>" & fecini & " AND format(tov.tov_fecemi,'yyyymmdd')<=" & fecfin & " AND dev.dev_canmer>0 AND pro.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") ORDER BY tov.tov_fecemi, pro.pro_codigo"
'-------> Fin traer documento traspaso entrada
  
'-------> Traer documento ventas cafeteria
vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(b.tvc_fecing, 'yyyymmdd') AS fecpro, c.pro_codigo AS codpro, " & _
              "a.dvp_candig AS cansto, a.dvp_precos AS propon, 'S' AS tipmov, " & _
              "0 as numdoc, '' AS tipdoc, b.tvc_cencos as rutcli, '100' AS orden FROM b_detventascafpro a, b_totventascaf b, b_productos c " & _
              "WHERE b.tvc_cencos=a.dvp_cencos " & _
              "AND   b.tvc_fecing=a.dvp_fecing " & _
              "AND   a.dvp_codmer=c.pro_codigo " & _
              "AND   c.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") " & _
              "AND   c.pro_ctrsto=1 " & _
              "AND   a.dvp_precos<>0 " & _
              "AND   b.tvc_fecing>cdate('" & fg_Ctod1(fecini) & "') AND b.tvc_fecing<=cdate('" & fg_Ctod1(fecfin) & "') " & _
              "AND   b.tvc_codbod=" & codbod & " " & _
              ""
'-------> Fin traer documento ventas cafeteria

'-------> Traer Documento Proveedor
Dim pctimp As Double, pctdes  As Double, Precio As Double
'         "AND   format(toc.toc_fecemi, 'yyyymmdd')>'" & fecini & "' " & _
'         "AND   format(toc.toc_fecemi, 'yyyymmdd')<='" & fecfin & "' " & _

'RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, de.dec_numdoc, " & _
'         "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
'         "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
'         "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
'         "WHERE toc.toc_rutpro=de.dec_rutpro " & _
'         "AND   toc.toc_tipdoc=de.dec_tipdoc " & _
'         "AND   toc.toc_numdoc=de.dec_numdoc " & _
'         "AND   de.dec_codmer=pro.pro_codigo " & _
'         "AND   de.dec_mueinv='S' and toc.toc_tipdoc<>'SN' " & _
'         "AND   de.dec_canrec>0 " & _
'         "AND   toc.toc_codbod=" & codbod & " " & _
'         "AND   format(toc.toc_fecrem, 'yyyymmdd')>'" & fecini & "' " & _
'         "AND   format(toc.toc_fecrem, 'yyyymmdd')<='" & fecfin & "' " & _
'         "AND   pro.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") " & _
'         "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic

RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, de.dec_numdoc, " & _
         "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
         "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
         "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
         "WHERE toc.toc_rutpro=de.dec_rutpro " & _
         "AND   toc.toc_tipdoc=de.dec_tipdoc " & _
         "AND   toc.toc_numdoc=de.dec_numdoc " & _
         "AND   de.dec_codmer=pro.pro_codigo " & _
         "AND   de.dec_mueinv='S' and toc.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo <> 'SN') " & _
         "AND   de.dec_canrec>0 " & _
         "AND   toc.toc_codbod=" & codbod & " " & _
         "AND   format(toc.toc_fecrem, 'yyyymmdd')>'" & fecini & "' " & _
         "AND   format(toc.toc_fecrem, 'yyyymmdd')<='" & fecfin & "' " & _
         "AND   pro.pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ") " & _
         "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic

If Not RS1.EOF Then
   Do While Not RS1.EOF
      pctimp = 0: Precio = 0
      RS2.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
               "FROM  b_detcomprasimp a, a_impuesto b " & _
               "WHERE a.imd_rutdoc='" & RS1!dec_rutpro & "' " & _
               "AND   a.imd_tipdoc='" & RS1!dec_tipdoc & "' " & _
               "AND   a.imd_numdoc=" & RS1!dec_numdoc & " " & _
               "AND   a.imd_numlin=" & RS1!dec_numlin & " " & _
               "AND   a.imd_codpro='" & RS1!pro_codigo & "' " & _
               "AND   a.imd_codimp=b.imp_codigo " & _
               "AND   b.imp_inccos=1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
      If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
      pctdes = 0
      If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
      If RS1!dec_prefle > 0 Then
         Precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec)) + (RS1!dec_prefle / RS1!dec_canrec)
      Else
         Precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec))
      End If
      vg_db.Execute "INSERT INTO " & aAp & " VALUES (" & Val(Format(RS1!toc_fecrem, "yyyymmdd")) & ", " & _
                    "'" & Trim(RS1!pro_codigo) & "', " & RS1!dec_canrec & ", " & Precio & ", '" & "E+" & "', " & RS1!dec_numdoc & ", '" & Trim(RS1!dec_tipdoc) & "', '" & Trim(RS1!dec_rutpro) & "', '20')"
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
'-------> Fin traer Documento Proveedor
   
'-------> Procesar Información Precio Promedio Ponderado
Dim auxCanmer As Double, auxPropon As Double, propon As Double, auxfec As Long, fecuco As String, upreco As Double
Dim auxcodpro As String, auxtipdoc As String
RS1.Open "SELECT * FROM " & aAp & " ORDER BY codpro, fecpro, orden, tipmov", vg_db, adOpenForwardOnly ', adOpenStatic
If Not RS1.EOF Then
   auxCanmer = 0: auxPropon = 0: propon = 0: auxfec = 0: auxcodpro = "": auxtipdoc = "": fecuco = "": upreco = 0
   Do While Not RS1.EOF
      If RS1!codpro <> auxcodpro Then
         For i = 1 To UBound(vecpro)
             If RS1!codpro = vecpro(i, 1) Then vecpro(i, 3) = 1: Exit For
         Next i
         If Trim(auxcodpro) <> "" Then
            '-------> Actualizar maestro producto y ingrediente
'*            vg_db.Execute "UPDATE b_contlistprepro SET cpp_propon=(" & propon & " ), cpp_upreco=" & IIf(upreco = 0, 0, upreco) & ", cpp_fecuco=" & IIf(Trim(fecuco) = "", "Null", fecuco) & " WHERE cpp_cencos='" & MuestraCasino(1) & "' AND cpp_codpro='" & auxcodpro & "'"
'*            vg_db.Execute "UPDATE b_contlistpreing INNER JOIN (b_productos INNER JOIN b_productosing ON b_productos.pro_codigo = b_productosing.pri_codpro) ON b_contlistpreing.cpi_coding = b_productosing.pri_coding SET b_contlistpreing.cpi_precos=((" & propon & " )/b_productos.pro_facing), b_contlistpreing.cpi_feccos=" & Format(Date, "yyyymmdd") & " " & _
'*                          "WHERE b_productosing.pri_codpro='" & auxcodpro & "' AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "'"
            '-------> Fin actualizar maestro producto y ingrediente
         End If
         auxcodpro = RS1!codpro:  auxCanmer = 0: auxPropon = 0: propon = 0: upreco = 0: fecuco = ""
      End If
      If RS1!tipmov = "S" Then
         '-------> Actualizar ajuste-salida y devolución producción-mermas
         If propon > 0 Then
            If RS1!Orden = "100" Then
               '-------> Actualizar cafeteria
               vg_db.Execute "UPDATE b_totventascaf INNER JOIN b_detventascafpro ON (b_totventascaf.tvc_fecing = b_detventascafpro.dvp_fecing) AND (b_totventascaf.tvc_cencos = b_detventascafpro.dvp_cencos) " & _
                             "SET b_detventascafpro.dvp_precos=" & propon & " WHERE b_totventascaf.tvc_cencos='" & RS1!rutcli & "' AND b_totventascaf.tvc_fecing=cdate('" & fg_Ctod1(RS1!fecpro) & "') AND b_totventascaf.tvc_codbod=" & codbod & " AND b_detventascafpro.dvp_codmer='" & RS1!codpro & "'"
            Else
               '-------> Actualizar encabezado y detalle ventas
               vg_db.Execute "UPDATE b_totventas INNER JOIN b_detventas ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli) SET b_detventas.dev_precos = " & propon & ", b_detventas.dev_predoc = " & propon & ", b_detventas.dev_ptotal=(" & propon & " * " & RS1!cansto & ") " & _
                             "WHERE b_totventas.tov_numdoc=" & RS1!NumDoc & " AND b_totventas.tov_tipdoc='" & RS1!tipdoc & "' AND b_totventas.tov_rutcli='" & RS1!rutcli & "' AND b_detventas.dev_codmer='" & RS1!codpro & "' AND b_totventas.tov_estdoc<>'A' AND b_totventas.tov_codbod=" & codbod & ""
               RS2.Open "SELECT SUM(b.dev_ptotal) AS ptotal FROM b_totventas a, b_detventas b WHERE a.tov_numdoc=b.dev_numdoc AND a.tov_tipdoc=b.dev_tipdoc AND a.tov_rutcli=b.dev_rutcli AND a.tov_rutcli='" & RS1!rutcli & "' AND a.tov_tipdoc='" & RS1!tipdoc & "' AND a.tov_numdoc=" & RS1!NumDoc & " AND a.tov_codbod=" & codbod & " GROUP BY a.tov_rutcli", vg_db, adOpenForwardOnly ', adOpenStatic
               If Not RS2.EOF Then
                  vg_db.Execute "UPDATE b_totventas SET b_totventas.tov_totdoc=" & RS2!ptotal & " " & _
                                "WHERE b_totventas.tov_estdoc<>'A' AND b_totventas.tov_rutcli='" & RS1!rutcli & "' AND b_totventas.tov_tipdoc='" & RS1!tipdoc & "' AND b_totventas.tov_numdoc=" & RS1!NumDoc & " AND b_totventas.tov_codbod=" & codbod & " AND " & RS2!ptotal & " > 0 AND Not IsNull(" & RS2!ptotal & ")"
               End If
               RS2.Close: Set RS2 = Nothing
            End If
            If RS1!Orden = "50" Then auxCanmer = (auxCanmer + RS1!cansto) Else auxCanmer = (auxCanmer - RS1!cansto)
         End If
         '-------> Fin actualizar ajuste-salida y devolución producción-mermas-cafeteria
      Else
         propon = Round(((auxPropon * IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer)) + (RS1!propon * IIf(RS1!Orden = "000" And RS1!cansto <= 0, 1, RS1!cansto))) / (IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer) + IIf(RS1!Orden = "000" And RS1!cansto <= 0, 1, RS1!cansto)), 0)
         auxPropon = propon: auxCanmer = (IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer) + IIf(RS1!Orden = "000" And RS1!cansto <= 0, 1, RS1!cansto))
         If RS1!tipmov = "E+" Then fecuco = CDate(fg_Ctod1(RS1!fecpro)): upreco = RS1!propon
      End If
      RS1.MoveNext
   Loop
'*   vg_db.Execute "UPDATE b_contlistprepro SET cpp_propon=(" & propon & " ), cpp_upreco=" & IIf(upreco = 0, 0, upreco) & ", cpp_fecuco=" & IIf(Trim(fecuco) = "", "Null", fecuco) & " WHERE cpp_cencos='" & MuestraCasino(1) & "' AND cpp_codpro='" & auxcodpro & "'"
'*   vg_db.Execute "UPDATE b_contlistpreing INNER JOIN (b_productos INNER JOIN b_productosing ON b_productos.pro_codigo = b_productosing.pri_codpro) ON b_contlistpreing.cpi_coding = b_productosing.pri_coding SET b_contlistpreing.cpi_precos=((" & propon & " )/b_productos.pro_facing), b_contlistpreing.cpi_feccos=" & Format(Date, "yyyymmdd") & " " & _
'*                 "WHERE b_productosing.pri_codpro='" & auxcodpro & "' AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "'"
End If
RS1.Close: Set RS1 = Nothing
'-------> actualizar precio que no estan en documento de entrada
For i = 1 To UBound(vecpro)
    If vecpro(i, 3) = 0 Then
'*       vg_db.Execute "UPDATE b_contlistprepro SET cpp_propon=0, cpp_upreco=0, cpp_fecuco=NULL WHERE cpp_cencos='" & MuestraCasino(1) & "' AND cpp_codpro='" & vecpro(i, 1) & "'"
'*       vg_db.Execute "UPDATE b_contlistpreing INNER JOIN (b_productos INNER JOIN b_productosing ON b_productos.pro_codigo = b_productosing.pri_codpro) ON b_contlistpreing.cpi_coding = b_productosing.pri_coding SET b_contlistpreing.cpi_precos=0, b_contlistpreing.cpi_feccos=" & Format(Date, "yyyymmdd") & " " & _
'*                     "WHERE b_productosing.pri_codpro='" & vecpro(i, 1) & "' AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "'"
    End If
Next i
fg_descarga
vg_db.CommitTrans

Exit Function
Error_RecalPrecioDoc:
    fg_descarga
    MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
    vg_db.RollbackTrans
    Resume Next
End Function

Function fg_pone_espacio(ByVal cadena As String, ByVal cuanto As Integer) As String
'-------> pone ceros a la izquierda
fg_pone_espacio = ""
If cadena <> "" Then
   Do While Len(cadena) < cuanto
      cadena = cadena + " "
   Loop
   fg_pone_espacio = cadena
End If
End Function

Function RecalcularStock(codbod As Long, codpro As String, fecini As Long, fecfin As Long)
Dim auxpro As String, tanali As Double, i As Long, j As Long
On Local Error GoTo Error_RecalcularStock
fg_carga ""

'-------> 00 Inventario inicial
'-------> 10 Ajuste de Entrada
'-------> 20 Proveedores Entrada
'-------> 30 Traspaso Entrada
'-------> 40 Prodccion Salida
'-------> 50 Produccion Entrada
'-------> 60 Traspaso Salida
'-------> 70 Mermas Salida
'-------> 80 Venta Directa Salida
'-------> 90 Ajuste Salida
    V_Acceso.Label1(0).Visible = True
    V_Acceso.Label1(1).Visible = True
    V_Acceso.Label1(1).Caption = "Recalcular Stock : "
    RS3.Open "SELECT distinct a.* FROM b_productos a, b_totventas b, b_detventas c " & _
              "WHERE b.tov_rutcli=c.dev_rutcli " & _
              "AND   b.tov_tipdoc=c.dev_tipdoc " & _
              "AND   b.tov_numdoc=c.dev_numdoc " & _
              "AND   c.dev_codmer=a.pro_codigo " & _
              "AND   b.tov_codbod=" & codbod & " " & _
              "AND  format(b.tov_fecemi, 'yyyymmdd')>'" & fecini & "' " & _
              "AND  format(b.tov_fecemi, 'yyyymmdd')<='" & fecfin & "' " & _
              "AND (a.pro_codigo='" & codpro & "' or '" & codpro & "'='') ORDER BY a.pro_nombre", vg_db, adOpenStatic
    Do While Not RS3.EOF
        V_Acceso.Label1(0).Caption = Trim(RS3!pro_nombre)
        DoEvents
        codpro = RS3!pro_codigo
        V_Acceso.vaSpread1.MaxRows = 0: V_Acceso.vaSpread1.MaxCols = 12
        Dim fectom As Long, fectin As Long
        fectom = 0: fectin = fecini
        '------- Traer Inventario del mes anterior
        RS1.Open "SELECT tin.tin_fectom, pro.pro_codigo, pro.pro_nombre, tin.tin_stofis, tin_propon " & _
                 "FROM b_tomainv tin, b_productos pro " & _
                 "WHERE tin.tin_codpro=pro.pro_codigo " & _
                 "AND  (tin.tin_codbod=" & codbod & " OR " & codbod & "=0) " & _
                 "AND   tin.tin_codpro='" & codpro & "' " & _
                 "AND   tin.tin_fectom<=" & fecini & " " & _
                 "ORDER BY tin.tin_fectom DESC", vg_db, adOpenStatic
        If Not RS1.EOF Then
            fectin = RS1!tin_fectom
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = RS1!tin_fectom
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Inventario Inicial"
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = ""
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = "....."
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!tin_stofis
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!tin_propon
            V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "00"
        End If
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer inventario del mes anterior
        '------- Traer ajuste inventario
        RS1.Open "SELECT tov.tov_fecemi, pro.pro_codigo, pro.pro_nombre, aju.aju_nombre, aju.aju_tipo, tov.tov_numdoc, dev.dev_canmin, dev.dev_precos " & _
                 "FROM b_totventas tov, b_detventas dev, b_productos pro, a_tipoajuste aju " & _
                 "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
                 "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
                 "AND   tov.tov_numdoc=dev.dev_numdoc " & _
                 "AND   tov.tov_codser=aju.aju_codigo " & _
                 "AND   dev.dev_codmer=pro.pro_codigo " & _
                 "AND   tov.tov_codbod=" & codbod & " " & _
                 "AND   dev.dev_codmer='" & codpro & "' " & _
                 "AND   format(tov.tov_fecemi, 'yyyymmdd')>" & fectin & " " & _
                 "AND   format(tov.tov_fecemi, 'yyyymmdd')<=" & fecfin & " " & _
                 "AND   tov.tov_tipdoc='AI' " & _
                 "AND   tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!tov_fecemi, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = IIf(RS1!aju_tipo = "A", "Ajuste Inventario (+)", "Ajuste Inventario (-)")
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = ""
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = RS1!tov_numdoc
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dev_canmin
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!dev_precos
            V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = IIf(RS1!aju_tipo = "A", "25", "90")
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin ajuste inventario
        '------- Traer Documento Proveedor
        Dim pctimp As Double, pctdes  As Double
'                 "AND  format(toc.toc_fecemi, 'yyyymmdd')>'" & fectin & "' " & _
'                 "AND  format(toc.toc_fecemi, 'yyyymmdd')<='" & fecfin & "' " & _

'        RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, pro.pro_nombre, de.dec_numdoc, " & _
'                 "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
'                 "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
'                 "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
'                 "WHERE toc.toc_rutpro=de.dec_rutpro AND toc.toc_tipdoc=de.dec_tipdoc " & _
'                 "AND   toc.toc_numdoc=de.dec_numdoc AND de.dec_codmer=pro.pro_codigo " & _
'                 "AND   toc.toc_codbod=" & codbod & " " & _
'                 "AND   pro.pro_codigo='" & codpro & "' " & _
'                 "AND   de.dec_mueinv='S' AND toc.toc_tipdoc<>'SN' AND de.dec_canrec>0 " & _
'                 "AND  format(toc.toc_fecrem, 'yyyymmdd')>'" & fectin & "' " & _
'                 "AND  format(toc.toc_fecrem, 'yyyymmdd')<='" & fecfin & "' " & _
'                 "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic
        
        RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, pro.pro_nombre, de.dec_numdoc, " & _
                 "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
                 "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
                 "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
                 "WHERE toc.toc_rutpro=de.dec_rutpro AND toc.toc_tipdoc=de.dec_tipdoc " & _
                 "AND   toc.toc_numdoc=de.dec_numdoc AND de.dec_codmer=pro.pro_codigo " & _
                 "AND   toc.toc_codbod=" & codbod & " " & _
                 "AND   pro.pro_codigo='" & codpro & "' " & _
                 "AND   de.dec_mueinv='S' AND toc.toc_tipdoc IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo <> 'SN') AND de.dec_canrec>0 " & _
                 "AND  format(toc.toc_fecrem, 'yyyymmdd')>'" & fectin & "' " & _
                 "AND  format(toc.toc_fecrem, 'yyyymmdd')<='" & fecfin & "' " & _
                 "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic
        
        Do While Not RS1.EOF
            pctimp = 0
            RS2.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
                     "FROM  b_detcomprasimp a, a_impuesto b " & _
                     "WHERE a.imd_rutdoc='" & RS1!dec_rutpro & "' " & _
                     "AND   a.imd_tipdoc='" & RS1!dec_tipdoc & "' " & _
                     "AND   a.imd_numdoc=" & RS1!dec_numdoc & " " & _
                     "AND   a.imd_numlin=" & RS1!dec_numlin & " " & _
                     "AND   a.imd_codpro='" & RS1!pro_codigo & "' " & _
                     "AND   a.imd_codimp=b.imp_codigo " & _
                     "AND   b.imp_inccos=1  GROUP BY b.imp_inccos", vg_db, adOpenStatic
            If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
            pctdes = 0: tanali = 0
            If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
            
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!toc_fecrem, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Entrada"
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = Trim(RS1!dec_tipdoc)
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = RS1!dec_numdoc
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dec_canrec
            V_Acceso.vaSpread1.Col = 8
            If RS1!dec_prefle > 0 Then
               V_Acceso.vaSpread1.text = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec)) + (RS1!dec_prefle / RS1!dec_canrec)
            Else
               V_Acceso.vaSpread1.text = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec))
            End If
            V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "10"
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer Documento Proveedor
        '------- Traer documento traspaso entrada
        RS1.Open "SELECT tov.tov_fecemi, tov.tov_codser, pro.pro_codigo, pro.pro_nombre, tov.tov_numdoc, dev.dev_canmer, dev.dev_precos " & _
                 "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE  tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                 "AND    tov.tov_numdoc=dev.dev_numdoc AND dev.dev_codmer=pro.pro_codigo " & _
                 "AND    tov.tov_codbod=" & codbod & " AND pro.pro_codigo='" & codpro & "' " & _
                 "AND    tov.tov_tipdoc='TR' AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' " & _
                 "AND    format(tov.tov_fecemi, 'yyyymmdd')>" & fectin & " " & _
                 "AND    format(tov.tov_fecemi,'yyyymmdd')<=" & fecfin & " AND dev.dev_mueinv='S' " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!tov_fecemi, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            If RS1!tov_codser = 1 Then
                V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Entrada"
                V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "20"
            Else
                V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Salida"
                V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "60"
            End If
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = "TR"
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = RS1!tov_numdoc
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dev_canmer
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!dev_precos
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer documento traspaso entrada
        '------- Traer devoluciones de entrada y salida
        RS1.Open "SELECT tov.tov_fecpro, pro.pro_codigo, pro.pro_nombre, tov.tov_numdoc, tov.tov_tipdoc, dev.dev_canmer, dev.dev_precos " & _
                 "FROM b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                 "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_codmer=pro.pro_codigo " & _
                 "AND   tov.tov_codbod=" & codbod & " " & _
                 "AND   pro.pro_codigo='" & codpro & "' AND dev.dev_mueinv='S' " & _
                 "AND  (tov.tov_tipdoc='SP' OR tov.tov_tipdoc='DP') " & _
                 "AND   dev.dev_canmer<>0 AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' " & _
                 "AND   tov.tov_fecpro>cdate('" & fg_Ctod1(fectin) & "') " & _
                 "AND   tov.tov_fecpro<=cdate('" & fg_Ctod1(fecfin) & "') " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!tov_fecpro, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            If RS1!tov_tipdoc = "SP" Then
                V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Salida"
                V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "40"
            Else
                V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Entrada"
                V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "50"
            End If
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = Trim(RS1!tov_tipdoc)
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = RS1!tov_numdoc
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dev_canmer
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!dev_precos
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer devoluciones de entrada y salida
        '------- Traer Ventas Cafeterias
        RS1.Open "SELECT tvc.tvc_fecing, pro.pro_codigo, pro.pro_nombre, 0 as tvc_numdoc, 'VC' as tvc_tipdoc, dvp.dvp_candig, dvp.dvp_precos " & _
                 "FROM b_totventascaf tvc, b_detventascafpro dvp, b_productos pro " & _
                 "WHERE tvc.tvc_cencos=dvp.dvp_cencos " & _
                 "AND   tvc.tvc_fecing=dvp.dvp_fecing " & _
                 "AND   dvp.dvp_codmer=pro.pro_codigo " & _
                 "AND   tvc.tvc_fecing>cdate('" & fg_Ctod1(fectin) & "') " & _
                 "AND   tvc.tvc_fecing<=cdate('" & fg_Ctod1(fecfin) & "') " & _
                 "AND   tvc.tvc_codbod=" & codbod & " " & _
                 "AND   pro.pro_codigo='" & codpro & "' AND pro.pro_ctrsto=1 " & _
                 "AND   tvc.tvc_estado='C' AND dvp.dvp_candig<>0 " & _
                 "ORDER BY tvc.tvc_fecing, pro.pro_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!tvc_fecing, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            If RS1!tvc_tipdoc = "VC" Then
               V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Salida"
               V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "40"
            End If
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = Trim(RS1!tvc_tipdoc)
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = IIf(RS1!tvc_numdoc = 0, "", RS1!tvc_numdoc)
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dvp_candig
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!dvp_precos
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer Ventas Cafeterias
        '------- Traer Mermas
        RS1.Open "SELECT tov.tov_fecemi, pro.pro_codigo, pro.pro_nombre, tov.tov_numdoc, tov.tov_tipdoc, dev.dev_canmer, dev.dev_precos " & _
                 "FROM  b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                 "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_codmer=pro.pro_codigo " & _
                 "AND   tov.tov_codbod=" & codbod & " " & _
                 "AND   pro.pro_codigo='" & codpro & "' AND tov.tov_tipdoc='ME' AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' " & _
                 "AND   format(tov.tov_fecemi, 'yyyymmdd')>" & fectin & " " & _
                 "AND   format(tov.tov_fecemi, 'yyyymmdd')<=" & fecfin & " " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!tov_fecemi, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Salida"
            V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "70"
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = Trim(RS1!tov_tipdoc)
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = RS1!tov_numdoc
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dev_canmer
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!dev_precos
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer Mermas
        '------- Traer ventas directas
        RS1.Open "SELECT tov.tov_fecemi, pro.pro_codigo, pro.pro_nombre, tov.tov_numdoc, tov.tov_tipdoc, dev.dev_canmer, dev.dev_precos " & _
                 "FROM  b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE tov.tov_rutcli=dev.dev_rutcli AND tov.tov_tipdoc=dev.dev_tipdoc " & _
                 "AND   tov.tov_numdoc=dev.dev_numdoc AND dev.dev_codmer=pro.pro_codigo " & _
                 "AND   tov.tov_codbod=" & codbod & " AND pro.pro_codigo='" & codpro & "' " & _
                 "AND  (tov.tov_tipdoc='FA' OR tov.tov_tipdoc='FE' or tov.tov_tipdoc='GD') AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' AND dev.dev_mueinv='S' " & _
                 "AND   tov.tov_fecemi>cdate('" & fg_Ctod1(fectin) & "') " & _
                 "AND   tov.tov_fecemi<=cdate('" & fg_Ctod1(fecfin) & "') " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_nombre", vg_db, adOpenStatic
        Do While Not RS1.EOF
            V_Acceso.vaSpread1.MaxRows = V_Acceso.vaSpread1.MaxRows + 1
            V_Acceso.vaSpread1.Row = V_Acceso.vaSpread1.MaxRows
            V_Acceso.vaSpread1.Col = 1: V_Acceso.vaSpread1.text = Trim(RS1!pro_nombre)
            V_Acceso.vaSpread1.Col = 2: V_Acceso.vaSpread1.text = Format(RS1!tov_fecemi, "yyyymmdd")
            V_Acceso.vaSpread1.Col = 3: V_Acceso.vaSpread1.text = Trim(RS1!pro_codigo)
            V_Acceso.vaSpread1.Col = 4: V_Acceso.vaSpread1.text = "Salida"
            V_Acceso.vaSpread1.Col = 9: V_Acceso.vaSpread1.text = "80"
            V_Acceso.vaSpread1.Col = 5: V_Acceso.vaSpread1.text = Trim(RS1!tov_tipdoc)
            V_Acceso.vaSpread1.Col = 6: V_Acceso.vaSpread1.text = RS1!tov_numdoc
            V_Acceso.vaSpread1.Col = 7: V_Acceso.vaSpread1.text = RS1!dev_canmer
            V_Acceso.vaSpread1.Col = 8: V_Acceso.vaSpread1.text = RS1!dev_precos
            RS1.MoveNext
        Loop
        RS1.Close: Set RS1 = Nothing
        '------- Fin traer ventas directas
        Dim SortKeys, SortKeyOrder As Variant
        SortKeys = Array(1, 2, 9)
        SortKeyOrder = Array(1, 1, 1)
        ' Sort data in first five columns and rows by column 1 and 3
        V_Acceso.vaSpread1.Sort 1, 1, V_Acceso.vaSpread1.MaxCols, V_Acceso.vaSpread1.MaxRows, 0, SortKeys, SortKeyOrder
        Dim CosMov As Double, CanMov As Double, canbod As Double, pmp As Long, FecMov As Long, tipdoc As String
        Dim SBod As Double, SMov As Double
        tipdoc = "": FecMov = 0: CosMov = 0: CanMov = 0: canbod = 0: pmp = 0
        i = 1
        V_Acceso.vaSpread1.Row = 1: V_Acceso.vaSpread1.Col = 7: CanMov = Val(V_Acceso.vaSpread1.Value)
        V_Acceso.vaSpread1.Col = 4
        If V_Acceso.vaSpread1.MaxRows > 1 Or (V_Acceso.vaSpread1.MaxRows = 1 And Trim(V_Acceso.vaSpread1.text) <> "Inventario Inicial") Or CanMov <> 0 Then
'            .StartTable
'            .TableCell(tcCols) = 9: .TableCell(tcRows) = (v_acceso.vaSpread1.MaxRows + 2) '10000
'            .TableCell(tcColWidth, , 1) = 1000: .TableCell(tcAlign, , 1) = taLeftTop
'            .TableCell(tcColWidth, , 2) = 2000: .TableCell(tcAlign, , 2) = taLeftTop
'            .TableCell(tcColWidth, , 3) = 1000: .TableCell(tcAlign, , 3) = taLeftTop
'            .TableCell(tcColWidth, , 4) = 1200: .TableCell(tcAlign, , 4) = taLeftTop
'            .TableCell(tcColWidth, , 5) = 1200: .TableCell(tcAlign, , 5) = taRightTop
'            .TableCell(tcColWidth, , 6) = 1200: .TableCell(tcAlign, , 6) = taRightTop
'            .TableCell(tcColWidth, , 7) = 1200: .TableCell(tcAlign, , 7) = taRightTop
'            .TableCell(tcColWidth, , 8) = 1200: .TableCell(tcAlign, , 8) = taRightTop
'            .TableCell(tcColWidth, , 9) = 700: .TableCell(tcAlign, , 9) = taRightTop
            V_Acceso.vaSpread1.Row = 1
            '.TableCell(tcFontBold, i) = True: .TableCell(tcColSpan, i, 2) = 9
            V_Acceso.vaSpread1.Col = 3 ': .TableCell(tcText, i, 1) = Trim(v_acceso.vaSpread1.Text)
            V_Acceso.vaSpread1.Col = 1 ': .TableCell(tcText, i, 2) = Trim(v_acceso.vaSpread1.Text)
            i = i + 1
            For j = 1 To V_Acceso.vaSpread1.MaxRows
                V_Acceso.vaSpread1.Row = j
                V_Acceso.vaSpread1.Col = 2: FecMov = Val(V_Acceso.vaSpread1.Value)
                V_Acceso.vaSpread1.Col = 4: tipdoc = Trim(V_Acceso.vaSpread1.text)
                V_Acceso.vaSpread1.Col = 7: CanMov = Val(V_Acceso.vaSpread1.Value)
                V_Acceso.vaSpread1.Col = 8: CosMov = Val(V_Acceso.vaSpread1.Value)
                If tipdoc = "Inventario Inicial" Then
                    canbod = CanMov
                    pmp = CosMov
                Else
                    If tipdoc = "Entrada" Or tipdoc = "Ajuste Inventario (+)" Then
                       canbod = canbod + CanMov
                    ElseIf tipdoc = "Salida" Or tipdoc = "Ajuste Inventario (-)" Then
                       canbod = canbod - CanMov
                    End If
                    'PMP = ((canpro * IIf(canbod < 0, (canbod * -1), canbod)) + (cosmov * canmov)) / (IIf(canbod < 0, (canbod * -1), canbod) + canmov)
                    SBod = IIf(canbod < 0, 0, canbod)
                    SMov = IIf(CanMov < 0, 0, CanMov)
                    If (SMov + SBod) = 0 Then
                        pmp = CosMov
                    Else
                        pmp = ((SMov * CosMov) + (SBod * pmp)) / (SMov + SBod)
                    End If
                End If
                If (FecMov >= fecini And FecMov <= fecfin) Or tipdoc = "Inventario Inicial" Then
                    V_Acceso.vaSpread1.Col = 2: '.TableCell(tcText, i, 1) = Mid(v_acceso.vaSpread1.Text, 7, 2) & "/" & Mid(v_acceso.vaSpread1.Text, 5, 2) & "/" & Mid(v_acceso.vaSpread1.Text, 1, 4)
                    V_Acceso.vaSpread1.Col = 4: '.TableCell(tcText, i, 2) = v_acceso.vaSpread1.Text
                    V_Acceso.vaSpread1.Col = 5: '.TableCell(tcText, i, 3) = v_acceso.vaSpread1.Text
                    V_Acceso.vaSpread1.Col = 6: '.TableCell(tcText, i, 4) = v_acceso.vaSpread1.Text
                    V_Acceso.vaSpread1.Col = 7: '.TableCell(tcText, i, 5) = Format(v_acceso.vaSpread1.Text, fg_Pict(6, 2))
                    V_Acceso.vaSpread1.Col = 8: '.TableCell(tcText, i, 6) = Format(v_acceso.vaSpread1.Text, fg_Pict(6, 2))
                    V_Acceso.vaSpread1.Col = 7: '.TableCell(tcText, i, 7) = Format(canbod, fg_Pict(6, 2))
                    i = i + 1
                End If
             Next j
             ValidaBod codbod, codpro
             vg_db.Execute "UPDATE b_bodegas SET bod_canmer=" & canbod & " WHERE bod_codbod=" & vg_codbod & " AND bod_codpro='" & codpro & "'"
        End If
        RS3.MoveNext
        DoEvents
    Loop
    RS3.Close: Set RS3 = Nothing
    V_Acceso.Label1(0).Visible = False
    V_Acceso.Label1(1).Visible = False
    fg_descarga
Exit Function
Error_RecalcularStock:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, "Actualizar Stock"
    Exit Function
End Function

Function ValidarProductoVigente()

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset

'-------> Actualizar codigo de compras a codigo pedido
If vg_tipbase = "1" Then
   
   vg_db.Execute "UPDATE b_contlistpreing INNER JOIN b_productos ON b_contlistpreing.cpi_codcom=b_productos.pro_codigo SET b_contlistpreing.cpi_codped=[b_productos].[pro_codigo] " & _
                 "WHERE  (b_productos.pro_fecven>" & Format(Date, "yyyymmdd") & " Or b_productos.pro_fecven<=0 OR b_productos.pro_codigo In (SELECT bod.bod_codpro FROM b_bodegas bod " & _
                 "WHERE bod.bod_codbod=" & vg_codbod & " AND bod.bod_canmer>0)) AND b_productos.pro_ctrsto=1 AND NOT (b_contlistpreing.cpi_codcom) IS NULL AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "'"

Else
   
   vg_db.Execute "UPDATE b_contlistpreing SET b_contlistpreing.cpi_codped = b.pro_codigo FROM b_contlistpreing a, b_productos b WHERE a.cpi_codcom = b.pro_codigo " & _
                 "AND (b.pro_fecven > " & Format(Date, "yyyymmdd") & " Or b.pro_fecven <= 0 OR b.pro_codigo In (SELECT bod.bod_codpro FROM b_bodegas bod " & _
                 "WHERE bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer > 0)) AND b.pro_ctrsto = 1 AND NOT (a.cpi_codcom) IS NULL AND a.cpi_cencos = '" & MuestraCasino(1) & "'"

End If
'-------> Actualizar codigo de pedido
Dim aAp As String
'-------> Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_contlistpreingped"
fg_CheckTmp aAp
vg_db.Execute ("SELECT DISTINCT a.cpi_codped, a.cpi_coding, " & _
              "(SELECT TOP 1 x.pri_codpro FROM b_productosing x, b_productos z WHERE a.cpi_coding = x.pri_coding AND x.pri_codpro = z.pro_codigo AND (z.pro_fecven > " & Format(Date, "yyyymmdd") & " OR z.pro_fecven <= 0)) AS pri_codpro " & _
              "INTO " & aAp & " FROM  b_contlistpreing a " & _
              "WHERE a.cpi_codped NOT IN (SELECT DISTINCT pri_codpro FROM b_productosing WHERE a.cpi_coding = pri_coding) " & _
              "AND   a.cpi_cencos = '" & MuestraCasino(1) & "'")
If vg_tipbase = "1" Then
   
   vg_db.Execute ("UPDATE b_contlistpreing INNER JOIN " & aAp & " ON b_contlistpreing.cpi_coding = " & aAp & ".cpi_coding SET b_contlistpreing.cpi_codped = " & aAp & ".pri_codpro " & _
                  "WHERE b_contlistpreing.cpi_cencos = '" & MuestraCasino(1) & "' AND " & aAp & ".pri_codpro IS NOT NULL")

Else
   
   vg_db.Execute ("UPDATE b_contlistpreing SET b_contlistpreing.cpi_codped = " & aAp & ".pri_codpro FROM b_contlistpreing, " & aAp & " WHERE b_contlistpreing.cpi_coding = " & aAp & ".cpi_coding " & _
                  "AND b_contlistpreing.cpi_cencos = '" & MuestraCasino(1) & "' AND " & aAp & ".pri_codpro IS NOT NULL")

End If

vg_db.Execute "DROP TABLE " & aAp & ""

'-------> Actualizar codigo de compras
'-------> Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_contlistpreingcom"
fg_CheckTmp aAp
vg_db.Execute ("SELECT DISTINCT a.cpi_codcom, a.cpi_coding, " & _
              "(SELECT TOP 1 x.pri_codpro FROM b_productosing x, b_productos z WHERE a.cpi_coding = x.pri_coding AND x.pri_codpro = z.pro_codigo AND (z.pro_fecven > " & Format(Date, "yyyymmdd") & " OR z.pro_fecven <= 0)) AS pri_codpro " & _
              "INTO " & aAp & " FROM  b_contlistpreing a " & _
              "WHERE a.cpi_codcom NOT IN (SELECT DISTINCT pri_codpro FROM b_productosing WHERE a.cpi_coding = pri_coding) " & _
              "AND   a.cpi_cencos = '" & MuestraCasino(1) & "'")


If vg_tipbase = "1" Then
   
   vg_db.Execute ("UPDATE b_contlistpreing INNER JOIN " & aAp & " ON b_contlistpreing.cpi_coding = " & aAp & ".cpi_coding SET b_contlistpreing.cpi_codcom = " & aAp & ".pri_codpro " & _
                  "WHERE b_contlistpreing.cpi_cencos = '" & MuestraCasino(1) & "' AND " & aAp & ".pri_codpro IS NOT NULL")

Else
   
   vg_db.Execute ("UPDATE b_contlistpreing SET b_contlistpreing.cpi_codcom = " & aAp & ".pri_codpro FROM b_contlistpreing, " & aAp & " WHERE b_contlistpreing.cpi_coding = " & aAp & ".cpi_coding " & _
                  "AND b_contlistpreing.cpi_cencos = '" & MuestraCasino(1) & "' AND " & aAp & ".pri_codpro IS NOT NULL")

End If

vg_db.Execute "DROP TABLE " & aAp & ""

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'-------> Consulta a productos que esten vencidos y actualizar pedido
RS1.Open "SELECT DISTINCT a.cpi_coding, a.cpi_codcom, a.cpi_codped, b.pro_codigo " & _
         "FROM b_contlistpreing a, b_productos b " & _
         "WHERE a.cpi_codped = b.pro_codigo AND a.cpi_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   b.pro_fecven < " & Format(Date, "yyyymmdd") & " " & _
         "AND   b.pro_fecven <> 0 AND (b.pro_codigo IN (SELECT bod.bod_codpro FROM b_bodegas bod WHERE bod.bod_codbod = " & vg_codbod & " AND bod.bod_canmer <= 0))", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Function
Do While Not RS1.EOF
   
   
   If RS2.State = 1 Then RS2.Close
   RS2.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS2.Open "SELECT a.* FROM b_productosing a, b_productos b " & _
            "WHERE a.pri_codpro = b.pro_codigo " & _
            "AND   a.pri_coding = '" & RS1!cpi_coding & "' " & _
            "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " Or b.pro_fecven <= 0) AND b.pro_ctrsto = 1", vg_db, adOpenStatic
   
   If Not RS2.EOF Then
      
      vg_db.Execute "UPDATE b_contlistpreing SET cpi_codcom = '" & RS2!pri_codpro & "' WHERE cpi_coding = '" & RS1!cpi_coding & "' AND cpi_cencos = '" & MuestraCasino(1) & "'"
   
   End If
   RS2.Close
   Set RS2 = Nothing
   RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing

End Function

Function GetItem(ByVal strItems, ByVal Index, Optional strSeparator = ";")
Dim j, strRet, lngSepLen, blnRightToLeft

    blnRightToLeft = False

    lngSepLen = Len(strSeparator)
    If (Right(strItems, lngSepLen) <> strSeparator) Then strItems = strItems & strSeparator
    
    If blnRightToLeft Then
        j = CountItems(strItems, strSeparator) - Index + 1
        'GetItem = GetItem(strItems, j, strSeparator, False)
        GetItem = GetItem(strItems, j, strSeparator)
    Else
        strRet = ""
        While (Index > 0) And (InStr(strItems, strSeparator) > 0)
                j = InStr(strItems, strSeparator)
                strRet = Left(strItems, j - 1)
                strItems = Right(strItems, Len(strItems) - j - lngSepLen + 1)
                Index = Index - 1
        Wend
        'GetItem = IIF(Index > 0, "", strRet)
        GetItem = ""
        If (Index <= 0) Then
            GetItem = strRet
        End If
    End If
    
End Function

Function CountItems(ByVal strItems, Optional strSeparator = ";")
Dim i, lngCount, lngSepLen

    'If IsMissing(strSeparator) Then strSeparator = ";"
    
    lngSepLen = Len(strSeparator)
    If (Right(strItems, lngSepLen) <> strSeparator) Then strItems = strItems & strSeparator
    
    lngCount = 0
    For i = 1 To Len(strItems)
        If (Mid(strItems, i, lngSepLen) = strSeparator) Then lngCount = lngCount + 1
    Next
    
    CountItems = lngCount
    
End Function

Function strGetParamDescription(ByRef strParamItem As String) As String
Dim strTmp As String
Dim strParamItemTmp As String
Dim intIndex As Integer

    strTmp = ""
    strParamItemTmp = strParamItem

    For intIndex = 1 To CountItems(strParamItem)
        
        If (Trim(strTmp) <> "") Then strTmp = strTmp & " / "
        
        strTmp = strTmp & Trim(GetItem(strParamItem, intIndex + 1))
        
        If (GetItem(strParamItem, intIndex) = "-1") Then
            strParamItemTmp = ""
            intIndex = CountItems(strParamItem)
        Else
            strParamItemTmp = Replace(strParamItemTmp, GetItem(strParamItem, intIndex + 1) & ";", "")
        End If
        
        intIndex = intIndex + 1
        
    Next

    strParamItem = strParamItemTmp
    If (strParamItemTmp <> "") Then strParamItem = Replace(Trim(Mid(strParamItemTmp, 1, Len(strParamItemTmp) - 1)), ";", ",")
    strGetParamDescription = strTmp

End Function

Function ValidarBodega(codbod, opcion As Integer) As Boolean
Dim RS1 As New ADODB.Recordset
ValidarBodega = False
Select Case opcion
Case 0
    '------- Validar bodega en bodega
    RS1.Open "SELECT DISTINCT bod_codbod FROM b_bodegas WHERE bod_codbod = " & codbod & " AND bod_canmer >= 0", vg_db, adOpenStatic
    If Not RS1.EOF Then ValidarBodega = True
    RS1.Close: Set RS1 = Nothing
    If ValidarBodega Then Exit Function
    '------- Validar bodega compras centralizada
    RS1.Open "SELECT DISTINCT toc_codbod FROM b_totcompras WHERE  toc_codbod = " & codbod & "", vg_db, adOpenStatic
    If Not RS1.EOF Then ValidarBodega = True
    RS1.Close: Set RS1 = Nothing
    If ValidarBodega Then Exit Function
    '------- Validar bodega en ajuste inventario - traspasos - salidas
    RS1.Open "SELECT DISTINCT tov_codbod FROM b_totventas WHERE  tov_codbod = " & codbod & "", vg_db, adOpenStatic
    If Not RS1.EOF Then ValidarBodega = True
    RS1.Close: Set RS1 = Nothing
End Select
End Function

Function MoverDatoNuevoContrato(estcie As Boolean, cencos As String)
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
'------- mod jpaz 07-02-2005 CierreAjuste
If estcie Then CierrePeriodo 0, vg_codbod, 1
'------ Si no existe calendario cierre de mes agregar datos al contrato - lista precio producto contrato
RS.Open "SELECT DISTINCT * FROM b_cierreperiodo WHERE cie_cencos = '" & cencos & "'", vg_db, adOpenStatic
If RS.EOF Then
   RS.Close: Set RS = Nothing
   RS.Open "SELECT DISTINCT cie_cencos FROM b_cierreperiodo WHERE cie_estado IN (1,2)", vg_db, adOpenStatic
   vg_db.BeginTrans
   If Not RS.EOF Then
      vg_db.Execute "INSERT INTO b_cierreperiodo (cie_cencos, cie_periodo, cie_fecini, cie_fecter, cie_estado, cie_proantali, cie_gdpenmesali, cie_gdpenmesantali, cie_sncpenmesali, cie_sncpenmesantali, cie_proantgrl, cie_gdpenmesgrl, cie_gdpenmesantgrl, cie_sncpenmesgrl, cie_sncpenmesantgrl, cie_proantdes, cie_gdpenmesdes, cie_gdpenmesantdes, cie_sncpenmesdes, cie_sncpenmesantdes) SELECT '" & cencos & "', cie_periodo, cie_fecini, cie_fecter, cie_estado, cie_proantali, cie_gdpenmesali, cie_gdpenmesantali, cie_sncpenmesali, cie_sncpenmesantali, cie_proantgrl, cie_gdpenmesgrl, cie_gdpenmesantgrl, cie_sncpenmesgrl, cie_sncpenmesantgrl, cie_proantdes, cie_gdpenmesdes, cie_gdpenmesantdes, cie_sncpenmesdes, cie_sncpenmesantdes FROM b_cierreperiodo WHERE cie_cencos='" & RS!cie_cencos & "' AND cie_estado IN (1,2)"
   End If
   vg_db.CommitTrans
   RS.Close: Set RS = Nothing
   
   '------- Mover lista precio ingrediente contrato
   vg_db.BeginTrans
   vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & cencos & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente"
   vg_db.CommitTrans
   
   '------- Mover parametro
   vg_db.BeginTrans
   vg_db.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor, par_cencos) SELECT DISTINCT par_codigo, par_nombre, par_tipo, par_valor, '" & cencos & "' FROM a_param WHERE par_cencos='" & vg_contra & "'"
   vg_db.Execute "UPDATE a_param SET par_valor='N' WHERE par_cencos='" & cencos & "' AND par_codigo='5etapas'"
   vg_db.Execute "UPDATE a_param SET par_valor='0' WHERE par_cencos='" & cencos & "' AND par_codigo='addreceta'"
   vg_db.Execute "UPDATE a_param SET par_valor='N' WHERE par_cencos='" & cencos & "' AND par_codigo='opgruvul'"
   vg_db.Execute "UPDATE a_param SET par_valor='N' WHERE par_cencos='" & cencos & "' AND par_codigo='modpac'"
   vg_db.Execute "UPDATE a_param SET par_valor='0' WHERE par_cencos='" & cencos & "' AND par_codigo='modpro'"
   vg_db.Execute "UPDATE a_param SET par_valor='0' WHERE par_cencos='" & cencos & "' AND par_codigo='modrec'"
   
   vg_db.Execute "UPDATE a_param SET par_valor='" & cencos & "' WHERE par_cencos='" & cencos & "' AND par_codigo='casino'"
   
   RS.Open "SELECT DISTINCT cie_cencos, cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos='" & cencos & "' AND cie_estado=1", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         RS1.Open "SELECT MAX(a.tin_fectom) AS tin_fectom FROM b_tomainv a, b_clientes b WHERE a.tin_codbod=b.cli_codbod AND b.cli_codigo='" & RS!cie_cencos & "'", vg_db, adOpenStatic
         If Not RS1.EOF And RS1!tin_fectom > 0 Then
            vg_db.Execute "UPDATE a_param  SET par_valor = '" & fg_Encripta(LimpiaDato(CDate(fg_Ctod1(CStr(RS1!tin_fectom))) + 1)) & "' WHERE par_codigo = 'ciediario' AND par_cencos = '" & cencos & "'"
            '-------> Insertar tabla b_productospmpdia
            vg_db.Execute "INSERT INTO b_productospmpdia(ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo) " & _
                          "SELECT DISTINCT '" & cencos & "', a.pro_codigo, " & RS1!tin_fectom & ", 0, 0 " & _
                          "FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & cencos & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1)"
         Else
            vg_db.Execute "UPDATE a_param  SET par_valor = '" & fg_Encripta(LimpiaDato(CDate(fg_Ctod1(CStr(RS!cie_fecter))) + 1)) & "' WHERE par_codigo = 'ciediario' AND par_cencos = '" & cencos & "'"
            '-------> Insertar tabla b_productospmpdia
            vg_db.Execute "INSERT INTO b_productospmpdia(ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo) " & _
                          "SELECT DISTINCT '" & cencos & "', a.pro_codigo, " & RS!cie_fecter & ", 0, 0 " & _
                          "FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & cencos & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1)"
         End If
         RS1.Close: Set RS1 = Nothing
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   vg_db.CommitTrans
   
   '------- Mover infcfcfoficte
   vg_db.BeginTrans
   vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie, inf_usuario) VALUES ('" & cencos & "', 'C', 1, 0, Null)"
   vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie, inf_usuario) VALUES ('" & cencos & "', 'F', 1, 0, Null)"
   vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie, inf_usuario) VALUES ('" & cencos & "', 'T', 1, 0, Null)"
   vg_db.CommitTrans
   
   '------- Mover parametro despacho
   vg_db.BeginTrans
   vg_db.Execute "INSERT INTO b_paramdesp (pad_codigo, pad_tipo, pad_cencos, pad_diario) SELECT DISTINCT pad_codigo, pad_tipo, '" & cencos & "', pad_diario FROM b_paramdesp"
   vg_db.CommitTrans
   Exit Function
End If
RS.Close: Set RS = Nothing
End Function

Function TraerCorrelativo(codbod As Long, tipdoc As String) As Long
Dim sql1 As String
Dim RS1  As New ADODB.Recordset

'------- Traer correlativo
TraerCorrelativo = 0
sql1 = IIf(vg_tipbase = "1", " HOLDLOCK ", " WITH (HOLDLOCK) ")
RS1.Open "SELECT * FROM b_parametros  " & sql1 & " WHERE par_codbod = " & codbod & " AND par_tipdoc = '" & tipdoc & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
   TraerCorrelativo = RS1!par_correlativo + 1
Else
   vg_db.Execute "INSERT INTO b_parametros VALUES ('" & tipdoc & "', " & codbod & ",0)"
   TraerCorrelativo = 1
End If
RS1.Close: Set RS1 = Nothing
End Function

Function isNetwork(ByRef NetworkName As String) As Boolean
Dim ret As Long
'Si la Api retorna 0 quiere decir que no hay ningun tipo de conexión de Red
If IsNetworkAlive(ret) = 0 Then
   isNetwork = False
   NetworkName = ""
   'MsgBox("El sistema no está conectado a una NetWork!", vbInformation)
Else
   ' hay conexión , y muestra el tipo
   isNetwork = True
   NetworkName = IIf(ret = NETWORK_ALIVE_AOL, "AOL", _
                 IIf(ret = NETWORK_ALIVE_LAN, "LAN", "WAN"))
   'MsgBox("El sistema está conectado a: " + _
   '       IIf(Ret = NETWORK_ALIVE_AOL, "AOL", _
   '       IIf(Ret = NETWORK_ALIVE_LAN, "LAN", "WAN")) + " network", vbInformation)
End If
End Function

Function isInternetConnected(ByRef CONNECTION_LAN As Boolean, ByRef CONNECTION_MODEM As Boolean, ByRef CONNECTION_PROXY As Boolean) As Boolean
Dim lngFlags As Long
CONNECTION_LAN = False: CONNECTION_MODEM = False: CONNECTION_PROXY = False
If InternetGetConnectedState(lngFlags, 0) Then
'    If lngFlags And Flags.INTERNET_CONNECTION_LAN Then CONNECTION_LAN = True
    If lngFlags And INTERNET_CONNECTION_LAN Then CONNECTION_LAN = True
    If lngFlags And INTERNET_CONNECTION_MODEM Then CONNECTION_MODEM = True
    If lngFlags And INTERNET_CONNECTION_PROXY Then CONNECTION_PROXY = True
End If
isInternetConnected = (CONNECTION_LAN Or CONNECTION_MODEM Or CONNECTION_PROXY)
End Function

Function FechaHora() As String
FechaHora = Date & " " & Time & " - "
End Function

Function fg_buscarcodtip(StrMensaje As String, codtip As String) As String
Dim StrMensaje1 As String, StrMensaje2 As String
StrMensaje1 = "": StrMensaje2 = "": fg_buscarcodtip = StrMensaje
If Len(StrMensaje) <> 0 Then
   Do While InStr(StrMensaje, ";") <> 0 And InStr(StrMensaje, ";") <> 1
      If StrMensaje <> "" Then
         StrMensaje1 = Mid(StrMensaje, 1, InStr(StrMensaje, ";") - 1)
         If codtip <> StrMensaje1 Then StrMensaje2 = StrMensaje2 & StrMensaje1 & ";"
         StrMensaje = Mid(StrMensaje, InStr(StrMensaje, ";") + 1)
      End If
   Loop
   If Trim(StrMensaje2) <> "" Then fg_buscarcodtip = Mid(StrMensaje2, 1, Len(StrMensaje2) - 1)
End If
End Function

Function ConsultaProcess(ByVal processName As String) As Boolean

On Error GoTo ErrHandler

Dim oWMI
Dim ret
Dim sService
Dim oWMIServices
Dim oWMIService
Dim oServices
Dim oService
Dim servicename
Set oWMI = GetObject("winmgmts:")
Set oServices = oWMI.InstancesOf("win32_process")
ConsultaProcess = False
For Each oService In oServices
    
    DoEvents
    servicename = LCase(Trim(CStr(oService.Name) & ""))
    
    If InStr(1, servicename, LCase(processName), vbTextCompare) > 0 Then
        
        'ret = oService.Terminate
        ConsultaProcess = True
        Exit Function
    
    End If

Next

Set oServices = Nothing
Set oWMI = Nothing
ErrHandler:
Err.Clear

End Function

Function CalcularPMPDiaSql(Formu As Form, op As Boolean, progrl As Boolean) As Boolean

On Local Error GoTo Error_CalcularPMPDia

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String, aAp1 As String, aAp2 As String, aAp3 As String, aAp4 As String, aAp5 As String, aAp6 As String
Dim fecini As Long, fecfin As Long, i As Long, FecInv As Long, est As Boolean, fecpro As Date, fecter As Date
Dim estpro As Integer, sql1 As String
Dim fecpro1 As String, fecter1 As String, fecdiad As String
'00 Inventario inicial
'10 Ajuste de Entrada
'20 Proveedores Entrada
'30 Traspaso Entrada
'40 Produccion Salida
'50 Produccion Entrada
'60 Traspaso Salida
'70 Mermas Salida
'80 Venta Directa Salida
'90 Ajuste Salida
'100 Venta Cafeteria

Dim auxCanmer As Double, auxPropon As Double, propon As Double, auxfec As Long, fecuco As String, upreco As Double
If op Then Formu.Label1(1).Visible = True
If op Then Formu.Label1(1).Caption = "Procesando Información"
If op Then Formu.Bar1(0).Visible = True: Formu.Bar1(0).Value = 0: Formu.Bar1(0).max = 2
If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1
'fecpro = CDate(vg_ciedia)
'fecter = dEoM(CDate(vg_ciedia))

CalcularPMPDiaSql = True

fecpro1 = CDate(vg_ciedia)
fecter1 = dEoM(CDate(vg_ciedia))
fecdiad = Format(CDate(vg_ciedia), "dd/mm/yyyy")

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgp_s_cierrediario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(CDate(vg_ciedia), "yyyymmdd") & ", '" & fecdiad & "', '" & fecpro1 & "',  '" & fecter1 & "', " & vg_DCa & "")
If RS1.EOF Then
   
   fg_descarga
   RS1.Close
   Set RS1 = Nothing
   If op Then MsgBox "Error Proceso de Cierre Día", vbInformation + vbOKOnly, MsgTitulo
   Exit Function

End If
   
If RS1!procesa = "1" Then
   
   fg_descarga
   RS1.Close
   Set RS1 = Nothing
   If op Then MsgBox "Error Proceso de Cierre Día", vbInformation + vbOKOnly, MsgTitulo
   Exit Function

End If
RS1.Close
Set RS1 = Nothing

If Not progrl Then
   
   '-------> Borrar tablas b_productospmpdia distinto al ultimo día cierre que tengan precio cero y saldo en cero
   vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <> " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"
   Exit Function

End If

'-------> Grabar log_cierrediario
vg_db.Execute "INSERT INTO log_cierrediario VALUES ('" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "', '" & Format(vg_ciedia, "yyyymmdd") & "', '" & Trim(vg_NUsr) & "', '1.- Cierre Diario', '" & MuestraCasino(1) & "')"

'-------> Borrar tablas b_productospmpdia distinto al ultimo día cierre que tengan precio cero y saldo en cero
vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <> " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"

'-------> Crear tabla y mover datos
If op Then Formu.Label1(1).Caption = "Actua. datos anexo"
If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

estpro = 2

Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String, DBO As String

'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile ' borrar base datos si existe
If Dir(mdir & MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".mdb") <> "" Then Kill mdir & MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".mdb" ' borrar base datos si existe

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open

'-------> Ini : Generar tabla Costo minuta & Food Cost
dbE.Execute "CREATE TABLE B_CostoMinutaRealizadoFoodCostN (IdCeco char(10), Fecha_Minuta int, IdRegimen int, IdServicio int, Raciones_Teorica int, " & _
            "Costo_Teorico_Alim float, Costo_Teorico_Desec float, Raciones_Real int, Costo_Real_Alim float, Costo_Real_Desec float, Raciones_Vendidas int, " & _
            "Costo_Realizado_Alim float, Costo_Realizado_Desec float, Venta_Dia float, Venta_Contado float, Glosa_Venta_Especial char(100)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraerCostoFoodCostMinutaCierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_CostoMinutaRealizadoFoodCostN values (' " & RS(0) & "', " & RS(3) & ", " & RS(1) & ", " & RS(2) & ", " & RS(4) & " " & _
                  ", " & RS(5) & ", " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", " & RS(14) & ", '" & RS(15) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Costo minuta & Food Cost

'-------> Ini : Generar tabla Insumos & Food Cost A13
dbE.Execute "CREATE TABLE B_A13InsumosFCost (IdCeco char(10), Periodo int, FechaIni Int, FechaFin int, FechaCierre int, " & _
            "Glosa char(200), Alimentos Float, Lim_Desc Float, Total Float, Porcentaje Float, Id int)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_VtaCtoServInsumosFoodCostGastoA13CierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(CDate(vg_ciedia), "yyyymmdd") & ", " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_A13InsumosFCost values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & Format(CDate(vg_ciedia), "yyyymmdd") & ", '" & RS(4) & "' " & _
                  ", " & RS(5) & ", '" & RS(6) & "', '" & RS(7) & "', " & RS(8) & ", " & RS(9) & ")"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Insumos & Food Cost A13

'-------> Ini : Generar tabla detalle mermas
dbE.Execute "CREATE TABLE B_DetalleMermasNK (IdCeco char(10), Periodo int, Fecha_Minuta int, IdRegimen int, IdServicio int, FechaCierre int, " & _
            "IdReceta int, IdEstServicio int, NumLin int, CostoRecetaAlimento float, CostoRecetaDesechable float, CantidadMerma float, CantidadRacionReal float, MermaxKilo float) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleMermasCierreDiario '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_DetalleMermasNK values (' " & RS(0) & "', " & Format(vg_ciedia, "yyyymm") & ", " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & Format(CDate(vg_ciedia), "yyyymmdd") & ", " & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ")"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla detalle mermas

'-------> Ini : Generar tabla mermas desconche - pan - produccion
dbE.Execute "CREATE TABLE B_MermaDesconche (IdCeco char(10), IdRegimen int, IdServicio int, Fecha_Merma int, " & _
            "Considera_Merma char(1), Merma_Desconche float, Merma_Pan float, Merma_Produccion float, Fecha_Modificacion datetime, Fecha_Creacion datetime, Usuario char(20)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_MermaDesconcheCierreDiario '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_MermaDesconche values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", '" & RS(8) & "', '" & RS(9) & "', '" & RS(10) & "')"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla mermas desconche - pan - produccion

'-------> Ini : Generar tabla consumo proyectado & real
dbE.Execute "CREATE TABLE B_ConsumoProyectadoReal (IdCeco char(10), IdRegimen int, NoRegimen char(50), IdServicio int, NoServicio char(50), NumDoc int, Periodo int, Fecha int, " & _
            "Codigo_Producto char(20), Descripcion_Producto char(100), Unidad char(5), Cantidad_Teorica float, Cantidad_Planificada float, Cantidad_Realizada float, PMP float, Racion_Teorica int, Usuario_Mod_Racion_Real char(20), Racion_Real int, Fecha_Mod_Racion_Real datetime, Usuario_Salida_Produccion char(20), Racion_Salida_Produccion int, Fecha_Mod_Salida_Produccion datetime, Cantidad_Devolucion float)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CalcularSalidaProducciónMinutaTeorica '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymmdd") & ", '1', " & vg_codbod & ", " & Format(vg_ciedia, "yyyymmdd") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_ConsumoProyectadoReal values ('" & MuestraCasino(1) & "', " & RS(0) & ", '" & RS(1) & "', " & RS(2) & ", '" & RS(3) & "', " & RS(5) & "," & _
                  "" & Format(vg_ciedia, "yyyymm") & ", " & Format(vg_ciedia, "yyyymmdd") & ", '" & RS(6) & "', '" & RS(7) & "', '" & RS(8) & "', " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", '" & RS(14) & "', " & RS(15) & ", '" & RS(16) & "', '" & RS(17) & "', " & RS(18) & ", '" & RS(19) & "', " & RS(20) & ")"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla consumo proyectado & real

'-------> Ini : Generar tabla a_tipoajuste 31/03/2022
dbE.Execute "CREATE TABLE a_tipoajuste (aju_codigo int, aju_nombre char(30), aju_tipaju int, aju_tipo char(1))"
dbE.Execute "INSERT INTO a_tipoajuste SELECT aju_codigo, aju_nombre, aju_tipaju, aju_tipo " & _
            "FROM a_tipoajuste IN " & DBO & ""
'-------> Fin : Generar tabla a_tipoajuste  31/03/2022

'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN " & DBO & ""
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN " & DBO & ""
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN " & DBO & ""
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN " & DBO & " " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale, 'dd/mm/yyyy') = CDATE('" & vg_ciedia & "')"
            
'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
dbE.Execute "CREATE TABLE a_bodega (bod_codigo int, bod_nombre char(25), bod_ubicac char(35))"
dbE.Execute "INSERT INTO a_bodega SELECT bod_codigo, bod_nombre, bod_ubicac FROM a_bodega a, b_clientes b IN " & DBO & " " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo = 0"

'-------> Generar tabla proveedor
dbE.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo datetime, prv_origen char(1))"
dbE.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen FROM b_proveedor IN " & DBO & ""

'-------> Generar tabla compras
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " AND (toc_fecrem = cdate('" & vg_ciedia & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"
            
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0') " & _
            "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
            "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
            "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)"

'            "union all "
'dbE.Execute "INSERT INTO b_totcompras " & _
'            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
'            "FROM b_totcompras IN " & DBO & " " & _
'            "WHERE toc_codbod = " & vg_codbod & " " & _
'            "AND   toc_tipdoc = 'GD' " & _
'            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras "
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R "
dbE.Execute "DROP TABLE b_totcompras_R"

'dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc int, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   (b.toc_fecrem = cdate('" & vg_ciedia & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
'dbE.Execute "INSERT INTO b_detcompras " & _
'            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
'            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.dec_rutpro = b.toc_rutpro " & _
'            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
'            "AND   a.dec_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'SN' " & _
'            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

'            "union all "
'dbE.Execute "INSERT INTO b_detcompras " & _
'            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
'            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.dec_rutpro = b.toc_rutpro " & _
'            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
'            "AND   a.dec_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'GD' " & _
'            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"

'dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND  (b.toc_fecrem = cdate('" & vg_ciedia & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle impuesto
'dbE.Execute "INSERT INTO b_detcomprasimp " & _
'            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
'            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
'            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
'            "AND   a.imd_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'SN' " & _
'            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

'dbE.Execute "INSERT INTO b_detcomprasimp " & _
'            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
'            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
'            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
'            "AND   a.imd_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'GD' " & _
'            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R "
dbE.Execute "DROP TABLE b_detcomprasimp_R"

'-------> Actualizar estado documento
vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " AND (CONVERT(VARCHAR(8),toc_fecrem,112) = CONVERT(VARCHAR(8),CONVERT(DATE, CONVERT(VARCHAR(10),('" & vg_ciedia & "')),103),112) OR isnull(EnvioDocSGPADM,'0') = '0')")

'vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
'               "WHERE toc_codbod = " & vg_codbod & " " & _
'               "AND   toc_tipdoc = 'SN' " & _
'               "AND  (isnull(toc_docsnc,'')='' or isnull(EnvioDocSGPADM,'0')='0') " & _
'               "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
'               "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
'               "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)")

vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " " & _
               "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
               "AND  (isnull(toc_docsnc,'')='' or isnull(EnvioDocSGPADM,'0')='0') " & _
               "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
               "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
               "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)")

'vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
'               "WHERE toc_codbod = " & vg_codbod & " " & _
'               "AND   toc_tipdoc = 'GD' " & _
'               "AND  (isnull(toc_docaso,'')='' or isnull(EnvioDocSGPADM,'0')='0')")

vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " " & _
               "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
               "AND  (isnull(toc_docaso,'')='' or isnull(EnvioDocSGPADM,'0')='0')")

'-------> Generar tabla b_ocsacrecibido
'            "AND   b.toc_fecemi = CDATE('" & vg_ciedia & "')"

dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc datetime, ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc FROM b_ocsacrecibido a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & vg_ciedia & "')"


'-------> Generar tabla ventas
'dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc int, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc double, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
            "FROM b_totventas  IN " & DBO & " " & _
            "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = cdate('" & vg_ciedia & "') AND tov_tipdoc IN ('DP','SP')) " & _
            "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = cdate('" & vg_ciedia & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"

'-------> Crear estructura tabla detalle ventas
'dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc int, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc double, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & vg_ciedia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & vg_ciedia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
'dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & vg_ciedia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & vg_ciedia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla ventas servicios especiales
dbE.Execute "CREATE TABLE b_totventaserviciosespeciales (tos_IdCeco char(10), tos_Tipo_Documento char(2), tos_Numero_Documento int, tos_Fecha_Produccion datetime, tos_IdBodega int, tos_Venta_servicio_Especiales char(100), Tos_Comensales double, tos_Precio_Servicio double, tos_Total_Documento double, tos_Estado_Documento char(1), tos_Fecha_Creacion datetime, tos_Fecha_Modificacion datetime, tos_Periodo int, tos_Usuario char(20), tos_Documento_Asociado int)"
dbE.Execute "INSERT INTO b_totventaserviciosespeciales SELECT tos_IdCeco, tos_Tipo_Documento, tos_Numero_Documento, tos_Fecha_Produccion, tos_IdBodega, tos_venta_servicio_Especiales, Tos_Comensales, tos_Precio_servicio, tos_Total_Documento, tos_Estado_Documento, tos_Fecha_Creacion, tos_Fecha_modificacion, tos_Periodo, tos_Usuario, tos_Documento_Asociado " & _
            "FROM b_totventaserviciosespeciales  IN " & DBO & " " & _
            "WHERE tos_IdBodega = " & vg_codbod & " AND tos_fecha_Produccion = cdate('" & vg_ciedia & "') AND tos_tipo_documento IN ('DE','SE') "

'-------> Crear estructura tabla detalle ventas servicios especiales
dbE.Execute "CREATE TABLE b_detventaserviciosespeciales (des_IdCeco char(10), des_Tipo_Documento char(2), des_Numero_Documento int, des_Numero_Linea int, des_IdProducto char(20), des_Cantidad_Mercaderia double, des_Cantidad_Devolver double, des_Precio_Documento double, des_Total_Documento double, des_Descripcion char(255), des_Mueve_Inventario char(1), des_Actualiza_Precio char(1))"
dbE.Execute "INSERT INTO b_detventaserviciosespeciales SELECT a.des_IdCeco, a.des_tipo_documento, a.des_numero_documento, a.des_numero_linea, a.des_IdProducto, a.des_cantidad_mercaderia, a.des_cantidad_devolver, a.des_Precio_Documento, a.des_Total_documento, a.des_Descripcion, a.des_Mueve_Inventario, a.des_Actualiza_Precio " & _
            "FROM b_detventaserviciosespeciales a, b_totventaserviciosespeciales b IN " & DBO & " " & _
            "WHERE a.des_IdCeco = b.tos_IdCeco " & _
            "AND   a.des_tipo_documento = b.tos_tipo_documento " & _
            "AND   a.des_numero_documento = b.tos_numero_documento " & _
            "AND   b.tos_fecha_produccion = cdate('" & vg_ciedia & "') AND b.tos_tipo_documento IN ('DE','SE') " & _
            "AND   b.tos_IdBodega = " & vg_codbod & ""

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & vg_ciedia & "')"

dbE.Execute "CREATE TABLE b_totventascaf (tvc_cencos char(10), tvc_fecing datetime, tvc_codbod int, tvc_estado char(1))"
dbE.Execute "INSERT INTO b_totventascaf SELECT tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado FROM b_totventascaf IN " & DBO & " " & _
            "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   tvc_fecing = cdate('" & vg_ciedia & "') " & _
            "AND   tvc_codbod = " & vg_codbod & " " & _
            "AND   tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascaf (dvc_cencos char(10), dvc_fecing datetime, dvc_numlin int, dvc_articulo char(20), dvc_canart double, dvc_precio double, dvc_tippag char(2), dvc_rutcli char(10), dvc_cencli char(10), dvc_tipdoc char(2), dvc_numdoc int, dvc_fecdoc datetime)"
dbE.Execute "INSERT INTO b_detventascaf SELECT a.dvc_cencos, a.dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, a.dvc_fecdoc FROM b_detventascaf a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvc_cencos = b.tvc_cencos " & _
            "AND   a.dvc_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & vg_ciedia & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascafpro (dvp_cencos char(10), dvp_fecing datetime, dvp_codmer char(20), dvp_cancal double, dvp_candig double, dvp_precos double)"
dbE.Execute "INSERT INTO b_detventascafpro SELECT a.dvp_cencos, a.dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos FROM b_detventascafpro a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvp_cencos = b.tvc_cencos " & _
            "AND   a.dvp_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & vg_ciedia & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
dbE.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
dbE.Execute "INSERT INTO a_regimen SELECT reg_codigo, reg_nombre FROM a_regimen IN " & DBO & ""

'-------> Generar tabla servicio
dbE.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(30), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob datetime, ser_horent datetime, ser_horpda datetime)"
dbE.Execute "INSERT INTO a_servicio SELECT ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_activo, ser_horcob, ser_horent, ser_horpda FROM a_servicio IN " & DBO & ""

'-------> Generar tabla estructura de servicio
dbE.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double, ess_cencos char(10))"
dbE.Execute "INSERT INTO a_estservicio SELECT ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin, ess_cencos FROM a_estservicio IN " & DBO & " where ess_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla sector
dbE.Execute "CREATE TABLE a_sector (sec_codigo int, sec_nombre char(50), sec_orden int)"
dbE.Execute "INSERT INTO a_sector SELECT sec_codigo, sec_nombre, sec_orden FROM a_sector IN " & DBO & ""

'-------> Generar tabla servicio rac
dbE.Execute "CREATE TABLE a_serviciorac (sra_codser int, sra_coditem int, sra_serdia int, sra_raciones int, sra_cencos char(10))"
dbE.Execute "INSERT INTO a_serviciorac SELECT sra_codser, sra_coditem, sra_serdia, sra_raciones, sra_cencos FROM a_serviciorac IN " & DBO & " where sra_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then

dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'-------> Si tipo de minuta es distinto simap puede generar minuta cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
   
   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
               "AND   a.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & ""

End If
    
dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'-------> Si tipo de minuta es distinto simap puede generar minuta detalle cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
    
    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
                "WHERE b.min_codigo = a.mid_codigo " & _
                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                "AND   b.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & ""
End If
'Else
'   dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
'               "WHERE a.min_codigo = b.mid_codigo " & _
'               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'               "AND   a.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & " " & _
'               "AND   b.mid_tipmin = '2'"
'
'    dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
'                "WHERE b.min_codigo = a.mid_codigo " & _
'                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
'                "AND   b.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & " " & _
'                "AND   b.mid_tipmin = '2'"
'End If
'-------> Generar tabla minutafija
dbE.Execute "CREATE TABLE b_minutafijadia (mfd_cencos char(10), mfd_codreg int, mfd_codser int, mfd_fecha int, mfd_codpro char(20), mfd_tipmin char(1), mfd_canpro double, mfd_cospro double)"
dbE.Execute "INSERT INTO b_minutafijadia SELECT mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro FROM b_minutafijadia IN " & DBO & " " & _
            "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mfd_fecha  = " & Format(vg_ciedia, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
dbE.Execute "CREATE TABLE b_minutaraciones (mir_cencos char(10), mir_codreg int, mir_codser int, mir_fecmin int, mir_rutcli char(10), mir_nrorac int, mir_nroguia int, mir_codcli char(10))"
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) SELECT mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli FROM b_minutaraciones IN " & DBO & " " & _
            "WHERE mir_cencos          = '" & MuestraCasino(1) & "' " & _
            "AND   mid(mir_fecmin,1,6) = " & Format(vg_ciedia, "yyyymm") & ""

'-------> Insertar mermas
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) " & _
"SELECT  bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin , " & _
"        'MERMAS' as mir_rutcli , " & _
"        SUM(bm2.mid_nummer) As mermas, " & _
"        0,                              " & _
"        ''                             " & _
"FROM    b_minuta AS bm, " & _
"        b_minutadet AS bm2 IN " & DBO & " " & _
"WHERE   bm.min_codigo = bm2.mid_codigo and bm.min_cencos = '" & MuestraCasino(1) & "' " & _
"        AND mid(bm.min_fecmin, 1, 6) = " & Format(vg_ciedia, "yyyymm") & " " & _
"        AND bm2.mid_tipmin = '2' " & _
"GROUP BY bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin"

'-------> Generar tabla precio venta
dbE.Execute "CREATE TABLE b_preciovta (prv_cencos char(10), prv_codreg int, prv_codser int, prv_fecvig int, prv_rutcli char(10), prv_preven double)"
dbE.Execute "INSERT INTO b_preciovta SELECT prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli, prv_preven FROM b_preciovta IN " & DBO & " " & _
            "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
dbE.Execute "CREATE TABLE b_ventacontado (vtc_codigo int, vtc_cencos char(10), vtc_codreg int, vtc_codser int, vtc_fecvta int, vtc_forpag int, vtc_totmon double, vtc_opccli char(1))"
dbE.Execute "INSERT INTO b_ventacontado SELECT vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli FROM b_ventacontado IN " & DBO & " " & _
            "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   vtc_fecvta = " & Format(vg_ciedia, "yyyymmdd") & ""

dbE.Execute "CREATE TABLE b_ventacontadodet (vtd_codigo int, vtd_numlin int, vtd_codcli char(10), vtd_codcco char(10), vtd_descripcion char(50), vtd_detmon double)"
dbE.Execute "INSERT INTO b_ventacontadodet SELECT a.vtd_codigo, a.vtd_numlin, a.vtd_codcli, a.vtd_codcco, a.vtd_descripcion, a.vtd_detmon FROM b_ventacontadodet a, b_ventacontado b IN " & DBO & " " & _
            "WHERE b.vtc_codigo = a.vtd_codigo " & _
            "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.vtc_fecvta = " & Format(vg_ciedia, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
dbE.Execute "CREATE TABLE b_clientecencos (clc_codigo char(10), clc_codcli char(10), clc_nombre char(50))"
dbE.Execute "INSERT INTO b_clientecencos SELECT clc_codigo, clc_codcli, clc_nombre FROM b_clientecencos IN " & DBO & " where clc_codcli = '" & MuestraCasino(1) & "' "

'-------> Generar tabla cliente
dbE.Execute "CREATE TABLE b_clientes (cli_codigo char(10), cli_nombre char(50), cli_direccion char(50), cli_comuna char(15), cli_ciudad char(15), cli_fono1 char(15), cli_fono2 char(15), cli_fax char(15), cli_percon char(50), cli_giro char(50), cli_email char(50), cli_tipo int, cli_codbod int, cli_codtis int, cli_codseg int, cli_codcli char(10), cli_clisap char(1), cli_socsap char(4), cli_cievta char(1), cli_ciedia int, cli_activo char(1), cli_sobrec char(1), cli_codmun int, cli_ccisac int, cli_cecsac char(4), cli_codreg int, id_tipo_vale char(100))"
dbE.Execute "INSERT INTO b_clientes SELECT cli_codigo, cli_nombre, cli_direccion, cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, cli_giro, cli_email, cli_tipo, cli_codbod, cli_codtis, cli_codseg, cli_codcli, cli_clisap, cli_socsap, cli_cievta, cli_ciedia, cli_activo, cli_sobrec, cli_codmun, cli_ccisac, cli_cecsac, cli_codreg FROM b_clientes IN " & DBO & ""

'-------> Generar tabla bodegas
vg_db.Execute ("delete paso_b_bodegas where bod_Cencos = '" & MuestraCasino(1) & "'")
vg_db.Execute ("insert into paso_b_bodegas (bod_cencos, bod_codbod, bod_codpro, bod_canmer) SELECT distinct '" & MuestraCasino(1) & "', bod_codbod, bod_codpro, round(bod_canmer,2) FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " ")
dbE.Execute "CREATE TABLE b_bodegas (bod_codbod int, bod_codpro char(20), bod_canmer double)"
dbE.Execute "INSERT INTO b_bodegas SELECT distinct bod_codbod, bod_codpro, bod_canmer FROM paso_b_bodegas IN " & DBO & " WHERE bod_cencos = '" & MuestraCasino(1) & "' and bod_codbod = " & vg_codbod & ""

'-------> Generar tabla precio cafeteria
dbE.Execute "CREATE TABLE b_totpreciocaf (tpc_codigo char(20), tpc_nombre char(50), tpc_precio double, tpc_cencos char(10), tpc_activo char(1))"
dbE.Execute "INSERT INTO b_totpreciocaf SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_cencos, tpc_activo FROM b_totpreciocaf  IN " & DBO & " WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

dbE.Execute "CREATE TABLE b_detpreciocaf (dpc_codigo char(20), dpc_codmer char(20), dpc_cantidad double, dpc_cencos char(10))"
dbE.Execute "INSERT INTO b_detpreciocaf SELECT a.dpc_codigo, a.dpc_codmer, a.dpc_cantidad, a.dpc_cencos FROM b_detpreciocaf a, b_totpreciocaf b IN " & DBO & " " & _
            "WHERE b.tpc_codigo = a.dpc_codigo " & _
            "AND   b.tpc_cencos = a.dpc_cencos " & _
            "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

RS1.Open "SELECT DISTINCT b.cie_fecini, b.cie_fecter FROM b_tomainv a, b_cierreperiodo b " & _
         "WHERE b.cie_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   a.tin_fectom = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
         "AND   a.tin_ciemes = b.cie_periodo", vg_db, adOpenStatic
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom >= " & RS1!cie_fecini & " AND tin_fectom <= " & RS1!cie_fecter & " " & _
               "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
   dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf FROM b_totventas IN " & DBO & " " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
    '-------> Generar tabla ajuste inventario detalle
    dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
                "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
                "WHERE a.dev_rutcli = b.tov_rutcli " & _
                "AND   a.dev_tipdoc = b.tov_tipdoc " & _
                "AND   a.dev_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
    
    '-------> Generar tabla ajuste inventario detalle impuesto
    dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
                "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
                "WHERE a.imd_rutdoc = b.tov_rutcli " & _
                "AND   a.imd_tipdoc = b.tov_tipdoc " & _
                "AND   a.imd_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
Else
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom = " & Format(vg_ciedia, "yyyymmdd") & " " & _
               "AND   tin_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
'vg_db.Execute "SELECT * INTO log_cierrediario IN '" & cDBI & "' FROM log_cierrediario WHERE feccie >= cdate('" & vg_ciedia & "') AND feccie <= cdate('" & vg_ciedia & "') + 2"
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN " & DBO & " " & _
            "WHERE feccie >= cdate('" & vg_ciedia & "') " & _
            "AND   feccie <= cdate('" & vg_ciedia & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN " & DBO & " " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN " & DBO & " " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"


'-------> generar tabla a_derechosperfil
dbE.Execute "CREATE TABLE a_derechosperfil (dpe_cecori char(10), dpe_codper int, dpe_codopc int, dpe_deracc int, dpe_deragr int, dpe_dermod int, dpe_dereli int, dpe_derimp int)"
dbE.Execute "INSERT INTO a_derechosperfil (dpe_cecori, dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT distinct '" & MuestraCasino(1) & "', dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp FROM a_derechosperfil IN " & DBO & " "

'-------> actualizar tabla a_opcsistema
Dim ciedia As String
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE a_opcsistema SET EnvioDocSGPADM = '" & ciedia & "' FROM a_opcsistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = ''")

'-------> generar tabla a_opcsistema
dbE.Execute "CREATE TABLE a_opcsistema (opc_cecori char(10), opc_codigo int, opc_nombre char(50))"
dbE.Execute "INSERT INTO a_opcsistema (opc_cecori, opc_codigo, opc_nombre) SELECT distinct '" & MuestraCasino(1) & "', opc_codigo, opc_nombre FROM a_opcsistema IN " & DBO & "  where EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> generar tabla a_perfil
dbE.Execute "CREATE TABLE a_perfil (per_cecori char(10), per_codigo int, per_nombre char(30))"
dbE.Execute "INSERT INTO a_perfil (per_cecori, per_codigo, per_nombre) SELECT distinct '" & MuestraCasino(1) & "', per_codigo, per_nombre FROM a_perfil IN " & DBO & " "

'-------> generar tabla a_usuarios
dbE.Execute "CREATE TABLE a_usuarios (usu_cecori char(10), usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios (usu_cecori, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios IN " & DBO & " "

'-------> generar tabla a_usuarios_eliminado
dbE.Execute "CREATE TABLE a_usuarios_eliminado (usu_cecori char(10), Fecha datetime, usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios_eliminado (usu_cecori, Fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios_eliminado IN " & DBO & " "

'-------> generar tabla b_usuariocontratos
dbE.Execute "CREATE TABLE b_usuariocontratos (uco_cecori char(10), uco_codusu char(20), uco_codcon char(10))"
dbE.Execute "INSERT INTO b_usuariocontratos (uco_cecori, uco_codusu, uco_codcon) SELECT distinct '" & MuestraCasino(1) & "', uco_codusu, uco_codcon FROM b_usuariocontratos IN " & DBO & " "

'-------> actualizar tabla log_sistema
ciedia = ""
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE Log_Sistema SET EnvioDocSGPADM = '" & ciedia & "' FROM Log_Sistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = '' and Loc_Id in (2,20,21,22,23)")
               
'-------> generar tabla Log_Sistema
dbE.Execute "CREATE TABLE Log_Sistema (cecori char(10), Fecha datetime, Usuario_Id char(20), Loc_Id int, Opcion_Sistema char(14), Dato_Nuevo memo, Dato_Anterior memo, Detalle_Operacion memo)"
dbE.Execute "INSERT INTO Log_Sistema (cecori, Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion) SELECT distinct '" & MuestraCasino(1) & "', Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion FROM Log_Sistema IN " & DBO & " where Loc_Id in (2,20,21,22,23) and EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
'dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_proveedor', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_CostoMinutaRealizadoFoodCostN', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_A13InsumosFCost', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_DetalleMermasNK', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_MermaDesconche', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_tipoajuste', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_ConsumoProyectadoReal', 0, '', 0, 0, 0)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_derechosperfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_opcsistema', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_perfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios_eliminado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_usuariocontratos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('Log_Sistema', 0, '', 0, 0, 0)"

'-------> Permite cambiar la estructura del campo fechas a la tabla a_servicio
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horcob char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horent char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horpda char(14)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_totcompras
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecemi char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecven char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecrem char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecdig char(24)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecemi char(10)"
'dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecpro char(10)"
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

'-------> Permite cambiar la estructura del campo fechas tabla b_ocsacrecibido
'dbE.Execute "ALTER TABLE b_ocsacrecibido ALTER COLUMN ocr_fecoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_totventascaf ALTER COLUMN tvc_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecing char(10)"
'dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecdoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_detventascafpro ALTER COLUMN dvp_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE log_cierrediario ALTER COLUMN feccie char(10)"

dbE.Close
DoEvents: Formu.Bar1.Value = Formu.Bar1.Value + 1

If op Then
    
    If ValidaMDBCierreDiario(dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx") Then

        CalcularPMPDiaSql = False
   
        MsgBox "El Archivo de Cierre Diario se ha Generado con Error." & VgLinea & "Debe Cerrar El Día Nuevamente...", vbCritical + vbOKOnly, "Valida MDB Cierre Diario"
        Formu.Label1(1).Caption = "ERROR: Es Necesario Volver a Ejecutar el Cierre Diario."
   
        Exit Function

    End If

End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"


'-------> Actualizando cerrando periodo y abriendo proximo día
If op Then vg_db.Execute "UPDATE a_param SET par_valor = '" & fg_Encripta(LimpiaDato(CDate(vg_ciedia) + 1)) & "' WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'"

'-------> Grabar log_enviocierrediario
If op Then
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   sql1 = IIf(vg_tipbase = "1", " CDATE('vg_ciedia') ", " '" & Format(vg_ciedia, "yyyymmdd") & "' ")
   RS1.Open "SELECT DISTINCT fecha FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & "", vg_db, adOpenStatic
   If RS1.EOF Then
      
      vg_db.Execute "INSERT INTO log_enviocierrediario VALUES ('" & MuestraCasino(1) & "', " & sql1 & ", '0', '')"
   
   Else
      
      vg_db.Execute "UPDATE log_enviocierrediario SET estenv = '0', fecsub = '' WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & ""
   
   End If
   RS1.Close: Set RS1 = Nothing

End If

Formu.Bar1(0).Visible = False: Formu.Bar1(0).Value = 0
Formu.Label1(1).Visible = False

CalcularPMPDiaSql = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgp_p_ReducirLog '" & dir_bkpsql & "', '" & vg_SqlBase & "'")
         
If Not RS.EOF Then
            
   If RS(0) > 0 Then
               
      MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
      Exit Function
            
   End If
         
End If
         
RS.Close
Set RS = Nothing

fg_descarga

Exit Function
Error_CalcularPMPDia:
        
        CalcularPMPDiaSql = False
        fg_descarga
        MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
'       vg_db.RollbackTrans
       Resume Next

End Function

Function GeneraMDBCierreMes(Formu As Form) As Boolean

On Local Error GoTo Error_GeneraMDBCierreMes

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim i As Long
Dim sql1 As String
Dim FecCie As String

FecCie = CDate(vg_ciedia) - 1

GeneraMDBCierreMes = True

'-------> Grabar log_cierrediario
vg_db.Execute "INSERT INTO log_cierrediario VALUES ('" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "', '" & Format(FecCie, "yyyymmdd") & "', '" & Trim(vg_NUsr) & "', '1.- Cierre Diario', '" & MuestraCasino(1) & "')"

'-------> Crear tabla y mover datos
Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String, DBO As String

'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(FecCie, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile ' borrar base datos si existe
If Dir(mdir & MuestraCasino(1) & Format(FecCie, "yyyymmdd") & ".mdb") <> "" Then Kill mdir & MuestraCasino(1) & Format(FecCie, "yyyymmdd") & ".mdb" ' borrar base datos si existe

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open

'-------> Ini : Generar tabla Costo minuta & Food Cost
dbE.Execute "CREATE TABLE B_CostoMinutaRealizadoFoodCostN (IdCeco char(10), Fecha_Minuta int, IdRegimen int, IdServicio int, Raciones_Teorica int, " & _
            "Costo_Teorico_Alim float, Costo_Teorico_Desec float, Raciones_Real int, Costo_Real_Alim float, Costo_Real_Desec float, Raciones_Vendidas int, " & _
            "Costo_Realizado_Alim float, Costo_Realizado_Desec float, Venta_Dia float, Venta_Contado float, Glosa_Venta_Especial char(100)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraerCostoFoodCostMinutaCierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(FecCie, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_CostoMinutaRealizadoFoodCostN values (' " & RS(0) & "', " & RS(3) & ", " & RS(1) & ", " & RS(2) & ", " & RS(4) & " " & _
                  ", " & RS(5) & ", " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", " & RS(14) & ", '" & RS(15) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Costo minuta & Food Cost

'-------> Ini : Generar tabla Insumos & Food Cost A13
dbE.Execute "CREATE TABLE B_A13InsumosFCost (IdCeco char(10), Periodo int, FechaIni Int, FechaFin int, FechaCierre int, " & _
            "Glosa char(200), Alimentos Float, Lim_Desc Float, Total Float, Porcentaje Float, Id int)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_VtaCtoServInsumosFoodCostGastoA13CierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(CDate(FecCie), "yyyymmdd") & ", " & Format(FecCie, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_A13InsumosFCost values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & Format(CDate(FecCie), "yyyymmdd") & ", '" & RS(4) & "' " & _
                  ", " & RS(5) & ", '" & RS(6) & "', '" & RS(7) & "', " & RS(8) & ", " & RS(9) & ")"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Insumos & Food Cost A13

'-------> Ini : Generar tabla detalle mermas
dbE.Execute "CREATE TABLE B_DetalleMermasNK (IdCeco char(10), Periodo int, Fecha_Minuta int, IdRegimen int, IdServicio int, FechaCierre int, " & _
            "IdReceta int, IdEstServicio int, NumLin int, CostoRecetaAlimento float, CostoRecetaDesechable float, CantidadMerma float, CantidadRacionReal float, MermaxKilo float) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleMermasCierreDiario '" & MuestraCasino(1) & "', " & Format(FecCie, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_DetalleMermasNK values (' " & RS(0) & "', " & Format(FecCie, "yyyymm") & ", " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & Format(CDate(FecCie), "yyyymmdd") & ", " & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ")"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla detalle mermas

'-------> Ini : Generar tabla mermas desconche - pan - produccion
dbE.Execute "CREATE TABLE B_MermaDesconche (IdCeco char(10), IdRegimen int, IdServicio int, Fecha_Merma int, " & _
            "Considera_Merma char(1), Merma_Desconche float, Merma_Pan float, Merma_Produccion float, Fecha_Modificacion datetime, Fecha_Creacion datetime, Usuario char(20)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_MermaDesconcheCierreDiario '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_MermaDesconche values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", '" & RS(8) & "', '" & RS(9) & "', '" & RS(10) & "')"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla mermas desconche - pan - produccion

'-------> Ini : Generar tabla consumo proyectado & real
dbE.Execute "CREATE TABLE B_ConsumoProyectadoReal (IdCeco char(10), IdRegimen int, NoRegimen char(50), IdServicio int, NoServicio char(50), NumDoc int, Periodo int, Fecha int, " & _
            "Codigo_Producto char(20), Descripcion_Producto char(100), Unidad char(5), Cantidad_Teorica float, Cantidad_Planificada float, Cantidad_Realizada float, PMP float, Racion_Teorica int, Usuario_Mod_Racion_Real char(20), Racion_Real int, Fecha_Mod_Racion_Real datetime, Usuario_Salida_Produccion char(20), Racion_Salida_Produccion int, Fecha_Mod_Salida_Produccion datetime, Cantidad_Devolucion float)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CalcularSalidaProducciónMinutaTeorica '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymmdd") & ", '1', " & vg_codbod & ", " & Format(vg_ciedia, "yyyymmdd") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
        
      dbE.Execute "insert into B_ConsumoProyectadoReal values ('" & MuestraCasino(1) & "', " & RS(0) & ", '" & RS(1) & "', " & RS(2) & ", '" & RS(3) & "', " & RS(5) & "," & _
                  "" & Format(vg_ciedia, "yyyymm") & ", " & Format(vg_ciedia, "yyyymmdd") & ", '" & RS(6) & "', '" & RS(7) & "', '" & RS(8) & "', " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", '" & RS(14) & "', " & RS(15) & ", '" & RS(16) & "', '" & RS(17) & "', " & RS(18) & ", '" & RS(19) & "', '" & RS(20) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla consumo proyectado & real

'-------> Ini : Generar tabla a_tipoajuste 31/03/2022
dbE.Execute "CREATE TABLE a_tipoajuste (aju_codigo int, aju_nombre char(30), aju_tipaju int, aju_tipo char(1))"
dbE.Execute "INSERT INTO a_tipoajuste SELECT aju_codigo, aju_nombre, aju_tipaju, aju_tipo " & _
            "FROM a_tipoajuste IN " & DBO & ""
'-------> Fin : Generar tabla a_tipoajuste  31/03/2022

'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN " & DBO & ""
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN " & DBO & ""
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN " & DBO & ""
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN " & DBO & " " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale, 'dd/mm/yyyy') = CDATE('" & FecCie & "')"
            
'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
dbE.Execute "CREATE TABLE a_bodega (bod_codigo int, bod_nombre char(25), bod_ubicac char(35))"
dbE.Execute "INSERT INTO a_bodega SELECT bod_codigo, bod_nombre, bod_ubicac FROM a_bodega a, b_clientes b IN " & DBO & " " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo = 0"

'-------> Generar tabla proveedor
dbE.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo datetime, prv_origen char(1))"
dbE.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen FROM b_proveedor IN " & DBO & ""

'-------> Generar tabla compras
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " AND (toc_fecrem = cdate('" & FecCie & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"
            
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0') " & _
            "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
            "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
            "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras "
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R "
dbE.Execute "DROP TABLE b_totcompras_R"

'dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc int, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   (b.toc_fecrem = cdate('" & FecCie & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"

dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND  (b.toc_fecrem = cdate('" & FecCie & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R "
dbE.Execute "DROP TABLE b_detcomprasimp_R"

'-------> Actualizar estado documento
vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " AND (CONVERT(VARCHAR(8),toc_fecrem,112) = CONVERT(VARCHAR(8),CONVERT(DATE, CONVERT(VARCHAR(10),('" & FecCie & "')),103),112) OR isnull(EnvioDocSGPADM,'0') = '0')")

vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " " & _
               "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
               "AND  (isnull(toc_docsnc,'')='' or isnull(EnvioDocSGPADM,'0')='0') " & _
               "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
               "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
               "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)")

vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " " & _
               "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
               "AND  (isnull(toc_docaso,'')='' or isnull(EnvioDocSGPADM,'0')='0')")

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc datetime, ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc FROM b_ocsacrecibido a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & FecCie & "')"


'-------> Generar tabla ventas
dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc double, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
            "FROM b_totventas  IN " & DBO & " " & _
            "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = cdate('" & FecCie & "') AND tov_tipdoc IN ('DP','SP')) " & _
            "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = cdate('" & FecCie & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"

'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc double, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla ventas servicios especiales
dbE.Execute "CREATE TABLE b_totventaserviciosespeciales (tos_IdCeco char(10), tos_Tipo_Documento char(2), tos_Numero_Documento int, tos_Fecha_Produccion datetime, tos_IdBodega int, tos_Venta_servicio_Especiales char(100), Tos_Comensales double, tos_Precio_Servicio double, tos_Total_Documento double, tos_Estado_Documento char(1), tos_Fecha_Creacion datetime, tos_Fecha_Modificacion datetime, tos_Periodo int, tos_Usuario char(20), tos_Documento_Asociado int)"
dbE.Execute "INSERT INTO b_totventaserviciosespeciales SELECT tos_IdCeco, tos_Tipo_Documento, tos_Numero_Documento, tos_Fecha_Produccion, tos_IdBodega, tos_venta_Servicio_Especiales, Tos_Comensales, tos_Precio_servicio, tos_Total_Documento, tos_Estado_Documento, tos_Fecha_Creacion, tos_Fecha_modificacion, tos_Periodo, tos_Usuario, tos_Documento_Asociado " & _
            "FROM b_totventaserviciosespeciales  IN " & DBO & " " & _
            "WHERE tos_IdBodega = " & vg_codbod & " AND tos_fecha_Produccion = cdate('" & FecCie & "') AND tos_tipo_documento IN ('DE','SE') "

'-------> Crear estructura tabla detalle ventas servicios especiales
dbE.Execute "CREATE TABLE b_detventaserviciosespeciales (des_IdCeco char(10), des_Tipo_Documento char(2), des_Numero_Documento int, des_Numero_Linea int, des_IdProducto char(20), des_Cantidad_Mercaderia double, des_Cantidad_Devolver double, des_Precio_Documento double, des_Total_Documento double, des_Descripcion char(255), des_Mueve_Inventario char(1), des_Actualiza_Precio char(1))"
dbE.Execute "INSERT INTO b_detventaserviciosespeciales SELECT a.des_IdCeco, a.des_tipo_documento, a.des_numero_documento, a.des_numero_linea, a.des_IdProducto, a.des_cantidad_mercaderia, a.des_cantidad_devolver, a.des_Precio_Documento, a.des_Total_documento, a.des_Descripcion, a.des_Mueve_Inventario, a.des_Actualiza_Precio " & _
            "FROM b_detventaserviciosespeciales a, b_totventaserviciosespeciales b IN " & DBO & " " & _
            "WHERE a.des_IdCeco = b.tos_IdCeco " & _
            "AND   a.des_tipo_documento = b.tos_tipo_documento " & _
            "AND   a.des_numero_documento = b.tos_numero_documento " & _
            "AND   b.tos_fecha_produccion = cdate('" & FecCie & "') AND b.tos_tipo_documento IN ('DE','SE') " & _
            "AND   b.tos_IdBodega = " & vg_codbod & ""

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & FecCie & "')"

dbE.Execute "CREATE TABLE b_totventascaf (tvc_cencos char(10), tvc_fecing datetime, tvc_codbod int, tvc_estado char(1))"
dbE.Execute "INSERT INTO b_totventascaf SELECT tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado FROM b_totventascaf IN " & DBO & " " & _
            "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   tvc_fecing = cdate('" & FecCie & "') " & _
            "AND   tvc_codbod = " & vg_codbod & " " & _
            "AND   tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascaf (dvc_cencos char(10), dvc_fecing datetime, dvc_numlin int, dvc_articulo char(20), dvc_canart double, dvc_precio double, dvc_tippag char(2), dvc_rutcli char(10), dvc_cencli char(10), dvc_tipdoc char(2), dvc_numdoc int, dvc_fecdoc datetime)"
dbE.Execute "INSERT INTO b_detventascaf SELECT a.dvc_cencos, a.dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, a.dvc_fecdoc FROM b_detventascaf a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvc_cencos = b.tvc_cencos " & _
            "AND   a.dvc_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & FecCie & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascafpro (dvp_cencos char(10), dvp_fecing datetime, dvp_codmer char(20), dvp_cancal double, dvp_candig double, dvp_precos double)"
dbE.Execute "INSERT INTO b_detventascafpro SELECT a.dvp_cencos, a.dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos FROM b_detventascafpro a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvp_cencos = b.tvc_cencos " & _
            "AND   a.dvp_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & FecCie & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
dbE.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
dbE.Execute "INSERT INTO a_regimen SELECT reg_codigo, reg_nombre FROM a_regimen IN " & DBO & ""

'-------> Generar tabla servicio
dbE.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(30), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob datetime, ser_horent datetime, ser_horpda datetime)"
dbE.Execute "INSERT INTO a_servicio SELECT ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_activo, ser_horcob, ser_horent, ser_horpda FROM a_servicio IN " & DBO & ""

'-------> Generar tabla estructura de servicio
dbE.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double, ess_cencos char(10))"
dbE.Execute "INSERT INTO a_estservicio SELECT ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin, ess_cencos FROM a_estservicio IN " & DBO & " where ess_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla sector
dbE.Execute "CREATE TABLE a_sector (sec_codigo int, sec_nombre char(50), sec_orden int)"
dbE.Execute "INSERT INTO a_sector SELECT sec_codigo, sec_nombre, sec_orden FROM a_sector IN " & DBO & ""

'-------> Generar tabla servicio rac
dbE.Execute "CREATE TABLE a_serviciorac (sra_codser int, sra_coditem int, sra_serdia int, sra_raciones int, sra_cencos char(10))"
dbE.Execute "INSERT INTO a_serviciorac SELECT sra_codser, sra_coditem, sra_serdia, sra_raciones, sra_cencos FROM a_serviciorac IN " & DBO & " where sra_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then

dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'-------> Si tipo de minuta es distinto simap puede generar minuta cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
   
   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
               "AND   a.min_fecmin = " & Format(FecCie, "yyyymmdd") & ""

End If
    
dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'-------> Si tipo de minuta es distinto simap puede generar minuta detalle cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
    
    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
                "WHERE b.min_codigo = a.mid_codigo " & _
                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                "AND   b.min_fecmin = " & Format(FecCie, "yyyymmdd") & ""
End If
'-------> Generar tabla minutafija
dbE.Execute "CREATE TABLE b_minutafijadia (mfd_cencos char(10), mfd_codreg int, mfd_codser int, mfd_fecha int, mfd_codpro char(20), mfd_tipmin char(1), mfd_canpro double, mfd_cospro double)"
dbE.Execute "INSERT INTO b_minutafijadia SELECT mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro FROM b_minutafijadia IN " & DBO & " " & _
            "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mfd_fecha  = " & Format(FecCie, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
dbE.Execute "CREATE TABLE b_minutaraciones (mir_cencos char(10), mir_codreg int, mir_codser int, mir_fecmin int, mir_rutcli char(10), mir_nrorac int, mir_nroguia int, mir_codcli char(10))"
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) SELECT mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli FROM b_minutaraciones IN " & DBO & " " & _
            "WHERE mir_cencos          = '" & MuestraCasino(1) & "' " & _
            "AND   mid(mir_fecmin,1,6) = " & Format(FecCie, "yyyymm") & ""

'-------> Insertar mermas
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) " & _
"SELECT  bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin , " & _
"        'MERMAS' as mir_rutcli , " & _
"        SUM(bm2.mid_nummer) As mermas, " & _
"        0,                              " & _
"        ''                             " & _
"FROM    b_minuta AS bm, " & _
"        b_minutadet AS bm2 IN " & DBO & " " & _
"WHERE   bm.min_codigo = bm2.mid_codigo and bm.min_cencos = '" & MuestraCasino(1) & "' " & _
"        AND mid(bm.min_fecmin, 1, 6) = " & Format(FecCie, "yyyymm") & " " & _
"        AND bm2.mid_tipmin = '2' " & _
"GROUP BY bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin"

'-------> Generar tabla precio venta
dbE.Execute "CREATE TABLE b_preciovta (prv_cencos char(10), prv_codreg int, prv_codser int, prv_fecvig int, prv_rutcli char(10), prv_preven double)"
dbE.Execute "INSERT INTO b_preciovta SELECT prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli, prv_preven FROM b_preciovta IN " & DBO & " " & _
            "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
dbE.Execute "CREATE TABLE b_ventacontado (vtc_codigo int, vtc_cencos char(10), vtc_codreg int, vtc_codser int, vtc_fecvta int, vtc_forpag int, vtc_totmon double, vtc_opccli char(1))"
dbE.Execute "INSERT INTO b_ventacontado SELECT vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli FROM b_ventacontado IN " & DBO & " " & _
            "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   vtc_fecvta = " & Format(FecCie, "yyyymmdd") & ""

dbE.Execute "CREATE TABLE b_ventacontadodet (vtd_codigo int, vtd_numlin int, vtd_codcli char(10), vtd_codcco char(10), vtd_descripcion char(50), vtd_detmon double)"
dbE.Execute "INSERT INTO b_ventacontadodet SELECT a.vtd_codigo, a.vtd_numlin, a.vtd_codcli, a.vtd_codcco, a.vtd_descripcion, a.vtd_detmon FROM b_ventacontadodet a, b_ventacontado b IN " & DBO & " " & _
            "WHERE b.vtc_codigo = a.vtd_codigo " & _
            "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.vtc_fecvta = " & Format(FecCie, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
dbE.Execute "CREATE TABLE b_clientecencos (clc_codigo char(10), clc_codcli char(10), clc_nombre char(50))"
dbE.Execute "INSERT INTO b_clientecencos SELECT clc_codigo, clc_codcli, clc_nombre FROM b_clientecencos IN " & DBO & " where clc_codcli = '" & MuestraCasino(1) & "' "

'-------> Generar tabla cliente
dbE.Execute "CREATE TABLE b_clientes (cli_codigo char(10), cli_nombre char(50), cli_direccion char(50), cli_comuna char(15), cli_ciudad char(15), cli_fono1 char(15), cli_fono2 char(15), cli_fax char(15), cli_percon char(50), cli_giro char(50), cli_email char(50), cli_tipo int, cli_codbod int, cli_codtis int, cli_codseg int, cli_codcli char(10), cli_clisap char(1), cli_socsap char(4), cli_cievta char(1), cli_ciedia int, cli_activo char(1), cli_sobrec char(1), cli_codmun int, cli_ccisac int, cli_cecsac char(4), cli_codreg int, id_tipo_vale char(100))"
dbE.Execute "INSERT INTO b_clientes SELECT cli_codigo, cli_nombre, cli_direccion, cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, cli_giro, cli_email, cli_tipo, cli_codbod, cli_codtis, cli_codseg, cli_codcli, cli_clisap, cli_socsap, cli_cievta, cli_ciedia, cli_activo, cli_sobrec, cli_codmun, cli_ccisac, cli_cecsac, cli_codreg FROM b_clientes IN " & DBO & ""

'-------> Generar tabla bodegas
vg_db.Execute ("delete paso_b_bodegas where bod_Cencos = '" & MuestraCasino(1) & "'")
vg_db.Execute ("insert into paso_b_bodegas (bod_cencos, bod_codbod, bod_codpro, bod_canmer) SELECT distinct '" & MuestraCasino(1) & "', bod_codbod, bod_codpro, round(bod_canmer,2) FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " ")
dbE.Execute "CREATE TABLE b_bodegas (bod_codbod int, bod_codpro char(20), bod_canmer double)"
dbE.Execute "INSERT INTO b_bodegas SELECT distinct bod_codbod, bod_codpro, bod_canmer FROM paso_b_bodegas IN " & DBO & " WHERE bod_cencos = '" & MuestraCasino(1) & "' and bod_codbod = " & vg_codbod & ""

'-------> Generar tabla precio cafeteria
dbE.Execute "CREATE TABLE b_totpreciocaf (tpc_codigo char(20), tpc_nombre char(50), tpc_precio double, tpc_cencos char(10), tpc_activo char(1))"
dbE.Execute "INSERT INTO b_totpreciocaf SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_cencos, tpc_activo FROM b_totpreciocaf  IN " & DBO & " WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

dbE.Execute "CREATE TABLE b_detpreciocaf (dpc_codigo char(20), dpc_codmer char(20), dpc_cantidad double, dpc_cencos char(10))"
dbE.Execute "INSERT INTO b_detpreciocaf SELECT a.dpc_codigo, a.dpc_codmer, a.dpc_cantidad, a.dpc_cencos FROM b_detpreciocaf a, b_totpreciocaf b IN " & DBO & " " & _
            "WHERE b.tpc_codigo = a.dpc_codigo " & _
            "AND   b.tpc_cencos = a.dpc_cencos " & _
            "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

RS1.Open "SELECT DISTINCT b.cie_fecini, b.cie_fecter FROM b_tomainv a, b_cierreperiodo b " & _
         "WHERE b.cie_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   a.tin_fectom = " & Format(CDate(FecCie), "yyyymmdd") & " " & _
         "AND   a.tin_ciemes = b.cie_periodo", vg_db, adOpenStatic
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom >= " & RS1!cie_fecini & " AND tin_fectom <= " & RS1!cie_fecter & " " & _
               "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
   dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf FROM b_totventas IN " & DBO & " " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
    '-------> Generar tabla ajuste inventario detalle
    dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
                "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
                "WHERE a.dev_rutcli = b.tov_rutcli " & _
                "AND   a.dev_tipdoc = b.tov_tipdoc " & _
                "AND   a.dev_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
    
    '-------> Generar tabla ajuste inventario detalle impuesto
    dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
                "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
                "WHERE a.imd_rutdoc = b.tov_rutcli " & _
                "AND   a.imd_tipdoc = b.tov_tipdoc " & _
                "AND   a.imd_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
Else
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom = " & Format(FecCie, "yyyymmdd") & " " & _
               "AND   tin_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN " & DBO & " " & _
            "WHERE feccie >= cdate('" & FecCie & "') " & _
            "AND   feccie <= cdate('" & FecCie & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"

'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN " & DBO & " " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"

'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN " & DBO & " " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla a_derechosperfil
dbE.Execute "CREATE TABLE a_derechosperfil (dpe_cecori char(10), dpe_codper int, dpe_codopc int, dpe_deracc int, dpe_deragr int, dpe_dermod int, dpe_dereli int, dpe_derimp int)"
dbE.Execute "INSERT INTO a_derechosperfil (dpe_cecori, dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT distinct '" & MuestraCasino(1) & "', dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp FROM a_derechosperfil IN " & DBO & " "

'-------> actualizar tabla a_opcsistema
Dim ciedia As String
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE a_opcsistema SET EnvioDocSGPADM = '" & ciedia & "' FROM a_opcsistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = ''")

'-------> generar tabla a_opcsistema
dbE.Execute "CREATE TABLE a_opcsistema (opc_cecori char(10), opc_codigo int, opc_nombre char(50))"
dbE.Execute "INSERT INTO a_opcsistema (opc_cecori, opc_codigo, opc_nombre) SELECT distinct '" & MuestraCasino(1) & "', opc_codigo, opc_nombre FROM a_opcsistema IN " & DBO & "  where EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> generar tabla a_perfil
dbE.Execute "CREATE TABLE a_perfil (per_cecori char(10), per_codigo int, per_nombre char(30))"
dbE.Execute "INSERT INTO a_perfil (per_cecori, per_codigo, per_nombre) SELECT distinct '" & MuestraCasino(1) & "', per_codigo, per_nombre FROM a_perfil IN " & DBO & " "

'-------> generar tabla a_usuarios
dbE.Execute "CREATE TABLE a_usuarios (usu_cecori char(10), usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios (usu_cecori, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios IN " & DBO & " "

'-------> generar tabla a_usuarios_eliminado
dbE.Execute "CREATE TABLE a_usuarios_eliminado (usu_cecori char(10), Fecha datetime, usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios_eliminado (usu_cecori, Fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios_eliminado IN " & DBO & " "

'-------> generar tabla b_usuariocontratos
dbE.Execute "CREATE TABLE b_usuariocontratos (uco_cecori char(10), uco_codusu char(20), uco_codcon char(10))"
dbE.Execute "INSERT INTO b_usuariocontratos (uco_cecori, uco_codusu, uco_codcon) SELECT distinct '" & MuestraCasino(1) & "', uco_codusu, uco_codcon FROM b_usuariocontratos IN " & DBO & " "

'-------> actualizar tabla log_sistema
ciedia = ""
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE Log_Sistema SET EnvioDocSGPADM = '" & ciedia & "' FROM Log_Sistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = '' and Loc_Id in (2,20,21,22,23)")
               
'-------> generar tabla Log_Sistema
dbE.Execute "CREATE TABLE Log_Sistema (cecori char(10), Fecha datetime, Usuario_Id char(20), Loc_Id int, Opcion_Sistema char(14), Dato_Nuevo memo, Dato_Anterior memo, Detalle_Operacion memo)"
dbE.Execute "INSERT INTO Log_Sistema (cecori, Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion) SELECT distinct '" & MuestraCasino(1) & "', Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion FROM Log_Sistema IN " & DBO & " where Loc_Id in (2,20,21,22,23) and EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_CostoMinutaRealizadoFoodCostN', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_A13InsumosFCost', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_DetalleMermasNK', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_MermaDesconche', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_tipoajuste', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_ConsumoProyectadoReal', 0, '', 0, 0, 0)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_derechosperfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_opcsistema', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_perfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios_eliminado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_usuariocontratos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('Log_Sistema', 0, '', 0, 0, 0)"


'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

dbE.Close
DoEvents

    If ValidaMDBCierreDiario(dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx") Then

        GeneraMDBCierreMes = False
   
        MsgBox "El Archivo de Cierre Diario se ha Generado con Error." & VgLinea & "Debe Cerrar El Día Nuevamente...", vbCritical + vbOKOnly, "Valida MDB Cierre Diario"
        Formu.Label1(1).Caption = "ERROR: Es Necesario Volver a Ejecutar el Cierre Diario."
   
        Exit Function

    End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"

'-------> Grabar log_enviocierrediario
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
   
sql1 = IIf(vg_tipbase = "1", " CDATE('FecCie') ", " '" & Format(FecCie, "yyyymmdd") & "' ")
RS1.Open "SELECT DISTINCT fecha FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & "", vg_db, adOpenStatic
If RS1.EOF Then
      
   vg_db.Execute "INSERT INTO log_enviocierrediario VALUES ('" & MuestraCasino(1) & "', " & sql1 & ", '0', '')"
   
Else
      
   vg_db.Execute "UPDATE log_enviocierrediario SET estenv = '0', fecsub = '' WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & ""
   
End If
RS1.Close: Set RS1 = Nothing

GeneraMDBCierreMes = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgp_p_ReducirLog '" & dir_bkpsql & "', '" & vg_SqlBase & "'")
         
If Not RS.EOF Then
            
   If RS(0) > 0 Then
               
      MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
      Exit Function
            
   End If
         
End If
         
RS.Close
Set RS = Nothing

fg_descarga

Exit Function
Error_GeneraMDBCierreMes:
        
       GeneraMDBCierreMes = False
       fg_descarga
       MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
       Resume Next

End Function

Function GeneraMDBInventario(Formu As Form, FecInv As Date) As Boolean

On Local Error GoTo Error_GeneraMDBInventario

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim i As Long
Dim sql1 As String
Dim FecCie As String

FecCie = CDate(FecInv)

GeneraMDBInventario = True

'-------> Grabar log_cierrediario
vg_db.Execute "INSERT INTO log_cierrediario VALUES ('" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "', '" & Format(FecCie, "yyyymmdd") & "', '" & Trim(vg_NUsr) & "', '1.- Cierre Diario', '" & MuestraCasino(1) & "')"

'-------> Crear tabla y mover datos
Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String, DBO As String

'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(FecCie, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile ' borrar base datos si existe
If Dir(mdir & MuestraCasino(1) & Format(FecCie, "yyyymmdd") & ".mdb") <> "" Then Kill mdir & MuestraCasino(1) & Format(FecCie, "yyyymmdd") & ".mdb"  ' borrar base datos si existe

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open

'-------> Ini : Generar tabla Costo minuta & Food Cost
dbE.Execute "CREATE TABLE B_CostoMinutaRealizadoFoodCostN (IdCeco char(10), Fecha_Minuta int, IdRegimen int, IdServicio int, Raciones_Teorica int, " & _
            "Costo_Teorico_Alim float, Costo_Teorico_Desec float, Raciones_Real int, Costo_Real_Alim float, Costo_Real_Desec float, Raciones_Vendidas int, " & _
            "Costo_Realizado_Alim float, Costo_Realizado_Desec float, Venta_Dia float, Venta_Contado float, Glosa_Venta_Especial char(100)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraerCostoFoodCostMinutaCierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(FecCie, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_CostoMinutaRealizadoFoodCostN values (' " & RS(0) & "', " & RS(3) & ", " & RS(1) & ", " & RS(2) & ", " & RS(4) & " " & _
                  ", " & RS(5) & ", " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", " & RS(14) & ", '" & RS(15) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Costo minuta & Food Cost

'-------> Ini : Generar tabla Insumos & Food Cost A13
dbE.Execute "CREATE TABLE B_A13InsumosFCost (IdCeco char(10), Periodo int, FechaIni Int, FechaFin int, FechaCierre int, " & _
            "Glosa char(200), Alimentos Float, Lim_Desc Float, Total Float, Porcentaje Float, Id int)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_VtaCtoServInsumosFoodCostGastoA13CierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(CDate(FecCie), "yyyymmdd") & ", " & Format(FecCie, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_A13InsumosFCost values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & Format(CDate(FecCie), "yyyymmdd") & ", '" & RS(4) & "' " & _
                  ", " & RS(5) & ", '" & RS(6) & "', '" & RS(7) & "', " & RS(8) & ", " & RS(9) & ")"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Insumos & Food Cost A13

'-------> Ini : Generar tabla detalle mermas
dbE.Execute "CREATE TABLE B_DetalleMermasNK (IdCeco char(10), Periodo int, Fecha_Minuta int, IdRegimen int, IdServicio int, FechaCierre int, " & _
            "IdReceta int, IdEstServicio int, NumLin int, CostoRecetaAlimento float, CostoRecetaDesechable float, CantidadMerma float, CantidadRacionReal float, MermaxKilo float) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleMermasCierreDiario '" & MuestraCasino(1) & "', " & Format(FecCie, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_DetalleMermasNK values (' " & RS(0) & "', " & Format(FecCie, "yyyymm") & ", " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & Format(CDate(FecCie), "yyyymmdd") & ", " & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ")"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla detalle mermas

'-------> Ini : Generar tabla mermas desconche - pan - produccion
dbE.Execute "CREATE TABLE B_MermaDesconche (IdCeco char(10), IdRegimen int, IdServicio int, Fecha_Merma int, " & _
            "Considera_Merma char(1), Merma_Desconche float, Merma_Pan float, Merma_Produccion float, Fecha_Modificacion datetime, Fecha_Creacion datetime, Usuario char(20)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_MermaDesconcheCierreDiario '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_MermaDesconche values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", '" & RS(8) & "', '" & RS(9) & "', '" & RS(10) & "')"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla mermas desconche - pan - produccion

'-------> Ini : Generar tabla consumo proyectado & real
dbE.Execute "CREATE TABLE B_ConsumoProyectadoReal (IdCeco char(10), IdRegimen int, NoRegimen char(50), IdServicio int, NoServicio char(50), NumDoc int, Periodo int, Fecha int, " & _
            "Codigo_Producto char(20), Descripcion_Producto char(100), Unidad char(5), Cantidad_Teorica float, Cantidad_Planificada float, Cantidad_Realizada float, PMP float, Racion_Teorica int, Usuario_Mod_Racion_Real char(20), Racion_Real int, Fecha_Mod_Racion_Real datetime, Usuario_Salida_Produccion char(20), Racion_Salida_Produccion int, Fecha_Mod_Salida_Produccion datetime, Cantidad_Devolucion float)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CalcularSalidaProducciónMinutaTeorica '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymmdd") & ", '1', " & vg_codbod & ", " & Format(vg_ciedia, "yyyymmdd") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_ConsumoProyectadoReal values ('" & MuestraCasino(1) & "', " & RS(0) & ", '" & RS(1) & "', " & RS(2) & ", '" & RS(3) & "', " & RS(5) & "," & _
                  "" & Format(vg_ciedia, "yyyymm") & ", " & Format(vg_ciedia, "yyyymmdd") & ", '" & RS(6) & "', '" & RS(7) & "', '" & RS(8) & "', " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", '" & RS(14) & "', " & RS(15) & ", '" & RS(16) & "', '" & RS(17) & "', " & RS(18) & ", '" & RS(19) & "', '" & RS(20) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla consumo proyectado & real

'-------> Ini : Generar tabla a_tipoajuste 31/03/2022
dbE.Execute "CREATE TABLE a_tipoajuste (aju_codigo int, aju_nombre char(30), aju_tipaju int, aju_tipo char(1))"
dbE.Execute "INSERT INTO a_tipoajuste SELECT aju_codigo, aju_nombre, aju_tipaju, aju_tipo " & _
            "FROM a_tipoajuste IN " & DBO & ""
'-------> Fin : Generar tabla a_tipoajuste  31/03/2022

'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN " & DBO & ""
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN " & DBO & ""
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN " & DBO & ""
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN " & DBO & " " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale, 'dd/mm/yyyy') = CDATE('" & FecCie & "')"
            
'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
dbE.Execute "CREATE TABLE a_bodega (bod_codigo int, bod_nombre char(25), bod_ubicac char(35))"
dbE.Execute "INSERT INTO a_bodega SELECT bod_codigo, bod_nombre, bod_ubicac FROM a_bodega a, b_clientes b IN " & DBO & " " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo = 0"

'-------> Generar tabla proveedor
dbE.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo datetime, prv_origen char(1))"
dbE.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen FROM b_proveedor IN " & DBO & ""

'-------> Generar tabla compras
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " AND (toc_fecrem = cdate('" & FecCie & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"
            
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0') " & _
            "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
            "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
            "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras "
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R "
dbE.Execute "DROP TABLE b_totcompras_R"

'dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc int, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   (b.toc_fecrem = cdate('" & FecCie & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"

dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND  (b.toc_fecrem = cdate('" & FecCie & "') OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in ( select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso) OR EnvioDocSGPADM is null OR EnvioDocSGPADM = '0')"

dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R "
dbE.Execute "DROP TABLE b_detcomprasimp_R"

'-------> Actualizar estado documento
vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " AND (CONVERT(VARCHAR(8),toc_fecrem,112) = CONVERT(VARCHAR(8),CONVERT(DATE, CONVERT(VARCHAR(10),('" & FecCie & "')),103),112) OR isnull(EnvioDocSGPADM,'0') = '0')")

vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " " & _
               "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
               "AND  (isnull(toc_docsnc,'')='' or isnull(EnvioDocSGPADM,'0')='0') " & _
               "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
               "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
               "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)")

vg_db.Execute ("UPDATE b_totcompras SET EnvioDocSGPADM = '1' FROM b_totcompras " & _
               "WHERE toc_codbod = " & vg_codbod & " " & _
               "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
               "AND  (isnull(toc_docaso,'')='' or isnull(EnvioDocSGPADM,'0')='0')")

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc datetime, ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc FROM b_ocsacrecibido a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & FecCie & "')"


'-------> Generar tabla ventas
dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc double, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
            "FROM b_totventas  IN " & DBO & " " & _
            "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = cdate('" & FecCie & "') AND tov_tipdoc IN ('DP','SP')) " & _
            "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = cdate('" & FecCie & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"

'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc double, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & FecCie & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla ventas servicios especiales
dbE.Execute "CREATE TABLE b_totventaserviciosespeciales (tos_IdCeco char(10), tos_Tipo_Documento char(2), tos_Numero_Documento int, tos_Fecha_Produccion datetime, tos_IdBodega int, tos_Venta_servicio_Especiales char(100), Tos_Comensales double, tos_Precio_Servicio double, tos_Total_Documento double, tos_Estado_Documento char(1), tos_Fecha_Creacion datetime, tos_Fecha_Modificacion datetime, tos_Periodo int, tos_Usuario char(20), tos_Documento_Asociado int)"
dbE.Execute "INSERT INTO b_totventaserviciosespeciales SELECT tos_IdCeco, tos_Tipo_Documento, tos_Numero_Documento, tos_Fecha_Produccion, tos_IdBodega, tos_venta_Servicio_Especiales, Tos_Comensales, tos_Precio_servicio, tos_Total_Documento, tos_Estado_Documento, tos_Fecha_Creacion, tos_Fecha_modificacion, tos_Periodo, tos_Usuario, tos_Documento_Asociado " & _
            "FROM b_totventaserviciosespeciales  IN " & DBO & " " & _
            "WHERE tos_IdBodega = " & vg_codbod & " AND tos_fecha_Produccion = cdate('" & FecCie & "') AND tos_tipo_documento IN ('DE','SE') "

'-------> Crear estructura tabla detalle ventas servicios especiales
dbE.Execute "CREATE TABLE b_detventaserviciosespeciales (des_IdCeco char(10), des_Tipo_Documento char(2), des_Numero_Documento int, des_Numero_Linea int, des_IdProducto char(20), des_Cantidad_Mercaderia double, des_Cantidad_Devolver double, des_Precio_Documento double, des_Total_Documento double, des_Descripcion char(255), des_Mueve_Inventario char(1), des_Actualiza_Precio char(1))"
dbE.Execute "INSERT INTO b_detventaserviciosespeciales SELECT a.des_IdCeco, a.des_tipo_documento, a.des_numero_documento, a.des_numero_linea, a.des_IdProducto, a.des_cantidad_mercaderia, a.des_cantidad_devolver, a.des_Precio_Documento, a.des_Total_documento, a.des_Descripcion, a.des_Mueve_Inventario, a.des_Actualiza_Precio " & _
            "FROM b_detventaserviciosespeciales a, b_totventaserviciosespeciales b IN " & DBO & " " & _
            "WHERE a.des_IdCeco = b.tos_IdCeco " & _
            "AND   a.des_tipo_documento = b.tos_tipo_documento " & _
            "AND   a.des_numero_documento = b.tos_numero_documento " & _
            "AND   b.tos_fecha_produccion = cdate('" & FecCie & "') AND b.tos_tipo_documento IN ('DE','SE') " & _
            "AND   b.tos_IdBodega = " & vg_codbod & ""

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & FecCie & "')"

dbE.Execute "CREATE TABLE b_totventascaf (tvc_cencos char(10), tvc_fecing datetime, tvc_codbod int, tvc_estado char(1))"
dbE.Execute "INSERT INTO b_totventascaf SELECT tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado FROM b_totventascaf IN " & DBO & " " & _
            "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   tvc_fecing = cdate('" & FecCie & "') " & _
            "AND   tvc_codbod = " & vg_codbod & " " & _
            "AND   tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascaf (dvc_cencos char(10), dvc_fecing datetime, dvc_numlin int, dvc_articulo char(20), dvc_canart double, dvc_precio double, dvc_tippag char(2), dvc_rutcli char(10), dvc_cencli char(10), dvc_tipdoc char(2), dvc_numdoc int, dvc_fecdoc datetime)"
dbE.Execute "INSERT INTO b_detventascaf SELECT a.dvc_cencos, a.dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, a.dvc_fecdoc FROM b_detventascaf a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvc_cencos = b.tvc_cencos " & _
            "AND   a.dvc_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & FecCie & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascafpro (dvp_cencos char(10), dvp_fecing datetime, dvp_codmer char(20), dvp_cancal double, dvp_candig double, dvp_precos double)"
dbE.Execute "INSERT INTO b_detventascafpro SELECT a.dvp_cencos, a.dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos FROM b_detventascafpro a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvp_cencos = b.tvc_cencos " & _
            "AND   a.dvp_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & FecCie & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
dbE.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
dbE.Execute "INSERT INTO a_regimen SELECT reg_codigo, reg_nombre FROM a_regimen IN " & DBO & ""

'-------> Generar tabla servicio
dbE.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(30), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob datetime, ser_horent datetime, ser_horpda datetime)"
dbE.Execute "INSERT INTO a_servicio SELECT ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_activo, ser_horcob, ser_horent, ser_horpda FROM a_servicio IN " & DBO & ""

'-------> Generar tabla estructura de servicio
dbE.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double, ess_cencos char(10))"
dbE.Execute "INSERT INTO a_estservicio SELECT ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin, ess_cencos FROM a_estservicio IN " & DBO & " where ess_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla sector
dbE.Execute "CREATE TABLE a_sector (sec_codigo int, sec_nombre char(50), sec_orden int)"
dbE.Execute "INSERT INTO a_sector SELECT sec_codigo, sec_nombre, sec_orden FROM a_sector IN " & DBO & ""

'-------> Generar tabla servicio rac
dbE.Execute "CREATE TABLE a_serviciorac (sra_codser int, sra_coditem int, sra_serdia int, sra_raciones int, sra_cencos char(10))"
dbE.Execute "INSERT INTO a_serviciorac SELECT sra_codser, sra_coditem, sra_serdia, sra_raciones, sra_cencos FROM a_serviciorac IN " & DBO & " where sra_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then

dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'-------> Si tipo de minuta es distinto simap puede generar minuta cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
   
   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
               "AND   a.min_fecmin = " & Format(FecCie, "yyyymmdd") & ""

End If
    
dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'-------> Si tipo de minuta es distinto simap puede generar minuta detalle cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
    
    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
                "WHERE b.min_codigo = a.mid_codigo " & _
                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                "AND   b.min_fecmin = " & Format(FecCie, "yyyymmdd") & ""
End If
'-------> Generar tabla minutafija
dbE.Execute "CREATE TABLE b_minutafijadia (mfd_cencos char(10), mfd_codreg int, mfd_codser int, mfd_fecha int, mfd_codpro char(20), mfd_tipmin char(1), mfd_canpro double, mfd_cospro double)"
dbE.Execute "INSERT INTO b_minutafijadia SELECT mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro FROM b_minutafijadia IN " & DBO & " " & _
            "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mfd_fecha  = " & Format(FecCie, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
dbE.Execute "CREATE TABLE b_minutaraciones (mir_cencos char(10), mir_codreg int, mir_codser int, mir_fecmin int, mir_rutcli char(10), mir_nrorac int, mir_nroguia int, mir_codcli char(10))"
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) SELECT mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli FROM b_minutaraciones IN " & DBO & " " & _
            "WHERE mir_cencos          = '" & MuestraCasino(1) & "' " & _
            "AND   mid(mir_fecmin,1,6) = " & Format(FecCie, "yyyymm") & ""

'-------> Insertar mermas
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) " & _
"SELECT  bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin , " & _
"        'MERMAS' as mir_rutcli , " & _
"        SUM(bm2.mid_nummer) As mermas, " & _
"        0,                              " & _
"        ''                             " & _
"FROM    b_minuta AS bm, " & _
"        b_minutadet AS bm2 IN " & DBO & " " & _
"WHERE   bm.min_codigo = bm2.mid_codigo and bm.min_cencos = '" & MuestraCasino(1) & "' " & _
"        AND mid(bm.min_fecmin, 1, 6) = " & Format(FecCie, "yyyymm") & " " & _
"        AND bm2.mid_tipmin = '2' " & _
"GROUP BY bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin"

'-------> Generar tabla precio venta
dbE.Execute "CREATE TABLE b_preciovta (prv_cencos char(10), prv_codreg int, prv_codser int, prv_fecvig int, prv_rutcli char(10), prv_preven double)"
dbE.Execute "INSERT INTO b_preciovta SELECT prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli, prv_preven FROM b_preciovta IN " & DBO & " " & _
            "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
dbE.Execute "CREATE TABLE b_ventacontado (vtc_codigo int, vtc_cencos char(10), vtc_codreg int, vtc_codser int, vtc_fecvta int, vtc_forpag int, vtc_totmon double, vtc_opccli char(1))"
dbE.Execute "INSERT INTO b_ventacontado SELECT vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli FROM b_ventacontado IN " & DBO & " " & _
            "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   vtc_fecvta = " & Format(FecCie, "yyyymmdd") & ""

dbE.Execute "CREATE TABLE b_ventacontadodet (vtd_codigo int, vtd_numlin int, vtd_codcli char(10), vtd_codcco char(10), vtd_descripcion char(50), vtd_detmon double)"
dbE.Execute "INSERT INTO b_ventacontadodet SELECT a.vtd_codigo, a.vtd_numlin, a.vtd_codcli, a.vtd_codcco, a.vtd_descripcion, a.vtd_detmon FROM b_ventacontadodet a, b_ventacontado b IN " & DBO & " " & _
            "WHERE b.vtc_codigo = a.vtd_codigo " & _
            "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.vtc_fecvta = " & Format(FecCie, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
dbE.Execute "CREATE TABLE b_clientecencos (clc_codigo char(10), clc_codcli char(10), clc_nombre char(50))"
dbE.Execute "INSERT INTO b_clientecencos SELECT clc_codigo, clc_codcli, clc_nombre FROM b_clientecencos IN " & DBO & " where clc_codcli = '" & MuestraCasino(1) & "' "

'-------> Generar tabla cliente
dbE.Execute "CREATE TABLE b_clientes (cli_codigo char(10), cli_nombre char(50), cli_direccion char(50), cli_comuna char(15), cli_ciudad char(15), cli_fono1 char(15), cli_fono2 char(15), cli_fax char(15), cli_percon char(50), cli_giro char(50), cli_email char(50), cli_tipo int, cli_codbod int, cli_codtis int, cli_codseg int, cli_codcli char(10), cli_clisap char(1), cli_socsap char(4), cli_cievta char(1), cli_ciedia int, cli_activo char(1), cli_sobrec char(1), cli_codmun int, cli_ccisac int, cli_cecsac char(4), cli_codreg int, id_tipo_vale char(100))"
dbE.Execute "INSERT INTO b_clientes SELECT cli_codigo, cli_nombre, cli_direccion, cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, cli_giro, cli_email, cli_tipo, cli_codbod, cli_codtis, cli_codseg, cli_codcli, cli_clisap, cli_socsap, cli_cievta, cli_ciedia, cli_activo, cli_sobrec, cli_codmun, cli_ccisac, cli_cecsac, cli_codreg FROM b_clientes IN " & DBO & ""

'-------> Generar tabla bodegas
vg_db.Execute ("delete paso_b_bodegas where bod_Cencos = '" & MuestraCasino(1) & "'")
vg_db.Execute ("insert into paso_b_bodegas (bod_cencos, bod_codbod, bod_codpro, bod_canmer) SELECT distinct '" & MuestraCasino(1) & "', bod_codbod, bod_codpro, round(bod_canmer,2) FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " ")
dbE.Execute "CREATE TABLE b_bodegas (bod_codbod int, bod_codpro char(20), bod_canmer double)"
dbE.Execute "INSERT INTO b_bodegas SELECT distinct bod_codbod, bod_codpro, bod_canmer FROM paso_b_bodegas IN " & DBO & " WHERE bod_cencos = '" & MuestraCasino(1) & "' and bod_codbod = " & vg_codbod & ""

'-------> Generar tabla precio cafeteria
dbE.Execute "CREATE TABLE b_totpreciocaf (tpc_codigo char(20), tpc_nombre char(50), tpc_precio double, tpc_cencos char(10), tpc_activo char(1))"
dbE.Execute "INSERT INTO b_totpreciocaf SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_cencos, tpc_activo FROM b_totpreciocaf  IN " & DBO & " WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

dbE.Execute "CREATE TABLE b_detpreciocaf (dpc_codigo char(20), dpc_codmer char(20), dpc_cantidad double, dpc_cencos char(10))"
dbE.Execute "INSERT INTO b_detpreciocaf SELECT a.dpc_codigo, a.dpc_codmer, a.dpc_cantidad, a.dpc_cencos FROM b_detpreciocaf a, b_totpreciocaf b IN " & DBO & " " & _
            "WHERE b.tpc_codigo = a.dpc_codigo " & _
            "AND   b.tpc_cencos = a.dpc_cencos " & _
            "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

Set RS1 = vg_db.Execute("SELECT DISTINCT a.tin_fectom FROM b_tomainv a with (nolock) " & _
         "inner join b_clientes b with (nolock) on b.cli_codbod = a.tin_codbod and isnull(b.cli_tipo,0) = 0 and  isnull(b.cli_activo,'') = '1' " & _
         "WHERE b.cli_codigo =   '" & MuestraCasino(1) & "' AND a.tin_fectom = " & Format(CDate(FecCie), "yyyymmdd") & "")
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom = " & RS1!tin_fectom & " " & _
               "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
   dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf FROM b_totventas IN " & DBO & " " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!tin_fectom) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!tin_fectom) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
    '-------> Generar tabla ajuste inventario detalle
    dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
                "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
                "WHERE a.dev_rutcli = b.tov_rutcli " & _
                "AND   a.dev_tipdoc = b.tov_tipdoc " & _
                "AND   a.dev_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!tin_fectom) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!tin_fectom) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
    
    '-------> Generar tabla ajuste inventario detalle impuesto
    dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
                "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
                "WHERE a.imd_rutdoc = b.tov_rutcli " & _
                "AND   a.imd_tipdoc = b.tov_tipdoc " & _
                "AND   a.imd_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!tin_fectom) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!tin_fectom) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN " & DBO & " " & _
            "WHERE feccie >= cdate('" & FecCie & "') " & _
            "AND   feccie <= cdate('" & FecCie & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"

'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN " & DBO & " " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"

'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"

'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN " & DBO & " " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & FecCie & "')"


'-------> generar tabla a_derechosperfil
dbE.Execute "CREATE TABLE a_derechosperfil (dpe_cecori char(10), dpe_codper int, dpe_codopc int, dpe_deracc int, dpe_deragr int, dpe_dermod int, dpe_dereli int, dpe_derimp int)"
dbE.Execute "INSERT INTO a_derechosperfil (dpe_cecori, dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT distinct '" & MuestraCasino(1) & "', dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp FROM a_derechosperfil IN " & DBO & " "

'-------> actualizar tabla a_opcsistema
Dim ciedia As String
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE a_opcsistema SET EnvioDocSGPADM = '" & ciedia & "' FROM a_opcsistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = ''")

'-------> generar tabla a_opcsistema
dbE.Execute "CREATE TABLE a_opcsistema (opc_cecori char(10), opc_codigo int, opc_nombre char(50))"
dbE.Execute "INSERT INTO a_opcsistema (opc_cecori, opc_codigo, opc_nombre) SELECT distinct '" & MuestraCasino(1) & "', opc_codigo, opc_nombre FROM a_opcsistema IN " & DBO & "  where EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> generar tabla a_perfil
dbE.Execute "CREATE TABLE a_perfil (per_cecori char(10), per_codigo int, per_nombre char(30))"
dbE.Execute "INSERT INTO a_perfil (per_cecori, per_codigo, per_nombre) SELECT distinct '" & MuestraCasino(1) & "', per_codigo, per_nombre FROM a_perfil IN " & DBO & " "

'-------> generar tabla a_usuarios
dbE.Execute "CREATE TABLE a_usuarios (usu_cecori char(10), usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios (usu_cecori, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios IN " & DBO & " "

'-------> generar tabla a_usuarios_eliminado
dbE.Execute "CREATE TABLE a_usuarios_eliminado (usu_cecori char(10), Fecha datetime, usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios_eliminado (usu_cecori, Fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios_eliminado IN " & DBO & " "

'-------> generar tabla b_usuariocontratos
dbE.Execute "CREATE TABLE b_usuariocontratos (uco_cecori char(10), uco_codusu char(20), uco_codcon char(10))"
dbE.Execute "INSERT INTO b_usuariocontratos (uco_cecori, uco_codusu, uco_codcon) SELECT distinct '" & MuestraCasino(1) & "', uco_codusu, uco_codcon FROM b_usuariocontratos IN " & DBO & " "

'-------> actualizar tabla log_sistema
ciedia = ""
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE Log_Sistema SET EnvioDocSGPADM = '" & ciedia & "' FROM Log_Sistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = '' and Loc_Id in (2,20,21,22,23)")
               
'-------> generar tabla Log_Sistema
dbE.Execute "CREATE TABLE Log_Sistema (cecori char(10), Fecha datetime, Usuario_Id char(20), Loc_Id int, Opcion_Sistema char(14), Dato_Nuevo memo, Dato_Anterior memo, Detalle_Operacion memo)"
dbE.Execute "INSERT INTO Log_Sistema (cecori, Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion) SELECT distinct '" & MuestraCasino(1) & "', Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion FROM Log_Sistema IN " & DBO & " where Loc_Id in (2,20,21,22,23) and EnvioDocSGPADM = cdate('" & vg_ciedia & "') "



'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_CostoMinutaRealizadoFoodCostN', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_A13InsumosFCost', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_DetalleMermasNK', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_MermaDesconche', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_tipoajuste', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_ConsumoProyectadoReal', 0, '', 0, 0, 0)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_derechosperfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_opcsistema', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_perfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios_eliminado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_usuariocontratos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('Log_Sistema', 0, '', 0, 0, 0)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

dbE.Close
DoEvents

If ValidaMDBCierreDiario(dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx") Then

        GeneraMDBInventario = False
   
        
'        MsgBox "El Archivo de Cierre Diario se ha Generado con Error." & VgLinea & "Debe Cerrar El Día Nuevamente...", vbCritical + vbOKOnly, "Valida MDB Cierre Diario"
'        Formu.Label1(1).Caption = "ERROR: Es Necesario Volver a Ejecutar el Cierre Diario."
   
   Exit Function

End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"

'-------> Grabar log_enviocierrediario
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

sql1 = IIf(vg_tipbase = "1", " CDATE('FecCie') ", " '" & Format(FecCie, "yyyymmdd") & "' ")
RS1.Open "SELECT DISTINCT fecha FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & "", vg_db, adOpenStatic
If RS1.EOF Then

   vg_db.Execute "INSERT INTO log_enviocierrediario VALUES ('" & MuestraCasino(1) & "', " & sql1 & ", '0', '')"

Else

   vg_db.Execute "UPDATE log_enviocierrediario SET estenv = '0', fecsub = '' WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & ""

End If
RS1.Close: Set RS1 = Nothing

GeneraMDBInventario = True

fg_descarga

Exit Function
Error_GeneraMDBInventario:
        
       GeneraMDBInventario = False
       fg_descarga
       MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
       Resume Next

End Function

Function CalcularPMPDiaAccess(Formu As Form, op As Boolean, progrl As Boolean)

On Local Error GoTo Error_CalcularPMPDia
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String, aAp1 As String, aAp2 As String, aAp3 As String, aAp4 As String, aAp5 As String, aAp6 As String
Dim fecini As Long, fecfin As Long, i As Long, FecInv As Long, est As Boolean, fecpro As Date, fecter As Date
Dim estpro As Integer, sql1 As String
'00 Inventario inicial
'10 Ajuste de Entrada
'20 Proveedores Entrada
'30 Traspaso Entrada
'40 Produccion Salida
'50 Produccion Entrada
'60 Traspaso Salida
'70 Mermas Salida
'80 Venta Directa Salida
'90 Ajuste Salida
'100 Venta Cafeteria
'120 Ajuste de inventario Entrada
'130 Ajuste de inventario Salida
'-------> buscar ultimo inventario
estpro = 1
FecInv = 0
RS1.Open "SELECT MAX(tin_fectom) AS tin_fectom FROM b_tomainv WHERE tin_fectom<" & Format(CDate(vg_ciedia), "yyyymmdd") & " AND tin_codbod=" & vg_codbod & "", vg_db, adOpenStatic
If Not RS1.EOF And Not IsNull(RS1!tin_fectom) Then FecInv = IIf(RS1!tin_fectom > 0, RS1!tin_fectom, 0)
RS1.Close: Set RS1 = Nothing
If FecInv < 1 Then FecInv = Format(CDate(vg_ciedia), "yyyymmdd")
DoEvents
'vg_db.BeginTrans
'-------> Creo tabla temporal producto PMP
aAp1 = Trim(vg_NUsr) & "_tmp_ProdPMPDia"
fg_CheckTmp aAp1
RS1.Open "SELECT '' AS pmp_cencos, 0 AS pmp_fecha, '' AS pmp_codpro, round(0, 2) AS pmp_precio, round(0, " & vg_DCa & ") AS pmp_saldo INTO " & aAp1 & "", vg_db, adOpenStatic
Set RS1 = Nothing
'-------> Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_CalcularPMPDia"
fg_CheckTmp aAp
fg_carga ""
If op Then Formu.Label1(1).Caption = "Actua. Documento"
If FecInv > 0 Then
   RS1.Open "SELECT count(*) as Nreg FROM b_productospmpdia " & _
            "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " ", vg_db, adOpenStatic
   If TipoDato(RS1!nreg, 0) = 0 Then
      '-------> Traer Inventario primer inventario
      vg_db.Execute "SELECT DISTINCT " & Format(CDate(fg_Ctod1(FecInv)) + 1, "yyyymmdd") & " AS fecpro, tin_codpro AS codpro, round(tin_stofis, " & vg_DCa & ") AS cansto, round(tin_propon, 2) AS propon, " & _
                    "'E' AS tipmov, 0 AS numdoc, 'E' AS tipdoc, 'E' AS rutcli, '000' AS orden INTO " & aAp & " FROM b_tomainv " & _
                    "WHERE tin_fectom = " & FecInv & " " & _
                    "AND   tin_propon > 0 " & _
                    "AND   tin_codbod = " & vg_codbod & ""
      DoEvents
   Else
      '-------> Fin traer Inventario primer inventario
      vg_db.Execute "SELECT " & Format(vg_ciedia, "yyyymmdd") & " AS fecpro, ppd_codpro AS codpro, round(ppd_saldo, " & vg_DCa & ") AS cansto, round(ppd_propon, 2) AS propon, " & _
                    "'E' AS tipmov, 0 AS numdoc, 'E' AS tipdoc, 'E' AS rutcli, '000' AS orden " & _
                    "INTO " & aAp & " FROM b_productospmpdia " & _
                    "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                    "AND   ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " "
      DoEvents
   End If
   RS1.Close: Set RS1 = Nothing
   '-------> Traer productos pmp dìa si no existe toma de inventario

   '-------> Traer salida y devolución produción
   vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecpro, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
                 "dev.dev_canmer AS cansto, round(dev.dev_precos, 2) AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, " & _
                 "tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, IIF(tov.tov_tipdoc='SP','40','50') AS orden " & _
                 "FROM b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND   tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND   tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND   dev.dev_codmer = pro.pro_codigo " & _
                 "AND   pro.pro_ctrsto = 1 " & _
                 "AND  (tov.tov_tipdoc = 'SP' or tov.tov_tipdoc = 'DP') " & _
                 "AND   tov.tov_estdoc <> 'A' " & _
                 "AND   tov.tov_codbod = " & vg_codbod & " " & _
                 "AND   tov.tov_fecpro = CDATE('" & vg_ciedia & "') "
   DoEvents
   '-------> Fin traer salida y devolución produción
   
   '-------> Traer mermas
   vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
                 "dev.dev_canmer AS cansto, round(dev.dev_precos, 2) AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, " & _
                 "tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '70' AS orden " & _
                 "FROM  b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND   tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND   tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND   dev.dev_codmer = pro.pro_codigo " & _
                 "AND   tov.tov_tipdoc = 'ME' and tov.tov_estdoc <> 'A' " & _
                 "AND   tov.tov_codbod = " & vg_codbod & " " & _
                 "AND  tov.tov_fecemi = CDATE('" & vg_ciedia & "')"
   DoEvents
   '-------> Fin traer mermas
       
   '-------> Traer documento traspaso entrada
   vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
                 "dev.dev_canmer AS cansto, round(dev.dev_precos, 2) AS propon, 'E' AS tipmov, " & _
                 "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '30' AS orden " & _
                 "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE  tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND    tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND    tov.tov_codbod = " & vg_codbod & " " & _
                 "AND    tov.tov_codreg = 0 " & _
                 "AND    dev.dev_codmer = pro.pro_codigo " & _
                 "AND    tov.tov_codser <> 0 " & _
                 "AND    pro.pro_ctrsto = 1 " & _
                 "AND    tov.tov_tipdoc = 'TR' AND tov.tov_estdoc<>'A' " & _
                 "AND    tov.tov_fecemi = CDATE('" & vg_ciedia & "') AND dev.dev_canmer > 0 ORDER BY tov.tov_fecemi, pro.pro_codigo"
   DoEvents
   '-------> Fin traer documento traspaso entrada
     
   '-------> Traer documento traspaso salida
   vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
                 "dev.dev_canmer AS cansto, round(dev.dev_precos, 2) AS propon, 'S' AS tipmov, " & _
                 "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '60' AS orden " & _
                 "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE  tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND    tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND    tov.tov_codbod = " & vg_codbod & " " & _
                 "AND    tov.tov_codreg = 0 " & _
                 "AND    dev.dev_codmer = pro.pro_codigo " & _
                 "AND    tov.tov_codser = 0 " & _
                 "AND    pro.pro_ctrsto = 1 " & _
                 "AND    tov.tov_tipdoc = 'TR' AND tov.tov_estdoc<>'A' " & _
                 "AND    tov.tov_fecemi = CDATE('" & vg_ciedia & "') AND dev.dev_canmer > 0 ORDER BY tov.tov_fecemi, pro.pro_codigo"
   DoEvents
   '-------> Fin traer documento traspaso salida
     
   
   '-------> Traer documento ventas cafeteria
   vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(b.tvc_fecing, 'yyyymmdd') AS fecpro, c.pro_codigo AS codpro, " & _
                 "a.dvp_candig AS cansto, round(a.dvp_precos, 2) AS propon, 'S' AS tipmov, " & _
                 "0 as numdoc, '' AS tipdoc, b.tvc_cencos as rutcli, '100' AS orden FROM b_detventascafpro a, b_totventascaf b, b_productos c " & _
                 "WHERE b.tvc_cencos = a.dvp_cencos " & _
                 "AND   b.tvc_fecing = a.dvp_fecing " & _
                 "AND   a.dvp_codmer = c.pro_codigo " & _
                 "AND   c.pro_ctrsto = 1 " & _
                 "AND   a.dvp_precos <> 0 " & _
                 "AND   b.tvc_fecing = CDATE('" & vg_ciedia & "') " & _
                 "AND   b.tvc_codbod = " & vg_codbod & ""
   DoEvents
   '-------> Fin traer documento ventas cafeteria
   
   '-------> Traer documento ajuste de inventario de entrada
   vg_db.Execute "INSERT INTO " & aAp & " " & _
                 "SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
                 "round(dev.dev_canmer, " & vg_DCa & ") AS cansto, round(dev.dev_precos, 2) AS propon, 'E' AS tipmov, " & _
                 "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '120' AS orden " & _
                 "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                 "Where tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND    tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND    tov.tov_codbod = " & vg_codbod & " " & _
                 "AND    tov.tov_codreg = 1 " & _
                 "AND    dev.dev_codmer = pro.pro_codigo " & _
                 "AND    tov.tov_codser <> 0 " & _
                 "AND    pro.pro_ctrsto = 1 " & _
                 "AND    tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' " & _
                 "AND    tov.tov_fecemi = CDATE('" & vg_ciedia & "') " & _
                 "AND    dev.dev_canmer > 0 " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_codigo"
   DoEvents
   '-------> Fin traer documento traspaso entrada

   '-------> Traer documento ajuste de inventario de entrada
   vg_db.Execute "INSERT INTO " & aAp & " " & _
                 "SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd')  AS fecpro, pro.pro_codigo AS codpro, " & _
                 "round(dev.dev_canmer, " & vg_DCa & ") AS cansto, round(dev.dev_precos, 2) AS propon, 'S' AS tipmov, " & _
                 "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '130' AS orden " & _
                 "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                 "Where tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND    tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND    tov.tov_codbod = " & vg_codbod & " " & _
                 "AND    tov.tov_codreg = 0 " & _
                 "AND    dev.dev_codmer = pro.pro_codigo " & _
                 "AND    tov.tov_codser <> 0 " & _
                 "AND    pro.pro_ctrsto = 1 " & _
                 "AND    tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' " & _
                 "AND    tov.tov_fecemi = CDATE('" & vg_ciedia & "') " & _
                 "AND    dev.dev_canmer > 0 " & _
                 "ORDER BY tov.tov_fecemi, pro.pro_codigo"
   DoEvents
   '-------> Fin traer documento traspaso entrada
   
   '-------> Traer documento guia ventas
   vg_db.Execute "INSERT INTO " & aAp & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
                 "dev.dev_canmer AS cansto, round(dev.dev_precos, 2) AS propon, 'S' AS tipmov, " & _
                 "tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli, '80' AS orden " & _
                 "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                 "WHERE  tov.tov_rutcli = dev.dev_rutcli " & _
                 "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
                 "AND    tov.tov_numdoc = dev.dev_numdoc " & _
                 "AND    tov.tov_codbod = " & vg_codbod & " " & _
                 "AND    dev.dev_codmer = pro.pro_codigo " & _
                 "AND    pro.pro_ctrsto = 1 " & _
                 "AND   (tov.tov_tipdoc = 'FA'  OR tov.tov_tipdoc='FE' or tov.tov_tipdoc='GD') AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' AND dev.dev_mueinv='S' " & _
                 "AND    tov.tov_fecemi = CDATE('" & vg_ciedia & "') AND dev.dev_canmer > 0 ORDER BY tov.tov_fecemi, pro.pro_codigo"
   DoEvents
   '-------> FIN traer documento guia ventas
   
   '-------> Traer Documento Proveedor
'            "AND   a.toc_fecemi = CDATE('" & vg_ciedia & "') " & _

   Dim pctimp As Double, pctdes  As Double, Precio As Double
   RS1.Open "SELECT a.toc_fecrem, c.pro_codigo, b.dec_numdoc, " & _
            "b.dec_codmer, b.dec_canmer, b.dec_precom, b.dec_pctdes, b.dec_numlin, " & _
            "b.dec_canrec, b.dec_prerec, b.dec_prefle, b.dec_rutpro, b.dec_tipdoc " & _
            "FROM b_totcompras a, b_detcompras b, b_productos c " & _
            "WHERE a.toc_rutpro = b.dec_rutpro " & _
            "AND   a.toc_tipdoc = b.dec_tipdoc " & _
            "AND   a.toc_numdoc = b.dec_numdoc " & _
            "AND   b.dec_codmer = c.pro_codigo " & _
            "AND   b.dec_mueinv = 'S' and a.toc_tipdoc <> 'SN' " & _
            "AND   b.dec_canrec > 0 " & _
            "AND   a.toc_codbod = " & vg_codbod & " " & _
            "AND   a.toc_fecrem = CDATE('" & vg_ciedia & "') " & _
            "ORDER BY a.toc_fecdig, c.pro_nombre", vg_db, adOpenStatic
   If Not RS1.EOF Then
      Do While Not RS1.EOF
         DoEvents
         pctimp = 0: Precio = 0
         RS2.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
                  "FROM  b_detcomprasimp a, a_impuesto b " & _
                  "WHERE a.imd_rutdoc = '" & RS1!dec_rutpro & "' " & _
                  "AND   a.imd_tipdoc = '" & RS1!dec_tipdoc & "' " & _
                  "AND   a.imd_numdoc = " & RS1!dec_numdoc & " " & _
                  "AND   a.imd_numlin = " & RS1!dec_numlin & " " & _
                  "AND   a.imd_codpro = '" & RS1!pro_codigo & "' " & _
                  "AND   a.imd_codimp = b.imp_codigo " & _
                  "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
         If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
         pctdes = 0
         If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
         If RS1!dec_prefle > 0 Then
            Precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec)) + (RS1!dec_prefle / RS1!dec_canrec)
         Else
            Precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec))
         End If
         vg_db.Execute "INSERT INTO " & aAp & " VALUES (" & Val(Format(RS1!toc_fecrem, "yyyymmdd")) & ", " & _
                       "'" & Trim(RS1!pro_codigo) & "', " & RS1!dec_canrec & ", " & Round(Precio, 2) & ", '" & "E+" & "', " & RS1!dec_numdoc & ", '" & Trim(RS1!dec_tipdoc) & "', '" & Trim(RS1!dec_rutpro) & "', '20')"
         RS1.MoveNext
      Loop
   End If
   RS1.Close: Set RS1 = Nothing
   '-------> Fin traer Documento Proveedor
   '-------> Procesar Información Precio Promedio Ponderado
   Dim auxCanmer As Double, auxPropon As Double, propon As Double, auxfec As Long, fecuco As String, upreco As Double
   Dim auxcodpro As String, auxtipdoc As String, nreg As Long
   '-------> Traer numero registo
   i = 1
   RS1.Open "SELECT COUNT(*) AS nreg FROM " & aAp & "", vg_db, adOpenForwardOnly
   If Not RS1.EOF Then nreg = RS1!nreg
   RS1.Close: Set RS1 = Nothing
   '-------> Traer datos, para actualizar documentos de salida-mermas-traspaso salida-cafeteria
   RS1.Open "SELECT * FROM " & aAp & " ORDER BY codpro, fecpro, orden, tipmov", vg_db, adOpenForwardOnly
   If op Then Formu.Bar1(0).Visible = True: Formu.Bar1(0).Value = 0: Formu.Bar1(0).max = 10
   If op Then Formu.Label1(1).Visible = True
   If Not RS1.EOF Then
      auxCanmer = 0: auxPropon = 0: propon = 0: auxcodpro = "": auxtipdoc = "": fecuco = "": upreco = 0
      Do While Not RS1.EOF
         DoEvents
         If RS1!codpro <> auxcodpro Then
            If Trim(auxcodpro) <> "" Then vg_db.Execute "INSERT INTO " & aAp1 & " VALUES ('" & MuestraCasino(1) & "', " & Format(CDate(vg_ciedia), "yyyymmdd") & ", '" & auxcodpro & "', " & Round(propon, 2) & ", " & auxCanmer & ")"
            auxcodpro = RS1!codpro:  auxCanmer = 0: auxPropon = 0: propon = 0: upreco = 0: fecuco = ""
         End If
         If RS1!tipmov = "S" Then
            If propon > 0 Then If RS1!Orden = "50" Then auxCanmer = (auxCanmer + RS1!cansto) Else auxCanmer = (auxCanmer - RS1!cansto)
         Else
            If RS1!Orden = "120" Then
               propon = RS1!propon
               auxPropon = propon
            Else
               propon = Round(((auxPropon * IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer)) + (RS1!propon * IIf(RS1!Orden = "000" And RS1!cansto <= 0, 1, RS1!cansto))) / (IIf(auxCanmer < 0, (auxCanmer * -1), auxCanmer) + IIf(RS1!Orden = "000" And RS1!cansto <= 0, 1, RS1!cansto)), 2)
               auxPropon = propon
            End If
            auxCanmer = auxCanmer + RS1!cansto
         End If
         RS1.MoveNext: i = i + 1
      Loop
      vg_db.Execute "INSERT INTO " & aAp1 & " VALUES ('" & MuestraCasino(1) & "', " & Format(CDate(vg_ciedia), "yyyymmdd") & ", '" & auxcodpro & "', " & Round(propon, 2) & ", " & auxCanmer & ")"
   End If
   RS1.Close: Set RS1 = Nothing
End If
If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Productos pmp x día
If op Then Formu.Label1(1).Caption = "Actua. PMP Día"
'-------> Respaldar productos pmp despues de la fecha cierre
vg_db.Execute "SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo, ppd_upreco, ppd_fecuco " & _
              "INTO " & vg_NUsr & "_Respprodpmppostfeccie " & _
              "FROM b_productospmpdia " & _
              "WHERE ppd_cencos='" & MuestraCasino(1) & "' AND ppd_fecdia  > " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND ppd_propon <> 0 "
'-------> Respaldar información, para rescatar ultimo precio compra
vg_db.Execute "SELECT DISTINCT ppd_cencos, ppd_codpro, " & Format(CDate(vg_ciedia), "yyyymmdd") & " AS ppd_fecdia, ppd_propon, ppd_saldo, ppd_upreco, ppd_fecuco INTO " & vg_NUsr & "_Respaldoproductospmp FROM b_productospmpdia WHERE ppd_cencos='" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ""
vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia " & _
              "WHERE ppd_cencos='" & MuestraCasino(1) & "' AND ppd_fecdia>=" & Format(CDate(vg_ciedia), "yyyymmdd") & " AND ppd_codpro IN (SELECT DISTINCT pro_codigo FROM b_productos)"
fecpro = CDate(vg_ciedia)
fecter = dEoM(CDate(vg_ciedia))
Do While fecpro <= fecter
   DoEvents
   vg_db.Execute "INSERT INTO b_productospmpdia(ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo) " & _
                 "SELECT DISTINCT '" & MuestraCasino(1) & "', a.pro_codigo, " & Format(fecpro, "yyyymmdd") & ", 0, 0 " & _
                 "FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1)"
   fecpro = fecpro + 1
Loop
'-------> Actualizar precio que la fecha sea mayor
vg_db.Execute "UPDATE b_productospmpdia INNER JOIN " & aAp1 & " ON (b_productospmpdia.ppd_codpro = " & aAp1 & ".pmp_codpro) AND (b_productospmpdia.ppd_cencos = " & aAp1 & ".pmp_cencos) AND b_productospmpdia.ppd_fecdia = " & aAp1 & ".pmp_fecha SET b_productospmpdia.ppd_propon = round(" & aAp1 & ".pmp_precio, 2), b_productospmpdia.ppd_saldo = " & aAp1 & ".pmp_saldo "
'-------> Actualizar uprecom- ufeccom
vg_db.Execute "UPDATE b_productospmpdia INNER JOIN " & vg_NUsr & "_Respaldoproductospmp ON (b_productospmpdia.ppd_fecdia = " & vg_NUsr & "_Respaldoproductospmp.ppd_fecdia) AND (b_productospmpdia.ppd_codpro = " & vg_NUsr & "_Respaldoproductospmp.ppd_codpro) AND (b_productospmpdia.ppd_cencos = " & vg_NUsr & "_Respaldoproductospmp.ppd_cencos) SET b_productospmpdia.ppd_upreco = " & vg_NUsr & "_Respaldoproductospmp.ppd_upreco, b_productospmpdia.ppd_fecuco = " & vg_NUsr & "_Respaldoproductospmp.ppd_fecuco"
'-------> Actualizar precio despues a la fecha cierre
vg_db.Execute "UPDATE b_productospmpdia INNER JOIN " & vg_NUsr & "_Respprodpmppostfeccie ON (b_productospmpdia.ppd_fecdia = " & vg_NUsr & "_Respprodpmppostfeccie.ppd_fecdia) AND (b_productospmpdia.ppd_codpro = " & vg_NUsr & "_Respprodpmppostfeccie.ppd_codpro) AND (b_productospmpdia.ppd_cencos = " & vg_NUsr & "_Respprodpmppostfeccie.ppd_cencos) SET b_productospmpdia.ppd_propon = " & vg_NUsr & "_Respprodpmppostfeccie.ppd_propon, b_productospmpdia.ppd_upreco = " & vg_NUsr & "_Respprodpmppostfeccie.ppd_upreco, b_productospmpdia.ppd_fecuco = " & vg_NUsr & "_Respprodpmppostfeccie.ppd_fecuco"
'-------> Borrar tabla temporal
vg_db.Execute "DROP TABLE " & vg_NUsr & "_Respaldoproductospmp"
vg_db.Execute "DROP TABLE " & vg_NUsr & "_Respprodpmppostfeccie"
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1
If Not progrl Then ' Corta proceso ya que se necesita recalcular promedio ponderado
   '-------> Borrar tablas b_productospmpdia distinto al ultimo día cierre que tengan precio cero y saldo en cero
   vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <> " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"

   '-------> Borrar tablas temporales
   If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
   If Trim(aAp1) <> "" Then vg_db.Execute "DROP TABLE " & aAp1 & ""
   If Trim(aAp2) <> "" Then vg_db.Execute "DROP TABLE " & aAp2 & ""
   If Trim(aAp3) <> "" Then vg_db.Execute "DROP TABLE " & aAp3 & ""
   If Trim(aAp4) <> "" Then vg_db.Execute "DROP TABLE " & aAp4 & ""
   If Trim(aAp5) <> "" Then vg_db.Execute "DROP TABLE " & aAp5 & ""
   If Trim(aAp6) <> "" Then vg_db.Execute "DROP TABLE " & aAp6 & ""
   Exit Function
End If

''-------> Actualizar productospmpproductos que esten en cero y fueron ingresado despues fecha cierre
'   '-------> Creo tabla temporal producto PMP
'   aAp5 = Trim(vg_NUsr) & "_tmp_Prodfechamayorcierre"
'   fg_CheckTmp aAp5
'   vg_db.Execute "SELECT 0 AS fecpro, '' AS codpro, Round(0, 2) AS propon INTO " & aAp5 & ""
'
'   '-------> Traer Documento Proveedor
'   RS1.Open "SELECT a.toc_fecrem, c.pro_codigo, b.dec_numdoc, " & _
'            "b.dec_codmer, b.dec_canmer, b.dec_precom, b.dec_pctdes, b.dec_numlin, " & _
'            "b.dec_canrec, b.dec_prerec, b.dec_prefle, b.dec_rutpro, b.dec_tipdoc " & _
'            "FROM b_totcompras a, b_detcompras b, b_productos c, b_productospmpdia d " & _
'            "WHERE a.toc_rutpro = b.dec_rutpro " & _
'            "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'            "AND   a.toc_numdoc = b.dec_numdoc " & _
'            "AND   b.dec_codmer = c.pro_codigo " & _
'            "AND   b.dec_mueinv = 'S' and a.toc_tipdoc <> 'SN' " & _
'            "AND   b.dec_canrec > 0 " & _
'            "AND   a.toc_codbod = " & vg_codbod & " " & _
'            "AND   a.toc_fecemi > CDATE('" & vg_ciedia & "') " & _
'            "AND   c.pro_codigo = d.ppd_codpro " & _
'            "AND   d.ppd_cencos = '" & MuestraCasino(1) & "' " & _
'            "AND   d.ppd_fecdia >= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'            "AND   d.ppd_propon < 1 " & _
'            "ORDER BY a.toc_fecdig, c.pro_nombre", vg_db, adOpenStatic
'   If Not RS1.EOF Then
'      Do While Not RS1.EOF
'         DoEvents
'         pctimp = 0: precio = 0
'         RS2.Open "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
'                  "FROM  b_detcomprasimp a, a_impuesto b " & _
'                  "WHERE a.imd_rutdoc = '" & RS1!dec_rutpro & "' " & _
'                  "AND   a.imd_tipdoc = '" & RS1!dec_tipdoc & "' " & _
'                  "AND   a.imd_numdoc = " & RS1!dec_numdoc & " " & _
'                  "AND   a.imd_numlin = " & RS1!dec_numlin & " " & _
'                  "AND   a.imd_codpro = '" & RS1!pro_codigo & "' " & _
'                  "AND   a.imd_codimp = b.imp_codigo " & _
'                  "AND   b.imp_inccos = 1 GROUP BY b.imp_inccos", vg_db, adOpenStatic
'         If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
'         pctdes = 0
'         If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
'         If RS1!dec_prefle > 0 Then
'            precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec)) + (RS1!dec_prefle / RS1!dec_canrec)
'         Else
'            precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec))
'         End If
'         vg_db.Execute "INSERT INTO " & aAp5 & " VALUES (" & Val(Format(RS1!toc_fecrem, "yyyymmdd")) & ", " & _
'                       "'" & Trim(RS1!pro_codigo) & "', " & round(precio, 2) & ")"
'         RS1.MoveNext
'      Loop
'   End If
'   RS1.Close: Set RS1 = Nothing
'   '-------> Fin traer Documento Proveedor
'
'   '-------> Traer documento traspaso entrada
'   vg_db.Execute "INSERT INTO " & aAp5 & " SELECT FORMAT(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, " & _
'                 "Round(dev.dev_precos, 2) AS propon " & _
'                 "FROM   b_totventas tov, b_detventas dev, b_productos pro, b_productospmpdia d " & _
'                 "WHERE  tov.tov_rutcli = dev.dev_rutcli " & _
'                 "AND    tov.tov_tipdoc = dev.dev_tipdoc " & _
'                 "AND    tov.tov_numdoc = dev.dev_numdoc " & _
'                 "AND    tov.tov_codbod = " & vg_codbod & " " & _
'                 "AND    tov.tov_codreg = 0 " & _
'                 "AND    dev.dev_codmer = pro.pro_codigo " & _
'                 "AND    tov.tov_codser <> 0 " & _
'                 "AND    pro.pro_ctrsto = 1 " & _
'                 "AND    tov.tov_tipdoc = 'TR' AND tov.tov_estdoc<>'A' " & _
'                 "AND    tov.tov_fecemi > CDATE('" & vg_ciedia & "') AND dev.dev_canmer > 0 " & _
'                 "AND    pro.pro_codigo = d.ppd_codpro " & _
'                 "AND    d.ppd_cencos = '" & MuestraCasino(1) & "' " & _
'                 "AND    d.ppd_fecdia >= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
'                 "AND    d.ppd_propon < 1 ORDER BY tov.tov_fecemi, pro.pro_codigo"
'   '-------> Fin Traer documento traspaso entrada
'   RS1.Open "SELECT DISTINCT * FROM " & aAp5 & " WHERE propon > 0 ORDER BY codpro, fecpro", vg_db, adOpenForwardOnly
'   If Not RS1.EOF Then
'      Do While Not RS1.EOF
'         DoEvents
''         vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & RS1!propon & ", ppd_upreco = " & RS1!propon & ", ppd_fecuco = cdate('" & fg_Ctod1(RS1!fecpro) & "') WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(fg_Ctod1(RS1!fecpro)) - 1, "yyyymmdd") & " AND ppd_codpro = '" & RS1!CodPro & "' AND ppd_propon < 1"
'         vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & RS1!propon & ", ppd_upreco = " & RS1!propon & ", ppd_fecuco = cdate('" & fg_Ctod1(RS1!fecpro) & "') WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & RS1!fecpro & " AND ppd_codpro = '" & RS1!codpro & "' AND ppd_propon < 1"
'         RS1.MoveNext
'      Loop
'   End If
'   RS1.Close: Set RS1 = Nothing
''-------> Fin actualizar productospmpproductos que esten en cero y fueron ingresado

'-------> Actualizar encabezado y detalle ventas
If op Then Formu.Label1(1).Caption = "Actua. Salida y Devolución Producción"
vg_db.Execute "UPDATE b_productos INNER JOIN (b_totventas INNER JOIN (b_detventas INNER JOIN b_productospmpdia ON b_detventas.dev_codmer = b_productospmpdia.ppd_codpro) ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli) AND (b_totventas.tov_rutcli = b_productospmpdia.ppd_cencos)) ON (b_productos.pro_codigo = b_productospmpdia.ppd_codpro) AND (b_productos.pro_codigo = b_detventas.dev_codmer) SET b_detventas.dev_precos = Round(b_productospmpdia.ppd_propon, 2), b_detventas.dev_predoc = Round(b_productospmpdia.ppd_propon, 2), b_detventas.dev_ptotal = Round(b_detventas.dev_canmer*b_productospmpdia.ppd_propon, 2) " & _
              "WHERE (((b_totventas.tov_estdoc)<>'A') AND ((b_totventas.tov_codbod)=" & vg_codbod & ") AND ((b_totventas.tov_fecpro)>=CDate('" & vg_ciedia & "')) AND ((b_totventas.tov_tipdoc) In ('SP','DP')) AND ((b_productospmpdia.ppd_fecdia)=" & Format(vg_ciedia, "yyyymmdd") & ") AND ((b_productos.pro_ctrsto)=1))"
vg_db.Execute "SELECT b.dev_rutcli, b.dev_tipdoc, b.dev_numdoc, SUM(b.dev_ptotal) AS ptotal INTO " & vg_NUsr & "_tmpDoc FROM b_totventas a, b_detventas b, b_productos c WHERE a.tov_numdoc = b.dev_numdoc AND a.tov_tipdoc = b.dev_tipdoc AND a.tov_rutcli = b.dev_rutcli AND b.dev_codmer = c.pro_codigo AND c.pro_ctrsto = 1 AND a.tov_codbod = " & vg_codbod & " AND a.tov_fecpro >= CDate('" & vg_ciedia & "') AND a.tov_estdoc <> 'A' AND a.tov_tipdoc In ('SP','DP') GROUP BY b.dev_rutcli, b.dev_tipdoc, b.dev_numdoc"
vg_db.Execute "UPDATE b_totventas INNER JOIN " & vg_NUsr & "_tmpDoc ON b_totventas.tov_numdoc = " & vg_NUsr & "_tmpDoc.dev_numdoc AND b_totventas.tov_tipdoc = " & vg_NUsr & "_tmpDoc.dev_tipdoc AND b_totventas.tov_rutcli = " & vg_NUsr & "_tmpDoc.dev_rutcli SET b_totventas.tov_totdoc = " & vg_NUsr & "_tmpDoc.ptotal " & _
              "WHERE b_totventas.tov_codbod = " & vg_codbod & " AND b_totventas.tov_fecpro >= CDate('" & vg_ciedia & "') AND b_totventas.tov_estdoc <> 'A' AND b_totventas.tov_tipdoc In ('SP','DP') "
vg_db.Execute "DROP TABLE " & vg_NUsr & "_tmpDoc"
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

''-------> Actualizar encabezado y detalle traspaso de salida
'If op Then Formu.Label1(1).Caption = "Actua. Traspaso Salida"
'vg_db.Execute "UPDATE b_productos INNER JOIN (b_totventas INNER JOIN (b_detventas INNER JOIN b_productospmpdia ON b_detventas.dev_codmer = b_productospmpdia.ppd_codpro) ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli) AND (b_totventas.tov_rutcli = b_productospmpdia.ppd_cencos)) ON (b_productos.pro_codigo = b_productospmpdia.ppd_codpro) AND (b_productos.pro_codigo = b_detventas.dev_codmer) SET b_detventas.dev_precos = Round(b_productospmpdia.ppd_propon, 2), b_detventas.dev_predoc = Round(b_productospmpdia.ppd_propon, 2), b_detventas.dev_ptotal = Round(b_detventas.dev_canmer*b_productospmpdia.ppd_propon, 2) " & _
'              "WHERE b_totventas.tov_estdoc<>'A' AND b_totventas.tov_codbod = " & vg_codbod & " AND b_totventas.tov_fecemi >= CDate('" & vg_ciedia & "') AND b_totventas.tov_tipdoc = 'TR' AND b_totventas.tov_codser = 0 AND b_productospmpdia.ppd_fecdia = " & Format(vg_ciedia, "yyyymmdd") & " AND b_productos.pro_ctrsto=1"
'vg_db.Execute "SELECT b.dev_rutcli, b.dev_tipdoc, b.dev_numdoc, SUM(b.dev_ptotal) AS ptotal INTO " & vg_NUsr & "_tmpDoc FROM b_totventas a, b_detventas b, b_productos c WHERE a.tov_numdoc = b.dev_numdoc AND a.tov_tipdoc = b.dev_tipdoc AND a.tov_rutcli = b.dev_rutcli AND b.dev_codmer = c.pro_codigo AND c.pro_ctrsto = 1 AND a.tov_codbod = " & vg_codbod & " AND a.tov_fecemi >= CDate('" & vg_ciedia & "') AND a.tov_estdoc <> 'A' AND a.tov_tipdoc = 'TR' AND a.tov_codser = 0 GROUP BY b.dev_rutcli, b.dev_tipdoc, b.dev_numdoc"
'vg_db.Execute "UPDATE b_totventas INNER JOIN " & vg_NUsr & "_tmpDoc ON b_totventas.tov_numdoc = " & vg_NUsr & "_tmpDoc.dev_numdoc AND b_totventas.tov_tipdoc = " & vg_NUsr & "_tmpDoc.dev_tipdoc AND b_totventas.tov_rutcli = " & vg_NUsr & "_tmpDoc.dev_rutcli SET b_totventas.tov_totdoc = " & vg_NUsr & "_tmpDoc.ptotal " & _
'              "WHERE b_totventas.tov_codbod=" & vg_codbod & " AND b_totventas.tov_fecemi >= CDate('" & vg_ciedia & "') AND b_totventas.tov_estdoc <> 'A' AND b_totventas.tov_tipdoc = 'TR' AND b_totventas.tov_codser = 0"
'vg_db.Execute "DROP TABLE " & vg_NUsr & "_tmpDoc"
'DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Actualizar Mermas
If op Then Formu.Label1(1).Caption = "Actua. Mermas"
vg_db.Execute "UPDATE b_productos INNER JOIN (b_totventas INNER JOIN (b_detventas INNER JOIN b_productospmpdia ON b_detventas.dev_codmer = b_productospmpdia.ppd_codpro) ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli) AND (b_totventas.tov_rutcli = b_productospmpdia.ppd_cencos)) ON (b_productos.pro_codigo = b_productospmpdia.ppd_codpro) AND (b_productos.pro_codigo = b_detventas.dev_codmer) SET b_detventas.dev_precos = Round(b_productospmpdia.ppd_propon, 2), b_detventas.dev_predoc = Round(b_productospmpdia.ppd_propon, 2), b_detventas.dev_ptotal = Round(b_detventas.dev_canmer*b_productospmpdia.ppd_propon, 2) " & _
              "WHERE b_totventas.tov_estdoc<>'A' AND b_totventas.tov_codbod = " & vg_codbod & " AND b_totventas.tov_fecemi >= CDate('" & vg_ciedia & "') AND b_totventas.tov_tipdoc = 'ME' AND b_productospmpdia.ppd_fecdia = " & Format(vg_ciedia, "yyyymmdd") & " AND b_productos.pro_ctrsto = 1"
vg_db.Execute "SELECT b.dev_rutcli, b.dev_tipdoc, b.dev_numdoc, SUM(b.dev_ptotal) AS ptotal INTO " & vg_NUsr & "_tmpDoc FROM b_totventas a, b_detventas b, b_productos c WHERE a.tov_numdoc = b.dev_numdoc AND a.tov_tipdoc = b.dev_tipdoc AND a.tov_rutcli = b.dev_rutcli AND b.dev_codmer = c.pro_codigo AND c.pro_ctrsto = 1 AND a.tov_codbod = " & vg_codbod & " AND a.tov_fecemi >= CDate('" & vg_ciedia & "') AND a.tov_estdoc <> 'A' AND a.tov_tipdoc = 'ME' GROUP BY b.dev_rutcli, b.dev_tipdoc, b.dev_numdoc"
vg_db.Execute "UPDATE b_totventas INNER JOIN " & vg_NUsr & "_tmpDoc ON b_totventas.tov_numdoc = " & vg_NUsr & "_tmpDoc.dev_numdoc AND b_totventas.tov_tipdoc = " & vg_NUsr & "_tmpDoc.dev_tipdoc AND b_totventas.tov_rutcli = " & vg_NUsr & "_tmpDoc.dev_rutcli SET b_totventas.tov_totdoc = " & vg_NUsr & "_tmpDoc.ptotal " & _
              "WHERE b_totventas.tov_codbod=" & vg_codbod & " AND b_totventas.tov_fecemi >= CDate('" & vg_ciedia & "') AND b_totventas.tov_estdoc <> 'A' AND b_totventas.tov_tipdoc = 'ME'"
vg_db.Execute "DROP TABLE " & vg_NUsr & "_tmpDoc"
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Actualizar cafeteria
If op Then Formu.Label1(1).Caption = "Actua. Cafeteria"
vg_db.Execute "UPDATE (b_totventascaf INNER JOIN b_detventascafpro ON (b_totventascaf.tvc_fecing = b_detventascafpro.dvp_fecing) AND (b_totventascaf.tvc_cencos = b_detventascafpro.dvp_cencos)) INNER JOIN b_productospmpdia ON (b_totventascaf.tvc_cencos = b_productospmpdia.ppd_cencos) AND (b_detventascafpro.dvp_codmer = b_productospmpdia.ppd_codpro) " & _
              "SET b_detventascafpro.dvp_precos = round(b_productospmpdia.ppd_propon, 2) " & _
              "WHERE b_totventascaf.tvc_codbod = " & vg_codbod & " AND b_totventascaf.tvc_fecing >= cdate('" & vg_ciedia & "') AND b_totventascaf.tvc_cencos = '" & MuestraCasino(1) & "' AND b_productospmpdia.ppd_fecdia=" & Format(vg_ciedia, "yyyymmdd") & ""
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

If op Then Formu.Label1(1).Caption = "Actua. Minuta Costo"
'-------> Actualizar minuta cambios día
vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto " & _
              "WHERE mic_cencos='" & MuestraCasino(1) & "' " & _
              "AND   mic_fecval>=" & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "AND   mic_fecval<=" & Format(CDate(fecter), "yyyymmdd") & " " & _
              "AND   mic_tipmin='2'"
DoEvents
  
'-------> tabla temporal costo ingredienteI
aAp6 = Trim(vg_NUsr) & "_tmp_CostoIngredienteI"
fg_CheckTmp aAp6

vg_db.Execute "SELECT DISTINCT e.ing_codigo, Round(AVG(g.ppd_propon/i.pro_facing), 2) AS ing_precos " & _
              "INTO " & aAp6 & " " & _
              "FROM  b_productospmpdia g, b_productosing h, b_productos i, b_ingrediente e, b_contlistpreing f " & _
              "WHERE f.cpi_coding = e.ing_codigo " & _
              "AND   f.cpi_cencos = '" & MuestraCasino(1) & "' " & _
              "and   e.ing_codigo = h.pri_coding " & _
              "AND   h.pri_codpro = i.pro_codigo " & _
              "AND   i.pro_codigo = g.ppd_codpro " & _
              "AND   g.ppd_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   g.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "AND   g.ppd_propon > 0 AND g.ppd_saldo <> 0 AND i.pro_ctrsto = 1 GROUP BY e.ing_codigo HAVING COUNT(e.ing_codigo) > 0"

vg_db.Execute "INSERT INTO " & aAp6 & " SELECT DISTINCT e.ing_codigo, Round(AVG(g.ppd_propon/i.pro_facing), 2) AS ing_precos " & _
              "FROM  b_productospmpdia g, b_productosing h, b_productos i, b_ingrediente e, b_contlistpreing f " & _
              "WHERE f.cpi_coding = e.ing_codigo " & _
              "AND   f.cpi_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   e.ing_codigo = h.pri_coding " & _
              "AND   e.ing_codigo NOT IN (SELECT DISTINCT ing_codigo FROM " & aAp6 & ") " & _
              "AND   h.pri_codpro = i.pro_codigo " & _
              "AND   i.pro_codigo = g.ppd_codpro " & _
              "AND   g.ppd_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   g.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "AND   g.ppd_propon > 0 AND g.ppd_saldo = 0 AND i.pro_ctrsto = 1 GROUP BY e.ing_codigo HAVING COUNT(e.ing_codigo) = 1"

vg_db.Execute "INSERT INTO " & aAp6 & " SELECT DISTINCT e.ing_codigo, Round(AVG(g.ppd_propon/i.pro_facing), 2) AS ing_precos " & _
              "FROM  b_productospmpdia g, b_productosing h, b_productos i, b_ingrediente e, b_contlistpreing f " & _
              "WHERE f.cpi_coding = e.ing_codigo " & _
              "AND   f.cpi_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   e.ing_codigo = h.pri_coding " & _
              "AND   e.ing_codigo NOT IN (SELECT DISTINCT ing_codigo FROM " & aAp6 & ") " & _
              "AND   h.pri_codpro = i.pro_codigo " & _
              "AND   i.pro_codigo = g.ppd_codpro " & _
              "AND   g.ppd_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   g.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "AND   g.ppd_propon > 0 AND g.ppd_saldo = 0 AND i.pro_ctrsto = 1 GROUP BY e.ing_codigo"

vg_db.Execute "UPDATE " & aAp6 & " SET ing_precos = 0 WHERE ing_precos IS NULL"
    
vg_db.Execute "INSERT INTO " & aAp6 & " SELECT DISTINCT cpi_coding AS ing_codigo, 0 AS ing_precos FROM b_contlistpreing  WHERE cpi_coding NOT IN (SELECT DISTINCT ing_codigo FROM " & aAp6 & ") AND cpi_cencos = '" & MuestraCasino(1) & "'"

vg_db.Execute "INSERT INTO b_minutacosto(mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) " & _
              "SELECT DISTINCT c.min_cencos, c.min_fecmin, '2', e.ing_codigo, h.ing_precos " & _
              "FROM  b_receta a, b_recetadet b, b_minuta c, b_minutadet d, b_ingrediente e, b_contlistpreing f, " & aAp6 & " h " & _
              "WHERE c.min_codigo = d.mid_codigo " & _
              "AND   d.mid_codrec = b.red_codigo " & _
              "AND   d.mid_tiprec = b.red_tiprec " & _
              "AND ((b.red_tiprec<>0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) " & _
              "AND   b.red_codigo = a.rec_codigo " & _
              "AND   b.red_codpro = e.ing_codigo " & _
              "AND   c.min_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   c.min_fecmin >= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "AND   c.min_fecmin <= " & Format(CDate(fecter), "yyyymmdd") & " " & _
              "AND   d.mid_tipmin = '2' " & _
              "AND   f.cpi_coding = e.ing_codigo " & _
              "AND   f.cpi_cencos = '" & MuestraCasino(1) & "' AND e.ing_codigo = h.ing_codigo"

DoEvents
vg_db.Execute "UPDATE b_minutacosto SET mic_cospro = 0 " & _
              "WHERE mic_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   mic_fecval >= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "AND   mic_fecval <= " & Format(CDate(fecter), "yyyymmdd") & " " & _
              "AND   mic_tipmin = '2' AND mic_cospro IS NULL"
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Actualizar minutas día
If op Then Formu.Label1(1).Caption = "Actua. Planificación Minutas"
RS1.Open "SELECT MAX(a.min_fecmin) AS min_fecmin FROM b_minuta a, b_minutadet b WHERE a.min_codigo=b.mid_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' AND b.mid_tipmin = '2' AND mid(a.min_fecmin,1,6) = " & Format(vg_ciedia, "yyyymm") & "", vg_db, adOpenForwardOnly
auxfec = Format(vg_ciedia, "yyyymmdd")
If Not RS1.EOF And RS1!min_fecmin > 0 Then
   '-------> Creo tabla temporal y chequeo si existe antes
   aAp2 = Trim(vg_NUsr) & "_tmp_Receta"
   fg_CheckTmp aAp2

   '-------> tabla temporal recetas
   vg_db.Execute "SELECT DISTINCT b.mid_codrec, b.mid_tiprec " & _
                 "INTO " & aAp2 & " " & _
                 "FROM b_recetadet AS a, b_minutadet AS b, b_minuta AS c " & _
                 "WHERE c.min_codigo = b.mid_codigo " & _
                 "AND   a.red_codigo = b.mid_codrec " & _
                 "AND   a.red_tiprec = b.mid_tiprec " & _
                 "AND   c.min_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   c.min_fecmin >= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                 "AND   c.min_fecmin <= " & RS1!min_fecmin & "  " & _
                 "AND   b.mid_tipmin = '2'"
   DoEvents

   '-------> tabla temporal costo ingrediente
   aAp4 = Trim(vg_NUsr) & "_tmp_CostoIngrediente"
   fg_CheckTmp aAp4
   vg_db.Execute "SELECT DISTINCT mic_codpro, mic_cospro " & _
                 "INTO " & aAp4 & " " & _
                 "FROM  b_minutacosto " & _
                 "WHERE mic_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   mic_fecval >= " & auxfec & " " & _
                 "AND   mic_fecval <= " & RS1!min_fecmin & " AND mic_tipmin='2'"
   
   '-------> tabla temporal costo alimentación
   aAp3 = Trim(vg_NUsr) & "_tmp_PreciosIns"
   fg_CheckTmp aAp3

   vg_db.Execute "SELECT DISTINCT f.mid_codrec, f.mid_tiprec, SUM(b.red_canpro*c.mic_cospro) AS cosrec " & _
                 "INTO " & aAp3 & " " & _
                 "FROM b_recetadet b, " & aAp4 & " c, b_contlistpreing d, b_productos e, " & aAp2 & " f " & _
                 "WHERE f.mid_codrec = b.red_codigo " & _
                 "AND   f.mid_tiprec = b.red_tiprec " & _
                 "AND   b.red_codpro = c.mic_codpro " & _
                 "AND   b.red_codpro = d.cpi_coding " & _
                 "AND   d.cpi_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND  (e.pro_codigo = d.cpi_codcom) " & _
                 "AND   e.pro_ctacon = '" & (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")) & "' " & _
                 "AND   b.red_cencos = IIf(f.mid_tiprec = 0, '0', '" & MuestraCasino(1) & "') " & _
                 "GROUP BY f.mid_codrec, f.mid_tiprec"
   
   vg_db.Execute "UPDATE " & aAp3 & " INNER JOIN (b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo) ON (" & aAp3 & ".mid_tiprec = b_minutadet.mid_tiprec) AND (" & aAp3 & ".mid_codrec = b_minutadet.mid_codrec) SET b_minutadet.mid_cosrec = " & aAp3 & ".cosrec, b_minutadet.mid_fecval = b_minuta.min_fecmin " & _
                 "WHERE b_minuta.min_cencos = '" & MuestraCasino(1) & "' AND b_minuta.min_fecmin >= " & auxfec & " AND b_minuta.min_fecmin <= " & RS1!min_fecmin & " AND b_minutadet.mid_tipmin = '2'"

   vg_db.Execute "DROP TABLE " & aAp3 & ""

   vg_db.Execute "SELECT DISTINCT f.mid_codrec, f.mid_tiprec, SUM(b.red_canpro*c.mic_cospro) AS cosrec " & _
                 "INTO " & aAp3 & " " & _
                 "FROM b_recetadet b, " & aAp4 & " c, b_contlistpreing d, b_productos e, " & aAp2 & " f " & _
                 "WHERE f.mid_codrec = b.red_codigo " & _
                 "AND   f.mid_tiprec = b.red_tiprec " & _
                 "AND   b.red_codpro = c.mic_codpro " & _
                 "AND   b.red_codpro = d.cpi_coding " & _
                 "AND   d.cpi_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND  (e.pro_codigo = d.cpi_codcom) " & _
                 "AND   e.pro_ctacon = '" & (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")) & "' " & _
                 "AND   b.red_cencos = IIf(f.mid_tiprec = 0, '0', '" & MuestraCasino(1) & "') " & _
                 "GROUP BY f.mid_codrec, f.mid_tiprec"

   vg_db.Execute "UPDATE " & aAp3 & " INNER JOIN (b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo) ON (" & aAp3 & ".mid_tiprec = b_minutadet.mid_tiprec) AND (" & aAp3 & ".mid_codrec = b_minutadet.mid_codrec) SET b_minutadet.mid_cosdes = " & aAp3 & ".cosrec, b_minutadet.mid_fecval = b_minuta.min_fecmin " & _
                 "WHERE b_minuta.min_cencos = '" & MuestraCasino(1) & "' AND b_minuta.min_fecmin >= " & auxfec & " AND b_minuta.min_fecmin <= " & RS1!min_fecmin & " AND b_minutadet.mid_tipmin = '2'"

   vg_db.Execute "DELETE " & aAp3 & " FROM " & aAp3 & ""
End If
RS1.Close: Set RS1 = Nothing
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Actualizar maestro producto y ingrediente
If op Then Formu.Label1(1).Caption = "Actua. Ingrediente"

vg_db.Execute "UPDATE b_contlistpreing INNER JOIN " & aAp6 & " ON b_contlistpreing.cpi_coding = " & aAp6 & ".ing_codigo " & _
              "SET b_contlistpreing.cpi_precos = " & aAp6 & ".ing_precos, b_contlistpreing.cpi_feccos = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
              "WHERE b_contlistpreing.cpi_cencos = '" & MuestraCasino(1) & "'"
'-------> Fin actualizar maestro producto y ingrediente
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Actualizar estructura fijas
If op Then Formu.Label1(1).Caption = "Actua. Estructura Fija"
vg_db.Execute "UPDATE b_minutafijadia SET mfd_cospro=0 WHERE mfd_tipmin='2' AND mfd_fecha>= " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND mfd_fecha<= " & Format(CDate(fecter), "yyyymmdd") & " AND mfd_cencos = '" & MuestraCasino(1) & "'"
vg_db.Execute "UPDATE b_minutafijadia INNER JOIN b_productospmpdia ON (b_minutafijadia.mfd_codpro = b_productospmpdia.ppd_codpro) AND (b_minutafijadia.mfd_fecha >= b_productospmpdia.ppd_fecdia) AND (b_minutafijadia.mfd_cencos = b_productospmpdia.ppd_cencos) SET b_minutafijadia.mfd_cospro = round(b_productospmpdia.ppd_propon, 2) " & _
              "WHERE b_minutafijadia.mfd_tipmin ='2' AND b_minutafijadia.mfd_fecha>=" & Format(CDate(vg_ciedia), "yyyymmdd") & " AND b_minutafijadia.mfd_fecha<=" & Format(CDate(fecter), "yyyymmdd") & " AND b_productospmpdia.ppd_fecdia= " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND b_productospmpdia.ppd_propon>0"
DoEvents: If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1

'-------> Borrar tablas b_productospmpdia distinto al ultimo día cierre que tengan precio cero y saldo en cero
vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <> " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"

'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
If Trim(aAp1) <> "" Then vg_db.Execute "DROP TABLE " & aAp1 & ""
If Trim(aAp2) <> "" Then vg_db.Execute "DROP TABLE " & aAp2 & ""
If Trim(aAp3) <> "" Then vg_db.Execute "DROP TABLE " & aAp3 & ""
If Trim(aAp4) <> "" Then vg_db.Execute "DROP TABLE " & aAp4 & ""
If Trim(aAp5) <> "" Then vg_db.Execute "DROP TABLE " & aAp5 & ""
If Trim(aAp6) <> "" Then vg_db.Execute "DROP TABLE " & aAp6 & ""

'-------> Grabar log_cierrediario
vg_db.Execute "INSERT INTO log_cierrediario VALUES ('" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "h:m:s") & "', '" & vg_ciedia & "', '" & Trim(vg_NUsr) & "', '1.- Cierre Diario', '" & MuestraCasino(1) & "')"

'-------> Crear tabla y mover datos
If op Then Formu.Label1(1).Caption = "Actua. datos anexo"
estpro = 2
Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String
'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile ' borrar base datos si existe
If Dir(mdir & MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".mdb") <> "" Then Kill mdir & MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".mdb" ' borrar base datos si existe

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open
'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN '" & cDBO & "'"
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN '" & cDBO & "'"
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN '" & cDBO & "'"
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN '" & cDBO & "' " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale,'dd/mm/yyyy') = CDATE('" & vg_ciedia & "')"
            
'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
vg_db.Execute "SELECT a.* INTO a_bodega IN '" & cdbi & "' FROM a_bodega a, b_clientes b " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo   = 0"

'-------> Generar tabla proveedor
vg_db.Execute "SELECT * INTO b_proveedor IN '" & cdbi & "' FROM b_proveedor"

'-------> Generar tabla compras
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
'dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'dd/mm/yyyy') AS toc_fecemi, format(toc_fecven,'dd/mm/yyyy') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = CDATE('" & vg_ciedia & "')"
dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'yyyymmdd') AS toc_fecemi, format(toc_fecven,'yyyymmdd') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, format(toc_fecrem,'yyyymmdd') toc_fecrem FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = CDATE('" & vg_ciedia & "')"
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
'dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'dd/mm/yyyy') AS toc_fecemi, format(toc_fecven,'dd/mm/yyyy') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem "
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'yyyymmdd') AS toc_fecemi, format(toc_fecven,'yyyymmdd') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, format(toc_fecrem ,'yyyymmdd')  toc_fecrem " & _
              "FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND   toc_tipdoc = 'SN' " & _
              "AND  (trim(toc_docsnc) <> '' or  not isnull(toc_docsnc)) "
            
'dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'dd/mm/yyyy') AS toc_fecemi, format(toc_fecven,'dd/mm/yyyy') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem "
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'yyyymmdd') AS toc_fecemi, format(toc_fecven,'yyyymmdd') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, format(toc_fecrem,'yyyymmdd') toc_fecrem " & _
              "FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND   toc_tipdoc = 'GD' " & _
              "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso)) "
  
dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras "
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R "
dbE.Execute "DROP TABLE b_totcompras_R"
  
dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras SELECT DISTINCT a.* FROM b_detcompras a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.dec_rutpro = b.toc_rutpro " & _
              "AND   a.dec_tipdoc = b.toc_tipdoc " & _
              "AND   a.dec_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_fecrem = CDATE('" & vg_ciedia & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
dbE.Execute "INSERT INTO b_detcompras SELECT DISTINCT a.* FROM b_detcompras a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.dec_rutpro = b.toc_rutpro " & _
              "AND   a.dec_tipdoc = b.toc_tipdoc " & _
              "AND   a.dec_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'SN' " & _
              "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc)) "

dbE.Execute "INSERT INTO b_detcompras SELECT DISTINCT a.* FROM b_detcompras a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.dec_rutpro = b.toc_rutpro " & _
              "AND   a.dec_tipdoc = b.toc_tipdoc " & _
              "AND   a.dec_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'GD' " & _
              "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso)) "

dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"


dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.* FROM b_detcomprasimp a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.imd_rutdoc = b.toc_rutpro " & _
              "AND   a.imd_tipdoc = b.toc_tipdoc " & _
              "AND   a.imd_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_fecrem = CDATE('" & vg_ciedia & "')"


'-------> Subir Solicitud y Guias Despachos cerradas detalle impuesto
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.* FROM b_detcomprasimp a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.imd_rutdoc = b.toc_rutpro " & _
              "AND   a.imd_tipdoc = b.toc_tipdoc " & _
              "AND   a.imd_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'SN' " & _
              "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc)) "

 
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.* FROM b_detcomprasimp a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.imd_rutdoc = b.toc_rutpro " & _
              "AND   a.imd_tipdoc = b.toc_tipdoc " & _
              "AND   a.imd_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'GD' " & _
              "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso)) "
              
dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R "
dbE.Execute "DROP TABLE b_detcomprasimp_R"


'-------> Generar tabla b_ocsacrecibido
dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc varchar(10), ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
            "SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, format(a.ocr_fecoc, 'yyyymmdd') , a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totcompras b IN '" & cDBO & "' " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & vg_ciedia & "')"


'-------> Generar tabla ventas
vg_db.Execute "SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, format(tov_fecemi,'yyyymmdd') AS tov_fecemi, format(tov_fecpro,'yyyymmdd') AS tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
              "INTO b_totventas IN '" & cdbi & "' FROM b_totventas " & _
              "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = CDATE('" & vg_ciedia & "') AND tov_tipdoc IN ('DP','SP')) " & _
              "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = CDATE('" & vg_ciedia & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"
  
'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc int, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = CDATE('" & vg_ciedia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = CDATE('" & vg_ciedia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = CDATE('" & vg_ciedia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = CDATE('" & vg_ciedia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla b_ocsacrecibido
'dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc FROM b_ocsacrecibido a, b_totventas b IN '" & cDBO & "'"
dbE.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
            "SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, format(a.ocr_fecoc, 'yyyymmdd') , a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & vg_ciedia & "')"

'-------> Generar tabla venta cafateria
'vg_db.Execute "SELECT tvc_cencos, format(tvc_fecing,'dd/mm/yyyy') AS tvc_fecing, tvc_codbod, tvc_estado INTO b_totventascaf IN '" & cDBI & "' FROM b_totventascaf "
vg_db.Execute "SELECT tvc_cencos, format(tvc_fecing,'yyyymmdd') AS tvc_fecing, tvc_codbod, tvc_estado INTO b_totventascaf IN '" & cdbi & "' FROM b_totventascaf " & _
              "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   tvc_fecing = CDATE('" & vg_ciedia & "') " & _
              "AND   tvc_codbod = " & vg_codbod & " " & _
              "AND   tvc_estado = 'C'"

'vg_db.Execute "SELECT a.dvc_cencos, format(a.dvc_fecing,'dd/mm/yyyy') AS dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, format(a.dvc_fecdoc,'dd/mm/yyyy') AS dvc_fecdoc INTO b_detventascaf IN '" & cDBI & "' FROM b_detventascaf a, b_totventascaf b "
vg_db.Execute "SELECT a.dvc_cencos, format(a.dvc_fecing,'yyyymmdd') AS dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, format(a.dvc_fecdoc,'yyyymmdd') AS dvc_fecdoc INTO b_detventascaf IN '" & cdbi & "' FROM b_detventascaf a, b_totventascaf b " & _
              "WHERE a.dvc_cencos = b.tvc_cencos " & _
              "AND   a.dvc_fecing = b.tvc_fecing " & _
              "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   b.tvc_fecing = CDATE('" & vg_ciedia & "') " & _
              "AND   b.tvc_codbod = " & vg_codbod & " " & _
              "AND   b.tvc_estado = 'C'"

'vg_db.Execute "SELECT a.dvp_cencos, format(a.dvp_fecing,'dd/mm/yyyy') AS dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos INTO b_detventascafpro IN '" & cDBI & "' FROM b_detventascafpro a, b_totventascaf b "
vg_db.Execute "SELECT a.dvp_cencos, format(a.dvp_fecing,'yyyymmdd') AS dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos INTO b_detventascafpro IN '" & cdbi & "' FROM b_detventascafpro a, b_totventascaf b " & _
              "WHERE a.dvp_cencos = b.tvc_cencos " & _
              "AND   a.dvp_fecing = b.tvc_fecing " & _
              "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   b.tvc_fecing = CDATE('" & vg_ciedia & "') " & _
              "AND   b.tvc_codbod = " & vg_codbod & " " & _
              "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
vg_db.Execute "SELECT * INTO a_regimen IN '" & cdbi & "' FROM a_regimen"

'-------> Generar tabla servicio
vg_db.Execute "SELECT * INTO a_servicio IN '" & cdbi & "' FROM a_servicio"

'-------> Generar tabla estructura de servicio
vg_db.Execute "SELECT * INTO a_estservicio IN '" & cdbi & "' FROM a_estservicio"

'-------> Generar tabla sector
vg_db.Execute "SELECT * INTO a_sector IN '" & cdbi & "' FROM a_sector"

'-------> Generar tabla servicio rac
vg_db.Execute "SELECT * INTO a_serviciorac IN '" & cdbi & "' FROM a_serviciorac"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then
    vg_db.Execute "SELECT DISTINCT a.* INTO b_minuta IN '" & cdbi & "' FROM b_minuta a, b_minutadet b " & _
                  "WHERE a.min_codigo = b.mid_codigo " & _
                  "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   a.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & ""

    vg_db.Execute "SELECT a.* INTO b_minutadet IN '" & cdbi & "' FROM b_minutadet a, b_minuta b " & _
                  "WHERE b.min_codigo = a.mid_codigo " & _
                  "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   b.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & ""
'Else
'    vg_db.Execute "SELECT DISTINCT a.* INTO b_minuta IN '" & cDBI & "' FROM b_minuta a, b_minutadet b " & _
'                  "WHERE a.min_codigo = b.mid_codigo " & _
'                  "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'                  "AND   a.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & " " & _
'                  "AND   b.mid_tipmin = '2'"
'
'    vg_db.Execute "SELECT a.* INTO b_minutadet IN '" & cDBI & "' FROM b_minutadet a, b_minuta b " & _
'                  "WHERE b.min_codigo = a.mid_codigo " & _
'                  "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
'                  "AND   b.min_fecmin = " & Format(vg_ciedia, "yyyymmdd") & " " & _
'                  "AND   b.mid_tipmin = '2'"
'End If
'-------> Generar tabla minutafija
vg_db.Execute "SELECT * INTO b_minutafijadia IN '" & cdbi & "' FROM b_minutafijadia " & _
              "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   mfd_fecha  = " & Format(vg_ciedia, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
vg_db.Execute "SELECT * INTO b_minutaraciones IN '" & cdbi & "' FROM b_minutaraciones " & _
              "WHERE mir_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   mid(mir_fecmin,1,6) = " & Format(vg_ciedia, "yyyymm") & ""

'-------> Generar tabla precio venta
vg_db.Execute "SELECT * INTO b_preciovta IN '" & cdbi & "' FROM b_preciovta " & _
              "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
vg_db.Execute "SELECT * INTO b_ventacontado IN '" & cdbi & "' FROM b_ventacontado " & _
              "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   vtc_fecvta = " & Format(vg_ciedia, "yyyymmdd") & ""

vg_db.Execute "SELECT a.* INTO b_ventacontadodet IN '" & cdbi & "' FROM b_ventacontadodet a, b_ventacontado b " & _
              "WHERE b.vtc_codigo = a.vtd_codigo " & _
              "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   b.vtc_fecvta = " & Format(vg_ciedia, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
vg_db.Execute "SELECT * INTO b_clientecencos IN '" & cdbi & "' FROM b_clientecencos"

'-------> Generar tabla cliente
vg_db.Execute "SELECT * INTO b_clientes IN '" & cdbi & "' FROM b_clientes"

'-------> Generar tabla bodegas
vg_db.Execute "SELECT bod_codbod, bod_codpro, round(bod_canmer, 2) AS bod_canmer INTO b_bodegas IN '" & cdbi & "' FROM b_bodegas WHERE bod_codbod = " & vg_codbod & ""

'-------> Generar tabla precio cafeteria
vg_db.Execute "SELECT * INTO b_totpreciocaf IN '" & cdbi & "' FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

vg_db.Execute "SELECT a.* INTO b_detpreciocaf IN '" & cdbi & "' FROM b_detpreciocaf a, b_totpreciocaf b " & _
              "WHERE b.tpc_codigo = a.dpc_codigo " & _
              "AND   b.tpc_cencos = a.dpc_cencos " & _
              "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

RS1.Open "SELECT DISTINCT b.cie_fecini, b.cie_fecter FROM b_tomainv a, b_cierreperiodo b " & _
         "WHERE b.cie_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   a.tin_fectom = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
         "AND   a.tin_ciemes = b.cie_periodo", vg_db, adOpenStatic
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   vg_db.Execute "SELECT * INTO b_tomainv IN '" & cdbi & "' FROM b_tomainv " & _
                 "WHERE tin_fectom >= " & RS1!cie_fecini & " AND tin_fectom <= " & RS1!cie_fecter & " " & _
                 "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
'   dbE.Execute "INSERT INTO b_totventas SELECT * FROM b_totventas IN '" & cDBO & "' "
   dbE.Execute "SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, format(tov_fecemi,'yyyymmdd') AS tov_fecemi, format(tov_fecpro,'yyyymmdd') AS tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
   '-------> Generar tabla ajuste inventario detalle
   dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
               "FROM b_detventas a, b_totventas b IN '" & cDBO & "' " & _
               "WHERE a.dev_rutcli = b.tov_rutcli " & _
               "AND   a.dev_tipdoc = b.tov_tipdoc " & _
               "AND   a.dev_numdoc = b.tov_numdoc " & _
               "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   b.tov_tipdoc IN ('AI') " & _
               "AND   b.tov_codbod = " & vg_codbod & ""
    
   '-------> Generar tabla ajuste inventario detalle impuesto
   dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
               "FROM b_detventasimp a, b_totventas b IN '" & cDBO & "' " & _
               "WHERE a.imd_rutdoc = b.tov_rutcli " & _
               "AND   a.imd_tipdoc = b.tov_tipdoc " & _
               "AND   a.imd_numdoc = b.tov_numdoc " & _
               "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   b.tov_tipdoc IN ('AI') " & _
               "AND   b.tov_codbod = " & vg_codbod & ""
Else
   '-------> Generar tabla toma inventario
   vg_db.Execute "SELECT * INTO b_tomainv IN '" & cdbi & "' FROM b_tomainv " & _
                 "WHERE tin_fectom = " & Format(vg_ciedia, "yyyymmdd") & " " & _
                 "AND   tin_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN '" & cDBO & "' " & _
            "WHERE feccie >= cdate('" & vg_ciedia & "') " & _
            "AND   feccie <= cdate('" & vg_ciedia & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN '" & cDBO & "' " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN '" & cDBO & "' " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN '" & cDBO & "' " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN '" & cDBO & "' " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN '" & cDBO & "' " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & vg_ciedia & "')"
'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
'dbe.Execute "INSERT INTO a_nomestenvio VALUES ('a_estenvio', 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
'dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_proveedor', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
'-------> Permite cambiar la estructura del campo fechas a la tabla a_servicio
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horcob char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horent char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horpda char(14)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_totcompras
dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecemi char(10)"
dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecven char(10)"
dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecrem char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecdig char(24)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecemi char(10)"
dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecpro char(10)"
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

'-------> Permite cambiar la estructura del campo fechas tabla b_ocsacrecibido
'dbE.Execute "ALTER TABLE b_ocsacrecibido ALTER COLUMN ocr_fecoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_totventascaf ALTER COLUMN tvc_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecing char(10)"
dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecdoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_detventascafpro ALTER COLUMN dvp_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE log_cierrediario ALTER COLUMN feccie char(10)"

dbE.Close
DoEvents: Formu.Bar1.Value = Formu.Bar1.Value + 1

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"

'-------> Actualizando cerrando periodo y abriendo proximo periodo
If op Then vg_db.Execute "UPDATE a_param SET par_valor='" & fg_Encripta(LimpiaDato(CDate(vg_ciedia) + 1)) & "' WHERE par_codigo='ciediario' AND par_cencos='" & MuestraCasino(1) & "'"
'-------> Grabar log_enviocierrediario
If op Then
   sql1 = IIf(vg_tipbase = "1", " CDATE('" & vg_ciedia & "') ", " '" & Format(vg_ciedia, "yyyymmdd") & "' ")
   RS1.Open "SELECT DISTINCT fecha FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & "", vg_db, adOpenStatic
   If RS1.EOF Then
      vg_db.Execute "INSERT INTO log_enviocierrediario VALUES ('" & MuestraCasino(1) & "', " & sql1 & ", '0', '')"
   Else
      vg_db.Execute "UPDATE log_enviocierrediario SET estenv = '0', fecsub = '' WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & ""
   End If
   RS1.Close: Set RS1 = Nothing
End If
Formu.Bar1(0).Visible = False: Formu.Bar1(0).Value = 0
Formu.Label1(1).Visible = False
fg_descarga
'If op Then MsgBox "Proceso de Cierre Día Finalizado", vbInformation + vbOKOnly, Msgtitulo
Exit Function
Error_CalcularPMPDia:
        fg_descarga
        MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
'       vg_db.RollbackTrans
       Resume Next
End Function

Function ProcesoCierreSql(Formu As Form, progrl As Boolean, fecdia) As Boolean

On Local Error GoTo Error_ProcesoCierreSql

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String, aAp1 As String, aAp2 As String, aAp3 As String, aAp4 As String, aAp5 As String, aAp6 As String
Dim fecini As Long, fecfin As Long, i As Long, FecInv As Long, est As Boolean, fecpro As Date, fecter As Date
Dim estpro As Integer, sql1 As String
Dim fecpro1 As String, fecter1 As String, fecdiad As String
Dim CierreDia As String
Dim auxCanmer As Double, auxPropon As Double, propon As Double, auxfec As Long, fecuco As String, upreco As Double
'-------> Grabar log_cierrediario
'vg_db.Execute "INSERT INTO log_cierrediario VALUES ('" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "', '" & Format(fecdia, "yyyymmdd") & "', '" & Trim(vg_NUsr) & "', '1.- Cierre Diario', '" & MuestraCasino(1) & "')"

ProcesoCierreSql = True
'-------> Crear tabla y mover datos
estpro = 2
Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String, DBO As String
'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(fecdia, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) = sourcefile Then Exit Function 'si existe sale
If Dir(mdir & MuestraCasino(1) & Format(fecdia, "yyyymmdd") & ".mdb") = MuestraCasino(1) & Format(fecdia, "yyyymmdd") & ".mdb" Then Exit Function 'si existe sale

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT DISTINCT par_nombre, par_valor FROM a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
If Not RS.EOF Then
       
       CierreDia = CDate(fg_Desencripta(TipoDato(RS!par_valor, ""))) - 1

End If
RS.Close
Set RS = Nothing

'-------> Ini : Generar tabla Costo minuta & Food Cost
dbE.Execute "CREATE TABLE B_CostoMinutaRealizadoFoodCostN (IdCeco char(10), Fecha_Minuta int, IdRegimen int, IdServicio int, Raciones_Teorica int, " & _
            "Costo_Teorico_Alim float, Costo_Teorico_Desec float, Raciones_Real int, Costo_Real_Alim float, Costo_Real_Desec float, Raciones_Vendidas int, " & _
            "Costo_Realizado_Alim float, Costo_Realizado_Desec float, Venta_Dia float, Venta_Contado float, Glosa_Venta_Especial char(100)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraerCostoFoodCostMinutaCierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(fecdia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_CostoMinutaRealizadoFoodCostN values (' " & RS(0) & "', " & RS(3) & ", " & RS(1) & ", " & RS(2) & ", " & RS(4) & " " & _
                  ", " & RS(5) & ", " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", " & RS(14) & ", '" & RS(15) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Costo minuta & Food Cost

'-------> Ini : Generar tabla Insumos & Food Cost A13
dbE.Execute "CREATE TABLE B_A13InsumosFCost (IdCeco char(10), Periodo int, FechaIni Int, FechaFin int, FechaCierre int, " & _
            "Glosa char(200), Alimentos Float, Lim_Desc Float, Total Float, Porcentaje Float, Id int)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'-------> comparar mes
Dim FecFinMes As String
If Format(CDate(CierreDia), "mm") <> Format(CDate(fecdia), "mm") Then

   FecFinMes = dEoM("25/" & Format(fecdia, "mm/yyyy"))

ElseIf Format(CDate(fecdia), "ddmmyyyy") < Format(CDate(CierreDia), "ddmmyyyy") Then

   FecFinMes = Format(CDate(CierreDia), "dd/mm/yyyy")

ElseIf Format(CDate(fecdia), "ddmmyyyy") = Format(CDate(CierreDia), "ddmmyyyy") Then

   FecFinMes = Format(CDate(CierreDia), "dd/mm/yyyy")

End If

Set RS = vg_db.Execute("sgp_Sel_VtaCtoServInsumosFoodCostGastoA13CierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(CDate(FecFinMes), "yyyymmdd") & ", " & Format(FecFinMes, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_A13InsumosFCost values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & Format(CDate(FecFinMes), "yyyymmdd") & ", '" & RS(4) & "' " & _
                  ", " & RS(5) & ", '" & RS(6) & "', '" & RS(7) & "', " & RS(8) & ", " & RS(9) & ")"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Insumos & Food Cost A13

'-------> Ini : Generar tabla detalle mermas
dbE.Execute "CREATE TABLE B_DetalleMermasNK (IdCeco char(10), Periodo int, Fecha_Minuta int, IdRegimen int, IdServicio int, FechaCierre int, " & _
            "IdReceta int, IdEstServicio int, NumLin int, CostoRecetaAlimento float, CostoRecetaDesechable float, CantidadMerma float, CantidadRacionReal float, MermaxKilo float) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleMermasCierreDiario '" & MuestraCasino(1) & "', " & Format(fecdia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_DetalleMermasNK values (' " & RS(0) & "', " & Format(fecdia, "yyyymm") & ", " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & Format(CDate(FecFinMes), "yyyymmdd") & ", " & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ")"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla detalle mermas

'-------> Ini : Generar tabla mermas desconche - pan - produccion
dbE.Execute "CREATE TABLE B_MermaDesconche (IdCeco char(10), IdRegimen int, IdServicio int, Fecha_Merma int, " & _
            "Considera_Merma char(1), Merma_Desconche float, Merma_Pan float, Merma_Produccion float, Fecha_Modificacion datetime, Fecha_Creacion datetime, Usuario char(20)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_MermaDesconcheCierreDiario '" & MuestraCasino(1) & "', " & Format(fecdia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_MermaDesconche values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", '" & RS(8) & "', '" & RS(9) & "', '" & RS(10) & "')"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla mermas desconche - pan - produccion

'-------> Ini : Generar tabla consumo proyectado & real
dbE.Execute "CREATE TABLE B_ConsumoProyectadoReal (IdCeco char(10), IdRegimen int, NoRegimen char(50), IdServicio int, NoServicio char(50), NumDoc int, Periodo int, Fecha int, " & _
            "Codigo_Producto char(20), Descripcion_Producto char(100), Unidad char(5), Cantidad_Teorica float, Cantidad_Planificada float, Cantidad_Realizada float, PMP float, Racion_Teorica int, Usuario_Mod_Racion_Real char(20), Racion_Real int, Fecha_Mod_Racion_Real datetime, Usuario_Salida_Produccion char(20), Racion_Salida_Produccion int, Fecha_Mod_Salida_Produccion datetime, Cantidad_Devolucion float)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CalcularSalidaProducciónMinutaTeorica '" & MuestraCasino(1) & "', " & Format(fecdia, "yyyymmdd") & ", '1', " & vg_codbod & ", " & Format(fecdia, "yyyymmdd") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
         
      dbE.Execute "insert into B_ConsumoProyectadoReal values ('" & MuestraCasino(1) & "', " & RS(0) & ", '" & RS(1) & "', " & RS(2) & ", '" & RS(3) & "', " & RS(5) & "," & _
                  "" & Format(fecdia, "yyyymm") & ", " & Format(fecdia, "yyyymmdd") & ", '" & RS(6) & "', '" & RS(7) & "', '" & RS(8) & "', " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", '" & RS(14) & "', " & RS(15) & ", '" & RS(16) & "', '" & RS(17) & "', " & RS(18) & ", '" & RS(19) & "', '" & RS(20) & "')"
           
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla consumo proyectado & real

'-------> Ini : Generar tabla a_tipoajuste 31/03/2022
dbE.Execute "CREATE TABLE a_tipoajuste (aju_codigo int, aju_nombre char(30), aju_tipaju int, aju_tipo char(1))"
dbE.Execute "INSERT INTO a_tipoajuste SELECT aju_codigo, aju_nombre, aju_tipaju, aju_tipo " & _
            "FROM a_tipoajuste IN " & DBO & ""
'-------> Fin : Generar tabla a_tipoajuste  31/03/2022

'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN " & DBO & ""
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN " & DBO & ""
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN " & DBO & ""
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN " & DBO & " " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale, 'dd/mm/yyyy') = CDATE('" & fecdia & "')"

'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
dbE.Execute "CREATE TABLE a_bodega (bod_codigo int, bod_nombre char(25), bod_ubicac char(35))"
dbE.Execute "INSERT INTO a_bodega SELECT bod_codigo, bod_nombre, bod_ubicac FROM a_bodega a, b_clientes b IN " & DBO & " " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo = 0"

'-------> Generar tabla proveedor
dbE.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo datetime, prv_origen char(1))"
dbE.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen FROM b_proveedor IN " & DBO & ""

'-------> Generar tabla compras
'            "WHERE toc_codbod = " & vg_codbod & " AND toc_fecemi = cdate('" & fecdia & "')"
            
'dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc int, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = cdate('" & fecdia & "')"
            
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
'dbE.Execute "INSERT INTO b_totcompras " & _
'            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
'            "FROM b_totcompras IN " & DBO & " " & _
'            "WHERE toc_codbod = " & vg_codbod & " " & _
'            "AND   toc_tipdoc = 'SN' " & _
'            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc))"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc))"

'dbE.Execute "INSERT INTO b_totcompras " & _
'            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
'            "FROM b_totcompras IN " & DBO & " " & _
'            "WHERE toc_codbod = " & vg_codbod & " " & _
'            "AND   toc_tipdoc = 'GD' " & _
'            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso))"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso))"

dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras "
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R "
dbE.Execute "DROP TABLE b_totcompras_R"

'dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc int, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = cdate('" & fecdia & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
'dbE.Execute "INSERT INTO b_detcompras " & _
'            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
'            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.dec_rutpro = b.toc_rutpro " & _
'            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
'            "AND   a.dec_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'SN' " & _
'            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

'dbE.Execute "INSERT INTO b_detcompras " & _
'            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
'            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.dec_rutpro = b.toc_rutpro " & _
'            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
'            "AND   a.dec_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'GD' " & _
'            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"
            
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"
            
dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"

'dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = cdate('" & fecdia & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle impuesto
'dbE.Execute "INSERT INTO b_detcomprasimp " & _
'            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
'            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
'            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
'            "AND   a.imd_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'SN' " & _
'            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

'dbE.Execute "INSERT INTO b_detcomprasimp " & _
'            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
'            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
'            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
'            "AND   a.imd_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'GD' " & _
'            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"

dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R "
dbE.Execute "DROP TABLE b_detcomprasimp_R"

'-------> Generar tabla b_ocsacrecibido
'            "AND   b.toc_fecemi = CDATE('" & fecdia & "')"

dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc datetime, ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc FROM b_ocsacrecibido a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & fecdia & "')"

'-------> Generar tabla ventas
'dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc int, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc double, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
            "FROM b_totventas  IN " & DBO & " " & _
            "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = cdate('" & fecdia & "') AND tov_tipdoc IN ('DP','SP')) " & _
            "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = cdate('" & fecdia & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"

'-------> Crear estructura tabla detalle ventas
'dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc int, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc double, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & fecdia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & fecdia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
'dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & fecdia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & fecdia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla ventas servicios especiales
dbE.Execute "CREATE TABLE b_totventaserviciosespeciales (tos_IdCeco char(10), tos_Tipo_Documento char(2), tos_Numero_Documento int, tos_Fecha_Produccion datetime, tos_IdBodega int, tos_Venta_servicio_Especiales char(100), Tos_Comensales double, tos_Precio_Servicio double, tos_Total_Documento double, tos_Estado_Documento char(1), tos_Fecha_Creacion datetime, tos_Fecha_Modificacion datetime, tos_Periodo int, tos_Usuario char(20), tos_Documento_Asociado int)"
dbE.Execute "INSERT INTO b_totventaserviciosespeciales SELECT tos_IdCeco, tos_Tipo_Documento, tos_Numero_Documento, tos_Fecha_Produccion, tos_IdBodega, tos_venta_Servicio_Especiales, Tos_Comensales, tos_Precio_servicio, tos_Total_Documento, tos_Estado_Documento, tos_Fecha_Creacion, tos_Fecha_modificacion, tos_Periodo, tos_Usuario, tos_Documento_Asociado " & _
            "FROM b_totventaserviciosespeciales  IN " & DBO & " " & _
            "WHERE tos_IdBodega = " & vg_codbod & " AND tos_fecha_Produccion = cdate('" & fecdia & "') AND tos_tipo_documento IN ('DE','SE') "

'-------> Crear estructura tabla detalle ventas servicios especiales
dbE.Execute "CREATE TABLE b_detventaserviciosespeciales (des_IdCeco char(10), des_Tipo_Documento char(2), des_Numero_Documento int, des_Numero_Linea int, des_IdProducto char(20), des_Cantidad_Mercaderia double, des_Cantidad_Devolver double, des_Precio_Documento double, des_Total_Documento double, des_Descripcion char(255), des_Mueve_Inventario char(1), des_Actualiza_Precio char(1))"
dbE.Execute "INSERT INTO b_detventaserviciosespeciales SELECT a.des_IdCeco, a.des_tipo_documento, a.des_numero_documento, a.des_numero_linea, a.des_IdProducto, a.des_cantidad_mercaderia, a.des_cantidad_devolver, a.des_Precio_Documento, a.des_Total_documento, a.des_Descripcion, a.des_Mueve_Inventario, a.des_Actualiza_Precio " & _
            "FROM b_detventaserviciosespeciales a, b_totventaserviciosespeciales b IN " & DBO & " " & _
            "WHERE a.des_IdCeco = b.tos_IdCeco " & _
            "AND   a.des_tipo_documento = b.tos_tipo_documento " & _
            "AND   a.des_numero_documento = b.tos_numero_documento " & _
            "AND   b.tos_fecha_produccion = cdate('" & fecdia & "') AND b.tos_tipo_documento IN ('DE','SE') " & _
            "AND   b.tos_IdBodega = " & vg_codbod & ""

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & fecdia & "')"

dbE.Execute "CREATE TABLE b_totventascaf (tvc_cencos char(10), tvc_fecing datetime, tvc_codbod int, tvc_estado char(1))"
dbE.Execute "INSERT INTO b_totventascaf SELECT tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado FROM b_totventascaf IN " & DBO & " " & _
            "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   tvc_fecing = cdate('" & fecdia & "') " & _
            "AND   tvc_codbod = " & vg_codbod & " " & _
            "AND   tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascaf (dvc_cencos char(10), dvc_fecing datetime, dvc_numlin int, dvc_articulo char(20), dvc_canart double, dvc_precio double, dvc_tippag char(2), dvc_rutcli char(10), dvc_cencli char(10), dvc_tipdoc char(2), dvc_numdoc int, dvc_fecdoc datetime)"
dbE.Execute "INSERT INTO b_detventascaf SELECT a.dvc_cencos, a.dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, a.dvc_fecdoc FROM b_detventascaf a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvc_cencos = b.tvc_cencos " & _
            "AND   a.dvc_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & fecdia & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascafpro (dvp_cencos char(10), dvp_fecing datetime, dvp_codmer char(20), dvp_cancal double, dvp_candig double, dvp_precos double)"
dbE.Execute "INSERT INTO b_detventascafpro SELECT a.dvp_cencos, a.dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos FROM b_detventascafpro a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvp_cencos = b.tvc_cencos " & _
            "AND   a.dvp_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & fecdia & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
dbE.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
dbE.Execute "INSERT INTO a_regimen SELECT reg_codigo, reg_nombre FROM a_regimen IN " & DBO & ""
'-------> Generar tabla servicio
dbE.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(30), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob datetime, ser_horent datetime, ser_horpda datetime)"
dbE.Execute "INSERT INTO a_servicio SELECT ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_activo, ser_horcob, ser_horent, ser_horpda FROM a_servicio IN " & DBO & ""

'-------> Generar tabla estructura de servicio
dbE.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double, ess_cencos char(10))"
dbE.Execute "INSERT INTO a_estservicio SELECT ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin, ess_cencos FROM a_estservicio IN " & DBO & " where ess_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla sector
dbE.Execute "CREATE TABLE a_sector (sec_codigo int, sec_nombre char(50), sec_orden int)"
dbE.Execute "INSERT INTO a_sector SELECT sec_codigo, sec_nombre, sec_orden FROM a_sector IN " & DBO & ""

'-------> Generar tabla servicio rac
dbE.Execute "CREATE TABLE a_serviciorac (sra_codser int, sra_coditem int, sra_serdia int, sra_raciones int, sra_cencos char(10))"
dbE.Execute "INSERT INTO a_serviciorac SELECT sra_codser, sra_coditem, sra_serdia, sra_raciones, sra_cencos FROM a_serviciorac IN " & DBO & " where sra_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then
   dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
   
'-------> Si tipo de minuta es distinto simap puede generar minuta detalle cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
   
   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
               "AND   a.min_fecmin = " & Format(fecdia, "yyyymmdd") & ""

End If

dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
    
'-------> Si tipo de minuta es distinto simap puede generar minuta detalle cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then

    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
                "WHERE b.min_codigo = a.mid_codigo " & _
                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                "AND   b.min_fecmin = " & Format(fecdia, "yyyymmdd") & ""

End If
'Else
'   dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
'               "WHERE a.min_codigo = b.mid_codigo " & _
'               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'               "AND   a.min_fecmin = " & Format(fecdia, "yyyymmdd") & " " & _
'               "AND   b.mid_tipmin = '2'"
'
'    dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
'                "WHERE b.min_codigo = a.mid_codigo " & _
'                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
'                "AND   b.min_fecmin = " & Format(fecdia, "yyyymmdd") & " " & _
'                "AND   b.mid_tipmin = '2'"
'End If
'-------> Generar tabla minutafija
dbE.Execute "CREATE TABLE b_minutafijadia (mfd_cencos char(10), mfd_codreg int, mfd_codser int, mfd_fecha int, mfd_codpro char(20), mfd_tipmin char(1), mfd_canpro double, mfd_cospro double)"
dbE.Execute "INSERT INTO b_minutafijadia SELECT mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro FROM b_minutafijadia IN " & DBO & " " & _
            "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mfd_fecha  = " & Format(fecdia, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
dbE.Execute "CREATE TABLE b_minutaraciones (mir_cencos char(10), mir_codreg int, mir_codser int, mir_fecmin int, mir_rutcli char(10), mir_nrorac int, mir_nroguia int, mir_codcli char(10))"
dbE.Execute "INSERT INTO b_minutaraciones SELECT mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli FROM b_minutaraciones IN " & DBO & " " & _
            "WHERE mir_cencos          = '" & MuestraCasino(1) & "' " & _
            "AND   mid(mir_fecmin,1,6) = " & Format(fecdia, "yyyymm") & ""

'-------> Insertar mermas
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) " & _
"SELECT  bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin , " & _
"        'MERMAS' as mir_rutcli , " & _
"        SUM(bm2.mid_nummer) As mermas, " & _
"        0,                              " & _
"        ''                             " & _
"FROM    b_minuta AS bm, " & _
"        b_minutadet AS bm2 IN " & DBO & " " & _
"WHERE   bm.min_codigo = bm2.mid_codigo and bm.min_cencos = '" & MuestraCasino(1) & "' " & _
"        AND mid(bm.min_fecmin, 1, 6) = " & Format(vg_ciedia, "yyyymm") & " " & _
"        AND bm2.mid_tipmin = '2' " & _
"GROUP BY bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin"

'-------> Generar tabla precio venta
dbE.Execute "CREATE TABLE b_preciovta (prv_cencos char(10), prv_codreg int, prv_codser int, prv_fecvig int, prv_rutcli char(10), prv_preven double)"
dbE.Execute "INSERT INTO b_preciovta SELECT prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli, prv_preven FROM b_preciovta IN " & DBO & " " & _
            "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
dbE.Execute "CREATE TABLE b_ventacontado (vtc_codigo int, vtc_cencos char(10), vtc_codreg int, vtc_codser int, vtc_fecvta int, vtc_forpag int, vtc_totmon double, vtc_opccli char(1))"
dbE.Execute "INSERT INTO b_ventacontado SELECT vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli FROM b_ventacontado IN " & DBO & " " & _
            "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   vtc_fecvta = " & Format(fecdia, "yyyymmdd") & ""

dbE.Execute "CREATE TABLE b_ventacontadodet (vtd_codigo int, vtd_numlin int, vtd_codcli char(10), vtd_codcco char(10), vtd_descripcion char(50), vtd_detmon double)"
dbE.Execute "INSERT INTO b_ventacontadodet SELECT a.vtd_codigo, a.vtd_numlin, a.vtd_codcli, a.vtd_codcco, a.vtd_descripcion, a.vtd_detmon FROM b_ventacontadodet a, b_ventacontado b IN " & DBO & " " & _
            "WHERE b.vtc_codigo = a.vtd_codigo " & _
            "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.vtc_fecvta = " & Format(fecdia, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
dbE.Execute "CREATE TABLE b_clientecencos (clc_codigo char(10), clc_codcli char(10), clc_nombre char(50))"
dbE.Execute "INSERT INTO b_clientecencos SELECT clc_codigo, clc_codcli, clc_nombre FROM b_clientecencos IN " & DBO & " where clc_codcli = '" & MuestraCasino(1) & "'"

'-------> Generar tabla cliente
dbE.Execute "CREATE TABLE b_clientes (cli_codigo char(10), cli_nombre char(50), cli_direccion char(50), cli_comuna char(15), cli_ciudad char(15), cli_fono1 char(15), cli_fono2 char(15), cli_fax char(15), cli_percon char(50), cli_giro char(50), cli_email char(50), cli_tipo int, cli_codbod int, cli_codtis int, cli_codseg int, cli_codcli char(10), cli_clisap char(1), cli_socsap char(4), cli_cievta char(1), cli_ciedia int, cli_activo char(1), cli_sobrec char(1), cli_codmun int, cli_ccisac int, cli_cecsac char(4), cli_codreg int, id_tipo_vale char(100))"
dbE.Execute "INSERT INTO b_clientes SELECT cli_codigo, cli_nombre, cli_direccion, cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, cli_giro, cli_email, cli_tipo, cli_codbod, cli_codtis, cli_codseg, cli_codcli, cli_clisap, cli_socsap, cli_cievta, cli_ciedia, cli_activo, cli_sobrec, cli_codmun, cli_ccisac, cli_cecsac, cli_codreg FROM b_clientes IN " & DBO & ""

'-------> Generar tabla bodegas
dbE.Execute "CREATE TABLE b_bodegas (bod_codbod int, bod_codpro char(20), bod_canmer double)"

If Format(CDate(fecdia), "ddmmyyyy") = Format(CDate(CierreDia), "ddmmyyyy") Then
   
   vg_db.Execute ("delete paso_b_bodegas where bod_Cencos = '" & MuestraCasino(1) & "'")
   vg_db.Execute ("insert into paso_b_bodegas (bod_cencos, bod_codbod, bod_codpro, bod_canmer) SELECT distinct '" & MuestraCasino(1) & "', bod_codbod, bod_codpro, round(bod_canmer,2) FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " ")
   dbE.Execute "INSERT INTO b_bodegas SELECT distinct bod_codbod, bod_codpro, bod_canmer FROM paso_b_bodegas IN " & DBO & " WHERE bod_cencos = '" & MuestraCasino(1) & "' and bod_codbod = " & vg_codbod & ""

End If

'-------> Generar tabla precio cafeteria
dbE.Execute "CREATE TABLE b_totpreciocaf (tpc_codigo char(20), tpc_nombre char(50), tpc_precio double, tpc_cencos char(10), tpc_activo char(1))"
dbE.Execute "INSERT INTO b_totpreciocaf SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_cencos, tpc_activo FROM b_totpreciocaf  IN " & DBO & " WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

dbE.Execute "CREATE TABLE b_detpreciocaf (dpc_codigo char(20), dpc_codmer char(20), dpc_cantidad double, dpc_cencos char(10))"
dbE.Execute "INSERT INTO b_detpreciocaf SELECT a.dpc_codigo, a.dpc_codmer, a.dpc_cantidad, a.dpc_cencos FROM b_detpreciocaf a, b_totpreciocaf b IN " & DBO & " " & _
            "WHERE b.tpc_codigo = a.dpc_codigo " & _
            "AND   b.tpc_cencos = a.dpc_cencos " & _
            "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

RS1.Open "SELECT DISTINCT b.cie_fecini, b.cie_fecter FROM b_tomainv a, b_cierreperiodo b " & _
         "WHERE b.cie_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   a.tin_fectom = " & Format(CDate(fecdia) - 1, "yyyymmdd") & " " & _
         "AND   a.tin_ciemes = b.cie_periodo", vg_db, adOpenStatic
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom >= " & RS1!cie_fecini & " AND tin_fectom <= " & RS1!cie_fecter & " " & _
               "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
   dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf FROM b_totventas IN " & DBO & " " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
    '-------> Generar tabla ajuste inventario detalle
    dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
                "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
                "WHERE a.dev_rutcli = b.tov_rutcli " & _
                "AND   a.dev_tipdoc = b.tov_tipdoc " & _
                "AND   a.dev_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
    
    '-------> Generar tabla ajuste inventario detalle impuesto
    dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
                "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
                "WHERE a.imd_rutdoc = b.tov_rutcli " & _
                "AND   a.imd_tipdoc = b.tov_tipdoc " & _
                "AND   a.imd_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
Else
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom = " & Format(fecdia, "yyyymmdd") & " " & _
               "AND   tin_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
'vg_db.Execute "SELECT * INTO log_cierrediario IN '" & cDBI & "' FROM log_cierrediario WHERE feccie >= cdate('" & fecdia & "') AND feccie <= cdate('" & fecdia & "') + 2"
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN " & DBO & " " & _
            "WHERE feccie >= cdate('" & fecdia & "') " & _
            "AND   feccie <= cdate('" & fecdia & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN " & DBO & " " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"
'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"

'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"

'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN " & DBO & " " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"

'-------> generar tabla a_derechosperfil
dbE.Execute "CREATE TABLE a_derechosperfil (dpe_cecori char(10), dpe_codper int, dpe_codopc int, dpe_deracc int, dpe_deragr int, dpe_dermod int, dpe_dereli int, dpe_derimp int)"
dbE.Execute "INSERT INTO a_derechosperfil (dpe_cecori, dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT distinct '" & MuestraCasino(1) & "', dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp FROM a_derechosperfil IN " & DBO & " "

'-------> actualizar tabla a_opcsistema
Dim ciedia As String
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE a_opcsistema SET EnvioDocSGPADM = '" & ciedia & "' FROM a_opcsistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = ''")

'-------> generar tabla a_opcsistema
dbE.Execute "CREATE TABLE a_opcsistema (opc_cecori char(10), opc_codigo int, opc_nombre char(50))"
dbE.Execute "INSERT INTO a_opcsistema (opc_cecori, opc_codigo, opc_nombre) SELECT distinct '" & MuestraCasino(1) & "', opc_codigo, opc_nombre FROM a_opcsistema IN " & DBO & "  where EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> generar tabla a_perfil
dbE.Execute "CREATE TABLE a_perfil (per_cecori char(10), per_codigo int, per_nombre char(30))"
dbE.Execute "INSERT INTO a_perfil (per_cecori, per_codigo, per_nombre) SELECT distinct '" & MuestraCasino(1) & "', per_codigo, per_nombre FROM a_perfil IN " & DBO & " "

'-------> generar tabla a_usuarios
dbE.Execute "CREATE TABLE a_usuarios (usu_cecori char(10), usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios (usu_cecori, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios IN " & DBO & " "

'-------> generar tabla a_usuarios_eliminado
dbE.Execute "CREATE TABLE a_usuarios_eliminado (usu_cecori char(10), Fecha datetime, usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios_eliminado (usu_cecori, Fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios_eliminado IN " & DBO & " "

'-------> generar tabla b_usuariocontratos
dbE.Execute "CREATE TABLE b_usuariocontratos (uco_cecori char(10), uco_codusu char(20), uco_codcon char(10))"
dbE.Execute "INSERT INTO b_usuariocontratos (uco_cecori, uco_codusu, uco_codcon) SELECT distinct '" & MuestraCasino(1) & "', uco_codusu, uco_codcon FROM b_usuariocontratos IN " & DBO & " "

'-------> actualizar tabla log_sistema
ciedia = ""
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE Log_Sistema SET EnvioDocSGPADM = '" & ciedia & "' FROM Log_Sistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = '' and Loc_Id in (2,20,21,22,23)")
               
'-------> generar tabla Log_Sistema
dbE.Execute "CREATE TABLE Log_Sistema (cecori char(10), Fecha datetime, Usuario_Id char(20), Loc_Id int, Opcion_Sistema char(14), Dato_Nuevo memo, Dato_Anterior memo, Detalle_Operacion memo)"
dbE.Execute "INSERT INTO Log_Sistema (cecori, Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion) SELECT distinct '" & MuestraCasino(1) & "', Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion FROM Log_Sistema IN " & DBO & " where Loc_Id in (2,20,21,22,23) and EnvioDocSGPADM = cdate('" & vg_ciedia & "') "


'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
'dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_proveedor', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_CostoMinutaRealizadoFoodCostN', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_A13InsumosFCost', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_DetalleMermasNK', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_MermaDesconche', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_tipoajuste', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_ConsumoProyectadoReal', 0, '', 0, 0, 0)"

'-------> Permite cambiar la estructura del campo fechas a la tabla a_servicio
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horcob char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horent char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horpda char(14)"


dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_derechosperfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_opcsistema', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_perfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios_eliminado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_usuariocontratos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('Log_Sistema', 0, '', 0, 0, 0)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_totcompras
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecemi char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecven char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecrem char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecdig char(24)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecemi char(10)"
'dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecpro char(10)"
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

'-------> Permite cambiar la estructura del campo fechas tabla b_ocsacrecibido
'dbE.Execute "ALTER TABLE b_ocsacrecibido ALTER COLUMN ocr_fecoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_totventascaf ALTER COLUMN tvc_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecing char(10)"
'dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecdoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_detventascafpro ALTER COLUMN dvp_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE log_cierrediario ALTER COLUMN feccie char(10)"

dbE.Close
DoEvents

If ValidaMDBCierreDiario(dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx") Then

        ProcesoCierreSql = False
   
        Exit Function

End If

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"
fg_descarga

Exit Function
Error_ProcesoCierreSql:
       
       fg_descarga
       ProcesoCierreSql = False
       MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
       Resume Next

End Function

Function ProcesoCierreAccess(Formu As Form, progrl As Boolean, fecdia As String)

On Local Error GoTo Error_ProcesoCierreAccess

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String, aAp1 As String, aAp2 As String, aAp3 As String, aAp4 As String, aAp5 As String, aAp6 As String
Dim fecini As Long, fecfin As Long, i As Long, FecInv As Long, est As Boolean, fecpro As Date, fecter As Date
Dim estpro As Integer, sql1 As String
'-------> Crear tabla y mover datos
estpro = 2
Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String
'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(fecdia, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) = sourcefile Then Exit Function 'salir si existe xxx
If Dir(mdir & MuestraCasino(1) & Format(fecdia, "yyyymmdd") & ".mdb") = MuestraCasino(1) & Format(fecdia, "yyyymmdd") & ".mdb" Then Exit Function 'salir si existe xxx

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open

'-------> Ini : Generar tabla a_tipoajuste 31/03/2022
dbE.Execute "CREATE TABLE a_tipoajuste (aju_codigo int, aju_nombre char(30), aju_tipaju int, aju_tipo char(1))"
dbE.Execute "INSERT INTO a_tipoajuste SELECT aju_codigo, aju_nombre, aju_tipaju, aju_tipo " & _
            "FROM a_tipoajuste IN " & cDBO & ""
'-------> Fin : Generar tabla a_tipoajuste  31/03/2022

'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN '" & cDBO & "'"
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN '" & cDBO & "'"
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN '" & cDBO & "'"
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN '" & cDBO & "' " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale,'dd/mm/yyyy') = CDATE('" & vg_ciedia & "')"

'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
vg_db.Execute "SELECT a.* INTO a_bodega IN '" & cdbi & "' FROM a_bodega a, b_clientes b " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo   = 0"

'-------> Generar tabla proveedor
vg_db.Execute "SELECT * INTO b_proveedor IN '" & cdbi & "' FROM b_proveedor"

'-------> Generar tabla compras
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
'dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'dd/mm/yyyy') AS toc_fecemi, format(toc_fecven,'dd/mm/yyyy') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = CDATE('" & fecdia & "')"
dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'yyyymmdd') AS toc_fecemi, format(toc_fecven,'yyyymmdd') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, format(toc_fecrem, 'yyyymmdd') toc_fecrem FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = CDATE('" & fecdia & "')"
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
'dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'dd/mm/yyyy') AS toc_fecemi, format(toc_fecven,'dd/mm/yyyy') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem "
dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'yyyymmdd') AS toc_fecemi, format(toc_fecven,'yyyymmdd') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, format(toc_fecrem,'yyyymmdd') toc_fecrem " & _
              "FROM b_totcompras IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND toc_tipdoc = 'SN' " & _
              "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc))"
            
'dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'dd/mm/yyyy') AS toc_fecemi, format(toc_fecven,'dd/mm/yyyy') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem "
dbE.Execute "INSERT INTO b_totcompras SELECT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, format(toc_fecemi,'yyyymmdd') AS toc_fecemi, format(toc_fecven,'yyyymmdd') AS toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, format(toc_fecrem ,'yyyymmdd') toc_fecrem " & _
              "FROM b_totcompras  IN '" & cDBO & "' WHERE toc_codbod = " & vg_codbod & " AND toc_tipdoc = 'GD' " & _
              "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso)) "
  
dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras"
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R"
dbE.Execute "DROP TABLE b_totcompras_R"

dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras SELECT DISTINCT a.* FROM b_detcompras a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.dec_rutpro = b.toc_rutpro " & _
              "AND   a.dec_tipdoc = b.toc_tipdoc " & _
              "AND   a.dec_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_fecrem = CDATE('" & fecdia & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
dbE.Execute "INSERT INTO b_detcompras SELECT DISTINCT a.* FROM b_detcompras a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.dec_rutpro = b.toc_rutpro " & _
              "AND   a.dec_tipdoc = b.toc_tipdoc " & _
              "AND   a.dec_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'SN' " & _
              "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc)) "

dbE.Execute "INSERT INTO b_detcompras SELECT DISTINCT a.* FROM b_detcompras a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.dec_rutpro = b.toc_rutpro " & _
              "AND   a.dec_tipdoc = b.toc_tipdoc " & _
              "AND   a.dec_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'GD' " & _
              "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso)) "
              
dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"

dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.* FROM b_detcomprasimp a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.imd_rutdoc = b.toc_rutpro " & _
              "AND   a.imd_tipdoc = b.toc_tipdoc " & _
              "AND   a.imd_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_fecrem = CDATE('" & fecdia & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle impuesto
dbE.Execute "insert INTO b_detcomprasimp SELECT DISTINCT a.* FROM b_detcomprasimp a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.imd_rutdoc = b.toc_rutpro " & _
              "AND   a.imd_tipdoc = b.toc_tipdoc " & _
              "AND   a.imd_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'SN' " & _
              "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc)) "
  
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.* FROM b_detcomprasimp a, b_totcompras b IN '" & cDBO & "' " & _
              "WHERE a.imd_rutdoc = b.toc_rutpro " & _
              "AND   a.imd_tipdoc = b.toc_tipdoc " & _
              "AND   a.imd_numdoc = b.toc_numdoc " & _
              "AND   b.toc_codbod = " & vg_codbod & " " & _
              "AND   b.toc_tipdoc = 'GD' " & _
              "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso)) "
  
dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R"
dbE.Execute "DROP TABLE b_detcomprasimp_R"
  
'-------> Generar tabla b_ocsacrecibido
dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc varchar(10), ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
            "SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, format(a.ocr_fecoc, 'yyyymmdd') , a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totcompras b IN '" & cDBO & "' " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & fecdia & "')"

'-------> Generar tabla ventas
vg_db.Execute "SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, format(tov_fecemi,'yyyymmdd') AS tov_fecemi, format(tov_fecpro,'yyyymmdd') AS tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
              "INTO b_totventas IN '" & cdbi & "' FROM b_totventas " & _
              "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = CDATE('" & fecdia & "') AND tov_tipdoc IN ('DP','SP')) " & _
              "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = CDATE('" & fecdia & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"
  
'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc int, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = CDATE('" & fecdia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = CDATE('" & fecdia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = CDATE('" & fecdia & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = CDATE('" & fecdia & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla ventas servicios especiales
dbE.Execute "CREATE TABLE b_totventaserviciosespeciales (tos_IdCeco char(10), tos_Tipo_Documento char(2), tos_Numero_Documento int, tos_Fecha_Produccion datetime, tos_IdBodega int, tos_Venta_servicio_Especiales char(100), Tos_Comensales double, tos_Precio_Servicio double, tos_Total_Documento double, tos_Estado_Documento char(1), tos_Fecha_Creacion datetime, tos_Fecha_Modificacion datetime, tos_Periodo int, tos_Usuario char(20), tos_Documento_Asociado int)"
dbE.Execute "INSERT INTO b_totventaserviciosespeciales SELECT tos_IdCeco, tos_Tipo_Documento, tos_Numero_Documento, tos_Fecha_Produccion, tos_IdBodega, tos_venta_Servicio_Especiales, Tos_Comensales, tos_Precio_servicio, tos_Total_Documento, tos_Estado_Documento, tos_Fecha_Creacion, tos_Fecha_modificacion, tos_Periodo, tos_Usuario, tos_Documento_Asociado " & _
            "FROM b_totventaserviciosespeciales  IN " & cDBO & " " & _
            "WHERE tos_IdBodega = " & vg_codbod & " AND tos_fecha_Produccion = cdate('" & fecdia & "') AND tos_tipo_documento IN ('DE','SE') "

'-------> Crear estructura tabla detalle ventas servicios especiales
dbE.Execute "CREATE TABLE b_detventaserviciosespeciales (des_IdCeco char(10), des_Tipo_Documento char(2), des_Numero_Documento int, des_Numero_Linea int, des_IdProducto char(20), des_Cantidad_Mercaderia double, des_Cantidad_Devolver double, des_Precio_Documento double, des_Total_Documento double, des_Descripcion char(255), des_Mueve_Inventario char(1), des_Actualiza_Precio char(1))"
dbE.Execute "INSERT INTO b_detventaserviciosespeciales SELECT a.des_IdCeco, a.des_tipo_documento, a.des_numero_documento, a.des_numero_linea, a.des_IdProducto, a.des_cantidad_mercaderia, a.des_cantidad_devolver, a.des_Precio_Documento, a.des_Total_documento, a.des_Descripcion, a.des_Mueve_Inventario, a.des_Actualiza_Precio " & _
            "FROM b_detventaserviciosespeciales a, b_totventaserviciosespeciales b IN " & cDBO & " " & _
            "WHERE a.des_IdCeco = b.tos_IdCeco " & _
            "AND   a.des_tipo_documento = b.tos_tipo_documento " & _
            "AND   a.des_numero_documento = b.tos_numero_documento " & _
            "AND   b.tos_fecha_produccion = cdate('" & fecdia & "') AND b.tos_tipo_documento IN ('DE','SE') " & _
            "AND   b.tos_IdBodega = " & vg_codbod & ""

'-------> Generar tabla b_ocsacrecibido
dbE.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
            "SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, format(a.ocr_fecoc, 'yyyymmdd') , a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN '" & cDBO & "' " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & fecdia & "')"

'-------> Generar tabla venta cafateria
'vg_db.Execute "SELECT tvc_cencos, format(tvc_fecing,'dd/mm/yyyy') AS tvc_fecing, tvc_codbod, tvc_estado INTO b_totventascaf IN '" & cDBI & "' FROM b_totventascaf "
vg_db.Execute "SELECT tvc_cencos, format(tvc_fecing,'yyyymmdd') AS tvc_fecing, tvc_codbod, tvc_estado INTO b_totventascaf IN '" & cdbi & "' FROM b_totventascaf " & _
              "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   tvc_fecing = CDATE('" & fecdia & "') " & _
              "AND   tvc_codbod = " & vg_codbod & " " & _
              "AND   tvc_estado = 'C'"

'vg_db.Execute "SELECT a.dvc_cencos, format(a.dvc_fecing,'dd/mm/yyyy') AS dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, format(a.dvc_fecdoc,'dd/mm/yyyy') AS dvc_fecdoc INTO b_detventascaf IN '" & cDBI & "' FROM b_detventascaf a, b_totventascaf b "
vg_db.Execute "SELECT a.dvc_cencos, format(a.dvc_fecing,'yyyymmdd') AS dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, format(a.dvc_fecdoc,'yyyymmdd') AS dvc_fecdoc INTO b_detventascaf IN '" & cdbi & "' FROM b_detventascaf a, b_totventascaf b " & _
              "WHERE a.dvc_cencos = b.tvc_cencos " & _
              "AND   a.dvc_fecing = b.tvc_fecing " & _
              "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   b.tvc_fecing = CDATE('" & fecdia & "') " & _
              "AND   b.tvc_codbod = " & vg_codbod & " " & _
              "AND   b.tvc_estado = 'C'"

'vg_db.Execute "SELECT a.dvp_cencos, format(a.dvp_fecing,'dd/mm/yyyy') AS dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos INTO b_detventascafpro IN '" & cDBI & "' FROM b_detventascafpro a, b_totventascaf b "
vg_db.Execute "SELECT a.dvp_cencos, format(a.dvp_fecing,'yyyymmdd') AS dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos INTO b_detventascafpro IN '" & cdbi & "' FROM b_detventascafpro a, b_totventascaf b " & _
              "WHERE a.dvp_cencos = b.tvc_cencos " & _
              "AND   a.dvp_fecing = b.tvc_fecing " & _
              "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   b.tvc_fecing = CDATE('" & fecdia & "') " & _
              "AND   b.tvc_codbod = " & vg_codbod & " " & _
              "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
vg_db.Execute "SELECT * INTO a_regimen IN '" & cdbi & "' FROM a_regimen"

'-------> Generar tabla servicio
vg_db.Execute "SELECT * INTO a_servicio IN '" & cdbi & "' FROM a_servicio"

'-------> Generar tabla estructura de servicio
vg_db.Execute "SELECT * INTO a_estservicio IN '" & cdbi & "' FROM a_estservicio"

'-------> Generar tabla sector
vg_db.Execute "SELECT * INTO a_sector IN '" & cdbi & "' FROM a_sector"

'-------> Generar tabla servicio rac
vg_db.Execute "SELECT * INTO a_serviciorac IN '" & cdbi & "' FROM a_serviciorac"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then
    vg_db.Execute "SELECT DISTINCT a.* INTO b_minuta IN '" & cdbi & "' FROM b_minuta a, b_minutadet b " & _
                  "WHERE a.min_codigo = b.mid_codigo " & _
                  "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   a.min_fecmin = " & Format(fecdia, "yyyymmdd") & ""

    vg_db.Execute "SELECT a.* INTO b_minutadet IN '" & cdbi & "' FROM b_minutadet a, b_minuta b " & _
                  "WHERE b.min_codigo = a.mid_codigo " & _
                  "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   b.min_fecmin = " & Format(fecdia, "yyyymmdd") & ""
'Else
'    vg_db.Execute "SELECT DISTINCT a.* INTO b_minuta IN '" & cDBI & "' FROM b_minuta a, b_minutadet b " & _
'                  "WHERE a.min_codigo = b.mid_codigo " & _
'                  "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'                  "AND   a.min_fecmin = " & Format(fecdia, "yyyymmdd") & " " & _
'                  "AND   b.mid_tipmin = '2'"
'
'    vg_db.Execute "SELECT a.* INTO b_minutadet IN '" & cDBI & "' FROM b_minutadet a, b_minuta b " & _
'                  "WHERE b.min_codigo = a.mid_codigo " & _
'                  "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
'                  "AND   b.min_fecmin = " & Format(fecdia, "yyyymmdd") & " " & _
'                  "AND   b.mid_tipmin = '2'"
'End If
'-------> Generar tabla minutafija
vg_db.Execute "SELECT * INTO b_minutafijadia IN '" & cdbi & "' FROM b_minutafijadia " & _
              "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   mfd_fecha  = " & Format(fecdia, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
vg_db.Execute "SELECT * INTO b_minutaraciones IN '" & cdbi & "' FROM b_minutaraciones " & _
              "WHERE mir_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   mid(mir_fecmin,1,6) = " & Format(fecdia, "yyyymm") & ""

'-------> Generar tabla precio venta
vg_db.Execute "SELECT * INTO b_preciovta IN '" & cdbi & "' FROM b_preciovta " & _
              "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
vg_db.Execute "SELECT * INTO b_ventacontado IN '" & cdbi & "' FROM b_ventacontado " & _
              "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   vtc_fecvta = " & Format(fecdia, "yyyymmdd") & ""

vg_db.Execute "SELECT a.* INTO b_ventacontadodet IN '" & cdbi & "' FROM b_ventacontadodet a, b_ventacontado b " & _
              "WHERE b.vtc_codigo = a.vtd_codigo " & _
              "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   b.vtc_fecvta = " & Format(fecdia, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
vg_db.Execute "SELECT * INTO b_clientecencos IN '" & cdbi & "' FROM b_clientecencos"

'-------> Generar tabla cliente
vg_db.Execute "SELECT * INTO b_clientes IN '" & cdbi & "' FROM b_clientes"

'-------> Generar tabla bodegas
vg_db.Execute "SELECT bod_codbod, bod_codpro, round(bod_canmer, 2) AS bod_canmer INTO b_bodegas IN '" & cdbi & "' FROM b_bodegas WHERE bod_codbod = " & vg_codbod & ""

'-------> Generar tabla precio cafeteria
vg_db.Execute "SELECT * INTO b_totpreciocaf IN '" & cdbi & "' FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

vg_db.Execute "SELECT a.* INTO b_detpreciocaf IN '" & cdbi & "' FROM b_detpreciocaf a, b_totpreciocaf b " & _
              "WHERE b.tpc_codigo = a.dpc_codigo " & _
              "AND   b.tpc_cencos = a.dpc_cencos " & _
              "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

RS1.Open "SELECT DISTINCT b.cie_fecini, b.cie_fecter FROM b_tomainv a, b_cierreperiodo b " & _
         "WHERE b.cie_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   a.tin_fectom = " & Format(CDate(fecdia) - 1, "yyyymmdd") & " " & _
         "AND   a.tin_ciemes = b.cie_periodo", vg_db, adOpenStatic
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   vg_db.Execute "SELECT * INTO b_tomainv IN '" & cdbi & "' FROM b_tomainv " & _
                 "WHERE tin_fectom >= " & RS1!cie_fecini & " AND tin_fectom <= " & RS1!cie_fecter & " " & _
                 "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
'   dbE.Execute "INSERT INTO b_totventas SELECT * FROM b_totventas IN '" & cDBO & "' "
    dbE.Execute "SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, format(tov_fecemi,'yyyymmdd') AS tov_fecemi, format(tov_fecpro,'yyyymmdd') AS tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
   '-------> Generar tabla ajuste inventario detalle
   dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
               "FROM b_detventas a, b_totventas b IN '" & cDBO & "' " & _
               "WHERE a.dev_rutcli = b.tov_rutcli " & _
               "AND   a.dev_tipdoc = b.tov_tipdoc " & _
               "AND   a.dev_numdoc = b.tov_numdoc " & _
               "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   b.tov_tipdoc IN ('AI') " & _
               "AND   b.tov_codbod = " & vg_codbod & ""
    
   '-------> Generar tabla ajuste inventario detalle impuesto
   dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
               "FROM b_detventasimp a, b_totventas b IN '" & cDBO & "' " & _
               "WHERE a.imd_rutdoc = b.tov_rutcli " & _
               "AND   a.imd_tipdoc = b.tov_tipdoc " & _
               "AND   a.imd_numdoc = b.tov_numdoc " & _
               "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   b.tov_tipdoc IN ('AI') " & _
               "AND   b.tov_codbod = " & vg_codbod & ""
Else
   '-------> Generar tabla toma inventario
   vg_db.Execute "SELECT * INTO b_tomainv IN '" & cdbi & "' FROM b_tomainv " & _
                 "WHERE tin_fectom = " & Format(fecdia, "yyyymmdd") & " " & _
                 "AND   tin_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN '" & cDBO & "' " & _
            "WHERE feccie >= cdate('" & fecdia & "') " & _
            "AND   feccie <= cdate('" & fecdia & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN '" & cDBO & "' " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN '" & cDBO & "' " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"
'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN '" & cDBO & "' " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"
'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN '" & cDBO & "' " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"
'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN '" & cDBO & "' " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & fecdia & "')"
'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
'dbe.Execute "INSERT INTO a_nomestenvio VALUES ('a_estenvio', 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
'dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_proveedor', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_tipoajuste', 0, '', 0, 0, 0)"
'-------> Permite cambiar la estructura del campo fechas a la tabla a_servicio
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horcob char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horent char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horpda char(14)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_totcompras
dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecemi char(10)"
dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecven char(10)"
dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecrem char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecdig char(24)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecemi char(10)"
dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecpro char(10)"
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

'-------> Permite cambiar la estructura del campo fechas tabla b_ocsacrecibido
'dbE.Execute "ALTER TABLE b_ocsacrecibido ALTER COLUMN ocr_fecoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_totventascaf ALTER COLUMN tvc_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecing char(10)"
dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecdoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
dbE.Execute "ALTER TABLE b_detventascafpro ALTER COLUMN dvp_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE log_cierrediario ALTER COLUMN feccie char(10)"

dbE.Close
DoEvents:

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"
fg_descarga

Exit Function
Error_ProcesoCierreAccess:
        fg_descarga
        MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
       Resume Next
End Function


Function TraerFechaCierre()

Dim RS As New ADODB.Recordset

vg_ciedia = ""
'RS.Open "SELECT DISTINCT par_nombre, par_valor FROM a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'ciediario'")

If Not RS.EOF Then
   
   vg_ciedia = fg_Desencripta(TipoDato(RS!par_valor, ""))
   Partida.StatusBar1.Panels(8).text = Trim(RS!par_nombre) & " : " & CDate(vg_ciedia) - 1

End If

RS.Close: Set RS = Nothing

If Trim(vg_ciedia) = "" Then MsgBox "No esta activo la fecha cierre día, Comunicase con departamento de informatica" & VgLinea & Space(40) & "Proceso cancelado ...", vbCritical + vbOKOnly, "Menú Principal": End

End Function

Function KillProcess(ByVal processName As String)
On Error GoTo ErrHandler
    Dim oWMI
    Dim ret
    Dim sService
    Dim oWMIServices
    Dim oWMIService
    Dim oServices
    Dim oService
    Dim servicename

    Set oWMI = GetObject("winmgmts:")
    Set oServices = oWMI.InstancesOf("win32_process")

    For Each oService In oServices
        servicename = _
            LCase(Trim(CStr(oService.Name) & ""))

        If InStr(1, servicename, _
            LCase(processName), vbTextCompare) > 0 Then
            ret = oService.Terminate
        End If
    Next

    Set oServices = Nothing
    Set oWMI = Nothing
    Exit Function
ErrHandler:
    Err.Clear
End Function

Function Cal_PMP(codcas As String, codbod As Long, codpro As String, fecemi As Date, preact As Double, canact) As Double
Dim pmp As Double, auxCanmer  As Double, auxPropon As Double
auxCanmer = 0
RS1.Open "SELECT SUM(bod.bod_canmer) AS canmer " & _
         "FROM b_productos pro, b_bodegas bod " & _
         "WHERE bod.bod_codpro = pro.pro_codigo " & _
         "AND   bod_codbod = " & codbod & " " & _
         "AND   pro.pro_codigo = '" & codpro & "'", vg_db, adOpenStatic
If Not RS1.EOF Then auxCanmer = IIf(IsNull(RS1!canmer) Or RS1!canmer = 0, 1, RS1!canmer)
RS1.Close: Set RS1 = Nothing
auxPropon = 0
RS1.Open "SELECT TOP 1 a.ppd_propon, max(a.ppd_fecdia) AS ppd_fecdia " & _
         "FROM   b_productospmpdia a, b_productos b " & _
         "WHERE  a.ppd_codpro = b.pro_codigo " & _
         "AND    b.pro_ctrsto = 1 " & _
         "AND    a.ppd_cencos = '" & codcas & "' " & _
         "AND    a.ppd_codpro = '" & codpro & "' " & _
         "AND    a.ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
         "AND    a.ppd_fecdia <= " & Format(CDate(fecemi), "yyyymmdd") & " " & _
         "GROUP BY a.ppd_propon ORDER BY max(a.ppd_fecdia) DESC", vg_db, adOpenStatic
If Not RS1.EOF Then auxPropon = IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon)
RS1.Close: Set RS1 = Nothing
If auxCanmer >= 0 Then
   pmp = Round(((auxPropon * auxCanmer) + (preact * canact)) / (auxCanmer + canact), vg_DPr)
Else
   pmp = Round(((auxPropon * (auxCanmer * -1)) + (preact * canact)) / ((auxCanmer * -1) + canact), vg_DPr)
End If
Cal_PMP = pmp
End Function

Function fg_FormatearCeldaGrillaNuemrica(Grilla As Object, Col As Long)
'-------> Editar celda Cantidad Solicitada
Grilla.Col = Col
Grilla.CellType = CellTypeNumber
Grilla.TypeNumberDecPlaces = 2
Grilla.TypeNumberMin = 1
Grilla.TypeNumberMax = 9999999
Grilla.TypeHAlign = 1
Grilla.TypeSpin = False
Grilla.TypeIntegerSpinInc = 1
Grilla.TypeIntegerSpinWrap = False
Grilla.text = Format(0, fg_Pict(6, 0))
Grilla.ForeColor = &HFF0000
End Function

Function fg_FormatearCeldaGrillaNuemricaDecimal(Grilla As Object, Col As Long)
'-------> Editar celda Cantidad Solicitada
Grilla.Col = Col
Grilla.CellType = CellTypeNumber
Grilla.TypeNumberDecPlaces = vg_RDCa
Grilla.TypeNumberMin = 1
Grilla.TypeNumberMax = 9999999
Grilla.TypeHAlign = 1
Grilla.TypeSpin = False
Grilla.TypeIntegerSpinInc = 1
Grilla.TypeIntegerSpinWrap = False
Grilla.text = Format(0, fg_Pict(6, 0))
Grilla.ForeColor = &HFF0000
End Function

Function ExtraeParentesis(ByVal text As String) As String
ExtraeParentesis = ""
Do While InStr(1, text, "(") <> 0
    text = Mid(text, 1, InStr(1, text, "(") - 1) & Mid(text, InStr(1, text, ")") + 1, Len(text) - (InStr(1, text, ")") - 1))
Loop
ExtraeParentesis = text
End Function

Function SemanaDeAño(Fecha As Long) As Integer
Dim RS As New ADODB.Recordset
Dim sql1 As String
SemanaDeAño = 0
'sql1 = IIf(vg_tipbase = "1", "CDate('" & fg_Ctod1(Fecha) & "')", " convert(datetime,'" & fg_Ctod1(Fecha) & "') ")
'RS.Open "select (case when datepart(dw," & sql1 & ") = 7 then (select DATEPART(week, (select DATEADD(day, -1, " & sql1 & ")))) Else (select DATEPART(week, " & sql1 & ")) End) ", vg_db, adOpenStatic
'If Not (RS.EOF And RS.BOF) Then
'   RS.MoveFirst
'   SemanaDeAño = RS.Fields(0)
'End If
'RS.Close
'Set RS = Nothing
SemanaDeAño = DatePart("ww", fg_Ctod1(Fecha), 2)
End Function

Function ValidarUsoOpcionesSistema(NameTemp As String) As String
'*****************---->Validar uso opciones de sistema <---------------------------
'------ Esta funcion crea una tabla temporal concatenando los parametros ingresaods
'------ para la minuta, de esta manera permanece una tabla temporal identificando
'------ que alguien se encuentra conectado a esa opcion, si alguien
'------ mas quiere acceder, se dara un aviso que esta en uso
'------ esta tabla temporal se destruye cuando se cierra este formulario (evento Unload)
'------ y tambien si el usuario cierra la sesion SQL Server la destruye automaticamente.
'----------------------------------------------------------------------
                
'Dim RSTempCheck As New ADODB.Recordset
'Dim RSTem As New ADODB.Recordset
'Dim RSinsert As New ADODB.Recordset
Dim RS As New ADODB.Recordset

ValidarUsoOpcionesSistema = "0"
Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaUsoOpcionesSistema_" & NameTemp & "'")
If RSTempCheck.EOF And RSTempCheck.BOF Then
   Set RSTem = vg_db.Execute("CREATE TABLE ##ValidaUsoOpcionesSistema_" & NameTemp & " (usu_codigo VarChar(20))")
   Set RSinsert = vg_db.Execute("INSERT INTO ##ValidaUsoOpcionesSistema_" & NameTemp & " (usu_codigo) values ('" & vg_NUsr & "')")
   ValidarUsoOpcionesSistema = "0"
Else
   Set RS = vg_db.Execute("SELECT usu_codigo from ##ValidaUsoOpcionesSistema_" & NameTemp & " ")
   If Not (RS.EOF = True And RS.BOF = True) Then
      RS.MoveFirst
      ValidarUsoOpcionesSistema = RS!usu_codigo
      RS.Close: Set RS = Nothing
      Exit Function
   End If
   RS.Close: Set RS = Nothing
End If
End Function

Function DropTeblaTmp(NameTable As String)
'*****************----> Destruye Tabla temporal<---------------------------
'---- Destruye tabla temporal, de manera que desbloquee el acceso a la minuta
Dim RSTempCheck As New ADODB.Recordset
Dim RSTem As New ADODB.Recordset

Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaUsoopcionessistema_" & NameTable & "'")
If Not (RSTempCheck.EOF = True And RSTempCheck.BOF = True) Then
   Set RSTem = vg_db.Execute("Drop Table ##ValidaUsoopcionessistema_" & NameTable & " ")
End If

RSTempCheck.Close
Set RSTempCheck = Nothing
Set RSTem = Nothing
vg_bloqueo_opciones = False
End Function

Function SacarCaracterEspecialesXml(caracteres As String) As String
Dim caracteresaux As String
caracteresaux = caracteres
caracteresaux = Replace(Trim(caracteresaux), Chr(34), "&quot;")
caracteresaux = Replace(Trim(caracteresaux), Chr(38), "&amp;")
caracteresaux = Replace(Trim(caracteresaux), Chr(39), "&apos;")
caracteresaux = Replace(Trim(caracteresaux), Chr(60), "&lt;")
caracteresaux = Replace(Trim(caracteresaux), Chr(62), "&gt;")
SacarCaracterEspecialesXml = caracteresaux

End Function

Function CalcularPMPDiaSqlPEL(Formu As Form, op As Boolean, progrl As Boolean)

On Local Error GoTo Error_CalcularPMPDiaPEL

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim aAp As String, aAp1 As String, aAp2 As String, aAp3 As String, aAp4 As String, aAp5 As String, aAp6 As String
Dim fecini As Long, fecfin As Long, i As Long, FecInv As Long, est As Boolean, fecpro As Date, fecter As Date
Dim estpro As Integer, sql1 As String
Dim fecpro1 As String, fecter1 As String, fecdiad As String

'00 Inventario inicial
'10 Ajuste de Entrada
'20 Proveedores Entrada
'30 Traspaso Entrada
'40 Produccion Salida
'50 Produccion Entrada
'60 Traspaso Salida
'70 Mermas Salida
'80 Venta Directa Salida
'90 Ajuste Salida
'100 Venta Cafeteria

Dim auxCanmer As Double, auxPropon As Double, propon As Double, auxfec As Long, fecuco As String, upreco As Double
If op Then Formu.Label1(1).Visible = True
If op Then Formu.Label1(1).Caption = "Procesando Información"
If op Then Formu.Bar1(0).Visible = True: Formu.Bar1(0).Value = 0: Formu.Bar1(0).max = 2
If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1
'fecpro = CDate(vg_ciedia)
'fecter = dEoM(CDate(vg_ciedia))

'-------> Reprocesar dia por integración PEL
Set RS = vg_db.Execute("SELECT  min(CONVERT(VARCHAR(8), lfs.Fecha, 112)) AS fecha " & _
                        "FROM    dbo.Log_FacturaSAP AS lfs " & _
                        "INNER JOIN dbo.b_clientes AS bc ON lfs.Ceco = bc.cli_codigo " & _
                        "WHERE   bc.cli_activo = '1' " & _
                        "AND CONVERT(VARCHAR(8), lfs.Fecha, 112) > ( SELECT  MAX(bt.tin_fectom) " & _
                                                    "FROM    dbo.b_tomainv AS bt " & _
                                                    "Where bt.tin_codbod = bc.cli_codbod ) " & _
                        "AND bc.cli_tipo = 0 " & _
                        "AND lfs.Estado = 3 " & _
                        "AND bc.cli_codigo = '" & MuestraCasino(1) & "'")
Dim FechaInicial As Date
Dim FechaFinal As Date
Dim FechaCierreAux As String
If Not RS.EOF Then
   If IsNull(RS!Fecha) Then
      
      fg_descarga
      RS.Close: Set RS = Nothing
      If op Then MsgBox "Proceso de cierre díario cancelado", vbInformation + vbOKOnly, MsgTitulo
      Exit Function
    
   End If
   FechaInicial = fg_Ctod1(RS!Fecha)
   FechaCierreAux = fg_Ctod1(RS!Fecha)
   FechaFinal = CDate(vg_ciedia)
   Do While FechaInicial <= FechaFinal
      
      DoEvents
        fecpro1 = FechaInicial 'CDate(vg_ciedia)
        FechaCierreAux = FechaInicial
        fecter1 = dEoM(FechaInicial) 'dEoM(CDate(vg_ciedia))
        fecdiad = FechaInicial 'Format(CDate(vg_ciedia), "dd/mm/yyyy")

        If op Then Formu.Label1(1).Visible = True
        If op Then Formu.Label1(1).Caption = "Procesando Información Día " & FechaInicial
        If op Then Formu.Bar1(0).Visible = True: Formu.Bar1(0).Value = 0: Formu.Bar1(0).max = 2
        If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1
        
        Set RS1 = vg_db.Execute("sgp_s_cierrediario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(FechaInicial, "yyyymmdd") & ", '" & fecdiad & "', '" & fecpro1 & "',  '" & fecter1 & "', " & vg_DCa & "")
        
        If RS1.EOF Then
            
            fg_descarga
            RS1.Close: Set RS1 = Nothing
            RS.Close: Set RS = Nothing
            If op Then MsgBox "Error Proceso de Cierre Día", vbInformation + vbOKOnly, MsgTitulo
            Exit Function
        
        End If
        
        If RS1!procesa = "1" Then
            
            fg_descarga
            RS1.Close: Set RS1 = Nothing
            RS.Close: Set RS = Nothing
            If op Then MsgBox "Error Proceso de Cierre Día", vbInformation + vbOKOnly, MsgTitulo
            Exit Function
        
        End If
        RS1.Close: Set RS1 = Nothing

If Not progrl Then
   '-------> Borrar tablas b_productospmpdia distinto al ultimo día cierre que tengan precio cero y saldo en cero
   vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <> " & Format(CDate(FechaCierreAux), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"
   Exit Function
End If
'-------> Grabar log_cierrediario
If FechaInicial = FechaFinal Then
   vg_db.Execute "INSERT INTO log_cierrediario VALUES ('" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "', '" & Format(FechaCierreAux, "yyyymmdd") & "', '" & Trim(vg_NUsr) & "', '1.- Cierre Diario', '" & MuestraCasino(1) & "')"
End If
'-------> Borrar tablas b_productospmpdia distinto al ultimo día cierre que tengan precio cero y saldo en cero
vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia <> " & Format(CDate(FechaCierreAux), "yyyymmdd") & " AND ppd_propon = 0 AND ppd_saldo = 0"

'-------> Crear tabla y mover datos
If op Then Formu.Label1(1).Caption = "Actua. datos anexo"
If op Then Formu.Bar1(0).Value = Formu.Bar1(0).Value + 1
estpro = 2
Dim cdbi As String, sourcefile As String, mdir As String, cDBO As String, DBO As String
'-------> Crear directorio si no existe
mdir = Dir(dir_trabajo & "\" & "EnvioWebReporting", vbDirectory)
If mdir = "" Then MkDir dir_trabajo & "\" & "EnvioWebReporting"
mdir = dir_trabajo & "EnvioWebReporting" & "\"
'-------> Fin crear directorio si no existe
    
'-------> Generar base padre
sourcefile = MuestraCasino(1) & Format(FechaCierreAux, "yyyymmdd") & ".xxx"
If Dir(mdir & sourcefile) <> "" Then Kill mdir & sourcefile ' borrar base datos si existe
If Dir(mdir & MuestraCasino(1) & Format(FechaCierreAux, "yyyymmdd") & ".mdb") <> "" Then Kill mdir & MuestraCasino(1) & Format(FechaCierreAux, "yyyymmdd") & ".mdb" ' borrar base datos si existe

cdbi = dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "xxx"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
'-------> Generar archivo mdb
'Set dbE = DBEngine(0).CreateDatabase(cDBI, dbLangGeneral)
: On Error Resume Next: Set dbE = DBEngine(0).CreateDatabase(dir_trabajo & "EnvioWebReporting\" & sourcefile, dbLangGeneral)
Set dbE = New ADODB.Connection
dbE.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbE.ConnectionTimeout = 3600
dbE.CommandTimeout = 3600
dbE.Open

'-------> Ini : Generar tabla Costo minuta & Food Cost
dbE.Execute "CREATE TABLE B_CostoMinutaRealizadoFoodCostN (IdCeco char(10), Fecha_Minuta int, IdRegimen int, IdServicio int, Raciones_Teorica int, " & _
            "Costo_Teorico_Alim float, Costo_Teorico_Desec float, Raciones_Real int, Costo_Real_Alim float, Costo_Real_Desec float, Raciones_Vendidas int, " & _
            "Costo_Realizado_Alim float, Costo_Realizado_Desec float, Venta_Dia float, Venta_Contado float, Glosa_Venta_Especial char(100)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraerCostoFoodCostMinutaCierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(FechaCierreAux, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_CostoMinutaRealizadoFoodCostN values (' " & RS(0) & "', " & RS(3) & ", " & RS(1) & ", " & RS(2) & ", " & RS(4) & " " & _
                  ", " & RS(5) & ", " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", " & RS(14) & ", '" & RS(15) & "')"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Costo minuta & Food Cost

'-------> Ini : Generar tabla Insumos & Food Cost A13
dbE.Execute "CREATE TABLE B_A13InsumosFCost (IdCeco char(10), Periodo int, FechaIni Int, FechaFin int, FechaCierre int, " & _
            "Glosa char(200), Alimentos Float, Lim_Desc Float, Total Float, Porcentaje Float, Id int)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_VtaCtoServInsumosFoodCostGastoA13CierreDiario '" & MuestraCasino(1) & "', " & vg_codbod & ", " & Format(CDate(FechaCierreAux), "yyyymmdd") & ", " & Format(FechaCierreAux, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_A13InsumosFCost values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & Format(CDate(FechaCierreAux), "yyyymmdd") & ", '" & RS(4) & "' " & _
                  ", " & RS(5) & ", '" & RS(6) & "', '" & RS(7) & "', " & RS(8) & ", " & RS(9) & ")"
      
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla Insumos & Food Cost A13

'-------> Ini : Generar tabla detalle mermas
dbE.Execute "CREATE TABLE B_DetalleMermasNK (IdCeco char(10), Periodo int, Fecha_Minuta int, IdRegimen int, IdServicio int, FechaCierre int, " & _
            "IdReceta int, IdEstServicio int, NumLin int, CostoRecetaAlimento float, CostoRecetaDesechable float, CantidadMerma float, CantidadRacionReal float, MermaxKilo float) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_DetalleMermasCierreDiario '" & MuestraCasino(1) & "', " & Format(FechaCierreAux, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_DetalleMermasNK values (' " & RS(0) & "', " & Format(FechaCierreAux, "yyyymm") & ", " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & Format(CDate(FechaCierreAux), "yyyymmdd") & ", " & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", " & RS(8) & ", " & RS(9) & ", " & RS(10) & ", " & RS(11) & ")"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla detalle mermas

'-------> Ini : Generar tabla mermas desconche - pan - produccion
dbE.Execute "CREATE TABLE B_MermaDesconche (IdCeco char(10), IdRegimen int, IdServicio int, Fecha_Merma int, " & _
            "Considera_Merma char(1), Merma_Desconche float, Merma_Pan float, Merma_Produccion float, Fecha_Modificacion datetime, Fecha_Creacion datetime, Usuario char(20)) "

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_MermaDesconcheCierreDiario '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymm") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_MermaDesconche values (' " & RS(0) & "', " & RS(1) & ", " & RS(2) & ", " & RS(3) & ", " & _
                  "" & RS(4) & ", '" & RS(5) & "', " & RS(6) & ", " & RS(7) & ", '" & RS(8) & "', '" & RS(9) & "', '" & RS(10) & "')"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla mermas desconche - pan - produccion

'-------> Ini : Generar tabla consumo proyectado & real
dbE.Execute "CREATE TABLE B_ConsumoProyectadoReal (IdCeco char(10), IdRegimen int, NoRegimen char(50), IdServicio int, NoServicio char(50), NumDoc int, Periodo int, Fecha int, " & _
            "Codigo_Producto char(20), Descripcion_Producto char(100), Unidad char(5), Cantidad_Teorica float, Cantidad_Planificada float, Cantidad_Realizada float, PMP float, Racion_Teorica int, Usuario_Mod_Racion_Real char(20), Racion_Real int, Fecha_Mod_Racion_Real datetime, Usuario_Salida_Produccion char(20), Racion_Salida_Produccion int, Fecha_Mod_Salida_Produccion datetime, Cantidad_Devolucion float)"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CalcularSalidaProducciónMinutaTeorica '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymmdd") & ", '1', " & vg_codbod & ", " & Format(vg_ciedia, "yyyymmdd") & "")
If Not RS.EOF Then

   Do While Not RS.EOF
      
      dbE.Execute "insert into B_ConsumoProyectadoReal values ('" & MuestraCasino(1) & "', " & RS(0) & ", '" & RS(1) & "', " & RS(2) & ", '" & RS(3) & "', " & RS(5) & "," & _
                  "" & Format(vg_ciedia, "yyyymm") & ", " & Format(vg_ciedia, "yyyymmdd") & ", '" & RS(6) & "', '" & RS(7) & "', '" & RS(8) & "', " & RS(9) & ", " & RS(10) & ", " & RS(11) & ", " & RS(12) & ", " & RS(13) & ", '" & RS(14) & "', " & RS(15) & ", '" & RS(16) & "', '" & RS(17) & "', " & RS(18) & ", '" & RS(19) & "', '" & RS(20) & "')"
   
      RS.MoveNext
      
   Loop
   
End If
RS.Close
Set RS = Nothing
'-------> Fin : Generar tabla consumo proyectado & real

'-------> Ini : Generar tabla a_tipoajuste 31/03/2022
dbE.Execute "CREATE TABLE a_tipoajuste (aju_codigo int, aju_nombre char(30), aju_tipaju int, aju_tipo char(1))"
dbE.Execute "INSERT INTO a_tipoajuste SELECT aju_codigo, aju_nombre, aju_tipaju, aju_tipo " & _
            "FROM a_tipoajuste IN " & DBO & ""
'-------> Fin : Generar tabla a_tipoajuste  31/03/2022

'-------> Generar tabla b_persona
dbE.Execute "CREATE TABLE b_persona (per_rut char(10), cli_codigo char(10), per_nombre char(100), per_codbarra char(100))"
dbE.Execute "INSERT INTO b_persona SELECT per_rut, cli_codigo, per_nombre, per_codbarra " & _
            "FROM b_persona IN " & DBO & ""
            
'-------> Generar tabla a_pto_atencion
dbE.Execute "CREATE TABLE a_pto_atencion (ate_codatencion int, ate_descripcion char(100))"
dbE.Execute "INSERT INTO a_pto_atencion SELECT ate_codatencion, ate_descripcion " & _
            "FROM a_pto_atencion IN " & DBO & ""
            
'-------> Generar tabla a_pto_lectura_vales
dbE.Execute "CREATE TABLE a_pto_lectura_vales (lec_codlecvales int, lec_nombrepc char(100), lec_ubicacion char(100), lec_activo int)"
dbE.Execute "INSERT INTO a_pto_lectura_vales SELECT lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo " & _
            "FROM a_pto_lectura_vales IN " & DBO & ""
            
'-------> Generar tabla b_detallelectura
dbE.Execute "CREATE TABLE b_detallelectura (cli_codigo char(10), cli_codigo_rutcliente char(10), reg_codigo int, ser_codigo int, ate_codatencion int, codigobarra char(100), fechahoraregistro datetime, fechahoravale datetime)"
dbE.Execute "INSERT INTO b_detallelectura SELECT cli_codigo, cli_codigo_rutcliente, reg_codigo, ser_codigo, ate_codatencion, codigobarra, fechahoraregistro, fechahoravale " & _
            "FROM b_detallelectura IN " & DBO & " " & _
            "WHERE cli_codigo = '" & MuestraCasino(1) & "' " & _
            "AND   format(fechahoravale, 'dd/mm/yyyy') = CDATE('" & FechaCierreAux & "')"
            
'-------> Generar tabla productopmpdia
dbE.Execute "CREATE TABLE b_productospmpdia (ppd_cencos char(10), ppd_codpro char(20), ppd_fecdia int, ppd_propon float, ppd_saldo float)"

'-------> Generar tabla bodega
dbE.Execute "CREATE TABLE a_bodega (bod_codigo int, bod_nombre char(25), bod_ubicac char(35))"
dbE.Execute "INSERT INTO a_bodega SELECT bod_codigo, bod_nombre, bod_ubicac FROM a_bodega a, b_clientes b IN " & DBO & " " & _
              "WHERE a.bod_codigo = b.cli_codbod " & _
              "AND   b.cli_codigo = '" & MuestraCasino(1) & "' " & _
              "AND   b.cli_tipo = 0"

'-------> Generar tabla proveedor
dbE.Execute "CREATE TABLE b_proveedor (prv_codigo char(10), prv_nombre char(50), prv_direccion char(50), prv_comuna char(15), prv_ciudad char(15), prv_fono1 char(12), prv_fono2 char(12), prv_fax char(12), prv_percon char(50), prv_giro char(50), prv_emapro char(50), prv_activo char(1), prv_fecumo datetime, prv_origen char(1))"
dbE.Execute "INSERT INTO b_proveedor SELECT prv_codigo, prv_nombre, prv_direccion, prv_comuna, prv_ciudad, prv_fono1, prv_fono2, prv_fax, prv_percon, prv_giro, prv_emapro, prv_activo, prv_fecumo, prv_origen FROM b_proveedor IN " & DBO & ""

'-------> Generar tabla compras
dbE.Execute "CREATE TABLE b_totcompras (toc_rutpro char(10), toc_tipdoc char(2), toc_numdoc double, toc_codbod int, toc_fecemi datetime, toc_fecven datetime, toc_desdoc double, toc_netdoc double, toc_exedoc double, toc_ivadoc double, toc_otrimp double, toc_totdoc double, toc_pendoc double, toc_estdoc char(1), toc_tipinf char(1), toc_numinf int, toc_docaso memo, toc_ordcom char(10), toc_fledoc double, toc_docsnc char(255), toc_envsap char(1), toc_fecdig datetime, toc_fecper int, toc_fecrem datetime)"
dbE.Execute "INSERT INTO b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " AND toc_fecrem = cdate('" & FechaCierreAux & "')"
            
'-------> Subir Solicitud y Guias Despachos cerradas encabezado
'dbE.Execute "INSERT INTO b_totcompras " & _
'            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
'            "FROM b_totcompras IN " & DBO & " " & _
'            "WHERE toc_codbod = " & vg_codbod & " " & _
'            "AND   toc_tipdoc = 'SN' " & _
'            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc)) " & _
'            "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
'            "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
'            "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(toc_docsnc) <> '' or not isnull(toc_docsnc)) " & _
            "AND   toc_rutpro not in (select toc_rutpro from b_totcompras) " & _
            "AND   toc_tipdoc not in (select toc_tipdoc from b_totcompras) " & _
            "AND   toc_numdoc not in (select toc_numdoc from b_totcompras)"

'            "union all "
'dbE.Execute "INSERT INTO b_totcompras " & _
'            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
'            "FROM b_totcompras IN " & DBO & " " & _
'            "WHERE toc_codbod = " & vg_codbod & " " & _
'            "AND   toc_tipdoc = 'GD' " & _
'            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso))"

dbE.Execute "INSERT INTO b_totcompras " & _
            "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem " & _
            "FROM b_totcompras IN " & DBO & " " & _
            "WHERE toc_codbod = " & vg_codbod & " " & _
            "AND   toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(toc_docaso) <> '' or not isnull(toc_docaso))"

dbE.Execute "SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem into b_totcompras_R FROM b_totcompras "
dbE.Execute "DELETE b_totcompras FROM b_totcompras "
dbE.Execute "insert into b_totcompras SELECT DISTINCT toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_estdoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem FROM b_totcompras_R "
dbE.Execute "DROP TABLE b_totcompras_R"

'dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc int, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "CREATE TABLE b_detcompras (dec_rutpro char(10), dec_tipdoc char(2), dec_numdoc double, dec_numlin int, dec_codmer char(20), dec_canmer double, dec_precom double, dec_pctdes double, dec_valdes double, dec_ptotal double, dec_descri memo, dec_canrec double, dec_prerec double, dec_mueinv char(1), dec_prefle double, dec_ptotrec double, dec_acepre char(1), dec_cmefac double, dec_pmefac double, dec_crefac double, dec_prefac double, dec_faccon double)"
dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = cdate('" & FechaCierreAux & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle
'dbE.Execute "INSERT INTO b_detcompras " & _
'            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
'            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.dec_rutpro = b.toc_rutpro " & _
'            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
'            "AND   a.dec_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'SN' " & _
'            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

'            "union all "
'dbE.Execute "INSERT INTO b_detcompras " & _
'            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
'            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.dec_rutpro = b.toc_rutpro " & _
'            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
'            "AND   a.dec_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'GD' " & _
'            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"

dbE.Execute "INSERT INTO b_detcompras " & _
            "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, a.dec_canmer, a.dec_precom, a.dec_pctdes, a.dec_valdes, a.dec_ptotal, a.dec_descri, a.dec_canrec, a.dec_prerec, a.dec_mueinv, a.dec_prefle, a.dec_ptotrec, a.dec_acepre, a.dec_cmefac, a.dec_pmefac, a.dec_crefac, a.dec_prefac, a.dec_faccon " & _
            "FROM b_detcompras a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.dec_rutpro = b.toc_rutpro " & _
            "AND   a.dec_tipdoc = b.toc_tipdoc " & _
            "AND   a.dec_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"

dbE.Execute "SELECT distinct dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon into b_detcompras_R FROM b_detcompras "
dbE.Execute "DELETE b_detcompras FROM b_detcompras "
dbE.Execute "insert into b_detcompras SELECT DISTINCT dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon FROM b_detcompras_R "
dbE.Execute "DROP TABLE b_detcompras_R"

'dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "CREATE TABLE b_detcomprasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detcomprasimp SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = cdate('" & FechaCierreAux & "')"

'-------> Subir Solicitud y Guias Despachos cerradas detalle impuesto
'dbE.Execute "INSERT INTO b_detcomprasimp " & _
'            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
'            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
'            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
'            "AND   a.imd_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'SN' " & _
'            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') " & _
            "AND  (trim(b.toc_docsnc) <> '' or not isnull(b.toc_docsnc))"

'dbE.Execute "INSERT INTO b_detcomprasimp " & _
'            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
'            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
'            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
'            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
'            "AND   a.imd_numdoc = b.toc_numdoc " & _
'            "AND   b.toc_codbod = " & vg_codbod & " " & _
'            "AND   b.toc_tipdoc = 'GD' " & _
'            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"

dbE.Execute "INSERT INTO b_detcomprasimp " & _
            "SELECT DISTINCT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detcomprasimp a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.toc_rutpro " & _
            "AND   a.imd_tipdoc = b.toc_tipdoc " & _
            "AND   a.imd_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD') " & _
            "AND  (trim(b.toc_docaso) <> '' or not isnull(b.toc_docaso))"

dbE.Execute "SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp into b_detcomprasimp_R FROM b_detcomprasimp"
dbE.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp "
dbE.Execute "insert into b_detcomprasimp SELECT DISTINCT imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp FROM b_detcomprasimp_R "
dbE.Execute "DROP TABLE b_detcomprasimp_R"

'-------> Generar tabla b_ocsacrecibido
'            "AND   b.toc_fecemi = CDATE('" & FechaCierreAux & "')"

dbE.Execute "CREATE TABLE b_ocsacrecibido (ocr_rutpro char(10), ocr_tipdoc char(2), ocr_numdoc int, ocr_numlin int, ocr_codprodsgp char(20), ocr_codprodsac char(20), ocr_cancom double, ocr_precom double, ocr_canrec double, ocr_fecoc datetime, ocr_canoc double, ocr_preoc double)"
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc FROM b_ocsacrecibido a, b_totcompras b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.toc_rutpro " & _
            "AND   a.ocr_tipdoc = b.toc_tipdoc " & _
            "AND   a.ocr_numdoc = b.toc_numdoc " & _
            "AND   b.toc_codbod = " & vg_codbod & " " & _
            "AND   b.toc_fecrem = CDATE('" & FechaCierreAux & "')"

'-------> Generar tabla ventas
'dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc int, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "CREATE TABLE b_totventas (tov_rutcli char(10), tov_tipdoc char(2), tov_numdoc double, tov_codbod int, tov_fecemi datetime, tov_fecpro datetime, tov_codreg int, tov_codser int, tov_desdoc double, tov_netdoc double, tov_exedoc double, tov_ivadoc double, tov_otrimp double, tov_totdoc double, tov_pendoc double, tov_estdoc char(1), tov_codcas char(10), tov_numinf int)"
dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf " & _
            "FROM b_totventas  IN " & DBO & " " & _
            "WHERE (tov_codbod = " & vg_codbod & " AND tov_fecpro = cdate('" & FechaCierreAux & "') AND tov_tipdoc IN ('DP','SP')) " & _
            "OR    (tov_codbod = " & vg_codbod & " AND tov_fecemi = cdate('" & FechaCierreAux & "') AND tov_tipdoc IN ('FA','GD','ME','TR'))"

'-------> Crear estructura tabla detalle ventas
'dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc int, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "CREATE TABLE b_detventas (dev_rutcli char(10), dev_tipdoc char(2), dev_numdoc double, dev_numlin int, dev_coding char(20), dev_codmer char(20), dev_canmin double, dev_canmer double, dev_porcen double, dev_precos double, dev_predoc double, dev_ptotal double, dev_descri char(250), dev_mueinv char(1), dev_codsec int, dev_acepre char(1))"
dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & FechaCierreAux & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
            "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
            "WHERE a.dev_rutcli = b.tov_rutcli " & _
            "AND   a.dev_tipdoc = b.tov_tipdoc " & _
            "AND   a.dev_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & FechaCierreAux & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Crear estructura tabla detalle ventas
'dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc int, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "CREATE TABLE b_detventasimp (imd_rutdoc char(10), imd_tipdoc char(2), imd_numdoc double, imd_numlin int, imd_codpro char(20), imd_codimp int, imd_pctimp double, imd_monimp double)"
dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecpro = cdate('" & FechaCierreAux & "') AND b.tov_tipdoc IN ('DP','SP') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
            "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
            "WHERE a.imd_rutdoc = b.tov_rutcli " & _
            "AND   a.imd_tipdoc = b.tov_tipdoc " & _
            "AND   a.imd_numdoc = b.tov_numdoc " & _
            "AND   b.tov_fecemi = cdate('" & FechaCierreAux & "') AND b.tov_tipdoc IN ('FA','GD','ME','TR') " & _
            "AND   b.tov_codbod = " & vg_codbod & ""

'-------> Generar tabla ventas servicios especiales
dbE.Execute "CREATE TABLE b_totventaserviciosespeciales (tos_IdCeco char(10), tos_Tipo_Documento char(2), tos_Numero_Documento int, tos_Fecha_Produccion datetime, tos_IdBodega int, tos_Venta_servicio_Especiales char(100), Tos_Comensales double, tos_Precio_Servicio double, tos_Total_Documento double, tos_Estado_Documento char(1), tos_Fecha_Creacion datetime, tos_Fecha_Modificacion datetime, tos_Periodo int, tos_Usuario char(20), tos_Documento_Asociado int)"
dbE.Execute "INSERT INTO b_totventaserviciosespeciales SELECT tos_IdCeco, tos_Tipo_Documento, tos_Numero_Documento, tos_Fecha_Produccion, tos_IdBodega, tos_venta_Servicio_Especiales, Tos_Comensales, tos_Precio_servicio, tos_Total_Documento, tos_Estado_Documento, tos_Fecha_Creacion, tos_Fecha_modificacion, tos_Periodo, tos_Usuario, tos_Documento_Asociado " & _
            "FROM b_totventaserviciosespeciales  IN " & DBO & " " & _
            "WHERE tos_IdBodega = " & vg_codbod & " AND tos_fecha_Produccion = cdate('" & FechaCierreAux & "') AND tos_tipo_documento IN ('DE','SE') "

'-------> Crear estructura tabla detalle ventas servicios especiales
dbE.Execute "CREATE TABLE b_detventaserviciosespeciales (des_IdCeco char(10), des_Tipo_Documento char(2), des_Numero_Documento int, des_Numero_Linea int, des_IdProducto char(20), des_Cantidad_Mercaderia double, des_Cantidad_Devolver double, des_Precio_Documento double, des_Total_Documento double, des_Descripcion char(255), des_Mueve_Inventario char(1), des_Actualiza_Precio char(1))"
dbE.Execute "INSERT INTO b_detventaserviciosespeciales SELECT a.des_IdCeco, a.des_tipo_documento, a.des_numero_documento, a.des_numero_linea, a.des_IdProducto, a.des_cantidad_mercaderia, a.des_cantidad_devolver, a.des_Precio_Documento, a.des_Total_documento, a.des_Descripcion, a.des_Mueve_Inventario, a.des_Actualiza_Precio " & _
            "FROM b_detventaserviciosespeciales a, b_totventaserviciosespeciales b IN " & DBO & " " & _
            "WHERE a.des_IdCeco = b.tos_IdCeco " & _
            "AND   a.des_tipo_documento = b.tos_tipo_documento " & _
            "AND   a.des_numero_documento = b.tos_numero_documento " & _
            "AND   b.tos_fecha_produccion = cdate('" & FechaCierreAux & "') AND b.tos_tipo_documento IN ('DE','SE') " & _
            "AND   b.tos_IdBodega = " & vg_codbod & ""


'-------> Generar tabla b_ocsacrecibido
dbE.Execute "INSERT INTO b_ocsacrecibido SELECT a.ocr_rutpro, a.ocr_tipdoc, a.ocr_numdoc, a.ocr_numlin, a.ocr_codprodsgp, a.ocr_codprodsac, a.ocr_cancom, a.ocr_precom, a.ocr_canrec, a.ocr_fecoc, a.ocr_canoc, a.ocr_preoc " & _
            "FROM b_ocsacrecibido a, b_totventas b IN " & DBO & " " & _
            "WHERE a.ocr_rutpro = b.tov_rutcli " & _
            "AND   a.ocr_tipdoc = b.tov_tipdoc " & _
            "AND   a.ocr_numdoc = b.tov_numdoc " & _
            "AND   b.tov_codbod = " & vg_codbod & " " & _
            "AND   b.tov_fecemi = CDATE('" & FechaCierreAux & "')"

dbE.Execute "CREATE TABLE b_totventascaf (tvc_cencos char(10), tvc_fecing datetime, tvc_codbod int, tvc_estado char(1))"
dbE.Execute "INSERT INTO b_totventascaf SELECT tvc_cencos, tvc_fecing, tvc_codbod, tvc_estado FROM b_totventascaf IN " & DBO & " " & _
            "WHERE tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   tvc_fecing = cdate('" & FechaCierreAux & "') " & _
            "AND   tvc_codbod = " & vg_codbod & " " & _
            "AND   tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascaf (dvc_cencos char(10), dvc_fecing datetime, dvc_numlin int, dvc_articulo char(20), dvc_canart double, dvc_precio double, dvc_tippag char(2), dvc_rutcli char(10), dvc_cencli char(10), dvc_tipdoc char(2), dvc_numdoc int, dvc_fecdoc datetime)"
dbE.Execute "INSERT INTO b_detventascaf SELECT a.dvc_cencos, a.dvc_fecing, a.dvc_numlin, a.dvc_articulo, a.dvc_canart, a.dvc_precio, a.dvc_tippag, a.dvc_rutcli, a.dvc_cencli, a.dvc_tipdoc, a.dvc_numdoc, a.dvc_fecdoc FROM b_detventascaf a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvc_cencos = b.tvc_cencos " & _
            "AND   a.dvc_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & FechaCierreAux & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

dbE.Execute "CREATE TABLE b_detventascafpro (dvp_cencos char(10), dvp_fecing datetime, dvp_codmer char(20), dvp_cancal double, dvp_candig double, dvp_precos double)"
dbE.Execute "INSERT INTO b_detventascafpro SELECT a.dvp_cencos, a.dvp_fecing, a.dvp_codmer, a.dvp_cancal, a.dvp_candig, a.dvp_precos FROM b_detventascafpro a, b_totventascaf b IN " & DBO & " " & _
            "WHERE a.dvp_cencos = b.tvc_cencos " & _
            "AND   a.dvp_fecing = b.tvc_fecing " & _
            "AND   b.tvc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.tvc_fecing = cdate('" & FechaCierreAux & "') " & _
            "AND   b.tvc_codbod = " & vg_codbod & " " & _
            "AND   b.tvc_estado = 'C'"

'-------> Generar tabla regimen
dbE.Execute "CREATE TABLE a_regimen (reg_codigo int, reg_nombre char(50))"
dbE.Execute "INSERT INTO a_regimen SELECT reg_codigo, reg_nombre FROM a_regimen IN " & DBO & ""
'-------> Generar tabla servicio
dbE.Execute "CREATE TABLE a_servicio (ser_codigo int, ser_nombre char(30), ser_orden int, ser_codsap char(20), ser_facturable char(1), ser_activo char(1), ser_horcob datetime, ser_horent datetime, ser_horpda datetime)"
dbE.Execute "INSERT INTO a_servicio SELECT ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_activo, ser_horcob, ser_horent, ser_horpda FROM a_servicio IN " & DBO & ""

'-------> Generar tabla estructura de servicio
dbE.Execute "CREATE TABLE a_estservicio (ess_codser int, ess_codigo int, ess_nombre char(30), ess_orden int, ess_codsec int, ess_racmin double, ess_cencos char(10))"
dbE.Execute "INSERT INTO a_estservicio SELECT ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin, ess_cencos FROM a_estservicio IN " & DBO & " where ess_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla sector
dbE.Execute "CREATE TABLE a_sector (sec_codigo int, sec_nombre char(50), sec_orden int)"
dbE.Execute "INSERT INTO a_sector SELECT sec_codigo, sec_nombre, sec_orden FROM a_sector IN " & DBO & ""

'-------> Generar tabla servicio rac
dbE.Execute "CREATE TABLE a_serviciorac (sra_codser int, sra_coditem int, sra_serdia int, sra_raciones int, sra_cencos char(10))"
dbE.Execute "INSERT INTO a_serviciorac SELECT sra_codser, sra_coditem, sra_serdia, sra_raciones, sra_cencos FROM a_serviciorac IN " & DBO & " where sra_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla minuta
'If Vg_MinSre = True Then
dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'-------> Si tipo de minuta es distinto simap puede generar minuta cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then

   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
               "WHERE a.min_codigo = b.mid_codigo " & _
               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
               "AND   a.min_fecmin = " & Format(FechaCierreAux, "yyyymmdd") & ""

End If

dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'-------> Si tipo de minuta es distinto simap puede generar minuta detalle cierre diario
If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
    
    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
                "WHERE b.min_codigo = a.mid_codigo " & _
                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
                "AND   b.min_fecmin = " & Format(FechaCierreAux, "yyyymmdd") & ""
End If
'Else
'   dbE.Execute "CREATE TABLE b_minuta (min_codigo int, min_cencos char(10), min_codreg int, min_codser int, min_fecmin int, min_indblo int, min_racteo int, min_racrea int)"
'   dbE.Execute "INSERT INTO b_minuta SELECT DISTINCT a.min_codigo, a.min_cencos, a.min_codreg, a.min_codser, a.min_fecmin, a.min_indblo, a.min_racteo, a.min_racrea FROM b_minuta a, b_minutadet b IN " & DBO & " " & _
'               "WHERE a.min_codigo = b.mid_codigo " & _
'               "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'               "AND   a.min_fecmin = " & Format(FechaCierreAux, "yyyymmdd") & " " & _
'               "AND   b.mid_tipmin = '2'"
'
'    dbE.Execute "CREATE TABLE b_minutadet (mid_codigo int, mid_tipmin char(1), mid_numlin int, mid_estser int, mid_codrec int, mid_numrac double, mid_descri char(50), mid_cosrec double, mid_fecval int, mid_tiprec int, mid_nummer double, mid_rec5eta char(1), mid_cosdes double)"
'    dbE.Execute "INSERT INTO b_minutadet SELECT a.mid_codigo, a.mid_tipmin, a.mid_numlin, a.mid_estser, a.mid_codrec, a.mid_numrac, a.mid_descri, a.mid_cosrec, a.mid_fecval, a.mid_tiprec, a.mid_nummer, a.mid_rec5eta, a.mid_cosdes FROM b_minutadet a, b_minuta b IN " & DBO & " " & _
'                "WHERE b.min_codigo = a.mid_codigo " & _
'                "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
'                "AND   b.min_fecmin = " & Format(FechaCierreAux, "yyyymmdd") & " " & _
'                "AND   b.mid_tipmin = '2'"
'End If
'-------> Generar tabla minutafija
dbE.Execute "CREATE TABLE b_minutafijadia (mfd_cencos char(10), mfd_codreg int, mfd_codser int, mfd_fecha int, mfd_codpro char(20), mfd_tipmin char(1), mfd_canpro double, mfd_cospro double)"
dbE.Execute "INSERT INTO b_minutafijadia SELECT mfd_cencos, mfd_codreg, mfd_codser, mfd_fecha, mfd_codpro, mfd_tipmin, mfd_canpro, mfd_cospro FROM b_minutafijadia IN " & DBO & " " & _
            "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mfd_fecha  = " & Format(FechaCierreAux, "yyyymmdd") & ""

'-------> Generar tabla minuta raciones
dbE.Execute "CREATE TABLE b_minutaraciones (mir_cencos char(10), mir_codreg int, mir_codser int, mir_fecmin int, mir_rutcli char(10), mir_nrorac int, mir_nroguia int, mir_codcli char(10))"
dbE.Execute "INSERT INTO b_minutaraciones SELECT mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli FROM b_minutaraciones IN " & DBO & " " & _
            "WHERE mir_cencos          = '" & MuestraCasino(1) & "' " & _
            "AND   mid(mir_fecmin,1,6) = " & Format(FechaCierreAux, "yyyymm") & ""

'-------> Actualizar mermas
dbE.Execute "INSERT INTO b_minutaraciones (mir_cencos, mir_codreg, mir_codser, mir_fecmin, mir_rutcli, mir_nrorac, mir_nroguia, mir_codcli) " & _
"SELECT  bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin , " & _
"        'MERMAS' as mir_rutcli , " & _
"        SUM(bm2.mid_nummer) As mermas, " & _
"        0,                              " & _
"        ''                             " & _
"FROM    b_minuta AS bm, " & _
"        b_minutadet AS bm2 IN " & DBO & " " & _
"WHERE   bm.min_codigo = bm2.mid_codigo and bm.min_cencos = '" & MuestraCasino(1) & "' " & _
"        AND mid(bm.min_fecmin, 1, 6) = " & Format(FechaCierreAux, "yyyymm") & " " & _
"        AND bm2.mid_tipmin = '2' " & _
"GROUP BY bm.min_cencos , " & _
"        bm.min_codreg , " & _
"        bm.min_codser , " & _
"        bm.min_fecmin"

'-------> Generar tabla precio venta
dbE.Execute "CREATE TABLE b_preciovta (prv_cencos char(10), prv_codreg int, prv_codser int, prv_fecvig int, prv_rutcli char(10), prv_preven double)"
dbE.Execute "INSERT INTO b_preciovta SELECT prv_cencos, prv_codreg, prv_codser, prv_fecvig, prv_rutcli, prv_preven FROM b_preciovta IN " & DBO & " " & _
            "WHERE prv_cencos = '" & MuestraCasino(1) & "'"

'-------> Generar tabla venta al contado
dbE.Execute "CREATE TABLE b_ventacontado (vtc_codigo int, vtc_cencos char(10), vtc_codreg int, vtc_codser int, vtc_fecvta int, vtc_forpag int, vtc_totmon double, vtc_opccli char(1))"
dbE.Execute "INSERT INTO b_ventacontado SELECT vtc_codigo, vtc_cencos, vtc_codreg, vtc_codser, vtc_fecvta, vtc_forpag, vtc_totmon, vtc_opccli FROM b_ventacontado IN " & DBO & " " & _
            "WHERE vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   vtc_fecvta = " & Format(FechaCierreAux, "yyyymmdd") & ""

dbE.Execute "CREATE TABLE b_ventacontadodet (vtd_codigo int, vtd_numlin int, vtd_codcli char(10), vtd_codcco char(10), vtd_descripcion char(50), vtd_detmon double)"
dbE.Execute "INSERT INTO b_ventacontadodet SELECT a.vtd_codigo, a.vtd_numlin, a.vtd_codcli, a.vtd_codcco, a.vtd_descripcion, a.vtd_detmon FROM b_ventacontadodet a, b_ventacontado b IN " & DBO & " " & _
            "WHERE b.vtc_codigo = a.vtd_codigo " & _
            "AND   b.vtc_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   b.vtc_fecvta = " & Format(FechaCierreAux, "yyyymmdd") & ""

'-------> Generar tabla centro costo cliente
dbE.Execute "CREATE TABLE b_clientecencos (clc_codigo char(10), clc_codcli char(10), clc_nombre char(50))"
dbE.Execute "INSERT INTO b_clientecencos SELECT clc_codigo, clc_codcli, clc_nombre FROM b_clientecencos IN " & DBO & " where clc_codcli = '" & MuestraCasino(1) & "'"

'-------> Generar tabla cliente
dbE.Execute "CREATE TABLE b_clientes (cli_codigo char(10), cli_nombre char(50), cli_direccion char(50), cli_comuna char(15), cli_ciudad char(15), cli_fono1 char(15), cli_fono2 char(15), cli_fax char(15), cli_percon char(50), cli_giro char(50), cli_email char(50), cli_tipo int, cli_codbod int, cli_codtis int, cli_codseg int, cli_codcli char(10), cli_clisap char(1), cli_socsap char(4), cli_cievta char(1), cli_ciedia int, cli_activo char(1), cli_sobrec char(1), cli_codmun int, cli_ccisac int, cli_cecsac char(4), cli_codreg int, id_tipo_vale char(100))"
dbE.Execute "INSERT INTO b_clientes SELECT cli_codigo, cli_nombre, cli_direccion, cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, cli_giro, cli_email, cli_tipo, cli_codbod, cli_codtis, cli_codseg, cli_codcli, cli_clisap, cli_socsap, cli_cievta, cli_ciedia, cli_activo, cli_sobrec, cli_codmun, cli_ccisac, cli_cecsac, cli_codreg FROM b_clientes IN " & DBO & ""

'-------> Generar tabla bodegas
vg_db.Execute ("delete paso_b_bodegas where bod_Cencos = '" & MuestraCasino(1) & "'")
vg_db.Execute ("insert into paso_b_bodegas (bod_cencos, bod_codbod, bod_codpro, bod_canmer) SELECT distinct '" & MuestraCasino(1) & "', bod_codbod, bod_codpro, round(bod_canmer,2) FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " ")
dbE.Execute "CREATE TABLE b_bodegas (bod_codbod int, bod_codpro char(20), bod_canmer double)"
dbE.Execute "INSERT INTO b_bodegas SELECT distinct bod_codbod, bod_codpro, bod_canmer FROM paso_b_bodegas IN " & DBO & " WHERE bod_cencos = '" & MuestraCasino(1) & "' and bod_codbod = " & vg_codbod & ""

'-------> Generar tabla precio cafeteria
dbE.Execute "CREATE TABLE b_totpreciocaf (tpc_codigo char(20), tpc_nombre char(50), tpc_precio double, tpc_cencos char(10), tpc_activo char(1))"
dbE.Execute "INSERT INTO b_totpreciocaf SELECT tpc_codigo, tpc_nombre, tpc_precio, tpc_cencos, tpc_activo FROM b_totpreciocaf  IN " & DBO & " WHERE tpc_cencos = '" & MuestraCasino(1) & "'"

dbE.Execute "CREATE TABLE b_detpreciocaf (dpc_codigo char(20), dpc_codmer char(20), dpc_cantidad double, dpc_cencos char(10))"
dbE.Execute "INSERT INTO b_detpreciocaf SELECT a.dpc_codigo, a.dpc_codmer, a.dpc_cantidad, a.dpc_cencos FROM b_detpreciocaf a, b_totpreciocaf b IN " & DBO & " " & _
            "WHERE b.tpc_codigo = a.dpc_codigo " & _
            "AND   b.tpc_cencos = a.dpc_cencos " & _
            "AND   b.tpc_cencos = '" & MuestraCasino(1) & "'"

RS1.Open "SELECT DISTINCT b.cie_fecini, b.cie_fecter FROM b_tomainv a, b_cierreperiodo b " & _
         "WHERE b.cie_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   a.tin_fectom = " & Format(CDate(FechaCierreAux) - 1, "yyyymmdd") & " " & _
         "AND   a.tin_ciemes = b.cie_periodo", vg_db, adOpenStatic
If Not RS1.EOF Then
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom >= " & RS1!cie_fecini & " AND tin_fectom <= " & RS1!cie_fecter & " " & _
               "AND   tin_codbod = " & vg_codbod & ""

   '-------> Generar tabla ajuste inventario encabezado
   dbE.Execute "INSERT INTO b_totventas SELECT tov_rutcli, tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_fecpro, tov_codreg, tov_codser, tov_desdoc, tov_netdoc, tov_exedoc, tov_ivadoc, tov_otrimp, tov_totdoc, tov_pendoc, tov_estdoc, tov_codcas, tov_numinf FROM b_totventas IN " & DBO & " " & _
               "WHERE tov_codbod = " & vg_codbod & " " & _
               "AND   tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
               "AND   tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
               "AND   tov_tipdoc IN ('AI')"
    
    '-------> Generar tabla ajuste inventario detalle
    dbE.Execute "INSERT INTO b_detventas SELECT a.dev_rutcli, a.dev_tipdoc, a.dev_numdoc, a.dev_numlin, a.dev_coding, a.dev_codmer, a.dev_canmin, a.dev_canmer, a.dev_porcen, a.dev_precos, a.dev_predoc, a.dev_ptotal, a.dev_descri, a.dev_mueinv, a.dev_codsec, a.dev_acepre " & _
                "FROM b_detventas a, b_totventas b IN " & DBO & " " & _
                "WHERE a.dev_rutcli = b.tov_rutcli " & _
                "AND   a.dev_tipdoc = b.tov_tipdoc " & _
                "AND   a.dev_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
    
    '-------> Generar tabla ajuste inventario detalle impuesto
    dbE.Execute "INSERT INTO b_detventasimp SELECT a.imd_rutdoc, a.imd_tipdoc, a.imd_numdoc, a.imd_numlin, a.imd_codpro, a.imd_codimp, a.imd_pctimp, a.imd_monimp " & _
                "FROM b_detventasimp a, b_totventas b IN " & DBO & " " & _
                "WHERE a.imd_rutdoc = b.tov_rutcli " & _
                "AND   a.imd_tipdoc = b.tov_tipdoc " & _
                "AND   a.imd_numdoc = b.tov_numdoc " & _
                "AND   b.tov_fecemi >= cdate('" & fg_Ctod1(RS1!cie_fecini) & "') " & _
                "AND   b.tov_fecemi <= cdate('" & fg_Ctod1(RS1!cie_fecter) & "') " & _
                "AND   b.tov_tipdoc IN ('AI') " & _
                "AND   b.tov_codbod = " & vg_codbod & ""
Else
   '-------> Generar tabla toma inventario
   dbE.Execute "CREATE TABLE b_tomainv (tin_fectom int, tin_codbod int, tin_codpro char(20), tin_stosis double, tin_stofis double, tin_propon double, tin_ciemes int, tin_envsap char(1), tin_autaju char(1))"
   dbE.Execute "INSERT INTO b_tomainv SELECT tin_fectom, tin_codbod, tin_codpro, tin_stosis, tin_stofis, tin_propon, tin_ciemes, tin_envsap, tin_autaju FROM b_tomainv IN " & DBO & " " & _
               "WHERE tin_fectom = " & Format(FechaCierreAux, "yyyymmdd") & " " & _
               "AND   tin_codbod = " & vg_codbod & ""
End If
RS1.Close: Set RS1 = Nothing

'-------> generar tabla log cierre diario
'vg_db.Execute "SELECT * INTO log_cierrediario IN '" & cDBI & "' FROM log_cierrediario WHERE feccie >= cdate('" & FechaCierreAux & "') AND feccie <= cdate('" & FechaCierreAux & "') + 2"
dbE.Execute "CREATE TABLE log_cierrediario (fecha datetime, feccie datetime, usuario char(20), tipocierre char(255))"
dbE.Execute "INSERT INTO log_cierrediario SELECT fecha, feccie, usuario, tipocierre FROM log_cierrediario IN " & DBO & " " & _
            "WHERE feccie >= cdate('" & FechaCierreAux & "') " & _
            "AND   feccie <= cdate('" & FechaCierreAux & "') + 2 " & _
            "AND   cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla a_param
dbE.Execute "CREATE TABLE a_param (par_codigo char(10), par_nombre char(40), par_tipo char(1), par_valor char(255), par_cencos char(10))"
dbE.Execute "INSERT INTO a_param SELECT par_codigo, par_nombre, par_tipo, par_valor, par_cencos FROM a_param IN " & DBO & " " & _
            "WHERE par_cencos = '" & MuestraCasino(1) & "'"
'-------> generar tabla sap_cfc
dbE.Execute "CREATE TABLE sap_cfc (cfc_codigo int, cfc_numlin int, cfc_nuedoc char(1), cfc_socied char(4), cfc_cladoc char(2), cfc_feccon char(8), cfc_fecdoc char(8), cfc_refere char(16), cfc_texcab char(25), cfc_mondoc char(5), cfc_clacon char(2), cfc_cueaux char(10), cfc_mtodoc char(11), cfc_asigna char(18), cfc_glosa char(50), cfc_ccosto char(10), cfc_codimp char(4), cfc_ctaimp char(10), cfc_monimp char(11), cfc_imprec char(2), cfc_otrimp char(2))"
dbE.Execute "INSERT INTO sap_cfc SELECT a.cfc_codigo, a.cfc_numlin, a.cfc_nuedoc, a.cfc_socied, a.cfc_cladoc, a.cfc_feccon, a.cfc_fecdoc, a.cfc_refere, a.cfc_texcab, a.cfc_mondoc, a.cfc_clacon, a.cfc_cueaux, a.cfc_mtodoc, a.cfc_asigna, a.cfc_glosa, a.cfc_ccosto, a.cfc_codimp, a.cfc_ctaimp, a.cfc_monimp, a.cfc_imprec, a.cfc_otrimp FROM sap_cfc a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.cfc_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '1' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FechaCierreAux & "')"
'-------> generar tabla sap_guiavta
dbE.Execute "CREATE TABLE sap_guiavta (gvt_codigo int, gvt_numlin int, gvt_pedvta char(10), gvt_codcli char(10), gvt_desmer char(10), gvt_censum char(10), gvt_fecent char(10), gvt_codmat char(18), gvt_desmat char(40), gvt_cantidad char(15), gvt_prevta char(13), gvt_tipmon char(5), gvt_glosa1 char(40), gvt_glosa2 char(40), gvt_glosa3 char(40))"
dbE.Execute "INSERT INTO sap_guiavta SELECT a.gvt_codigo, a.gvt_numlin, a.gvt_pedvta, a.gvt_codcli, a.gvt_desmer, a.gvt_censum, a.gvt_fecent, a.gvt_codmat, a.gvt_desmat, a.gvt_cantidad, a.gvt_prevta, a.gvt_tipmon, a.gvt_glosa1, a.gvt_glosa2, a.gvt_glosa3 FROM sap_guiavta a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.gvt_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND b.tipo_proceso = '4' AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FechaCierreAux & "')"
'-------> generar tabla sap_inv
dbE.Execute "CREATE TABLE sap_inv (inv_codigo int, inv_numlin int, inv_nuedoc char(1), inv_socied char(4), inv_cladoc char(2), inv_feccon char(8), inv_fecdoc char(8), inv_refere char(16), inv_texcab char(25), inv_mondoc char(5), inv_clacon char(2), inv_cueaux char(10), inv_mtodoc char(11), inv_asigna char(18), inv_glosa char(50), inv_ccosto char(10), inv_codimp char(4), inv_ctaimp char(10), inv_monimp char(11), inv_imprec char(2), inv_otrimp char(2))"
dbE.Execute "INSERT INTO sap_inv SELECT a.inv_codigo, a.inv_numlin, a.inv_nuedoc, a.inv_socied, a.inv_cladoc, a.inv_feccon, a.inv_fecdoc, a.inv_refere, a.inv_texcab, a.inv_mondoc, a.inv_clacon, a.inv_cueaux, a.inv_mtodoc, a.inv_asigna, a.inv_glosa, a.inv_ccosto, a.inv_codimp, a.inv_ctaimp, a.inv_monimp, a.inv_imprec, a.inv_otrimp FROM sap_inv a, log_procesos b IN " & DBO & " " & _
            "WHERE b.envio = a.inv_codigo AND b.cencos = '" & MuestraCasino(1) & "' AND (b.tipo_proceso = '2' OR b.tipo_proceso = '3') AND format(b.fecha, 'dd/mm/yyyy') = cdate('" & FechaCierreAux & "')"
'-------> generar tabla log_procesos
dbE.Execute "CREATE TABLE log_procesos (cencos char(10), numero int, fecha datetime, tipo_proceso char(1), rut char(10), tipo_documento char(10), num_documento char(10), num_cfc int, estado char(1), mensaje memo, envio int, anulado char(1))"
dbE.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio, anulado) SELECT '" & MuestraCasino(1) & "', numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mid(mensaje,1,500), envio, anulado FROM log_procesos IN " & DBO & " " & _
            "WHERE cencos = '" & MuestraCasino(1) & "' AND format(fecha, 'dd/mm/yyyy') = cdate('" & FechaCierreAux & "')"

'-------> generar tabla a_derechosperfil
dbE.Execute "CREATE TABLE a_derechosperfil (dpe_cecori char(10), dpe_codper int, dpe_codopc int, dpe_deracc int, dpe_deragr int, dpe_dermod int, dpe_dereli int, dpe_derimp int)"
dbE.Execute "INSERT INTO a_derechosperfil (dpe_cecori, dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp) SELECT distinct '" & MuestraCasino(1) & "', dpe_codper, dpe_codopc, dpe_deracc, dpe_deragr, dpe_dermod, dpe_dereli, dpe_derimp FROM a_derechosperfil IN " & DBO & " "

'-------> actualizar tabla a_opcsistema
Dim ciedia As String
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE a_opcsistema SET EnvioDocSGPADM = '" & ciedia & "' FROM a_opcsistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = ''")

'-------> generar tabla a_opcsistema
dbE.Execute "CREATE TABLE a_opcsistema (opc_cecori char(10), opc_codigo int, opc_nombre char(50))"
dbE.Execute "INSERT INTO a_opcsistema (opc_cecori, opc_codigo, opc_nombre) SELECT distinct '" & MuestraCasino(1) & "', opc_codigo, opc_nombre FROM a_opcsistema IN " & DBO & "  where EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> generar tabla a_perfil
dbE.Execute "CREATE TABLE a_perfil (per_cecori char(10), per_codigo int, per_nombre char(30))"
dbE.Execute "INSERT INTO a_perfil (per_cecori, per_codigo, per_nombre) SELECT distinct '" & MuestraCasino(1) & "', per_codigo, per_nombre FROM a_perfil IN " & DBO & " "

'-------> generar tabla a_usuarios
dbE.Execute "CREATE TABLE a_usuarios (usu_cecori char(10), usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios (usu_cecori, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios IN " & DBO & " "

'-------> generar tabla a_usuarios_eliminado
dbE.Execute "CREATE TABLE a_usuarios_eliminado (usu_cecori char(10), Fecha datetime, usu_codigo char(20), usu_nombre char(50), usu_password char(20), usu_perfil int, usu_telefono char(20), usu_email char(50), usu_oficina char(50), usu_depart char(50), usu_activo char(1), Fecha_Creacion datetime, Fecha_Modificacion datetime, Ticket memo)"
dbE.Execute "INSERT INTO a_usuarios_eliminado (usu_cecori, Fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket) SELECT distinct '" & MuestraCasino(1) & "', fecha, usu_codigo, usu_nombre, usu_password, usu_perfil, usu_telefono, usu_email, usu_oficina, usu_depart, usu_activo, Fecha_Creacion, Fecha_Modificacion, Ticket FROM a_usuarios_eliminado IN " & DBO & " "

'-------> generar tabla b_usuariocontratos
dbE.Execute "CREATE TABLE b_usuariocontratos (uco_cecori char(10), uco_codusu char(20), uco_codcon char(10))"
dbE.Execute "INSERT INTO b_usuariocontratos (uco_cecori, uco_codusu, uco_codcon) SELECT distinct '" & MuestraCasino(1) & "', uco_codusu, uco_codcon FROM b_usuariocontratos IN " & DBO & " "

'-------> actualizar tabla log_sistema
ciedia = ""
ciedia = Mid(vg_ciedia, 7, 4) & Mid(vg_ciedia, 4, 2) & Mid(vg_ciedia, 1, 2)
vg_db.Execute ("UPDATE Log_Sistema SET EnvioDocSGPADM = '" & ciedia & "' FROM Log_Sistema " & _
               "WHERE isnull(EnvioDocSGPADM,'') = '' and Loc_Id in (2,20,21,22,23)")
               
'-------> generar tabla Log_Sistema
dbE.Execute "CREATE TABLE Log_Sistema (cecori char(10), Fecha datetime, Usuario_Id char(20), Loc_Id int, Opcion_Sistema char(14), Dato_Nuevo memo, Dato_Anterior memo, Detalle_Operacion memo)"
dbE.Execute "INSERT INTO Log_Sistema (cecori, Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion) SELECT distinct '" & MuestraCasino(1) & "', Fecha, Usuario_Id, Loc_Id, Opcion_Sistema, Dato_Nuevo, Dato_Anterior, Detalle_Operacion FROM Log_Sistema IN " & DBO & " where Loc_Id in (2,20,21,22,23) and EnvioDocSGPADM = cdate('" & vg_ciedia & "') "

'-------> Crear estructura tabla estado envio
dbE.Execute "CREATE TABLE a_nomestenvio (ese_nomtab char(255), ese_estenv int, ese_observ char(255), ese_fecini date, ese_fecter date, ese_canreg int)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_nomestenvio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_persona', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_atencion', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_pto_lectura_vales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detallelectura', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_bodega', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_estservicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_regimen', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_sector', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_servicio', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_serviciorac', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_param', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_bodegas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientecencos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_clientes', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detcomprasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventascafpro', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventasimp', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_detventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minuta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutadet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutafijadia', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_minutaraciones', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_preciovta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_productospmpdia', 0, '', 0, 0, 0)"
'dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_proveedor', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_tomainv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totcompras', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totpreciocaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventas', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventascaf', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ventacontadodet', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_totventaserviciosespeciales', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_cierrediario', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_ocsacrecibido', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_cfc', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_guiavta', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('sap_inv', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('log_procesos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_CostoMinutaRealizadoFoodCostN', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_A13InsumosFCost', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_DetalleMermasNK', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_MermaDesconche', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_tipoajuste', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('B_ConsumoProyectadoReal', 0, '', 0, 0, 0)"

dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_derechosperfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_opcsistema', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_perfil', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('a_usuarios_eliminado', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('b_usuariocontratos', 0, '', 0, 0, 0)"
dbE.Execute "INSERT INTO a_nomestenvio VALUES ('Log_Sistema', 0, '', 0, 0, 0)"


'-------> Permite cambiar la estructura del campo fechas a la tabla a_servicio
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horcob char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horent char(14)"
'dbE.Execute "ALTER TABLE a_servicio ALTER COLUMN ser_horpda char(14)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_proveedor
dbE.Execute "ALTER TABLE b_proveedor ALTER COLUMN prv_fecumo char(10)"

'-------> Permite cambiar la estructura del campo fechas a la tabla b_totcompras
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecemi char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecven char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecrem char(10)"
'dbE.Execute "ALTER TABLE b_totcompras ALTER COLUMN toc_fecdig char(24)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecemi char(10)"
'dbE.Execute "ALTER TABLE b_totventas ALTER COLUMN tov_fecpro char(10)"
dbE.Execute "UPDATE b_totventas SET tov_fecpro = tov_fecemi WHERE trim(tov_fecpro) = '' OR (tov_fecpro) IS NULL"

'-------> Permite cambiar la estructura del campo fechas tabla b_ocsacrecibido
'dbE.Execute "ALTER TABLE b_ocsacrecibido ALTER COLUMN ocr_fecoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_totventascaf ALTER COLUMN tvc_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecing char(10)"
'dbE.Execute "ALTER TABLE b_detventascaf ALTER COLUMN dvc_fecdoc char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE b_detventascafpro ALTER COLUMN dvp_fecing char(10)"

'-------> Permite cambiar la estructura del campo fechas
'dbE.Execute "ALTER TABLE log_cierrediario ALTER COLUMN feccie char(10)"

dbE.Close
DoEvents: Formu.Bar1.Value = Formu.Bar1.Value + 1

Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
fso.MoveFile dir_trabajo & "EnvioWebReporting\" & sourcefile, dir_trabajo & "EnvioWebReporting\" & Mid(sourcefile, 1, Len(sourcefile) - 3) & "mdb"

'-------> Actualizando cerrando periodo y abriendo proximo día
If FechaInicial = FechaFinal Then
  If op Then vg_db.Execute "UPDATE a_param SET par_valor = '" & fg_Encripta(LimpiaDato(CDate(FechaCierreAux) + 1)) & "' WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'"
  '-------> Grabar log_enviocierrediario
  If op Then
     sql1 = IIf(vg_tipbase = "1", " CDATE('FechaCierreAux') ", " '" & Format(FechaCierreAux, "yyyymmdd") & "' ")
     RS1.Open "SELECT DISTINCT fecha FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & "", vg_db, adOpenStatic
     If RS1.EOF Then
        vg_db.Execute "INSERT INTO log_enviocierrediario VALUES ('" & MuestraCasino(1) & "', " & sql1 & ", '0', '')"
     Else
        vg_db.Execute "UPDATE log_enviocierrediario SET estenv = '0', fecsub = '' WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql1 & ""
     End If
     RS1.Close: Set RS1 = Nothing
   End If
End If

'-------> Update tabla Log_FacturaSAP
vg_db.Execute ("UPDATE dbo.Log_FacturaSAP SET Estado = 1, Observacion = 'Factura ingresada Exitosamente' WHERE Ceco = '" & MuestraCasino(1) & "' and CONVERT(VARCHAR(8),Fecha,112) =  " & Format(FechaCierreAux, "yyyymmdd") & " AND Estado = 3")
FechaInicial = FechaInicial + 1
Loop
End If
RS.Close: Set RS = Nothing
'-------> Fin reprocesar dia por integración PEL

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgp_p_ReducirLog '" & dir_bkpsql & "', '" & vg_SqlBase & "'")
         
If Not RS.EOF Then
            
   If RS(0) > 0 Then
               
      MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
      Exit Function
            
   End If
         
End If
         
RS.Close: Set RS = Nothing

Formu.Bar1(0).Visible = False: Formu.Bar1(0).Value = 0
Formu.Label1(1).Visible = False
fg_descarga
'If op Then MsgBox "Proceso de Cierre Día Finalizado", vbInformation + vbOKOnly, Msgtitulo
Exit Function
Error_CalcularPMPDiaPEL:
        fg_descarga
        MsgBox "Error: " & Err.Number & " - " & Err.Description, vbCritical, MsgTitulo
'       vg_db.RollbackTrans
       Resume Next
End Function

Function ValidaMDBCierreDiario(ByVal varRuta As String) As Boolean
    
'    On Error GoTo Man_Error
    
    Dim fso As New FileSystemObject, cdbi As String
    Dim RS0 As New ADODB.Recordset
    'TRUE: Archivo Incorrecto | FALSE: Archivo Correcto
    
    ValidaMDBCierreDiario = True
        
    If Not fso.FileExists(varRuta) Then
    
       MsgBox "No Se Encuentra el Archivo de Cierre Diario...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Function
    
    End If
    cdbi = varRuta
    Set dbI = New ADODB.Connection
    dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
    dbI.ConnectionTimeout = 3600
    dbI.CommandTimeout = 3600
    dbI.Open
       
    '---> Valida que el archivo se generó correctamente.
    Set RS0 = dbI.Execute("SELECT count(*) FROM B_ConsumoProyectadoReal")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida que el archivo se generó correctamente.
    Set RS0 = dbI.Execute("SELECT TOP 1 ese_nomtab FROM a_nomestenvio")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_bodega.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_bodega")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> ini : Valida archivo se generó correctamente a_bodega 31/03/2022
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_tipoajuste")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    '---> fin : Valida archivo se generó correctamente a_bodega 31/03/2022
    
    '---> Valida archivo se generó correctamente a_estservicio.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_estservicio")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_param.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_param")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_pto_atencion.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_pto_atencion")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_pto_lectura_vales.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_pto_lectura_vales")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_regimen.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_regimen")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_sector.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_sector")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
       
       Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_servicio.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_servicio")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_serviciorac.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_serviciorac")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_bodegas.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_bodegas")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_clientecencos.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_clientecencos")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_clientes.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_clientes")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detallelectura.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detallelectura")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detcompras.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detcompras")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detcomprasimp.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detcomprasimp")
    If RS0.EOF Then
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detpreciocaf.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detpreciocaf")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detventas.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detventas")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detventaserviciosespeciales.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detventaserviciosespeciales")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detventascafpro.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detventascafpro")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_detventasimp.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_detventasimp")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_minuta.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_minuta")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_minutadet.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_minutadet")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_minutafijadia.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_minutafijadia")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_minutaraciones.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_minutaraciones")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_ocsacrecibido.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_ocsacrecibido")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_persona.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_persona")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_preciovta.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_preciovta")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_productospmpdia.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_productospmpdia")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_proveedor.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_proveedor")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_tomainv.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_tomainv")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_totcompras.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_totcompras")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_totpreciocaf.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_totpreciocaf")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_totventas.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_totventas")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_totventaserviciosespeciales.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_totventaserviciosespeciales")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_totventascaf.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_totventascaf")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_ventacontado.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_ventacontado")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_ventacontadodet.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_ventacontadodet")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente log_cierrediario.
    Set RS0 = dbI.Execute("SELECT count(*) FROM log_cierrediario")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente log_procesos.
    Set RS0 = dbI.Execute("SELECT count(*) FROM log_procesos")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente sap_cfc.
    Set RS0 = dbI.Execute("SELECT count(*) FROM sap_cfc")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente sap_guiavta.
    Set RS0 = dbI.Execute("SELECT count(*) FROM sap_guiavta")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente sap_inv.
    Set RS0 = dbI.Execute("SELECT count(*) FROM sap_inv")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente B_CostoMinutaRealizadoFoodCostN.
    Set RS0 = dbI.Execute("SELECT count(*) FROM B_CostoMinutaRealizadoFoodCostN")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente B_A13InsumosFCost.
    Set RS0 = dbI.Execute("SELECT count(*) FROM B_A13InsumosFCost")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente B_DetalleMermasNK.
    Set RS0 = dbI.Execute("SELECT count(*) FROM B_DetalleMermasNK")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente B_MermaDesconche.
    Set RS0 = dbI.Execute("SELECT count(*) FROM B_MermaDesconche")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_derechosperfil.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_derechosperfil")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_opcsistema.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_opcsistema")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_perfil.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_perfil")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_usuarios.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_usuarios")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente a_usuarios_eliminado.
    Set RS0 = dbI.Execute("SELECT count(*) FROM a_usuarios_eliminado")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente b_usuariocontratos.
    Set RS0 = dbI.Execute("SELECT count(*) FROM b_usuariocontratos")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    '---> Valida archivo se generó correctamente Log_Sistema.
    Set RS0 = dbI.Execute("SELECT count(*) FROM Log_Sistema")
    If RS0.EOF Then
        
        Let ValidaMDBCierreDiario = True
        RS0.Close: Set RS0 = Nothing
        dbI.Close: Set dbI = Nothing
        Exit Function
    
    Else
        
        Let ValidaMDBCierreDiario = False
    
    End If
    RS0.Close: Set RS0 = Nothing
    
    dbI.Close: Set dbI = Nothing
    Exit Function
    
Man_Error:
    Let ValidaMDBCierreDiario = True
    
    If Err = -2147217865 Or Err = 3265 Then
        dbI.Close: Set dbI = Nothing
       Exit Function
    End If
    
    If Err = 3034 Then Exit Function
    
    If Err.Number = -2147467259 Then
        
    Else

    End If

End Function

Function encabezadoVta(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)

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

Function fg_TraerRelacionTipoDocumento(tipdoc As String) As String

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
fg_TraerRelacionTipoDocumento = ""

Set RS = vg_db.Execute("sgp_Sel_RelacionTipoDocumento '" & tipdoc & "'")
If Not RS.EOF Then

   fg_TraerRelacionTipoDocumento = RS!tdo_IdCodigo

End If
RS.Close
Set RS = Nothing

Exit Function
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Function
