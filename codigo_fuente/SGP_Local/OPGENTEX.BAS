Attribute VB_Name = "Módulo4"
Public Function MoverDatosArcTextoCosto()
Dim ConSql As ADODB.Recordset
Dim nEstructura As String * 20, nReceta As String * 50, cReceta As String, s As String, nPlan As String, nTotal As String
Dim MiFecha$, Dia%, Mes%, año%, mes1$, FechaM$, FechaP$, DiaSem%, FinDiaMes%, Fecha$
Dim IndPos As Long, WsNumDia As Long, WsDiaIni As Long, WsDiaFin As Long, WsTipoSemana As Long, WsAuxTipoSemana As Long
Dim Sw_Dia As Integer
Dim MatrizMenuReceta(8, 100) As String
Dim MatrizMenuCosto(8, 100) As Double
Dim MatrizMenuRacio(8, 100) As Long
Dim MatrizMenuGlosa(100) As String, MatrizNomDia(7) As String, MatrizNumDia(7) As String
Dim MatrizDiasInicio(7) As Long, MatrizDiasFin(7) As Long
vg_NomA = "mestudio.lst"
M_MINU13.CD.DialogTitle = "Guarda Costo Minutas"
M_MINU13.CD.Filter = "Archivos(*.lst)|*.lst|"
M_MINU13.CD.FilterIndex = 1
M_MINU13.CD.FileName = vg_NomA
M_MINU13.CD.Flags = &H40&
M_MINU13.CD.ShowSave
If M_MINU13.CD.FileName = "" Then Exit Function
If Dir(M_MINU13.CD.FileName) <> "" Then         ' El archivo ya existe, así que pregunta si el usuario desea
                                         ' sobrescribir el archivo.
   Response = MsgBox("¿Desea sobrescribir el archivo?", vbYesNo + vbQuestion + vbDefaultButton2)
   If Response = vbNo Then Exit Function
End If
vg_NomA = M_MINU13.CD.FileName
Open vg_NomA For Output As #1
For i = 1 To 100
    MatrizMenuReceta(8, i) = "/"
    MatrizMenuGlosa(i) = ""
Next i
M_MINU13.gauge.Value = 0
M_MINU13.gauge.Value = 0
M_MINU13.Picture1.Visible = True: M_MINU13.Label1.Visible = True: M_MINU13.gauge.Visible = True
M_MINU13.Label1.Caption = "Generando Archivo Texto": M_MINU13.Picture1.Refresh
fg_carga (ss)
Set ConSql = New ADODB.Recordset
ConSql.CursorType = adOpenDynamic
ConSql.LockType = adLockOptimistic
ConSql.CursorLocation = adUseClient
ConSql.Open "sdx_s_sabana 1, 0, '" + CCasino + "', " & CSegmento & ", " & CRegimen & ", " & CPVenta & ", " & CServicio & ", " & CiFechaIni & ", " & CiFechaFin & "", vg_db, , , adCmdStoredProc
If Not ConSql.EOF Then
   ConSql.MoveLast: TotalRegistro = ConSql.RecordCount: ContRegistro = 0
   ConSql.MoveFirst
   Fecha = CStr(Mid(ConSql!Fecha_Minutas, 7, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 5, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 1, 4))
   MiFecha = Format$(Fecha, "mmm dd yyyy")
   DiaSem = Weekday(DateValue(MiFecha))
   Select Case DiaSem
    Case 1
      WsNumDia = 7
    Case 2
      WsNumDia = 1
    Case 3
      WsNumDia = 2
    Case 4
      WsNumDia = 3
    Case 5
      WsNumDia = 4
    Case 6
      WsNumDia = 5
    Case 7
      WsNumDia = 6
   End Select
   WsDiaFin = 0
   WsDiaIni = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
   WsDiaFin = (7 + WsDiaIni - WsNumDia)
   WsTipoSemana = 1: WsAuxTipoSemana = 1
   For i = 1 To 7
       MatrizDiasInicio(i) = WsDiaIni
       MatrizDiasFin(i) = WsDiaFin
       WsDiaIni = WsDiaFin
       WsDiaIni = WsDiaIni + 1
       WsDiaFin = WsDiaFin + 7
   Next i
   nEstructura = M_MINU13.Combo2(0).Text
   nReceta = M_MINU13.fpText1(0).Text
   cReceta = M_MINU13.fpText1(1).Text
   Print #1, nEstructura & ";" & nReceta & ";" & cReceta
   nEstructura = "": nReceta = "": cReceta = ""
   Print #1, nEstructura & ";" & nReceta & ";" & cReceta
   Do While Not ConSql.EOF
      ContRegistro = ContRegistro + 1
      M_MINU13.gauge.Value = Val((ContRegistro / TotalRegistro) * 100)
      For i = 1 To 7
          If Val(Mid(ConSql!Fecha_Minutas, 7, 2)) >= MatrizDiasInicio(i) And Val(Mid(ConSql!Fecha_Minutas, 7, 2)) <= MatrizDiasFin(i) Then
             WsTipoSemana = i
             Exit For
          End If
      Next i
      If WsTipoSemana <> WsAuxTipoSemana Then
         If IndPos = 1 Then Sw_Dia = 1
         If IndPos = 2 Then Sw_Dia = 2
         If IndPos = 3 Then Sw_Dia = 3
         If IndPos = 4 Then Sw_Dia = 4
         If IndPos = 5 Then Sw_Dia = 5
         If IndPos = 6 Then Sw_Dia = 6
         If IndPos = 7 Then Sw_Dia = 7
         Select Case Sw_Dia
           Case 1
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
           Case 2
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
           Case 3
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
           Case 4
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
           Case 5
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(5) & " " & MatrizNumDia(5)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
           Case 6
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(5) & " " & MatrizNumDia(5)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(6) & " " & MatrizNumDia(6)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
           Case 7
             s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
             nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
             s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(5) & " " & MatrizNumDia(5)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(6) & " " & MatrizNumDia(6)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
             nReceta = MatrizNomDia(7) & " " & MatrizNumDia(7)
             s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
         End Select
         Print #1, s
         For j = 1 To 100
             If MatrizMenuReceta(8, j) = "" Then
                Select Case Sw_Dia
                  Case 1
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = MatrizMenuGlosa(j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    nReceta = MatrizMenuReceta(1, j)
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                  Case 2
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = MatrizMenuGlosa(j)
                    nReceta = MatrizMenuReceta(1, j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(2, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                  Case 3
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = Left(MatrizMenuGlosa(j), 18)
                    nReceta = MatrizMenuReceta(1, j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(2, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(3, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                  Case 4
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = MatrizMenuGlosa(j)
                    nReceta = MatrizMenuReceta(1, j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(2, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(3, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(4, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                  Case 5
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = MatrizMenuGlosa(j)
                    nReceta = MatrizMenuReceta(1, j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(2, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(3, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(4, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(5, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(5, j) > 0 Then cReceta = Format(MatrizMenuCosto(5, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(5, j) > 0 Then nPlan = Format(MatrizMenuRacio(5, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(5, j) > 0 And MatrizMenuRacio(5, j) > 0 Then nTotal = Format(MatrizMenuRacio(5, j) * MatrizMenuCosto(5, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                  Case 6
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = MatrizMenuGlosa(j)
                    nReceta = MatrizMenuReceta(1, j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(2, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(3, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(4, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(5, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(5, j) > 0 Then cReceta = Format(MatrizMenuCosto(5, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(5, j) > 0 Then nPlan = Format(MatrizMenuRacio(5, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(5, j) > 0 And MatrizMenuRacio(5, j) > 0 Then nTotal = Format(MatrizMenuRacio(5, j) * MatrizMenuCosto(5, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(6, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(6, j) > 0 Then cReceta = Format(MatrizMenuCosto(6, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(6, j) > 0 Then nPlan = Format(MatrizMenuRacio(6, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(6, j) > 0 And MatrizMenuRacio(6, j) > 0 Then nTotal = Format(MatrizMenuRacio(6, j) * MatrizMenuCosto(6, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                  Case 7
                    s = "": cReceta = "": nPlan = "": nTotal = ""
                    nEstructura = MatrizMenuGlosa(j)
                    nReceta = MatrizMenuReceta(1, j)
                    If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
                    s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(2, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(3, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(4, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(5, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(5, j) > 0 Then cReceta = Format(MatrizMenuCosto(5, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(5, j) > 0 Then nPlan = Format(MatrizMenuRacio(5, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(5, j) > 0 And MatrizMenuRacio(5, j) > 0 Then nTotal = Format(MatrizMenuRacio(5, j) * MatrizMenuCosto(5, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(6, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(6, j) > 0 Then cReceta = Format(MatrizMenuCosto(6, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(6, j) > 0 Then nPlan = Format(MatrizMenuRacio(6, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(6, j) > 0 And MatrizMenuRacio(6, j) > 0 Then nTotal = Format(MatrizMenuRacio(6, j) * MatrizMenuCosto(6, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                    nReceta = MatrizMenuReceta(7, j)
                    cReceta = "": nPlan = "": nTotal = ""
                    If MatrizMenuCosto(7, j) > 0 Then cReceta = Format(MatrizMenuCosto(7, j), fg_Pict(5, 2))
                    If MatrizMenuRacio(7, j) > 0 Then nPlan = Format(MatrizMenuRacio(7, j), fg_Pict(5, 0))
                    If MatrizMenuCosto(7, j) > 0 And MatrizMenuRacio(7, j) > 0 Then nTotal = Format(MatrizMenuRacio(7, j) * MatrizMenuCosto(7, j), fg_Pict(5, 2))
                    s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
                End Select
                Print #1, s
             End If
         Next j
         s = ""
         Print #1, s
         WsAuxTipoSemana = WsTipoSemana
         For i = 1 To 7
             MatrizNomDia(i) = ""
             MatrizNumDia(i) = ""
             For j = 1 To 100
                 MatrizMenuReceta(i, j) = ""
                 MatrizMenuReceta(8, j) = "/"
                 MatrizMenuCosto(i, j) = 0
                 MatrizMenuRacio(i, j) = 0
                 MatrizMenuGlosa(j) = ""
             Next j
         Next i
      End If
      Fecha = CStr(Mid(ConSql!Fecha_Minutas, 7, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 5, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 1, 4))
      j = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
      MiFecha = Format$(Fecha, "mmm dd yyyy")
      DiaSem = Weekday(DateValue(MiFecha))
      Select Case DiaSem
        Case 1
          IndPos = 7
          MatrizNomDia(IndPos) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
        Case 2
          IndPos = 1
          MatrizNomDia(IndPos) = "Lunes"
          MatrizNomDia(2) = "Martes"
          MatrizNomDia(3) = "Miercoles"
          MatrizNomDia(4) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(2) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(3) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(4) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 4
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 5
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 6
        Case 3
          IndPos = 2
          MatrizNomDia(IndPos) = "Martes"
          MatrizNomDia(3) = "Miercoles"
          MatrizNomDia(4) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(3) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(4) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 4
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 5
        Case 4
          IndPos = 3
          MatrizNomDia(IndPos) = "Miercoles"
          MatrizNomDia(4) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(4) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 4
        Case 5
          IndPos = 4
          MatrizNomDia(IndPos) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
        Case 6
          IndPos = 5
          MatrizNomDia(IndPos) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
        Case 7
          IndPos = 6
          MatrizNomDia(IndPos) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
      End Select
      Select Case ConSql!Mnitm_Type
        Case 1
          MatrizMenuRacio(IndPos, ConSql!Mnitm_No) = ConSql!Mnitm_Pld_Cnt
          MatrizMenuCosto(IndPos, ConSql!Mnitm_No) = CCur(ConSql!WsNumPorcion + MatrizMenuCosto(IndPos, ConSql!Mnitm_No))
          MatrizMenuReceta(IndPos, ConSql!Mnitm_No) = Fg_SacaParentesis(Trim(ConSql!Rcpe_Desc))
          MatrizMenuReceta(8, ConSql!Mnitm_No) = ""
        Case 6
          If ConSql!Rcpe_Desc <> "" Then
             MatrizMenuReceta(IndPos, ConSql!Mnitm_No) = ""
             MatrizMenuReceta(8, ConSql!Mnitm_No) = ""
             MatrizMenuGlosa(ConSql!Mnitm_No + 1) = Trim(ConSql!Rcpe_Desc)
          End If
      End Select
      ConSql.MoveNext
   Loop
   If IndPos = 1 Then Sw_Dia = 1
   If IndPos = 2 Then Sw_Dia = 2
   If IndPos = 3 Then Sw_Dia = 3
   If IndPos = 4 Then Sw_Dia = 4
   If IndPos = 5 Then Sw_Dia = 5
   If IndPos = 6 Then Sw_Dia = 6
   If IndPos = 7 Then Sw_Dia = 7
   Select Case Sw_Dia
     Case 1
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
     Case 2
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
     Case 3
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
     Case 4
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
     Case 5
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(5) & " " & MatrizNumDia(5)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
     Case 6
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(5) & " " & MatrizNumDia(5)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(6) & " " & MatrizNumDia(6)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
     Case 7
       s = "": nEstructura = "Estructura Servicio": nPlan = "Racion": cReceta = "Costo": nTotal = "Costo Total"
       nReceta = MatrizNomDia(1) & " " & MatrizNumDia(1)
       s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(2) & " " & MatrizNumDia(2)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(3) & " " & MatrizNumDia(3)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(4) & " " & MatrizNumDia(4)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(5) & " " & MatrizNumDia(5)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(6) & " " & MatrizNumDia(6)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
       nReceta = MatrizNomDia(7) & " " & MatrizNumDia(7)
       s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
   End Select
   Print #1, s
   For j = 1 To 100
       If MatrizMenuReceta(8, j) = "" Then
          Select Case Sw_Dia
            Case 1
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = MatrizMenuGlosa(j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
            Case 2
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = MatrizMenuGlosa(j)
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(2, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
              If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
              If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
            Case 3
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = Left(MatrizMenuGlosa(j), 18)
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(2, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
              If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
              If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(3, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
              If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
              If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
            Case 4
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = MatrizMenuGlosa(j)
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(2, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
              If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
              If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(3, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
              If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
              If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(4, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
              If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
              If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
            Case 5
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = MatrizMenuGlosa(j)
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(2, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
              If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
              If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(3, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
              If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
              If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(4, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
              If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
              If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(5, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(5, j) > 0 Then cReceta = Format(MatrizMenuCosto(5, j), fg_Pict(5, 2))
              If MatrizMenuRacio(5, j) > 0 Then nPlan = Format(MatrizMenuRacio(5, j), fg_Pict(5, 0))
              If MatrizMenuCosto(5, j) > 0 And MatrizMenuRacio(5, j) > 0 Then nTotal = Format(MatrizMenuRacio(5, j) * MatrizMenuCosto(5, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
            Case 6
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = MatrizMenuGlosa(j)
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(2, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
              If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
              If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(3, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
              If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
              If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(4, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
              If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
              If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(5, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(5, j) > 0 Then cReceta = Format(MatrizMenuCosto(5, j), fg_Pict(5, 2))
              If MatrizMenuRacio(5, j) > 0 Then nPlan = Format(MatrizMenuRacio(5, j), fg_Pict(5, 0))
              If MatrizMenuCosto(5, j) > 0 And MatrizMenuRacio(5, j) > 0 Then nTotal = Format(MatrizMenuRacio(5, j) * MatrizMenuCosto(5, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(6, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(6, j) > 0 Then cReceta = Format(MatrizMenuCosto(6, j), fg_Pict(5, 2))
              If MatrizMenuRacio(6, j) > 0 Then nPlan = Format(MatrizMenuRacio(6, j), fg_Pict(5, 0))
              If MatrizMenuCosto(6, j) > 0 And MatrizMenuRacio(6, j) > 0 Then nTotal = Format(MatrizMenuRacio(6, j) * MatrizMenuCosto(6, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
            Case 7
              s = "": cReceta = "": nPlan = "": nTotal = ""
              nEstructura = MatrizMenuGlosa(j)
              nReceta = MatrizMenuReceta(1, j)
              If MatrizMenuCosto(1, j) > 0 Then cReceta = Format(MatrizMenuCosto(1, j), fg_Pict(5, 2))
              If MatrizMenuRacio(1, j) > 0 Then nPlan = Format(MatrizMenuRacio(1, j), fg_Pict(5, 0))
              If MatrizMenuCosto(1, j) > 0 And MatrizMenuRacio(1, j) > 0 Then nTotal = Format(MatrizMenuRacio(1, j) * MatrizMenuCosto(1, j), fg_Pict(5, 2))
              s = nEstructura & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(2, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(2, j) > 0 Then cReceta = Format(MatrizMenuCosto(2, j), fg_Pict(5, 2))
              If MatrizMenuRacio(2, j) > 0 Then nPlan = Format(MatrizMenuRacio(2, j), fg_Pict(5, 0))
              If MatrizMenuCosto(2, j) > 0 And MatrizMenuRacio(2, j) > 0 Then nTotal = Format(MatrizMenuRacio(2, j) * MatrizMenuCosto(2, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(3, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(3, j) > 0 Then cReceta = Format(MatrizMenuCosto(3, j), fg_Pict(5, 2))
              If MatrizMenuRacio(3, j) > 0 Then nPlan = Format(MatrizMenuRacio(3, j), fg_Pict(5, 0))
              If MatrizMenuCosto(3, j) > 0 And MatrizMenuRacio(3, j) > 0 Then nTotal = Format(MatrizMenuRacio(3, j) * MatrizMenuCosto(3, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(4, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(4, j) > 0 Then cReceta = Format(MatrizMenuCosto(4, j), fg_Pict(5, 2))
              If MatrizMenuRacio(4, j) > 0 Then nPlan = Format(MatrizMenuRacio(4, j), fg_Pict(5, 0))
              If MatrizMenuCosto(4, j) > 0 And MatrizMenuRacio(4, j) > 0 Then nTotal = Format(MatrizMenuRacio(4, j) * MatrizMenuCosto(4, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(5, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(5, j) > 0 Then cReceta = Format(MatrizMenuCosto(5, j), fg_Pict(5, 2))
              If MatrizMenuRacio(5, j) > 0 Then nPlan = Format(MatrizMenuRacio(5, j), fg_Pict(5, 0))
              If MatrizMenuCosto(5, j) > 0 And MatrizMenuRacio(5, j) > 0 Then nTotal = Format(MatrizMenuRacio(5, j) * MatrizMenuCosto(5, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(6, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(6, j) > 0 Then cReceta = Format(MatrizMenuCosto(6, j), fg_Pict(5, 2))
              If MatrizMenuRacio(6, j) > 0 Then nPlan = Format(MatrizMenuRacio(6, j), fg_Pict(5, 0))
              If MatrizMenuCosto(6, j) > 0 And MatrizMenuRacio(6, j) > 0 Then nTotal = Format(MatrizMenuRacio(6, j) * MatrizMenuCosto(6, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
              nReceta = MatrizMenuReceta(7, j)
              cReceta = "": nPlan = "": nTotal = ""
              If MatrizMenuCosto(7, j) > 0 Then cReceta = Format(MatrizMenuCosto(7, j), fg_Pict(5, 2))
              If MatrizMenuRacio(7, j) > 0 Then nPlan = Format(MatrizMenuRacio(7, j), fg_Pict(5, 0))
              If MatrizMenuCosto(7, j) > 0 And MatrizMenuRacio(7, j) > 0 Then nTotal = Format(MatrizMenuRacio(7, j) * MatrizMenuCosto(7, j), fg_Pict(5, 2))
              s = s & ";" & nReceta & ";" & nPlan & ";" & cReceta & ";" & nTotal
          End Select
          Print #1, s
       End If
   Next j
   Print #1, ""
End If
Close #1
ConSql.Close: Set ConSql = Nothing
M_MINU13.Picture1.Visible = False: M_MINU13.Label1.Visible = False: M_MINU13.gauge.Visible = False
fg_descarga
MsgBox "Generación Archivo Texto Finalizo Sin Problema", vbInformation + vbOKOnly, "Lista Menú Estudio"
End Function
Public Function MoverDatosArcTextoMenuMecano()
Dim ConSql As ADODB.Recordset, ConSql1 As ADODB.Recordset
Dim MiFecha$, Dia%, Mes%, año%, mes1$, FechaM$, FechaP$, DiaSem%, FinDiaMes%, Fecha$
Dim Titulo1 As String, s As String
Dim IndPos As Long, WsNumDia As Long, WsDiaIni As Long, WsDiaFin As Long, WsTipoSemana As Long, WsAuxTipoSemana As Long, SwSalto As Long
Dim Sw_Dia As Integer
Dim MatrizMenuReceta(8, 100) As String
Dim MatrizMenuGlosa(100) As String
Dim MatrizNomDia(7) As String
Dim MatrizDiasInicio(7) As Long
Dim MatrizDiasFin(7) As Long
Dim MatrizNumDia(7) As String
vg_NomA = "mestudio.lst"
M_MINU13.CD.DialogTitle = "Guarda Costo Minutas"
M_MINU13.CD.Filter = "Archivos(*.lst)|*.lst|"
M_MINU13.CD.FilterIndex = 1
M_MINU13.CD.FileName = vg_NomA
M_MINU13.CD.Flags = &H40&
M_MINU13.CD.ShowSave
If M_MINU13.CD.FileName = "" Then Exit Function
If Dir(M_MINU13.CD.FileName) <> "" Then         ' El archivo ya existe, así que pregunta si el usuario desea
                                         ' sobrescribir el archivo.
   Response = MsgBox("¿Desea sobrescribir el archivo?", vbYesNo + vbQuestion + vbDefaultButton2)
   If Response = vbNo Then Exit Function
End If
vg_NomA = M_MINU13.CD.FileName
Open vg_NomA For Output As #1
For i = 1 To 100
    MatrizMenuReceta(8, i) = "/"
    MatrizMenuGlosa(i) = ""
Next i
M_MINU13.gauge.Value = 0
M_MINU13.Picture1.Visible = True: M_MINU13.Label1.Visible = True: M_MINU13.gauge.Visible = True
M_MINU13.Label1.Caption = "Generando Archivo Texto": M_MINU13.Picture1.Refresh
fg_carga (ss)
Set ConSql = New ADODB.Recordset
ConSql.CursorType = adOpenDynamic
ConSql.LockType = adLockOptimistic
ConSql.CursorLocation = adUseClient
ConSql.Open "sdx_s_minutas 1, 0, '" + CCasino + "', " & CSegmento & ", " & CRegimen & ", " & CPVenta & ", " & CServicio & ", " & CiFechaIni & ", " & CiFechaFin & "", vg_db, , , adCmdStoredProc
If Not ConSql.EOF Then
   ConSql.MoveLast: TotalRegistro = ConSql.RecordCount: ContRegistro = 0
   ConSql.MoveFirst
   Fecha = CStr(Mid(ConSql!Fecha_Minutas, 7, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 5, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 1, 4))
   MiFecha = Format$(Fecha, "mmm dd yyyy")
   DiaSem = Weekday(DateValue(MiFecha))
   Select Case DiaSem
    Case 1
      WsNumDia = 7
    Case 2
      WsNumDia = 1
    Case 3
      WsNumDia = 2
    Case 4
      WsNumDia = 3
    Case 5
      WsNumDia = 4
    Case 6
      WsNumDia = 5
    Case 7
      WsNumDia = 6
   End Select
   WsDiaFin = 0
   WsDiaIni = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
   WsDiaFin = (7 + WsDiaIni - WsNumDia)
   WsTipoSemana = 1: WsAuxTipoSemana = 1
   For i = 1 To 7
       MatrizDiasInicio(i) = WsDiaIni
       MatrizDiasFin(i) = WsDiaFin
       WsDiaIni = WsDiaFin
       WsDiaIni = WsDiaIni + 1
       WsDiaFin = WsDiaFin + 7
   Next i
   Do While Not ConSql.EOF
      ContRegistro = ContRegistro + 1
      M_MINU13.gauge.Value = Val((ContRegistro / TotalRegistro) * 100)
      For i = 1 To 7
          If Val(Mid(ConSql!Fecha_Minutas, 7, 2)) >= MatrizDiasInicio(i) And Val(Mid(ConSql!Fecha_Minutas, 7, 2)) <= MatrizDiasFin(i) Then
             WsTipoSemana = i
             Exit For
          End If
      Next i
      If WsTipoSemana <> WsAuxTipoSemana Then
         If SwSalto = 0 Then
            Titulo1 = "SETS MINUTAS" & " " & UCase(M_MINU13.fpText1(1).Text) & " " & "(" & Format$(M_MINU13.fpDateTime1(0).Text, "mmm yyyy") & ")"
            Print #1, Titulo1
            SwSalto = 1
         Else
            Print #1, ""
         End If
         If IndPos = 1 Then Sw_Dia = 1
         If IndPos = 2 Then Sw_Dia = 2
         If IndPos = 3 Then Sw_Dia = 3
         If IndPos = 4 Then Sw_Dia = 4
         If IndPos = 5 Then Sw_Dia = 5
         If IndPos = 6 Then Sw_Dia = 6
         If IndPos = 7 Then Sw_Dia = 7
         Select Case Sw_Dia
           Case 1
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";"
           Case 2
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";"
           Case 3
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";"
           Case 4
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";"
           Case 5
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";" & MatrizNomDia(5) & " " & MatrizNumDia(5) & ";"
           Case 6
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";" & MatrizNomDia(5) & " " & MatrizNumDia(5) & ";" & MatrizNomDia(6) & " " & MatrizNumDia(6) & ";"
           Case 7
             s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";" & MatrizNomDia(5) & " " & MatrizNumDia(5) & ";" & MatrizNomDia(6) & " " & MatrizNumDia(6) & ";" & MatrizNomDia(7) & " " & MatrizNumDia(7) & ";"
         End Select
         Print #1, s
         For j = 1 To 100
             If MatrizMenuReceta(8, j) = "" Then
                Select Case Sw_Dia
                  Case 1
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";"
                  Case 2
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";"
                  Case 3
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";"
                  Case 4
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";"
                  Case 5
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";" & MatrizMenuReceta(5, j) & ";"
                  Case 6
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";" & MatrizMenuReceta(5, j) & ";" & MatrizMenuReceta(6, j) & ";"
                  Case 7
                    s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";" & MatrizMenuReceta(5, j) & ";" & MatrizMenuReceta(6, j) & ";" & MatrizMenuReceta(7, j) & ";"
                End Select
                Print #1, s
             End If
         Next j
         Print #1, ""
         WsAuxTipoSemana = WsTipoSemana
         For i = 1 To 7
             MatrizNomDia(i) = ""
             MatrizNumDia(i) = ""
             For j = 1 To 100
                 MatrizMenuReceta(i, j) = ""
                 MatrizMenuReceta(8, j) = "/"
                 MatrizMenuGlosa(j) = ""
             Next j
         Next i
      End If
      Fecha = CStr(Mid(ConSql!Fecha_Minutas, 7, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 5, 2)) + "/" + CStr(Mid(ConSql!Fecha_Minutas, 1, 4))
      j = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
      MiFecha = Format$(Fecha, "mmm dd yyyy")
      DiaSem = Weekday(DateValue(MiFecha))
      Select Case DiaSem
        Case 1
          IndPos = 7
          MatrizNomDia(IndPos) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
        Case 2
          IndPos = 1
          MatrizNomDia(IndPos) = "Lunes"
          MatrizNomDia(2) = "Martes"
          MatrizNomDia(3) = "Miercoles"
          MatrizNomDia(4) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(2) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(3) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(4) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 4
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 5
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 6
        Case 3
          IndPos = 2
          MatrizNomDia(IndPos) = "Martes"
          MatrizNomDia(3) = "Miercoles"
          MatrizNomDia(4) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(3) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(4) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 4
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 5
        Case 4
          IndPos = 3
          MatrizNomDia(IndPos) = "Miercoles"
          MatrizNomDia(4) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(4) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 4
        Case 5
          IndPos = 4
          MatrizNomDia(IndPos) = "Jueves"
          MatrizNomDia(5) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(5) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 3
        Case 6
          IndPos = 5
          MatrizNomDia(IndPos) = "Viernes"
          MatrizNomDia(6) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(6) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 2
        Case 7
          IndPos = 6
          MatrizNomDia(IndPos) = "Sabado"
          MatrizNomDia(7) = "Domingo"
          MatrizNumDia(IndPos) = Val(Mid(ConSql!Fecha_Minutas, 7, 2))
          MatrizNumDia(7) = Val(Mid(ConSql!Fecha_Minutas, 7, 2)) + 1
      End Select
      Set ConSql1 = New ADODB.Recordset
      ConSql1.CursorType = adOpenDynamic
      ConSql1.LockType = adLockOptimistic
      ConSql1.CursorLocation = adUseClient
      ConSql1.Open "sdx_s_minutas 2, " & ConSql!Minutas_Indice & ", '" + " " + "', 0, 0, 0, 0, 0, 0", vg_db, , , adCmdStoredProc
      If Not ConSql1.EOF Then
         Do While Not ConSql1.EOF
            Select Case ConSql1!Mnitm_Type
              Case 1
                MatrizMenuReceta(IndPos, ConSql1!Mnitm_No) = Fg_SacaParentesis(Trim(ConSql1!Rcpe_Desc))
                MatrizMenuReceta(8, ConSql1!Mnitm_No) = ""
              Case 6
                MatrizMenuReceta(IndPos, ConSql1!Mnitm_No) = ""
                MatrizMenuReceta(8, ConSql1!Mnitm_No) = ""
                MatrizMenuGlosa(ConSql1!Mnitm_No + 1) = Trim(ConSql1!Mnitm_Desc)
            End Select
            ConSql1.MoveNext
         Loop
      End If
      ConSql1.Close: Set ConSql1 = Nothing
      ConSql.MoveNext
   Loop
   If IndPos = 1 Then Sw_Dia = 1
   If IndPos = 2 Then Sw_Dia = 2
   If IndPos = 3 Then Sw_Dia = 3
   If IndPos = 4 Then Sw_Dia = 4
   If IndPos = 5 Then Sw_Dia = 5
   If IndPos = 6 Then Sw_Dia = 6
   If IndPos = 7 Then Sw_Dia = 7
   If SwSalto = 0 Then
      Titulo1 = "SETS MINUTAS" & " " & UCase(M_MINU13.fpText1(1).Text) & " " & "(" & Format$(M_MINU13.fpDateTime1(0).Text, "mmm yyyy") & ")"
      Print #1, Titulo1
      Print #1, ""
      SwSalto = 1
   Else
      Print #1, ""
   End If
   Select Case IndPos
     Case 1
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";"
     Case 2
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";"
     Case 3
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";"
     Case 4
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";"
     Case 5
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";" & MatrizNomDia(5) & " " & MatrizNumDia(5) & ";"
     Case 6
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";" & MatrizNomDia(5) & " " & MatrizNumDia(5) & ";" & MatrizNomDia(6) & " " & MatrizNumDia(6) & ";"
     Case 7
       s = "Estructura Servicio" & ";" & MatrizNomDia(1) & " " & MatrizNumDia(1) & ";" & MatrizNomDia(2) & " " & MatrizNumDia(2) & ";" & MatrizNomDia(3) & " " & MatrizNumDia(3) & ";" & MatrizNomDia(4) & " " & MatrizNumDia(4) & ";" & MatrizNomDia(5) & " " & MatrizNumDia(5) & ";" & MatrizNomDia(6) & " " & MatrizNumDia(6) & ";" & MatrizNomDia(7) & " " & MatrizNumDia(7) & ";"
   End Select
   Print #1, s
   For j = 1 To 100
       If MatrizMenuReceta(8, j) = "" Then
          Select Case IndPos
            Case 1
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";"
            Case 2
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";"
            Case 3
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";"
            Case 4
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";"
            Case 5
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";" & MatrizMenuReceta(5, j) & ";"
            Case 6
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";" & MatrizMenuReceta(5, j) & ";" & MatrizMenuReceta(6, j) & ";"
            Case 7
              s = MatrizMenuGlosa(j) & ";" & MatrizMenuReceta(1, j) & ";" & MatrizMenuReceta(2, j) & ";" & MatrizMenuReceta(3, j) & ";" & MatrizMenuReceta(4, j) & ";" & MatrizMenuReceta(5, j) & ";" & MatrizMenuReceta(6, j) & ";" & MatrizMenuReceta(7, j) & ";"
          End Select
          Print #1, s
       End If
   Next j
   Print #1, ""
   WsAuxTipoSemana = WsTipoSemana
   For i = 1 To 7
       MatrizNomDia(i) = ""
       MatrizNumDia(i) = ""
       For j = 1 To 100
           MatrizMenuReceta(i, j) = ""
           MatrizMenuReceta(8, j) = "/"
           MatrizMenuGlosa(j) = ""
       Next j
   Next i
End If
Close #1
ConSql.Close: Set ConSql = Nothing
M_MINU13.Picture1.Visible = False: M_MINU13.Label1.Visible = False: M_MINU13.gauge.Visible = False
fg_descarga
MsgBox "Generación Archivo Texto Finalizo Sin Problema", vbInformation + vbOKOnly, "Lista Menú Estudio"
End Function

