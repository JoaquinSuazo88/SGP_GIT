Attribute VB_Name = "MVI"

'MVA - BLOQUEO DE COLUMNA PONDERACIONES POR PARAMETRO - 2013-01-11
Public Function BloqueaMinuta(minuta As String) As Boolean
Dim RS As New ADODB.Recordset
Dim sql As String

sql = "select par_valor from a_param where par_codigo = '" & minuta & "'"
RS.Open sql, vg_db, adOpenStatic

If Not RS.EOF Then
    If RS!par_valor = "1" Then
        BloqueaMinuta = False
    Else
        BloqueaMinuta = True
    End If
Else
    BloqueaMinuta = True
End If

End Function
'FIN MVA - BLOQUEO DE COLUMNA PONDERACIONES POR PARAMETRO - 2013-01-11

'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18
Public Function BloqueaBotonActua(clave As String) As Boolean
Dim RS As New ADODB.Recordset
Dim sql As String

sql = " select par_valor"
sql = sql & " from a_param"
sql = sql & " where par_codigo = 'paslimbas'"
'Sql = Sql & " and par_cencos ="

RS.Open sql, vg_db, adOpenStatic

If Not RS.EOF Then
    If fg_Desencripta(RS!par_valor) = clave Then
        BloqueaBotonActua = True
    Else
        BloqueaBotonActua = False
    End If
Else
    BloqueaBotonActua = False
End If

End Function
'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18

'-------> JPAZ - MVI - 2013-02-26 Función valida Minuta bloque
Public Function ValidarMinutaBloque(cencos As String, codigoregimen As Long, codigoservicio As Long, fechaminuta As Long) As Boolean
Dim RS As New ADODB.Recordset
ValidarMinutaBloque = False
RS.Open "SELECT DISTINCT a.min_cencos FROM b_minuta a inner join b_minutadet b on a.min_codigo = b.mid_codigo WHERE a.min_cencos = '" & cencos & "' AND a.min_codreg = " & codigoregimen & " and a.min_codser = " & codigoservicio & " and a.min_fecmin = " & fechaminuta & "", vg_db, adOpenStatic
If RS.EOF And GetParametro("minsre") = "1" And GetParametro("blockmicon") = "1" Then
   ValidarMinutaBloque = True
End If
RS.Close: Set RS = Nothing
End Function

'-------> JPAZ - MVI - 2013-02-26 Función valida Minuta bloque
Public Function ValidarMinutaBloqueAnexo() As Boolean
Dim RS As New ADODB.Recordset
ValidarMinutaBloqueAnexo = False
If GetParametro("minsre") = "1" And GetParametro("blockmicon") = "1" Then
   ValidarMinutaBloqueAnexo = True
End If
End Function

'-------> JPAZ - MVI - 2013-02-26 Función valida Minuta teorica
Public Function ValidarMinutaTeorica() As Boolean
Dim RS As New ADODB.Recordset
ValidarMinutaTeorica = False
If GetParametro("minsre") = "1" And GetParametro("blockmiteo") = "1" Then
   ValidarMinutaTeorica = True
End If
End Function

'-------> JPAZ - MVI - 2013-02-27 Función valida Minuta real
Public Function ValidarMinutaReal() As Boolean
Dim RS As New ADODB.Recordset
ValidarMinutaReal = False
If GetParametro("minsre") = "1" And GetParametro("blockmirea") = "1" Then
   ValidarMinutaReal = True
End If
End Function

'-------> JPAZ - MVI - 2013-02-27 Función valida password
Public Function ValidarMinutaPassword(cencos As String, password As String) As Boolean
Dim RS As New ADODB.Recordset
ValidarMinutaPassword = False

RS.Open "select isnull(par_valor,'') as par_valor from a_param where par_codigo = 'pasminblo' and par_cencos = '" & cencos & "'", vg_db, adOpenStatic
If Not RS.EOF Then
    If fg_Desencripta(RS!par_valor) = password Then
        ValidarMinutaPassword = True
    Else
        ValidarMinutaPassword = False
    End If
Else
    ValidarMinutaPassword = False
End If
RS.Close: Set RS = Nothing
End Function


'source: http://www.recursosvisualbasic.com.ar/htm/trucos-codigofuente-visual-basic/144-exportar-ado-excel.htm
' ------------------------------------------------------------------------------------
' \\ -- Función para exportar el recordset ADO a una hoja de Excel
' ------------------------------------------------------------------------------------
Public Sub Exportar_ADO_Excel(ByVal form As form, ByVal sql As String, ByVal sOutputPathXLS As String)
      
    On Error GoTo errSub
      
    Dim cn          As New ADODB.Connection
    Dim rec         As New ADODB.Recordset
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
      
    form.Enabled = False
      
   ' -- Abrir la base
    'cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sPathDB & ";"
          
    ' -- Abrir el Recordset pasándole la cadena sql
    rec.Open sql, vg_db
      
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
      
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
      
    Excel.Visible = True: Excel.UserControl = True
    iCol = rec.Fields.count
    For iCol = 1 To rec.Fields.count
        Hoja.Cells(1, iCol).Value = rec.Fields(iCol - 1).Name
    Next
      
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rec
    Else
  
        arrData = rec.GetRows
  
        iRec = UBound(arrData, 2) + 1
          
        For iCol = 0 To rec.Fields.count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
              
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rec.Fields.count).Value = GetData(arrData)
    End If
  
    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit
  
    ' -- Cierra el recordset y la base de datos y los objetos ADO
    rec.Close
    cn.Close
      
    Set rec = Nothing
    Set cn = Nothing
    ' -- guardar el libro
    Libro.SaveAs sOutputPathXLS
    Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    Excel.Quit
    Set Excel = Nothing
      
    'Exportar_ADO_Excel = True
    form.Enabled = True
    Exit Sub
errSub:
    'MsgBox Err.Description, vbCritical, "Error"
    'Exportar_ADO_Excel = False
    form.Enabled = True
End Sub

Private Function GetData(vValue As Variant) As Variant
    Dim X As Long, y As Long, xMax As Long, yMax As Long, T As Variant
      
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
      
    ReDim T(xMax, yMax)
    For X = 0 To xMax
        For y = 0 To yMax
            T(X, y) = vValue(y, X)
        Next y
    Next X
      
    GetData = T
End Function

Public Sub Dialogo(form As form, ByRef TheFile As Archivo, FilterIndex As Integer, TipArc As String)
   On Error GoTo Cero
   form.Cd.CancelError = True            ' Establecer CancelError a True
   form.Cd.Flags = cdlOFNHideReadOnly    ' Establecer los filtros
   form.Cd.Filter = TipArc ' "Archivos Excel|*.xls*|Archivos txt|*.txt|Archivos mdb|*.mdb|Archivos prn |*.prn|Archivos Todos |*.*|"
   form.Cd.FilterIndex = FilterIndex     ' Especificar el filtro predeterminado
   form.Cd.InitDir = dir_trabajo_Inf         ' Presentar el cuadro de diálogo Abrir
   form.Cd.ShowOpen
   On Error GoTo 0
   TheFile.Filename = form.Cd.Filename
   TheFile.FileTitle = form.Cd.FileTitle
  TheFile.Success = True
 

Exit Sub
Cero:
   TheFile.Filename = ""
   TheFile.FileTitle = ""
   TheFile.Success = False
End Sub

Function TipoArchivo(opcion As Long) As String
TipoArchivo = ""
Select Case opcion
Case 1
    TipoArchivo = "Todos los archivos (*.xls)|*.xls"
End Select
End Function
