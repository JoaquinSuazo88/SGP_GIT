VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "g_Proc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Txt As String, var As Variant

' Estas variables se utilizan de manera global
' para pasar datos desde otros formularios a alguno de
' los procedimientos de esta clase, se puede borrar sin problemas
' cada vez que se llama a una procedimiento, ya que solo es,
' para ser usada en ese momento

Public Sub CellEdite(Form As Form, Title As String, Campo As String, ByRef Spread As vaSpread)
'***--- Creado  :24-12-09
'***--- Autor   :Samuel Melendez
'***--- Esta Funcion llama al formulario B_CelEdi, el cual
'***--- permite ingresar cualquier Texto, y
'***--- edita la celda del SPread enviado por parametro
    Txt = ""
    Form.Label1.Caption = Campo
    Form.Caption = Title
    Form.fpText = Spread.text
    Form.Show 1
    Spread.Col = Spread.ActiveCol
    Spread.Row = Spread.ActiveRow
    If Trim(Txt) <> "" And Trim(Txt) <> "-1e" And Trim(Txt) <> Trim(Spread.text) Then
        Spread.text = Txt
'        M_Plami2.DeshacerInto Spread.Col, Spread.Row, Spread, DeshacerUltimoCammbio + 1
'        M_Plami2.ActualizaEstructuraInferior Spread, Txt
    End If
    'CellEdite = TXT
End Sub

Public Function fg_ExtraeServicio(ByVal CodigoSer As Integer, ByVal CodEstructura As Long) As String
'Creado Samuel Melendez 28/09/09
'** Esta funcion recibe el codigo de servicio y el codigo de la estructura
'** y devuelve el nombre de la estructura si la encuentra sino
'** devuelve vacio
Dim RSServ As New ADODB.Recordset
Set RSServ = vg_db.Execute("select ess_nombre from a_estservicio where ess_codser = " & CodigoSer & " and ess_codigo = " & CodEstructura & " ")

    If Not (RSServ.EOF And RSServ.BOF) Then
        fg_ExtraeServicio = RSServ!ess_nombre
    Else
        fg_ExtraeServicio = ""
    End If

RSServ.Close: Set RS = Nothing
End Function

Public Function fg_TraerNombre(ByVal Tabla As String, ByVal CampoFiltrar As String, ByVal DatoFiltrar As Integer, ByVal ReturnCampo As String) As String
' funcion creada para traer el nombre de un codigo en una tabla
' tomamdo en cuenta que el codigo que se consulta es unico en dicha tabla
Dim RsQuery As ADODB.Recordset
    
    Set RsQuery = vg_db.Execute("select " & ReturnCampo & " from " & Tabla & " where " & CampoFiltrar & " = " & DatoFiltrar & " ")
    
    If Not (RsQuery.EOF And RsQuery.BOF) Then
        RsQuery.MoveFirst
        fg_TraerNombre = RsQuery.Fields(0)
    Else
        fg_TraerNombre = ""
    End If
    RsQuery.Close
    Set RsQuery = Nothing
End Function
