VERSION 5.00
Object = "{8996B0A4-D7BE-101B-8650-00AA003A5593}#4.0#0"; "Cfx4032.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form G_TeoRea 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costo Plan. Teórico -Plan.  Real - Realizado"
   ClientHeight    =   6600
   ClientLeft      =   2100
   ClientTop       =   2145
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   11325
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   2865
      TabIndex        =   1
      Top             =   75
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
   End
   Begin ChartfxLibCtl.ChartFX Chart1 
      Height          =   6495
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11235
      _cx             =   19817
      _cy             =   11456
      Build           =   20
      TypeMask        =   1183322113
      Axis(0).Max     =   90
      nSer            =   4
      NumSer          =   4
      _Data_          =   "G_TeoRea.frx":0000
   End
End
Attribute VB_Name = "G_TeoRea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7080
Me.Width = 11415
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
End Sub

Sub LlenarGrafico(cencos As String, codreg As String, codser As String, fecini As Long, fecfin As Long, opcion As Integer, opcosto As Integer, tipmin As String, opcons As Integer, opgraf As Boolean)
Dim titgrafico As String, sql1 As String, sql2 As String, numreg As Long, numser As Long, aAp As String, sql3 As String, sql4 As String, sql5 As String
Dim i As Long, j As Long, inddia As Long, fecesf As Long, nrorac As Long, racteo As Long, racrea As Long
Dim totdoc As Double, tdiateo As Double, tdiarea As Double, vCosFij As Double, cospis As Double, costec As Double
If opcion = 0 Then Me.Caption = "Costo Plan. Teorico - Realizado": titgrafico = "Costo Plan. Teorico - Realizado " & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
If opcion = 1 Then Me.Caption = "Costo Plan. Real - Realizado": titgrafico = "Plan. Real - Realizado " & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
If opcion = 2 Then titgrafico = "Costo Plan. Teorico - Plan. Real - Realizado " & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
'-------> Traer contrato
RS1.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(cencos)), ""), vg_db, adOpenStatic
If Not RS1.EOF Then titgrafico = titgrafico & " " & RS1!cli_nombre
RS1.Close: Set RS1 = Nothing
'-------> Traer regimen
numser = 0
RS1.Open "SELECT COUNT(*) AS nreg FROM a_servicio WHERE ser_codigo IN (" & Mid(codser, 1, Len(codser) - 1) & ")", vg_db, adOpenStatic
If Not RS1.EOF And Not IsNull(RS1!nreg) Then numser = RS1!nreg
RS1.Close: Set RS1 = Nothing
If numser = 1 Then
   RS1.Open "SELECT ser_nombre FROM a_servicio WHERE ser_codigo IN (" & Mid(codser, 1, Len(codser) - 1) & ")", vg_db, adOpenStatic
   If Not RS1.EOF Then titgrafico = titgrafico & " " & RS1!ser_nombre
   RS1.Close: Set RS1 = Nothing
   numreg = 0
   RS1.Open "SELECT COUNT(*) AS nreg FROM a_regimen WHERE reg_codigo IN (" & Mid(codreg, 1, Len(codreg) - 1) & ")", vg_db, adOpenStatic
   If Not RS1.EOF And Not IsNull(RS1!nreg) Then numreg = RS1!nreg
   RS1.Close: Set RS1 = Nothing
   If numreg = 1 Or numser = 1 Then
      '-------> traer regimen
      RS1.Open "SELECT reg_nombre FROM a_regimen WHERE reg_codigo IN (" & Mid(codreg, 1, Len(codreg) - 1) & ")", vg_db, adOpenStatic
      If Not RS1.EOF Then titgrafico = RS1!reg_nombre & " " & titgrafico
      RS1.Close: Set RS1 = Nothing
      '-------> Traer costo patron
      RS1.Open "SELECT * FROM b_costopatron WHERE cpa_cencos = '" & cencos & "' AND cpa_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") AND cpa_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") AND cpa_anomes = " & Mid(fecini, 1, 6) & "", vg_db, adOpenStatic
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            DoEvents
            If Trim(RS1!cpa_descripcion) = "PISO" Then
               cospis = IIf(IsNull(RS1!cpa_valor), 0, RS1!cpa_valor)
            ElseIf Trim(RS1!cpa_descripcion) = "TECHO" Then
               costec = IIf(IsNull(RS1!cpa_valor), 0, RS1!cpa_valor)
            End If
            RS1.MoveNext
         Loop
      End If
      RS1.Close: Set RS1 = Nothing
   Else
      titgrafico = "Todos Regimen" & " " & titgrafico
   End If
Else
   titgrafico = titgrafico & " " & "Todos Servicios"
End If
If opcosto = 0 Then
   sql1 = "c.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') "
ElseIf opcosto = 1 Then
   sql1 = "c.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "') "
ElseIf opcosto = 2 Then
   sql1 = "(c.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) "
End If
'-------> Buscar nº días
sql3 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1 ,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1 ,6)) ")
RS1.Open "SELECT DISTINCT a.min_fecmin " & _
         "FROM  b_minuta a, b_minutadet b " & _
         "WHERE a.min_codigo = b.mid_codigo " & _
         "AND   a.min_cencos = '" & cencos & "' " & _
         "AND   a.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
         "AND   a.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
         "AND   " & sql3 & " = " & Val(Mid(fecini, 1, 6)) & " " & _
         "AND   b.mid_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ") " & _
         "ORDER BY a.min_fecmin DESC ", vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
inddia = Mid(RS1!min_fecmin, 7, 2)
RS1.Close: Set RS1 = Nothing
With Chart1
    .ToolBar = True
    .ToolBarObj.Moveable = False
    .ToolBarObj(0).Visible = False    'Cargar
    .ToolBarObj(1).Visible = False    'Grabar
    .ToolBarObj(2).Visible = True     'Copiar
    .ToolBarObj(3).Visible = False    'Separador
    .ToolBarObj(4).Visible = True     'Tipo de Grafico
    .ToolBarObj(5).Visible = False    'Color
    .ToolBarObj(6).Visible = False    'Separador
    .ToolBarObj(7).Visible = False    'Grilla vertical
    .ToolBarObj(8).Visible = False    'Grilla horizontal
    .ToolBarObj(9).Visible = False     'Cuadro de Leyenda
    .ToolBarObj(10).Visible = False   'Editor de datos
    .ToolBarObj(11).Visible = False   'Propiedades del grafico
    .ToolBarObj(12).Visible = True    'Separador
    .ToolBarObj(13).Visible = True    '2D/3D
    .ToolBarObj(14).Visible = False   'Rotar
    .ToolBarObj(15).Visible = True    'Profundizar
    .ToolBarObj(16).Visible = True    'Separador
    .ToolBarObj(17).Visible = False   'Zoom
    .ToolBarObj(18).Visible = False   'Preview
    .ToolBarObj(19).Visible = True    'Imprimir
    .ToolBarObj(20).Visible = False   'Separador
    .ToolBarObj(21).Visible = False   'Barras de Herramientas
    .AllowEdit = True
    .AllowResize = True
    .AllowDrag = True
    .MenuBar = False
    .ContextMenus = False             'Menus boton derecho
    .DblClk CHART_NONECLK, 1
    .OpenDataEx COD_VALUES, IIf(opgraf, 5, 3), inddia
    .TITLE(CHART_TOPTIT) = titgrafico
    .Fonts(CHART_TOPFT) = CF_ARIAL
    .SerLegBoxObj.Visible = True
    .SerLegBoxObj.Docked = 515
    .SerLegBoxObj.BorderStyle = 3
    .SerLegBoxObj.Style = 0
    .Axis(AXIS_Y).TITLE = IIf(opgraf = True, "Costo Bandeja ($)", "Costo Totales ($)")
    .Axis(AXIS_Y).max = 15000
    .Axis(AXIS_Y).Min = 1
    .Axis(AXIS_X).TITLE = "Día " '"Titulo Eje X"
    .Axis(AXIS_X).Visible = True
    .Axis(AXIS_X).ClearLabels
    j = 1
    For i = 0 To inddia
        If j <= inddia Then .Axis(AXIS_X).KeyLabel(i) = j
        j = j + 1
    Next i
    sql2 = IIf(opcons = 0, "SUM(isnull(c.mid_cosrec,0)*c.mid_numrac) AS cosmin", IIf(opcosto = 1, "SUM(isnull(c.mid_cosdes,0)*c.mid_numrac) AS cosmin", "SUM((isnull(c.mid_cosrec,0)+isnull(c.mid_cosdes,0))*c.mid_numrac) AS cosmin"))
    RS1.Open "SELECT c.mid_tipmin, b.min_fecmin, b.min_racteo, b.min_racrea, " & sql2 & " " & _
             "FROM  b_receta a, b_minuta b, b_minutadet c " & _
             "WHERE b.min_codigo = c.mid_codigo " & _
             "AND   c.mid_codrec = a.rec_codigo " & _
             "AND   b.min_cencos = '" & cencos & "' " & _
             "AND   b.min_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
             "AND   b.min_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
             "AND   b.min_fecmin >= " & fecini & " " & _
             "AND   b.min_fecmin <= " & fecfin & " " & _
             "AND   c.mid_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ") " & _
             "GROUP BY c.mid_tipmin, b.min_fecmin, b.min_racteo, b.min_racrea ORDER BY b.min_fecmin", vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
    For i = 0 To inddia
        DoEvents
        .Series(0).Yvalue(i) = 0
        .Series(1).Yvalue(i) = 0
        .Series(2).Yvalue(i) = 0
        If opgraf Then
           .Series(3).Yvalue(i) = 0
           .Series(4).Yvalue(i) = 0
        End If
    Next i
    auxfec = 0: tdiateo = 0: tdiarea = 0
    .Series(0).Visible = IIf(opcion = 0 Or opcion = 2, True, False)
    .Series(1).Visible = IIf(opcion = 1 Or opcion = 2, True, False)
    Dim estfij As Boolean
    estfij = False
    
    '-------> Traer salida & devolución
    sql3 = IIf(vg_tipbase = "1", " SUM(IIf(a.tov_tipdoc='SP',b.dev_ptotal,'-' & b.dev_ptotal)) AS totdoc ", " SUM(CASE WHEN a.tov_tipdoc = 'SP' THEN b.dev_ptotal ELSE (-1*b.dev_ptotal) END) AS totdoc ")
    sql4 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
    sql5 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
    RS2.Open "SELECT a.tov_fecpro, a.tov_codreg, a.tov_codser, " & sql3 & " " & _
             "FROM b_totventas a, b_detventas b, b_productos c " & _
             "WHERE a.tov_rutcli = b.dev_rutcli " & _
             "AND   a.tov_tipdoc = b.dev_tipdoc " & _
             "AND   a.tov_numdoc = b.dev_numdoc " & _
             "AND   b.dev_codmer = c.pro_codigo " & _
             "AND   " & sql1 & " " & _
             "AND   a.tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") AND a.tov_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
             "AND  (a.tov_tipdoc = 'SP' OR a.tov_tipdoc = 'DP') AND b.dev_canmer <> 0 " & _
             "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' AND a.tov_fecpro >= " & sql4 & " AND a.tov_fecpro <= " & sql5 & " GROUP BY a.tov_fecpro, a.tov_codreg, a.tov_codser", vg_db, adOpenStatic
    Do While Not RS2.EOF
       DoEvents
       nrorac = 0: totdoc = 0
       If opgraf = True Then
          RS3.Open "SELECT mir_fecmin, SUM(mir_nrorac) AS nrorac FROM b_minutaraciones " & _
                   "WHERE mir_cencos='" & cencos & "' " & _
                   "AND   mir_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                   "AND   mir_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
                   "AND  (mir_rutcli = 'PRODUCIDAS') " & _
                   "AND   mir_fecmin = " & Format(RS2!tov_fecpro, "yyyymmdd") & " GROUP BY mir_fecmin", vg_db, adOpenStatic
          DoEvents
          If Not RS3.EOF And Not IsNull(RS3!nrorac) Then nrorac = RS3!nrorac
          RS3.Close: Set RS3 = Nothing
       End If
       
       If opgraf Then
          If nrorac > 0 And RS2!totdoc > 0 Then totdoc = (RS2!totdoc / nrorac) Else totdoc = 0
       Else
          totdoc = IIf(IsNull(RS2!totdoc), 0, RS2!totdoc)
       End If
       .Series(2).Yvalue(Val(Mid(RS2!tov_fecpro, 1, 2)) - 1) = Format(totdoc, fg_Pict(6, 2))
       RS2.MoveNext
    Loop
    RS2.Close: Set RS2 = Nothing
    
    '-------> Buscar datos estructura fija día
    RS2.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
            "WHERE mfd_cencos='" & cencos & "' " & _
            "AND   mfd_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
            "AND   mfd_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
            "AND   mfd_fecha >= " & fecini & " AND mfd_fecha <= " & fecfin & " AND mfd_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ")", vg_db, adOpenStatic
    If Not RS2.EOF Then estfij = True
    RS2.Close: Set RS2 = Nothing
    fecesf = 0
    If Not estfij Then
       RS2.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
                "WHERE mif_cencos = '" & cencos & "' " & _
                "AND   mif_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                "AND   mif_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ")", vg_db, adOpenStatic
       If Not RS2.EOF Then fecesf = IIf(IsNull(RS2!fecval), 0, RS2!fecval)
       RS2.Close: Set RS2 = Nothing
    End If
    If Not estfij And fecesf > 0 And vg_tipbase = "1" Then
        '-------> Insert tabla productospmpdia
        aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPLlenarGrafico"
        fg_CheckTmp aAp
        vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                      "INTO " & aAp & " " & _
                      "FROM b_productospmpdia " & _
                      "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                      "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                      "AND   ppd_propon > 0 " & _
                      "GROUP BY ppd_cencos, ppd_codpro"
        vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
        vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
        vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    End If
    totdoc = 0: racteo = 0: racrea = 0
    Do While Not RS1.EOF
       DoEvents
       If RS1!min_fecmin <> auxfec Then
          If auxfec > 0 Then
             If opgraf Then .Series(3).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(cospis, fg_Pict(6, 2))
             If opgraf Then .Series(4).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(costec, fg_Pict(6, 2))
             If opgraf Then
                '-------> Mover teórico
                .Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(0, fg_Pict(6, 2))
                If racteo > 0 Then .Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format((tdiateo / racteo), fg_Pict(6, 2))
                '-------> Mover real
                .Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(0, fg_Pict(6, 2))
                If racrea > 0 Then .Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format((tdiarea / racrea), fg_Pict(6, 2))
             Else
                '-------> Mover teórico
                .Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiateo, fg_Pict(6, 2))
                '-------> Mover real
                .Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiarea, fg_Pict(6, 2))
             End If
             totdoc = 0: racteo = 0: racrea = 0
             '-------> Mover raciones
          End If
          
          auxfec = RS1!min_fecmin
          tdiateo = 0: tdiarea = 0
          j = j + 1
       End If
       vCosFij = 0
       If estfij Then
          '-------> Calcular datos desde tabla estructura fija día
          RS3.Open "SELECT SUM(a.mfd_canpro*a.mfd_cospro) AS cosfij " & _
                   "FROM b_minutafijadia a, b_productos c " & _
                   "WHERE a.mfd_codpro = c.pro_codigo " & _
                   "AND   a.mfd_cencos = '" & cencos & "' " & _
                   "AND   a.mfd_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                   "AND   a.mfd_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ")  " & _
                   "AND   a.mfd_fecha = " & RS1!min_fecmin & " AND a.mfd_tipmin = '" & RS1!mid_tipmin & "' " & _
                   "AND   " & sql1 & "", vg_db, adOpenStatic
          DoEvents
          If Not RS3.EOF And Not IsNull(RS3!cosfij) Then vCosFij = RS3!cosfij
          RS3.Close: Set RS3 = Nothing
       ElseIf Not estfij And fecesf > 0 Then
          '-------> Calcular datos desde tabla estructura fija
          If vg_tipbase = "1" Then
                RS3.Open "SELECT SUM(b.ppd_propon*a.mif_canpro) AS cosfij " & _
                         "FROM  b_minutafija a, b_productos c, " & aAp & " b " & _
                         "WHERE a.mif_codpro = c.pro_codigo AND c.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   a.mif_cencos = '" & cencos & "' " & _
                         "AND   a.mif_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                         "AND   a.mif_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
                         "AND   a.mif_fecval=" & fecesf & " " & _
                         "AND   a.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2)) - 2))) & " " & _
                         "AND   " & sql1 & "", vg_db, adOpenStatic
          Else
                RS3.Open "SELECT SUM(b.ppd_propon*a.mif_canpro) AS cosfij " & _
                         "FROM  b_minutafija a, b_productos c, b_productospmpdia b " & _
                         "WHERE a.mif_codpro = c.pro_codigo AND c.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                         "AND   b.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
                         "AND   a.mif_cencos = '" & cencos & "' " & _
                         "AND   a.mif_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
                         "AND   a.mif_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
                         "AND   a.mif_fecval = " & fecesf & " " & _
                         "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2), Len(fg_Fecha_Dia(RS1!min_fecmin & Right("0" & i, 2), 2)) - 2))) & " " & _
                         "AND   " & sql1 & "", vg_db, adOpenStatic
          End If
          DoEvents
          If Not RS3.EOF And Not IsNull(RS3!cosfij) Then vCosFij = RS3!cosfij
          RS3.Close: Set RS3 = Nothing
       End If
       If RS1!mid_tipmin = "1" Then
          tdiateo = Round(tdiateo + IIf(opcons = 0, IIf(RS1!min_racteo = 0 Or IsNull(RS1!min_racteo), 0, RS1!cosmin), (IIf(IsNull(RS1!min_racteo), 0, RS1!cosmin))) + vCosFij, 2)
          racteo = (racteo + IIf(Not IsNull(RS1!min_racteo), RS1!min_racteo, 0))
       ElseIf RS1!mid_tipmin = "2" Then
          tdiarea = Round(tdiarea + IIf(opcons = 0, IIf(RS1!min_racrea = 0 Or IsNull(RS1!min_racrea), 0, RS1!cosmin), (IIf(IsNull(RS1!min_racrea), 0, RS1!cosmin))) + vCosFij, 2)
          racrea = (racrea + IIf(Not IsNull(RS1!min_racrea), RS1!min_racrea, 0))
       End If
       RS1.MoveNext
    Loop
    'totdoc = 0
    ''Mover costo teórico
    '.Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiateo, fg_Pict(6, 2))
    '.Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiarea, fg_Pict(6, 2))
    
    If opgraf Then
       '-------> Mover teórico
       .Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(0, fg_Pict(6, 2))
       If racteo > 0 Then .Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format((tdiateo / racteo), fg_Pict(6, 2))
       '-------> Mover real
       .Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(0, fg_Pict(6, 2))
       If racrea > 0 Then .Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format((tdiarea / racrea), fg_Pict(6, 2))
    Else
       '-------> Mover teórico
       .Series(0).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiateo, fg_Pict(6, 2))
       '-------> Mover real
       .Series(1).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(tdiarea, fg_Pict(6, 2))
    End If
    totdoc = 0: racteo = 0: racrea = 0
    ''-------> Mover raciones
    'nrorac = 0
    'If opgraf = True Then
    '   RS2.Open "SELECT mir_fecmin, SUM(mir_nrorac) AS nrorac FROM b_minutaraciones " & _
    '            "WHERE mir_cencos='" & cencos & "' " & _
    '            "AND   mir_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") " & _
    '            "AND   mir_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
    '            "AND  (mir_rutcli='PRODUCIDAS') " & _
    '            "AND   mir_fecmin=" & auxfec & " GROUP BY mir_fecmin", vg_db, adOpenStatic
    '   If Not RS2.EOF And Not IsNull(RS2!nrorac) Then nrorac = RS2!nrorac
    '   RS2.Close: Set RS2 = Nothing
    'End If
    ''-------> Traer salida & devolución
    'RS2.Open "SELECT a.tov_codreg, a.tov_codser, SUM(IIf(a.tov_tipdoc='SP',b.dev_ptotal,'-' & b.dev_ptotal)) AS totdoc " & _
    '         "FROM b_totventas a, b_detventas b, b_productos c WHERE a.tov_rutcli = b.dev_rutcli " & _
    '         "AND  a.tov_tipdoc = b.dev_tipdoc AND a.tov_numdoc = b.dev_numdoc " & _
    '         "AND  b.dev_codmer = c.pro_codigo AND " & sql1 & " " & _
    '         "AND  a.tov_codreg IN (" & Mid(codreg, 1, Len(codreg) - 1) & ") AND a.tov_codser IN (" & Mid(codser, 1, Len(codser) - 1) & ") " & _
    '         "AND (a.tov_tipdoc = 'SP' OR a.tov_tipdoc = 'DP') AND b.dev_canmer <> 0 " & _
    '         "AND  a.tov_codbod = " & vg_codbod & " AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' AND a.tov_fecpro = cdate('" & fg_Ctod1(auxfec) & "') GROUP BY a.tov_codreg, a.tov_codser", vg_db, adOpenStatic
    'If Not RS2.EOF And Not IsNull(RS2!totdoc) Then
    '   If opgraf Then
    '      If nrorac > 0 Then totdoc = (RS2!totdoc / nrorac) Else totdoc = 0
    '   Else
    '      totdoc = RS2!totdoc
    '   End If
    'End If
    'RS2.Close: Set RS2 = Nothing
    '.Series(2).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(totdoc, fg_Pict(6, 2))
    RS1.Close: Set RS1 = Nothing
    If opgraf Then .Series(3).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(cospis, fg_Pict(6, 2))
    If opgraf Then .Series(4).Yvalue(Mid(auxfec, 7, 2) - 1) = Format(costec, fg_Pict(6, 2))
    .OpenDataEx COD_COLORS, 3, 0
    For i = 0 To 4
        DoEvents
        If i = 0 Then .Series(i).Legend = "Pla. Teórico"
        If i = 1 Then .Series(i).Legend = "Pla. Real"
        If i = 2 Then .Series(i).color = RGB(80, 240, 60): .Series(i).Legend = "Realizado"
        If i = 3 And opgraf Then .Series(i).color = RGB(40, 240, 202): .Series(i).Legend = "Cto. Piso"
        If i = 4 And opgraf Then .Series(i).Legend = "Cto. Techo"
    Next i
    .CloseData COD_VALUES
    .CloseData COD_COLORS
End With
'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    Me.Hide
    Unload Me
End Select
End Sub
