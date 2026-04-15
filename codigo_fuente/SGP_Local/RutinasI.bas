Attribute VB_Name = "RutinasI"
Option Explicit
Private Declare Function sendmessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wmsg As Long, ByVal wparam As Long, lparam As Any) As Long
Private Declare Function ShellExecute _
                            Lib "shell32.dll" _
                            Alias "ShellExecuteA" ( _
                            ByVal hWnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

Const WM_PASTE = &H302

Function CalcularCostoMinuta(codreg As Long, codser As Long, fecini As Long, fecfin As Long, tipmin As String) As Double
Dim RS1 As New ADODB.Recordset
Dim estfij As Boolean
Dim fecesf As Long, i As Long
Dim aAp  As String, sql1 As String
Dim vCosFij As Double
Dim fecpin As Date, fecpfi As Date
fecpin = fg_Ctod1(fecini)
fecpfi = fg_Ctod1(fecfin)
CalcularCostoMinuta = 0
RS1.Open "SELECT " & _
         "ROUND(SUM(c.mid_cosrec*c.mid_numrac),2) AS mid_cosrec, ROUND(SUM(c.mid_cosdes*c.mid_numrac),2) AS mid_cosdes " & _
         "FROM b_minuta b, b_minutadet c " & _
         "WHERE b.min_codigo = c.mid_codigo " & _
         "AND   b.min_cencos = '" & MuestraCasino(1) & "' " & _
         "AND   b.min_codreg = " & codreg & " " & _
         "AND   b.min_codser = " & codser & " " & _
         "AND   b.min_fecmin >= " & fecini & " " & _
         "AND   b.min_fecmin <= " & fecfin & " " & _
         "AND   c.mid_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ") " & _
         "" & _
         "", vg_db, adOpenStatic
If Not RS1.EOF Then CalcularCostoMinuta = IIf(IsNull(RS1!mid_cosrec), 0, RS1!mid_cosrec) + IIf(IsNull(RS1!mid_cosdes), 0, RS1!mid_cosdes)
RS1.Close: Set RS1 = Nothing
estfij = False
'-------> Buscar datos estructura fija día
RS1.Open "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
        "WHERE mfd_cencos='" & MuestraCasino(1) & "' " & _
        "AND   mfd_codreg = " & codreg & " " & _
        "AND   mfd_codser = " & codser & " " & _
        "AND   mfd_fecha >= " & fecini & " AND mfd_fecha <= " & fecfin & " AND mfd_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ")", vg_db, adOpenStatic
If Not RS1.EOF Then estfij = True
RS1.Close: Set RS1 = Nothing
fecesf = 0
If Not estfij Then
   RS1.Open "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija " & _
            "WHERE mif_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mif_codreg = " & codreg & " " & _
            "AND   mif_codser = " & codser & " ", vg_db, adOpenStatic
   If Not RS1.EOF Then fecesf = IIf(IsNull(RS1!fecval), 0, RS1!fecval)
   RS1.Close: Set RS1 = Nothing
End If
If Not estfij And fecesf > 0 And vg_tipbase = "1" Then
    '-------> Insert tabla productospmpdia
    aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPCalCtoMin"
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
vCosFij = 0
If estfij Then
   '-------> Calcular datos desde tabla estructura fija día
   RS1.Open "SELECT SUM(a.mfd_canpro*a.mfd_cospro) AS cosfij " & _
            "FROM b_minutafijadia a, b_productos c " & _
            "WHERE a.mfd_codpro = c.pro_codigo " & _
            "AND   a.mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   a.mfd_codreg = " & codreg & " " & _
            "AND   a.mfd_codser = " & codser & " " & _
            "AND   a.mfd_fecha  >= " & fecini & " AND a.mfd_fecha  <= " & fecfin & " AND mfd_tipmin IN (" & Mid(tipmin, 1, Len(tipmin)) & ") " & _
            "", vg_db, adOpenStatic
   If Not RS1.EOF And Not IsNull(RS1!cosfij) Then vCosFij = RS1!cosfij
   RS1.Close: Set RS1 = Nothing
ElseIf Not estfij And fecesf > 0 Then
   Do While fecpin <= fecpfi
      '-------> Calcular datos desde tabla estructura fija
      If vg_tipbase = "1" Then
         RS1.Open "SELECT SUM(b.ppd_propon*a.mif_canpro) AS cosfij " & _
                  "FROM  b_minutafija a, b_productos c, " & aAp & " b " & _
                  "WHERE a.mif_codpro = c.pro_codigo AND c.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   a.mif_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   a.mif_codreg = " & codreg & " " & _
                  "AND   a.mif_codser = " & codser & " " & _
                  "AND   a.mif_fecval=" & fecesf & " " & _
                  "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Year(fecpin) & fg_pone_cero(Str(Month(fecpin)), 2) & fg_pone_cero(Str(Day(fecpin)), 2), 2), Len(fg_Fecha_Dia(Year(fecpin) & fg_pone_cero(Str(Month(fecpin)), 2) & fg_pone_cero(Str(Day(fecpin)), 2), 2)) - 2))) & " " & _
                  "", vg_db, adOpenStatic
      Else
         RS1.Open "SELECT SUM(b.ppd_propon*a.mif_canpro) AS cosfij " & _
                  "FROM  b_minutafija a, b_productos c, b_productospmpdia b " & _
                  "WHERE a.mif_codpro = c.pro_codigo AND c.pro_codigo = b.ppd_codpro AND b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   b.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
                  "AND   a.mif_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   a.mif_codreg = " & codreg & " " & _
                  "AND   a.mif_codser = " & codser & " " & _
                  "AND   a.mif_fecval = " & fecesf & " " & _
                  "AND   a.mif_dianro = " & fg_NumDia(Trim(Left(fg_Fecha_Dia(Year(fecpin) & fg_pone_cero(Str(Month(fecpin)), 2) & fg_pone_cero(Str(Day(fecpin)), 2), 2), Len(fg_Fecha_Dia(Year(fecpin) & fg_pone_cero(Str(Month(fecpin)), 2) & fg_pone_cero(Str(Day(fecpin)), 2), 2)) - 2))) & " " & _
                  "", vg_db, adOpenStatic
                  
      End If
      If Not RS1.EOF And Not IsNull(RS1!cosfij) Then vCosFij = RS1!cosfij
      RS1.Close: Set RS1 = Nothing
      fecpin = fecpin + 1
   Loop
End If
CalcularCostoMinuta = (CalcularCostoMinuta + vCosFij)
End Function

Function TraerStock(codbod As Long, Fecha As Long) As Double
Dim RS1 As New ADODB.Recordset
Dim FecInv As Long
Dim sql1 As String
TraerStock = 0
sql1 = IIf(vg_tipbase = "1", " VAL(MID(tin_fectom,1,6)) ", " convert(int,substring(convert(varchar(8),tin_fectom),1,6)) ")
'-------> Buscar inventario anterior
RS1.Open "SELECT MAX(tin_fectom) AS fecinv " & _
         "FROM b_tomainv " & _
         "WHERE tin_codbod=" & codbod & " " & _
         "AND " & sql1 & " = " & Val(Format(BoM(Mid(Fecha, 7, 2) & "/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)), "yyyymm")) & " AND tin_ciemes <> 0", vg_db, adOpenStatic
If Not RS1.EOF And Not IsNull(RS1!FecInv) Then FecInv = RS1!FecInv
RS1.Close: Set RS1 = Nothing
  
RS1.Open "SELECT SUM(b.tin_stofis*b.tin_propon) AS cosinv " & _
         "FROM  b_productos a, b_tomainv b " & _
         "WHERE b.tin_codpro = a.pro_codigo " & _
         "AND  (a.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR a.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
         "AND   b.tin_codbod = " & codbod & " " & _
         "AND   b.tin_fectom = " & FecInv & " " & _
         "AND   b.tin_stofis <> 0 " & _
         "AND   b.tin_propon <> 0", vg_db, adOpenStatic
If Not RS1.EOF And Not IsNull(RS1!cosinv) Then TraerStock = RS1!cosinv
RS1.Close: Set RS1 = Nothing
End Function

Function TraerDocumentoSalida(codbod As Long, fecini As Long, fecfin As Long, tipdoc As String, tipfec As String, op As String) As Double
Dim RS1 As New ADODB.Recordset
If TraerPrimerDiaPeriodo(MuestraCasino(1), Mid(fecini, 1, 6)) <> 0 Then
   fecini = TraerPrimerDiaPeriodo(MuestraCasino(1), Mid(fecini, 1, 6))
ElseIf TraerPrimerDiaPeriodo(MuestraCasino(1), Mid(Format(BEoM(fg_Ctod1(fecini)), "yyyymmdd"), 1, 6)) <> 0 Then
   fecini = TraerPrimerDiaPeriodo(MuestraCasino(1), Mid(Format(BEoM(fg_Ctod1(fecini)), "yyyymmdd"), 1, 6))
End If
Dim sqlfini As String, sqlffin As String, sql1 As String
'------->
TraerDocumentoSalida = 0
sqlfini = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
sqlffin = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
sql1 = IIf(vg_tipbase = "1", " SUM(b.dev_precos*iif(d.aju_tipo='A',(b.dev_canmer),(b.dev_canmer*-1))) ", " SUM(b.dev_precos* CASE WHEN d.aju_tipo='A' THEN (b.dev_canmer) ELSE (b.dev_canmer*-1) END) ")

If op = "1" Then
   RS1.Open "SELECT SUM(b.dev_ptotal) AS ptotal FROM b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND  (c.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR c.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "')) " & _
            "AND   " & tipdoc & " AND b.dev_canmer <> 0 AND a.tov_codbod = " & codbod & " AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' " & _
            "AND   " & tipfec & " >= " & sqlfini & " AND " & tipfec & " <= " & sqlffin & " " & _
            "", vg_db, adOpenStatic
ElseIf op = "2" Then
   RS1.Open "SELECT " & sql1 & " AS ptotal " & _
            "FROM  b_totventas a, b_detventas b, b_productos c, a_tipoajuste d " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   c.pro_codigo = b.dev_codmer " & _
            "AND   a.tov_codser = d.aju_codigo " & _
            "AND   " & tipfec & " >= " & sqlfini & " AND " & tipfec & " <= " & sqlffin & "  " & _
            "AND   a.tov_codbod = " & codbod & " AND a.tov_tipdoc = 'AI' " & _
            "AND   a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' " & _
            "AND   b.dev_canmer > 0 " & _
            "AND   (c.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR c.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "')) " & _
            "", vg_db, adOpenStatic
ElseIf op = "3" Then
    '-------> Incremento por rebaje de Venta Cafetería
    RS1.Open "SELECT ROUND(SUM(a.dvp_precos*a.dvp_candig),0) AS ptotal " & _
             "FROM b_detventascafpro a, b_totventascaf b, b_productos c " & _
             "WHERE b.tvc_cencos = a.dvp_cencos " & _
             "AND   b.tvc_fecing = a.dvp_fecing " & _
             "AND   a.dvp_codmer = c.pro_codigo " & _
             "AND  (c.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos"), ";", "','") & "') OR c.pro_ctacon NOT IN ('" & fg_CambiaChar(GetParametro("ctagastos2"), ";", "','") & "')) " & _
             "AND   a.dvp_precos <> 0 AND b.tvc_estado = 'C' AND b.tvc_codbod = " & codbod & " " & _
             "AND   b.tvc_fecing >= " & sqlfini & " AND b.tvc_fecing <= " & sqlffin & " " & _
             "", vg_db, adOpenStatic
End If
If Not RS1.EOF And Not IsNull(RS1!ptotal) Then TraerDocumentoSalida = IIf(op = "2", (RS1!ptotal * -1), RS1!ptotal)
RS1.Close: Set RS1 = Nothing
End Function

Function ValidarInventarioRotativo(cencos As String) As Boolean
Dim RS1 As New ADODB.Recordset
ValidarInventarioRotativo = False
RS1.Open "SELECT * FROM b_casinoparametrostock WHERE cps_cencos = '" & cencos & "' AND cps_diario = 'S'", vg_db, adOpenStatic
If Not RS1.EOF Then ValidarInventarioRotativo = True Else ValidarInventarioRotativo = False
RS1.Close: Set RS1 = Nothing
'-------> Validar ActividadesDiariaInvRotativo
RS1.Open "SELECT * FROM b_casinotipoactividades WHERE cta_cencos = '" & cencos & "' AND cta_tipact = 10", vg_db, adOpenStatic
If Not RS1.EOF Then ValidarInventarioRotativo = True Else ValidarInventarioRotativo = False
RS1.Close: Set RS1 = Nothing
End Function

Function ValidarActividadesDiariaInvRotativo(cencos As String) As Boolean
Dim RS1 As New ADODB.Recordset
ValidarActividadesDiariaInvRotativo = False
RS1.Open "SELECT * FROM b_casinotipoactividades WHERE cta_cencos = '" & cencos & "' AND cta_tipact = 10", vg_db, adOpenStatic
If Not RS1.EOF Then ValidarActividadesDiariaInvRotativo = True
RS1.Close: Set RS1 = Nothing
End Function

Function fg_OcultarGrilla(Grilla As Object, Row As Long, Col As Long, op As Boolean)

Grilla.Col = Col: Grilla.ColHidden = op

End Function

Function CalcularInvRotPorInventario(codbod As Long, tipopc As String)
Dim RS1 As New ADODB.Recordset
Dim totgrl As Double, curva As Double, porinv As Double
Dim aAp As String, sql1 As String
'-------> Crear tabla temporal
aAp = Trim(vg_NUsr) & "_tmp_filtomainv"
fg_CheckTmp aAp
vg_db.Execute "CREATE TABLE " & aAp & " (tem_codigo varchar(20))"
totgrl = 0
porinv = TraerPorcentajeInventario(MuestraCasino(1))
sql1 = IIf(vg_tipbase = "1", " CDATE('" & vg_ciedia & "') ", " '" & Format(vg_ciedia, "yyyymmdd") & "'  ")
If tipopc = "1" Then
   RS1.Open "SELECT SUM(bod_canmer) AS totabc FROM b_bodegas WHERE bod_codbod = " & codbod & " AND bod_canmer <> 0", vg_db, adOpenStatic
Else
   RS1.Open "SELECT SUM(b.dev_canmer) AS totabc FROM b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctrsto = 1 " & _
            "AND  (c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
            "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
            "AND   a.tov_codbod = " & codbod & " AND b.dev_canmer > 0 " & _
            "AND ((a.tov_fecpro = " & sql1 & " AND a.tov_tipdoc = 'SP' AND (a.tov_estdoc = 'P' OR a.tov_estdoc <> 'A') AND (tov_fecpro) IS NOT NULL) " & _
            "OR   (a.tov_fecemi = " & sql1 & " AND a.tov_tipdoc = 'TR' AND a.tov_codser>0 AND (a.tov_estdoc <> 'A') AND (tov_fecemi) IS NOT NULL))", vg_db, adOpenStatic
End If
If Not RS1.EOF Then totgrl = IIf(IsNull(RS1!totabc), 0, RS1!totabc)
RS1.Close: Set RS1 = Nothing
curva = 0
If tipopc = "1" Then
   RS1.Open "SELECT bod_codbod, bod_codpro as codpro, bod_canmer AS canmer FROM b_bodegas WHERE bod_codbod = " & codbod & " AND bod_canmer <> 0 ORDER BY bod_canmer DESC", vg_db, adOpenStatic
Else
   RS1.Open "SELECT b.dev_codmer AS codpro, SUM(b.dev_canmer) AS canmer " & _
            "FROM b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctrsto = 1 " & _
            "AND  (c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
            "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
            "AND   a.tov_codbod = " & codbod & " AND b.dev_canmer > 0 " & _
            "AND ((a.tov_fecpro = " & sql1 & " AND a.tov_tipdoc = 'SP' AND (a.tov_estdoc = 'P' OR a.tov_estdoc <> 'A') AND (tov_fecpro) IS NOT NULL) " & _
            "OR   (a.tov_fecemi = " & sql1 & " AND a.tov_tipdoc = 'TR' AND a.tov_codser>0 AND (a.tov_estdoc <> 'A') AND (tov_fecemi) IS NOT NULL)) GROUP BY b.dev_codmer ORDER BY canmer DESC", vg_db, adOpenStatic
End If
Do While Not RS1.EOF
   If totgrl > 0 And Not IsNull(RS1!canmer) Then
      curva = curva + ((RS1!canmer / totgrl) * 100)
      vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo) VALUES ('" & RS1!codpro & "')"
   End If
   If curva > porinv Then
      Exit Do
   End If
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
vg_codigo = "|Ok|"
End Function

Function CalcularInvRotCuarvaABC(codbod As Long, tipopc As String)
Dim RS1 As New ADODB.Recordset
Dim totabc As Double, curvaa As Double, curvab As Double, curvac As Double, CurvaABC As Double, curva As Double
Dim indcur As Long, i As Long
Dim tipabc As String, aAp As String, sql1 As String
Dim vec_curvaabc() As Variant
'-------> Traer curva ABC
RS1.Open "SELECT * FROM a_curvaabc", vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      If RS1!abc_codigo = "A" Then curvaa = RS1!abc_porce
      If RS1!abc_codigo = "B" Then curvab = RS1!abc_porce
      If RS1!abc_codigo = "C" Then curvac = RS1!abc_porce
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
'-------> Calcular total mercaderia
totabc = 0
sql1 = IIf(vg_tipbase = "1", " CDATE('" & vg_ciedia & "') ", " '" & Format(vg_ciedia, "yyyymmdd") & "'  ")
If tipopc = "1" Then
   RS1.Open "SELECT SUM(bod_canmer) AS totabc FROM b_bodegas WHERE bod_codbod = " & codbod & " AND bod_canmer <> 0", vg_db, adOpenStatic
Else
   RS1.Open "SELECT SUM(b.dev_canmer) AS totabc FROM b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctrsto = 1 " & _
            "AND  (c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
            "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
            "AND   a.tov_codbod = " & codbod & " AND b.dev_canmer > 0 " & _
            "AND ((a.tov_fecpro = " & sql1 & " AND a.tov_tipdoc = 'SP' AND (a.tov_estdoc = 'P' OR a.tov_estdoc <> 'A') AND (tov_fecpro) IS NOT NULL) " & _
            "OR   (a.tov_fecemi = " & sql1 & " AND a.tov_tipdoc = 'TR' AND a.tov_codser>0 AND (a.tov_estdoc <> 'A') AND (tov_fecemi) IS NOT NULL))", vg_db, adOpenStatic
End If
If Not RS1.EOF Then totabc = IIf(IsNull(RS1!totabc), 0, RS1!totabc)
RS1.Close: Set RS1 = Nothing
'-------> traer cantidad productos
If tipopc = "1" Then
   RS1.Open "SELECT DISTINCT COUNT(bod_codpro) AS nreg FROM b_bodegas WHERE bod_codbod = " & codbod & " AND bod_canmer <> 0", vg_db, adOpenStatic
Else
   RS1.Open "SELECT COUNT(DISTINCT b.dev_codmer) AS nreg FROM b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctrsto = 1 " & _
            "AND  (c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
            "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
            "AND   a.tov_codbod = " & codbod & " AND b.dev_canmer > 0 " & _
            "AND ((a.tov_fecpro = " & sql1 & " AND a.tov_tipdoc = 'SP' AND (a.tov_estdoc = 'P' OR a.tov_estdoc <> 'A') AND (tov_fecpro) IS NOT NULL) " & _
            "OR   (a.tov_fecemi = " & sql1 & " AND a.tov_tipdoc = 'TR' AND a.tov_codser>0 AND (a.tov_estdoc <> 'A') AND (tov_fecemi) IS NOT NULL))", vg_db, adOpenStatic
End If
If Not RS1.EOF Then: ReDim vec_curvaabc(RS1!nreg + 50, 3)
RS1.Close: Set RS1 = Nothing
For i = 0 To UBound(vec_curvaabc)
    vec_curvaabc(i, 1) = ""
    vec_curvaabc(i, 2) = ""
    vec_curvaabc(i, 3) = 0
Next i
indcur = 1: CurvaABC = curvaa: curva = 0: i = 0: tipabc = "A"
If tipopc = "1" Then
   RS1.Open "SELECT bod_codbod, bod_codpro AS codpro, bod_canmer AS canmer FROM b_bodegas WHERE bod_codbod = " & codbod & " AND bod_canmer <> 0 ORDER BY bod_canmer DESC", vg_db, adOpenStatic
Else
   RS1.Open "SELECT b.dev_codmer AS codpro, SUM(b.dev_canmer) AS canmer " & _
            "FROM b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctrsto = 1 " & _
            "AND  (c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') OR c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctalimdes"), ";", "','") & "')) " & _
            "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
            "AND   a.tov_codbod = " & codbod & " AND b.dev_canmer > 0 " & _
            "AND ((a.tov_fecpro = " & sql1 & " AND a.tov_tipdoc = 'SP' AND (a.tov_estdoc = 'P' OR a.tov_estdoc <> 'A') AND (tov_fecpro) IS NOT NULL) " & _
            "OR   (a.tov_fecemi = " & sql1 & " AND a.tov_tipdoc = 'TR' AND a.tov_codser>0 AND (a.tov_estdoc <> 'A') AND (tov_fecemi) IS NOT NULL)) GROUP BY b.dev_codmer ORDER BY canmer DESC", vg_db, adOpenStatic
End If
Do While Not RS1.EOF
   If totabc > 0 And Not IsNull(RS1!canmer) Then
      curva = curva + ((RS1!canmer / totabc) * 100)
      vec_curvaabc(i, 1) = RS1!codpro
      vec_curvaabc(i, 2) = tipabc
      vec_curvaabc(i, 3) = RS1!canmer
      i = i + 1
   Else
      curva = curva + 0
   End If
   If curva > CurvaABC And CurvaABC <> curvac And curva <= 99 Then
      CurvaABC = IIf(indcur = 1, curvab, curvac): curva = 0
      tipabc = IIf(indcur = 1, "B", "C")
      curva = curva + ((RS1!canmer / totabc) * 100)
      vec_curvaabc(i, 1) = RS1!codpro
      vec_curvaabc(i, 2) = tipabc
      vec_curvaabc(i, 3) = RS1!canmer
      i = i + 1
      indcur = indcur + 1
   End If
   RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
'-------> Crear tabla temporal
aAp = Trim(vg_NUsr) & "_tmp_filtomainv"
fg_CheckTmp aAp
vg_db.Execute "CREATE TABLE " & aAp & " (tem_codigo varchar(20))"

'-------> Calcular monto total curva A
totabc = 0
For i = 0 To UBound(vec_curvaabc)
    If vec_curvaabc(i, 2) = "A" Then
       totabc = totabc + vec_curvaabc(i, 3)
    End If
Next i
'-------> calcular los productos que tenga el 10%
curva = 0
For i = 0 To UBound(vec_curvaabc)
    If vec_curvaabc(i, 2) = "A" Then
       If curva > 10 Then
          Exit For
       End If
       curva = curva + ((vec_curvaabc(i, 3) / totabc) * 100)
       vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo) VALUES ('" & vec_curvaabc(i, 1) & "')"
    End If
Next i

'-------> Calcular monto total curva B
totabc = 0
For i = 0 To UBound(vec_curvaabc)
    If vec_curvaabc(i, 2) = "B" Then
       totabc = totabc + vec_curvaabc(i, 3)
    End If
Next i
'-------> calcular los productos que tenga el 10%
curva = 0
For i = 0 To UBound(vec_curvaabc)
    If vec_curvaabc(i, 2) = "B" Then
       If curva > 10 Then
          Exit For
       End If
       curva = curva + ((vec_curvaabc(i, 3) / totabc) * 100)
       vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo) VALUES ('" & vec_curvaabc(i, 1) & "')"
    End If
Next i

'-------> Calcular monto total curva C
totabc = 0
For i = 0 To UBound(vec_curvaabc)
    If vec_curvaabc(i, 2) = "B" Then
       totabc = totabc + vec_curvaabc(i, 3)
    End If
Next i
'-------> calcular los productos que tenga el 10%
curva = 0
For i = 0 To UBound(vec_curvaabc)
    If vec_curvaabc(i, 2) = "C" Then
       If curva > 10 Then
          Exit For
       End If
       curva = curva + ((vec_curvaabc(i, 3) / totabc) * 100)
       vg_db.Execute "INSERT INTO " & aAp & " (tem_codigo) VALUES ('" & vec_curvaabc(i, 1) & "')"
    End If
Next i
vg_codigo = "|Ok|"
End Function

Function ValidarDatoCurvaABC() As Boolean
Dim RS1 As New ADODB.Recordset
'-------> Validar si exietn datos en tabla curva abc
ValidarDatoCurvaABC = True
RS1.Open "SELECT * FROM a_curvaabc", vg_db, adOpenStatic
If RS1.EOF Then ValidarDatoCurvaABC = False
RS1.Close: Set RS1 = Nothing
End Function

Function TraerParametroStock(cencos As String) As String
Dim RS1 As New ADODB.Recordset
TraerParametroStock = ""
RS1.Open "SELECT cps_invsto, cps_reqmen FROM b_casinoparametrostock WHERE cps_cencos = '" & cencos & "' AND cps_diario = 'S'", vg_db, adOpenStatic
If Not RS1.EOF Then
   If RS1!cps_invsto = "S" Then
      TraerParametroStock = "1"
   ElseIf RS1!cps_reqmen = "S" Then
      TraerParametroStock = "2"
   End If
End If
RS1.Close: Set RS1 = Nothing
End Function

Function TraerTipoInventarioRotativo(cencos As String) As String
Dim RS1 As New ADODB.Recordset
TraerTipoInventarioRotativo = ""
RS1.Open "SELECT cps_liscri FROM b_casinoparametrostock WHERE cps_cencos = '" & cencos & "' AND cps_diario = 'S'", vg_db, adOpenStatic
If Not RS1.EOF Then
   TraerTipoInventarioRotativo = IIf(IsNull(RS1!cps_liscri), "", RS1!cps_liscri)
End If
RS1.Close: Set RS1 = Nothing
End Function

Function TraerPorcentajeInventario(cencos As String) As Double
Dim RS1 As New ADODB.Recordset
TraerPorcentajeInventario = 0
RS1.Open "SELECT cps_porinv FROM b_casinoparametrostock WHERE cps_cencos = '" & cencos & "' AND cps_diario = 'S'", vg_db, adOpenStatic
If Not RS1.EOF Then
   TraerPorcentajeInventario = IIf(IsNull(RS1!cps_porinv), 0, RS1!cps_porinv)
End If
RS1.Close: Set RS1 = Nothing
End Function

Function ValidarCodInternoSac(cencos As String) As Boolean
Dim RS1 As New ADODB.Recordset
ValidarCodInternoSac = False
RS1.Open "SELECT cli_ccisac FROM b_clientes WHERE cli_codigo = '" & cencos & "'", vg_db, adOpenStatic
If Not RS1.EOF Then ValidarCodInternoSac = IIf(IsNull(RS1!cli_ccisac) Or RS1!cli_ccisac = 0, False, True)
RS1.Close: Set RS1 = Nothing
End Function

Function ValidarCentralCompraSac(cencos As String) As Boolean
Dim RS1 As New ADODB.Recordset
ValidarCentralCompraSac = False
RS1.Open "SELECT cli_cecsac FROM b_clientes WHERE cli_codigo = '" & cencos & "'", vg_db, adOpenStatic
If Not RS1.EOF Then ValidarCentralCompraSac = IIf(IsNull(RS1!cli_cecsac) Or Trim(RS1!cli_cecsac) = "", False, True)
RS1.Close: Set RS1 = Nothing
End Function

Function TraerNumeroDiasPeriodo(cencos As String, periodo As Long) As Long
Dim RS1 As New ADODB.Recordset
TraerNumeroDiasPeriodo = 0
RS1.Open "SELECT cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos = '" & cencos & "' AND cie_periodo = " & periodo & "", vg_db, adOpenStatic
If Not RS1.EOF Then TraerNumeroDiasPeriodo = (IIf(IsNull(RS1!cie_fecter), 0, RS1!cie_fecter) - IIf(IsNull(RS1!cie_fecini), 0, RS1!cie_fecini)) + 1
RS1.Close: Set RS1 = Nothing
End Function

Function TraerPrimerDiaPeriodo(cencos As String, periodo As Long) As Long
Dim RS1 As New ADODB.Recordset
TraerPrimerDiaPeriodo = 0
RS1.Open "SELECT cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos = '" & cencos & "' AND cie_periodo = " & periodo & "", vg_db, adOpenStatic
If Not RS1.EOF Then TraerPrimerDiaPeriodo = IIf(IsNull(RS1!cie_fecini), 0, RS1!cie_fecini)
RS1.Close: Set RS1 = Nothing
End Function

Function CalcularProrroteoGrlPerDep(diapor As Long, valor As Double, fecini As Long, fecfin As Long) As Double
CalcularProrroteoGrlPerDep = 0
If diapor > 0 Then CalcularProrroteoGrlPerDep = ((valor / diapor) * ((fecfin - fecini) + 1))
End Function

Function RetencionFuente(codref As Long) As Double
Dim RS As New ADODB.Recordset
RetencionFuente = 0
'-------> Traer Impuesto retencion en la fuente
RS.Open "SELECT ref_codigo, ref_portar " & _
        "FROM   b_retencionfuente " & _
        "WHERE  ref_codigo = " & IIf(IsNull(codref), 0, codref) & "", vg_db, adOpenForwardOnly
If Not RS.EOF Then
     RetencionFuente = IIf(IsNull(RS!ref_portar), 0, RS!ref_portar)
End If
RS.Close: Set RS = Nothing
End Function

Function RetencionIca(rut As String, codrei As Long) As Double
Dim RS As New ADODB.Recordset
Dim v_rut As String
RetencionIca = 0
'-------> Traer Impuesto retención ica si el contrato tiene asignado la retencion obligatorio
RS.Open "SELECT TOP 1 b.dri_portar " & _
        "FROM  b_retencionica a, b_detretencionica b, b_clientes c, a_municipio d " & _
        "WHERE a.rei_codigo = b.dri_codigo " & _
        "AND   b.dri_codmun = c.cli_codmun " & _
        "AND   c.cli_codmun = d.mun_codigo " & _
        "AND   d.mun_retobl = '1' " & _
        "AND   a.rei_codigo = " & IIf(IsNull(codrei), 0, codrei) & " " & _
        "AND   c.cli_codigo = '" & MuestraCasino(1) & "'", vg_db, adOpenForwardOnly
If Not RS.EOF Then
   RetencionIca = IIf(IsNull(RS!dri_portar), 0, RS!dri_portar)
End If
RS.Close: Set RS = Nothing
'-------> Traer impuesto retención ica si contrato no tiene asignado
If RetencionIca = 0 Then
   v_rut = fg_DespintaRut(rut)
   RS.Open "SELECT TOP 1 b.dri_portar " & _
           "FROM b_retencionica a, b_detretencionica b, b_proveedor c, a_municipio d " & _
           "WHERE a.rei_codigo = b.dri_codigo " & _
            "AND   b.dri_codmun = c.prv_codmun " & _
            "AND   c.prv_codmun = d.mun_codigo " & _
            "AND   a.rei_codigo = " & IIf(IsNull(codrei), 0, codrei) & " " & _
            "AND   c.prv_codigo = '" & rut & "'", vg_db, adOpenForwardOnly
   If Not RS.EOF Then
      RetencionIca = IIf(IsNull(RS!dri_portar), 0, RS!dri_portar)
   End If
   RS.Close: Set RS = Nothing
End If
End Function

Function ValidarCodIva(codiva As Long) As Boolean
Dim RS As New ADODB.Recordset
ValidarCodIva = False
Set RS = vg_db.Execute("SELECT par_valor FROM a_param a WHERE par_codigo = 'pariva'")
If Not RS.EOF Then
   If Trim(RS!par_valor) <> "" Then
      RS.Close: Set RS = Nothing
      Set RS = vg_db.Execute("SELECT b.imp_codigo, b.imp_nombre FROM a_impuesto b WHERE b.imp_codigo IN (" & fg_CambiaChar(GetParametro("pariva"), ";", ",") & ")")
      Do While Not RS.EOF
         If codiva = RS!imp_codigo Then ValidarCodIva = True: Exit Do
         RS.MoveNext
      Loop
      RS.Close: Set RS = Nothing
   Else
      RS.Close: Set RS = Nothing
   End If
Else
   RS.Close: Set RS = Nothing
End If
End Function

Function ValidarImpuestoAdicional(codiva As Variant) As Boolean
Dim RS As New ADODB.Recordset

ValidarImpuestoAdicional = False
Set RS = vg_db.Execute("SELECT imp_adicional FROM a_impuesto WHERE imp_codigo = " & codiva & "")
If Not RS.EOF Then
   ValidarImpuestoAdicional = IIf(IsNull(RS!imp_adicional) Or RS!imp_adicional = 0, True, False)
End If
RS.Close: Set RS = Nothing
End Function

Function CargarDatoCombo(Combo As Object, Index As Integer, TablaGen As String, SufGen As String, op As String, tipdat As String)

Dim RS As New ADODB.Recordset
Dim sql1 As String

Combo(Index).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If op = "Gen" Then
   
   If TablaGen = "a_tipodocumento" Then
   
      RS.Open "SELECT * FROM " & TablaGen & " where " & SufGen & "VisualizaDoc = 1 " & "ORDER BY " & SufGen & "orden", vg_db, adOpenStatic
   
   Else
   
      RS.Open "SELECT * FROM " & TablaGen & "  ORDER BY " & SufGen & "orden", vg_db, adOpenStatic
   
   End If

ElseIf Mid(op, 1, 6) = "PunVen" Then

'   RS.Open RutinaLectura.PuntoAtencion(6, 0, SufGen), vg_db, adOpenStatic
   Set RS = vg_db.Execute("sgp_Sel_PuntoAtencion '" & SufGen & "'")

ElseIf Mid(op, 1, 6) = "LecReg" Then
   
   RS.Open RutinaLectura.PtoLecturaValesServicios(1, CLng(vg_ptoate), SufGen, 0), vg_db, adOpenStatic
   Set RS = vg_db.Execute("sgp_Sel_PuntoLecturaValesxRegimen '" & SufGen & "', " & CLng(vg_ptoate) & "")

ElseIf Mid(op, 1, 6) = "LecSer" Then
'   RS.Open RutinaLectura.PtoLecturaValesServicios(2, CLng(vg_ptoate), SufGen, CLng(vg_codreg)), vg_db, adOpenStatic
   
   Set RS = vg_db.Execute("sgp_Sel_PuntoLecturaValesxServicio '" & SufGen & "', " & CLng(vg_ptoate) & ", " & CLng(vg_codreg) & "")

ElseIf Mid(op, 1, 6) = "TipDoc" Then
   
   sql1 = ""
   If Trim(Mid(op, 7, 1)) <> "" Then
      
      sql1 = IIf(Trim(Mid(op, 7, 1)) = "N", " WHERE tdo_codigo NOT IN ('CE', 'DE', 'FE') ", " WHERE tdo_codigo NOT IN ('NC', 'ND', 'FA') ")
   
   End If
   RS.Open "SELECT * FROM " & TablaGen & " " & sql1 & " ORDER BY " & SufGen & "orden", vg_db, adOpenStatic

ElseIf op = "CliBod" Then
   
   RS.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE a.bod_codigo = b.cli_codbod AND b.cli_codigo = '" & vg_contra & "' ORDER BY bod_nombre", vg_db, adOpenStatic

ElseIf op = "TipAju" Then
   
   RS.Open "SELECT aju_codigo, aju_nombre FROM a_tipoajuste WHERE aju_tipaju = 0 and aju_activo = '1' and aju_codigo > 9999 ORDER BY aju_nombre", vg_db, adOpenStatic

ElseIf op = "TipAju2" Then
   
   RS.Open "SELECT aju_codigo, aju_nombre FROM a_tipoajuste WHERE aju_tipaju = 0 and aju_activo = '1' ORDER BY aju_nombre", vg_db, adOpenStatic

ElseIf op = "TipPro" Then
   
   RS.Open "SELECT * FROM a_tipopro ORDER BY tip_nombre", vg_db, adOpenStatic

ElseIf op = "Prod" Then
   
   RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic

ElseIf op = "PtoLecVal" Then
   
   RS.Open RutinaLectura.PuntoAtencion(5, 0, SufGen), vg_db, adOpenStatic

ElseIf op = "TipoVales" Then
   
   RS.Open RutinaLectura.TipoVales(1, 0, vg_contra), vg_db, adOpenStatic

ElseIf op = "LugFis" Then
   
   Set RS = vg_db.Execute("SELECT IDLugarFisico, LugarFisico FROM LugarFisico_AX ORDER BY IDLugarFisico")

End If

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      If tipdat = "A" Then
         
         Combo(Index).AddItem Trim(RS(1)) & Space(150) & "(" & Trim(RS(0)) & ")"
      
      ElseIf tipdat = "AE" Then
         
         Combo(Index).AddItem Trim(RS(1)) & Space(150) & "(" & fg_pone_espacio(RS(0), 100) & ")"
      
      ElseIf tipdat = "NM" Then
      
         Combo(Index).AddItem Trim(RS(1)) & " - " & RS(0) & Space(150) & "(" & fg_pone_cero(Str(RS(0)), 10) & ")"
      
      Else
         
         Combo(Index).AddItem Trim(RS(1)) & Space(150) & "(" & fg_pone_cero(Str(RS(0)), 10) & ")"
      
      End If
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing

End Function

Function SetFpDouble(FpDouble As Object, Index As Integer, op As Integer, nValue As Variant)

FpDouble(Index).DecimalPlaces = IIf(op = 1, vg_RDCa, vg_DCa)
FpDouble(Index).Value = nValue

End Function

Function TraerCuentaIva(codimp As Long) As String

'-------> Traer variable impuesto iva
Dim RS As New ADODB.Recordset
Dim sql1 As String
TraerCuentaIva = ""
sql1 = IIf(vg_tipbase = "1", " trim(imp_codsap) ", " ltrim(imp_codsap) ")
RS.Open "SELECT imp_codsap FROM a_impuesto WHERE imp_codigo = " & codimp & " AND imp_adicional = 0 AND ((imp_codsap) IS NOT NULL OR " & sql1 & " <> '')", vg_db, adOpenStatic
If Not RS.EOF Then TraerCuentaIva = Trim(RS!imp_codsap)
RS.Close: Set RS = Nothing

End Function

Function TraerFolioDocumento(tipinf As String) As Long

Dim RS As New ADODB.Recordset
TraerFolioDocumento = 0
RS.Open "SELECT MAX(inf_numero) AS numero FROM a_infcfcfofi WHERE inf_cencos ='" & MuestraCasino(1) & "' AND inf_tipo = '" & tipinf & "'", vg_db, adOpenStatic
If Not RS.EOF Then TraerFolioDocumento = IIf(IsNull(RS!numero), 1, RS!numero)
RS.Close: Set RS = Nothing

End Function

Function MarcaPredeterminadoFormatoCompras()

'-------> Marcar como predeterminado formato de compras si no esta.
Dim aAp As String
Dim sql1 As String
Dim sql2 As String
Dim sql3 As String
Dim sql4 As String

'-------> Creo tabla temporal y chequeo si existe antes
aAp = Trim(vg_NUsr) & "_tmp_formatocomprassgp"
sql1 = IIf(vg_tipbase = "1", " AND cdate(a.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), a.foc_vigfin,101) >  '" & Date & "'")
sql2 = IIf(vg_tipbase = "1", " AND cdate(z.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), z.foc_vigfin,101) >  '" & Date & "'")
sql3 = IIf(vg_tipbase = "1", " AND cdate(y.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), y.foc_vigfin,101) >  '" & Date & "'")
sql4 = IIf(vg_tipbase = "1", " cdate(b_formatocompras.foc_vigfin) <  cdate('" & Date & "') ", " convert(varchar(10), a.foc_vigfin,101) <  '" & Date & "'")
If vg_tipbase = "1" Then
   vg_db.Execute "UPDATE b_formatocomprassgp INNER JOIN b_formatocompras ON b_formatocomprassgp.fcs_codsac = b_formatocompras.foc_codsac SET b_formatocomprassgp.fcs_sgppre = 0, b_formatocomprassgp.fcs_cenpre = 0 WHERE ( b_formatocompras.foc_flexec = 2 OR (" & sql4 & ")) AND (b_formatocompras.foc_vigfin) IS NOT NULL"
Else
   vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_sgppre = 0, fcs_cenpre = 0 FROM b_formatocomprassgp b, b_formatocompras a WHERE b.fcs_codsac = a.foc_codsac AND ( a.foc_flexec = 2 OR (" & sql4 & "))"
End If
fg_CheckTmp aAp
vg_db.Execute "SELECT DISTINCT a.fcs_codsgp INTO " & aAp & " FROM b_formatocomprassgp a  WHERE a.fcs_codsac IN (SELECT DISTINCT a.foc_codsac FROM b_formatocompras a, b_formatocomprassgp b WHERE a.foc_codsac = b.fcs_codsac AND   b.fcs_sgppre <> 1 AND (a.foc_flexec = 0 OR (a.foc_flexec = -1  " & sql1 & "))) AND a.fcs_sgppre = 0"
If vg_tipbase = "1" Then
   vg_db.Execute "UPDATE b_formatocomprassgp INNER JOIN " & aAp & " ON b_formatocomprassgp.fcs_codsgp = " & aAp & ".fcs_codsgp SET b_formatocomprassgp.fcs_sgppre = 1 " & _
                 "WHERE b_formatocomprassgp.fcs_sgppre <> 1 AND  b_formatocomprassgp.fcs_codsac = (SELECT TOP 1 x.fcs_codsac FROM b_formatocomprassgp x WHERE x.fcs_codsgp =  b_formatocomprassgp.fcs_codsgp) and  b_formatocomprassgp.fcs_codsgp = (SELECT TOP 1 z.fcs_codsgp FROM b_formatocomprassgp z, b_formatocompras y WHERE y.foc_codsac = b_formatocomprassgp.fcs_codsac and b_formatocomprassgp.fcs_sgppre <> 1 AND (y.foc_flexec = 0 OR (y.foc_flexec = -1  " & sql3 & ")))"
Else
   '   vg_db.Execute "UPDATE b_formatocomprassgp SET b_formatocomprassgp.fcs_sgppre = 1 FROM b_formatocomprassgp a, " & aAp & " b WHERE a.fcs_codsgp = b.fcs_codsgp AND a.fcs_sgppre <> 1 AND a.fcs_codsac = (SELECT TOP 1 x.fcs_codsac FROM b_formatocomprassgp x WHERE x.fcs_codsgp = a.fcs_codsgp)"
   '   vg_db.Execute "UPDATE b_formatocomprassgp SET b_formatocomprassgp.fcs_sgppre = 1 FROM b_formatocomprassgp a, " & aAp & " b WHERE a.fcs_codsgp = b.fcs_codsgp AND a.fcs_sgppre <> 1 AND a.fcs_codsac = (SELECT TOP 1 x.fcs_codsac FROM b_formatocomprassgp x WHERE x.fcs_codsgp = a.fcs_codsgp) and  a.fcs_codsgp = (SELECT TOP 1 z.fcs_codsgp FROM b_formatocomprassgp z WHERE a.fcs_sgppre <> 1)"
   vg_db.Execute "UPDATE b_formatocomprassgp SET b_formatocomprassgp.fcs_sgppre = 1 FROM b_formatocomprassgp a, " & aAp & " b WHERE a.fcs_codsgp = b.fcs_codsgp AND a.fcs_sgppre <> 1 AND a.fcs_codsac = (SELECT TOP 1 x.fcs_codsac FROM b_formatocomprassgp x, b_formatocompras z WHERE z.foc_codsac = x.fcs_codsac and x.fcs_codsgp = a.fcs_codsgp and a.fcs_sgppre <> 1 AND (z.foc_flexec = 0 OR (z.foc_flexec = -1  " & sql2 & ")))"
End If
vg_db.Execute "UPDATE b_formatocomprassgp SET b_formatocomprassgp.fcs_cenpre = 1 WHERE b_formatocomprassgp.fcs_sgppre = 1 AND (b_formatocomprassgp.fcs_cenpre = 0 OR (b_formatocomprassgp.fcs_cenpre) IS NULL) AND b_formatocomprassgp.fcs_codsac = (SELECT TOP 1 x.fcs_codsac FROM b_formatocomprassgp x, b_formatocompras z WHERE z.foc_codsac = x.fcs_codsac and x.fcs_codsgp = b_formatocomprassgp.fcs_codsgp AND x.fcs_sgppre = 1 AND (z.foc_flexec = 0 OR (z.foc_flexec = -1  " & sql2 & ")))"
vg_db.Execute "DROP TABLE " & aAp & ""
End Function

Function GenerarFolioCFC(cencos As String, tipdoc As String) As Long
Dim RS As New ADODB.Recordset
GenerarFolioCFC = 0
RS.Open "SELECT MAX(inf_numero) AS Mayor FROM a_infcfcfofi WHERE inf_cencos = '" & cencos & "' AND inf_tipo = '" & tipdoc & "'", vg_db, adOpenStatic
GenerarFolioCFC = TipoDato(RS!mayor, 0) + 1
RS.Close: Set RS = Nothing
vg_db.BeginTrans
vg_db.Execute "INSERT INTO a_infcfcfofi VALUES ('" & cencos & "', '" & tipdoc & "', " & GenerarFolioCFC & ", 0, NULL)"
vg_db.CommitTrans
End Function

Function SepararFolioDocumento(cencos As String, codbod As Long, numero As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim corre As Long, periodo
Dim sql1 As String
Dim sql2 As String

'-------> Traer Periodo

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open "SELECT * FROM b_cierreperiodo WHERE cie_cencos ='" & cencos & "' AND cie_estado = 1", vg_db, adOpenStatic
periodo = 0
If Not RS.EOF Then
   
   periodo = IIf(IsNull(RS!cie_periodo), 0, RS!cie_periodo)

End If
RS.Close
Set RS = Nothing

'-------> Asignar nuevo correlativo de folio si existe documento electronicos
corre = 0
'sql1 = IIf(vg_tipbase = "1", " IIF(toc_tipdoc <> 'FE' OR toc_tipdoc <> 'DE' OR toc_tipdoc <> 'CE','FA', '') AS facnor, IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', '') AS facele, COUNT(IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', 'FA')) AS nreg ", " (CASE WHEN toc_tipdoc <> 'FE' AND toc_tipdoc <> 'DE' AND toc_tipdoc <> 'CE' THEN null ELSE toc_tipdoc END) facnor, (CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' END) AS facele, COUNT(CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' ELSE 'FA' END) AS nreg ")
'If vg_tipbase = "1" Then
'
'   RS1.Open "SELECT TOP 1 (SELECT TOP 1 'FA' FROM  b_totcompras a WHERE a.toc_codbod = " & codbod & "  AND a.toc_numinf = " & numero & " AND a.toc_tipdoc NOT IN ('FE','DE','CE', 'SN') AND a.toc_envsap = '0' AND a.toc_fecper = " & periodo & ") AS facnor, " & _
'            "(SELECT TOP 1 'FE' FROM b_totcompras b WHERE b.toc_codbod = " & codbod & " AND b.toc_numinf = " & numero & " AND b.toc_tipdoc IN ('FE','DE','CE') AND b.toc_envsap = '0' AND b.toc_fecper = " & periodo & ") AS facele, " & _
'            "COUNT(IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', 'FA')) AS nreg " & _
'            "FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " AND toc_envsap = '0' AND toc_tipdoc not in ('SN') AND toc_fecper = " & periodo & "", vg_db, adOpenStatic
'
'Else
   
'   RS1.Open "SELECT TOP 1 (SELECT TOP 1 'FA' FROM b_totcompras as a with (nolock) inner join a_tipodocumento as at1 with (nolock) on AT1.tdo_codigo = A.toc_tipdoc WHERE a.toc_codbod = " & codbod & "  AND a.toc_numinf = " & numero & " AND a.toc_tipdoc NOT IN ('FE', 'CE','DE', 'SN') AND a.toc_envsap = '0' AND a.toc_fecper = " & periodo & ") facnor, " & _
'            "(SELECT TOP 1 'FE' FROM b_totcompras b WHERE b.toc_codbod = " & codbod & " AND b.toc_numinf = " & numero & " AND aT1.tdo_IdCodigo IN ('FE', 'CE','DE') AND b.toc_envsap = '0' AND b.toc_fecper = " & periodo & ") AS facele, " & _
'            "COUNT(CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' ELSE 'FA' END) AS nreg " & _
'            "FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " AND toc_envsap = '0' AND toc_tipdoc not in ('SN') AND toc_fecper = " & periodo & "", vg_db, adOpenStatic
''            "FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " AND toc_envsap = '0' AND toc_tipdoc not in ('SN') AND toc_fecper = " & periodo & " GROUP BY toc_tipdoc, toc_tipdoc", vg_db, adOpenStatic

    Set RS1 = vg_db.Execute("sgp_Sel_SepararFoliDocumento " & codbod & ", " & numero & ", " & periodo & "")
'End If

If Not RS1.EOF Then
   
   If RS1!nreg > 1 And (Not IsNull(RS1!facnor) And Not IsNull(RS1!facele) And Trim(RS1!facele) <> "") Then
   
      RS1.Close
      Set RS1 = Nothing
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      RS1.Open "SELECT MAX(inf_numero) AS Mayor FROM a_infcfcfofi WHERE inf_cencos = '" & cencos & "' AND inf_tipo = 'C'", vg_db, adOpenStatic
      corre = TipoDato(RS1!mayor, 0) + 1
      RS1.Close
      Set RS1 = Nothing
      
      vg_db.Execute "INSERT INTO a_infcfcfofi VALUES ('" & Trim(cencos) & "', 'C', " & corre & ", 0, NULL)"
      sql2 = IIf(vg_tipbase = "1", "  val(toc_docaso) ", " convert(int,toc_docaso) ")
'      vg_db.Execute "UPDATE b_totcompras SET toc_numinf = " & corre & " WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " AND toc_tipdoc IN ('SN') AND " & sql2 & " IN (SELECT DISTINCT toc_numdoc FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " AND toc_tipdoc IN ('FE','DE','CE')) AND toc_rutpro IN  (SELECT DISTINCT toc_rutpro FROM b_totcompras WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " AND toc_tipdoc IN ('FE','DE','CE'))"
      vg_db.Execute "UPDATE b_totcompras SET toc_numinf = " & corre & " WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " " & _
                    "AND toc_tipdoc IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND " & _
                    "" & sql2 & " IN (SELECT DISTINCT toc_numdoc FROM b_totcompras WHERE toc_codbod = " & codbod & " AND " & _
                    "toc_numinf = " & numero & " AND toc_tipdoc IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('FE','DE','CE'))) AND " & _
                    "toc_rutpro IN  (SELECT DISTINCT toc_rutpro FROM b_totcompras WHERE toc_codbod = " & codbod & " AND " & _
                    "toc_numinf = " & numero & " AND toc_tipdoc IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('FE','DE','CE')))"
      
      vg_db.Execute "UPDATE b_totcompras SET toc_numinf = " & corre & " WHERE toc_codbod = " & codbod & " AND toc_numinf = " & numero & " " & _
                    "AND toc_tipdoc IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('FE','DE','CE'))"
   
   Else
      
      RS1.Close
      Set RS1 = Nothing
   
   End If

Else
   
   RS1.Close
   Set RS1 = Nothing

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Function ValidarProductosSgpSac(codsac As String, codsgp As String) As Boolean
Dim sql1 As String
Dim RS As New ADODB.Recordset
ValidarProductosSgpSac = False
sql1 = IIf(vg_tipbase = "1", " AND cdate(a.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), a.foc_vigfin,101) >  '" & Date & "'")
RS.Open "SELECT DISTINCT a.foc_codsac, a.foc_nomsac " & _
        "FROM b_formatocompras a, b_formatocomprassgp b " & _
        "WHERE a.foc_codsac = b.fcs_codsac " & _
        "AND   b.fcs_codsgp IN ('" & codsgp & "') " & _
        "AND   b.fcs_codsac <> '" & codsac & "' " & _
        "AND  (a.foc_flexec = 0 OR (a.foc_flexec = -1 " & sql1 & "))", vg_db, adOpenStatic
If Not RS.EOF Then ValidarProductosSgpSac = True
RS.Close: Set RS = Nothing
End Function

Function ValidarDocumentoSap(rutpro As String, tipdoc As String, NumDoc As Long, codbod As Long, cencos) As Boolean

Dim RS As New ADODB.Recordset
Dim sql1 As String
sql1 = IIf(vg_tipbase = "1", " trim(b.inf_usuario)  ", " ltrim(b.inf_usuario) ")
ValidarDocumentoSap = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT DISTINCT a.toc_fecemi " & _
        "FROM   b_totcompras a, a_infcfcfofi b " & _
        "WHERE  a.toc_numinf  = b.inf_numero " & _
        "AND    b.inf_cencos  = '" & cencos & "' " & _
        "AND    b.inf_tipo    = 'C' " & _
        "AND   (b.inf_feccie  > 0   OR  (b.inf_feccie) IS NOT NULL) " & _
        "AND   (" & sql1 & "  <> '' OR (b.inf_usuario)IS NOT NULL) " & _
        "AND    a.toc_rutpro  = '" & rutpro & "' " & _
        "AND    a.toc_tipdoc  = '" & tipdoc & "' " & _
        "AND    a.toc_numdoc  = " & NumDoc & " " & _
        "AND    a.toc_codbod  = " & codbod & "", vg_db, adOpenStatic
'        "AND    a.toc_envsap  = '1'
If Not RS.EOF Then ValidarDocumentoSap = True
RS.Close
Set RS = Nothing

End Function

Function ValidarEnvioCorrelativo(cencos As String, Tipo As String, numfol As Long) As Boolean

Dim RS As New ADODB.Recordset
Dim sql1 As String

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

sql1 = IIf(vg_tipbase = "1", " trim(inf_usuario)  ", " ltrim(inf_usuario) ")
ValidarEnvioCorrelativo = False
RS.Open "SELECT min(inf_numero) AS inf_numero " & _
        "FROM   a_infcfcfofi " & _
        "WHERE  inf_cencos = '" & cencos & "' " & _
        "AND    inf_tipo   = '" & Tipo & "' " & _
        "AND   (inf_feccie = 0 OR (inf_feccie) IS NULL) " & _
        "AND   (" & sql1 & " = '' OR (inf_usuario) IS NULL)", vg_db, adOpenStatic

If Not RS.EOF Then
   
   If numfol > RS!inf_numero Then ValidarEnvioCorrelativo = True

End If
RS.Close: Set RS = Nothing

End Function

Function TraerFolioCFCPosterioesNoEnvio(cencos As String, Tipo As String) As Long
Dim RS As New ADODB.Recordset
Dim sql1 As String
sql1 = IIf(vg_tipbase = "1", " trim(inf_usuario)  ", " ltrim(inf_usuario) ")
TraerFolioCFCPosterioesNoEnvio = 0
RS.Open "SELECT min(inf_numero) AS inf_numero " & _
        "FROM   a_infcfcfofi " & _
        "WHERE  inf_cencos = '" & cencos & "' " & _
        "AND    inf_tipo   = '" & Tipo & "' " & _
        "AND   (inf_feccie = 0 OR (inf_feccie) IS NULL) " & _
        "AND   (" & sql1 & " = '' OR (inf_usuario) IS NULL)", vg_db, adOpenStatic
If Not RS.EOF Then
   TraerFolioCFCPosterioesNoEnvio = RS!inf_numero
End If
RS.Close: Set RS = Nothing
End Function

Function ValidarAccesoMinutaBloqueyBloqueo(ByVal Ceco As String, ByVal op As Integer) As Boolean
Dim RS As New ADODB.Recordset
Dim Sql As String
ValidarAccesoMinutaBloqueyBloqueo = False
Select Case op
Case 1
    Set RS = vg_db.Execute("select cli_tipominuta from b_clientes WITH (NOLOCK) where cli_codigo = '" & Ceco & "' and cli_tipo = 0 and cli_activo = '1'")
    If Not RS.EOF Then
       ValidarAccesoMinutaBloqueyBloqueo = IIf(RS!cli_tipominuta = 3, False, True)
    End If
    RS.Close: Set RS = Nothing
Case 2
    Set RS = vg_db.Execute("select cli_tipominuta from b_clientes WITH (NOLOCK) where cli_codigo = '" & Ceco & "' and cli_tipo = 0 and cli_activo = '1'")
    If Not RS.EOF Then
       ValidarAccesoMinutaBloqueyBloqueo = IIf(RS!cli_tipominuta = 1, True, False)
    End If
    RS.Close: Set RS = Nothing
End Select
End Function

Function GeneraInvAX(codigo_Inv As Long, periodo As String, Fecha As Long) As Boolean

On Error GoTo error

Dim RS          As New ADODB.Recordset
Dim RS1         As New ADODB.Recordset
Dim RSAx        As New ADODB.Recordset
Dim Ceco        As String
Dim periodommaa As String
Dim AnJes       As Object
Dim CencosAx    As String
Dim CencosSgp   As String
Dim hora        As String
Dim Glosa       As String
Dim i           As Long
Dim Sociedad    As String
Dim Contacto    As String
Dim XL          As Object 'Excel.Application
Dim FileName    As String
Dim MontoTotal  As Double

GeneraInvAX = False
'--> Validar homologación Ceco y cuentas contable AX
Ceco = MuestraCasino(1)
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT bc.cli_percon FROM Cecos_Sap_AX csa with (nolock) inner join b_clientes bc on bc.cli_codigo = csa.Cecos_Sap and bc.cli_socsap = csa.Sociedad_Sap and bc.cli_activo = '1' and bc.cli_tipo = 0 and bc.cli_codbod > 0 WHERE bc.cli_codigo = '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación Ceco con AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarCuentasAx '" & Ceco & "'")
If Not RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación cuentas contables AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Sociedad '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe sociedad AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
Sociedad = RS!Sociedad_AX
RS.Close
Set RS = Nothing

GeneraInvAX = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'Set RS = vg_db.Execute("sgp_Sel_GenerarArchivoInvAX_01 '" & Ceco & "',  " & codigo_Inv & "")
Set RS = vg_db.Execute("sgp_Sel_ResumenInvAX '" & Ceco & "',  " & codigo_Inv & "")
If Not RS.EOF Then

   Set AnJes = CreateObject("scripting.filesystemobject")
   
   If Not AnJes.FolderExists(dir_trabajo_Inf & "InformesAXInventario") Then
      
      Call AnJes.CreateFolder(dir_trabajo_Inf & "InformesAXInventario")
   
   End If

   CencosAx = Trim(RS("Cecos_AX"))
   CencosSgp = Trim(Ceco)
   hora = Format(Time$, "HHMMSS")
     
   periodommaa = IIf(periodo = "", 0, Mid(periodo, 5, 2) & Mid(periodo, 1, 4))
   
   If (AnJes.FileExists(dir_trabajo_Inf & "InformesAXInventario\InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".txt")) Then
       
       Kill (dir_trabajo_Inf & "InformesAXInventario\InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".txt")
   
   End If
     
   Open dir_trabajo_Inf & "InformesAXInventario\InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".txt" For Append As #9
   Close #9
   Open dir_trabajo_Inf & "InformesAXInventario\InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".txt" For Append As #9
    
   If RSAx.State = 1 Then RSAx.Close
   Set RSAx = vg_db.Execute("sgp_Ins_GrabaLogInvAX 'InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".txt" & "', '" & CencosSgp & "', " & periodo)
   If Not RSAx.EOF Then
      
      If RSAx(0) > 0 Then
         
         fg_descarga
         MsgBox RSAx(0) & " " & RSAx(1), vbCritical + vbOKOnly, MsgTitulo
         RSAx.Close: Set RSAx = Nothing
         Exit Function
     
     End If
   
   End If
   
   RSAx.Close
   Set RSAx = Nothing
   
   i = 0
   
   Glosa = ""
   Glosa = "" & ";" & "" & ";" & "INVENTARIO"
   Print #9, Glosa
   
   Glosa = ""
   Print #9, Glosa
   
   Glosa = ""
   Glosa = "" & ";" & "Fecha" & ";" & Mid(Fecha, 7, 2) & "/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4)
   Print #9, Glosa
   
   Glosa = ""
   Glosa = "" & ";" & "Ceco Optimum" & ";" & CencosAx
   Print #9, Glosa
   
   Glosa = ""
   Glosa = "" & ";" & "Sociedad" & ";" & Sociedad
   Print #9, Glosa
   
   Glosa = ""
   Glosa = "" & ";" & "Usuario/Responsable" & ";" & Contacto
   Print #9, Glosa
   
   Glosa = ""
   Print #9, Glosa
   
   Glosa = ""
   Glosa = "" & ";" & "Cuenta" & ";" & "Denominación Cuenta" & ";" & "Monto"
   Print #9, Glosa
   
   Do While Not RS.EOF

      Glosa = ""
      Glosa = "" & ";" & Glosa & Trim(RS("Cuentas_AX")) & ";" ' Cuenta
      Glosa = Glosa & Trim(RS("cta_nombre")) & ";" ' Denominación Cuenta
      Glosa = Glosa & RS("inv_mtodoc") & ";" ' Monto
      Print #9, Glosa
      
      MontoTotal = RS("inv_mtodoc") + MontoTotal
      RS.MoveNext
   
   Loop
   
   Glosa = ""
   Print #9, Glosa
   
   Glosa = ""
   Glosa = "" & ";" & "" & ";" & "Total Inventario" & ";" & MontoTotal
   Print #9, Glosa
   
   Close #9
   fg_descarga
   
   Dim ClaveCaratula As String
   
   ClaveCaratula = "123456"
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS1 = vg_db.Execute("select isnull(par_valor,'123456') as par_valor from a_param where par_codigo = 'parconcain' and par_cencos = '" & Ceco & "'")
   If Not RS1.EOF Then
   
     ClaveCaratula = fg_Desencripta(RS1!par_valor)
     
   End If
   RS1.Close
   Set RS1 = Nothing
   
   Set XL = CreateObject("Excel.application")
   FileName = dir_trabajo_Inf & "InformesAXInventario\InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".txt"
   XL.Workbooks.OpenText Mid((FileName), 1, Len((FileName)) - 3) & "txt", , 1, 1, , , , , , , True, ";"
   
   'Bloquear hoja
'Mod Ini 20240801   XL.ActiveSheet.Protect Password:=ClaveCaratula, DrawingObjects:=True, _
'Mod Ini 20240801              Contents:=True, Scenarios:=True, AllowFormattingCells:=True
   
'   XL.ActiveSheet.Protect Password:="", DrawingObjects:=True, _
'              Contents:=True, Scenarios:=True, AllowFormattingCells:=True
   
   XL.ActiveWorkbook.SaveAs FileName:=Mid((FileName), 1, Len((FileName)) - 3) & "xls", _
                                     FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                                     ReadOnlyRecommended:=False, CreateBackup:=False
   
   XL.Quit
   Set XL = Nothing
   
   If Dir(Mid((FileName), 1, Len((FileName)) - 3) & "txt") <> "" Then
   
      Kill Mid((FileName), 1, Len((FileName)) - 3) & "txt"
      
   End If
   
   Call MsgBox("Se generó archivo para contabilización de inventario en OPTIMUM." & Chr(13) & "Por favor envíe el archivo mediante correo a su ejecutivo contable, éste se encuentra en carpeta" & Chr(13) & dir_trabajo_Inf & "InformesAXInventario\InventarioAX_" & Ceco & "_" & periodommaa & "_" & hora & ".xls", vbInformation)

Else
   
   fg_descarga
   GeneraInvAX = False

End If

RS.Close
Set RS = Nothing

Exit Function
error:
    fg_descarga
    GeneraInvAX = False
    
    XL.Quit
    Set XL = Nothing
    
    MsgBox Err.Description, vbCritical

End Function

Function GeneraCfcAX(Folio As Long, periodo As String, LugarFisico As String) As Boolean

On Error GoTo error

Dim RS As New ADODB.Recordset
Dim RSAx As New ADODB.Recordset
Dim Ceco As String
Dim periodommaa As String
Dim AnJes As Object
Dim CencosAx  As String
Dim CencosSgp As String
Dim hora As String
Dim Glosa As String
Dim i As Long

GeneraCfcAX = False
           
'--> Validar homologación Ceco y cuentas contable AX
Ceco = MuestraCasino(1)
Set RS = vg_db.Execute("SELECT 1 FROM Cecos_Sap_AX csa with (nolock) inner join b_clientes bc on bc.cli_codigo = csa.Cecos_Sap and bc.cli_socsap = csa.Sociedad_Sap and bc.cli_activo = '1' and bc.cli_tipo = 0 and bc.cli_codbod > 0 WHERE bc.cli_codigo = '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación Ceco con AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
RS.Close
Set RS = Nothing

Set RS = vg_db.Execute("sgp_Sel_ValidarCuentasAx '" & Ceco & "'")
If Not RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación cuentas contables AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
RS.Close
Set RS = Nothing
           
GeneraCfcAX = True

Set RS = vg_db.Execute("sgp_Sel_GenerarArchivoCfcAX_01 '" & Ceco & "',  " & Folio & ", '" & periodo & "'")
If Not RS.EOF Then
    
    Set AnJes = CreateObject("scripting.filesystemobject")
    If Not AnJes.FolderExists(dir_trabajo_Inf & "InformesAXFacturacion") Then
       
       Call AnJes.CreateFolder(dir_trabajo_Inf & "InformesAXFacturacion")
    
    End If
    
    CencosSgp = Trim(Ceco)
    hora = Format(Time$, "HHMMSS")
    
    periodommaa = IIf(periodo = "", 0, Mid(periodo, 5, 2) & Mid(periodo, 1, 4))
    
    If (AnJes.FileExists(dir_trabajo_Inf & "InformesAXFacturacion\FacturacionAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".csv")) Then
    Kill (dir_trabajo_Inf & "InformesAXFacturacion\FacturacionAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".csv")
    End If
    
    Open dir_trabajo_Inf & "InformesAXFacturacion\FacturacionAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".csv" For Append As #9
    Close #9
    Open dir_trabajo_Inf & "InformesAXFacturacion\FacturacionAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".csv" For Append As #9
    
    If RSAx.State = 1 Then RSAx.Close
    
    Set RSAx = vg_db.Execute("sgp_Ins_GrabaLogFacAX 'FacturacionAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".csv" & "', '" & CencosSgp & "', " & periodo & ",  " & Folio & ", 'C'")
    If Not RSAx.EOF Then
       
       If RSAx(0) > 0 Then
          
          fg_descarga
          MsgBox RSAx(0) & " " & RSAx(1), vbCritical + vbOKOnly, MsgTitulo
          RSAx.Close: Set RSAx = Nothing
          Exit Function
       
       End If
    
    End If
    RSAx.Close: Set RSAx = Nothing
    
    Do While Not RS.EOF
    
       Glosa = ""
       Glosa = Glosa & Trim(RS("Cecos_AX")) & ";" ' Profit Center (*)
       Glosa = Glosa & RS("Cuentas_AX") & ";" ' Account number (*)
       Glosa = Glosa & (RS("cfc_mtodoc")) & ";"  ' Amount (*)
       Glosa = Glosa & Trim(RS("FecEnv")) & ";" ' Document  Date (*)
       Glosa = Glosa & "" & ";" ' ORDER (keep empty)
       Glosa = Glosa & Trim(Mid(RS("cfc_glosa"), 1, 60)) & ";" ' TEXT (*)
       Glosa = Glosa & RS("NumDoc") & ";" ' INVOICE (keep empty)
       Glosa = Glosa & "" & ";" ' EMPLOYEE (keep empty)
       Glosa = Glosa & RS("CURRENT1") & ";" ' Currency
       Glosa = Glosa & "" & ";" ' DUE DATE
       Glosa = Glosa & RS("GlosaImpuesto") & ";" ' Sales Tax Group (* for P&L)
       Glosa = Glosa & RS("imp_nombre") & ";" ' Sales Tax Group (* for P&L)
       Glosa = Glosa & RS("MonAdi") & ";"  ' Item Sales Tax Group (* for P&L)
       Glosa = Glosa & "" & ";" ' Tax Amount (* for P&L)
       Glosa = Glosa & RS("DetAsiento") & ";" ' Offset Account (*)
       Glosa = Glosa & RS("Optional16") & ";" ' Detail
       Glosa = Glosa & LugarFisico & ";"  ' Fecha de Recepcion (Nuevo)
       Glosa = Glosa & RS("Optional18") & ";" ' Factura referencia
       Glosa = Glosa & "" & ";" ' Tipo de documento
       Glosa = Glosa & "" & ";" ' Purch doc.Type
       Glosa = Glosa & "" & ";" ' Posting Profile
       Print #9, Glosa
    
       RS.MoveNext
   
   Loop
   Close #9
   fg_descarga
   Call MsgBox("Se generó archivo para contabilización CFC en OPTIMUM." & Chr(13) & "Por favor envíe el archivo mediante correo a su ejecutivo contable, éste se encuentra en carpeta" & Chr(13) & dir_trabajo_Inf & "InformesAXFacturacion\FacturacionAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".csv", vbInformation)

Else
   
   GeneraCfcAX = False
   fg_descarga

End If
RS.Close
Set RS = Nothing

Exit Function
error:
    fg_descarga
    GeneraCfcAX = False
    MsgBox Err.Description, vbCritical

End Function

Function GenerarTraspasoSalidaAX(Folio As Long, periodo As String) As Boolean

On Error GoTo error

Dim RS           As New ADODB.Recordset
Dim RSAx         As New ADODB.Recordset
Dim Ceco         As String
Dim periodommaa  As String
Dim AnJes        As Object
Dim CencosAx     As String
Dim NomCeco      As String
Dim ResCeco      As String
Dim CencosSgp    As String
Dim hora         As String
Dim Glosa        As String
Dim Sociedad     As String
Dim i            As Long
Dim MonTotal     As Double
Dim XL           As Object 'Excel.Application
Dim FileName     As String
Dim FileNameOslo As String
Dim NomHoja      As String
Dim EstCeco      As Boolean

GenerarTraspasoSalidaAX = False
           
'--> Validar homologación Ceco y cuentas contable AX
Ceco = MuestraCasino(1)
     
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT isnull(csa.Cecos_AX,'') as Cecos_AX, isnull(bc.cli_nombre,'') as cli_nombre, isnull(bc.cli_percon,'') as cli_percon FROM Cecos_Sap_AX csa with (nolock) inner join b_clientes bc on bc.cli_codigo = csa.Cecos_Sap and bc.cli_socsap = csa.Sociedad_Sap and bc.cli_activo = '1' and bc.cli_tipo = 0 and bc.cli_codbod > 0 WHERE bc.cli_codigo = '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación Ceco con AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If

CencosAx = RS!Cecos_AX
NomCeco = RS!cli_nombre
ResCeco = RS!cli_percon

RS.Close
Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarCuentasAx '" & Ceco & "'")
If Not RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación cuentas contables AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
RS.Close
Set RS = Nothing
           
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Sociedad '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe sociedad AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
Sociedad = RS!Sociedad_AX
RS.Close
Set RS = Nothing
           
GenerarTraspasoSalidaAX = True
MonTotal = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ResumenTraspasoSalidaAX '" & Ceco & "',  " & Folio & ", " & vg_codbod & "")
If Not RS.EOF Then
    
    Set AnJes = CreateObject("scripting.filesystemobject")
    If Not AnJes.FolderExists(dir_trabajo_Inf & "InformesAXFacturacionManual") Then
       
       Call AnJes.CreateFolder(dir_trabajo_Inf & "InformesAXFacturacionManual")
    
    End If
    
    CencosSgp = Trim(Ceco)
    hora = Format(Time$, "HHMMSS")
    
    periodommaa = IIf(periodo = "", 0, Mid(periodo, 5, 2) & Mid(periodo, 1, 4))
    FileName = "TraspasoSalidaAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora
    
    If (AnJes.FileExists(dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".txt")) Then
       
       Kill (dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".txt")
    
    End If
    
    Open dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".txt" For Append As #9
    Close #9
    Open dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".txt" For Append As #9
    
    If RSAx.State = 1 Then RSAx.Close

    EstCeco = True
    
    Set RSAx = vg_db.Execute("sgp_Ins_GrabaLogFacAX '" & FileName & ".txt" & "', '" & CencosSgp & "', " & periodo & ",  " & Folio & ", 'C'")
    If Not RSAx.EOF Then

       If RSAx(0) > 0 Then

          fg_descarga
          MsgBox RSAx(0) & " " & RSAx(1), vbCritical + vbOKOnly, MsgTitulo
          RSAx.Close: Set RSAx = Nothing
          Exit Function

       End If

    End If
    RSAx.Close: Set RSAx = Nothing
    
    Glosa = ""
    Print #9, Glosa
    Print #9, Glosa
    Print #9, Glosa
    
    Glosa = "" & ";" & "Profit center " & ";"
    Glosa = Glosa & CencosAx
    Print #9, Glosa
    
    Glosa = ""
    Glosa = "" & ";" & "Descripción PC" & ";"
    Glosa = Glosa & NomCeco & ";" & "Traspasos entre Profit center"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = "" & ";" & "Sociedad" & ";"
    Glosa = Glosa & Sociedad
    Print #9, Glosa
    
    Glosa = ""
    Glosa = "" & ";" & "Responsable" & ";"
    Glosa = Glosa & ResCeco
    Print #9, Glosa
    
    Glosa = ""
    Print #9, Glosa
     
    Glosa = ""
    Glosa = "" & ";" & "Profit Center Destino" & ";" & "Descripción Ceco" & ";" & "N° Documento/Guía" & ";" & "Monto Traspasado" & ";" & "CUENTA CONTABLE" & ";" & "DESCRIPCION CUENTA"
    Print #9, Glosa
    
    Do While Not RS.EOF
    
       Glosa = ""
       Glosa = "" & ";" & Glosa & IIf(Trim(RS("Ceco Correcto")) = "0", "", Trim(RS("Profit Center Destino"))) & ";" ' Profit Center Destino
       Glosa = Glosa & RS("Descripcion Ceco") & ";" ' Descripción Ceco
       Glosa = Glosa & (RS("Numero Documento")) & ";"  ' N° Documento/Guía
       Glosa = Glosa & RS("Monto Traspaso") & ";" ' Monto Traspasado
       Glosa = Glosa & RS("cuenta contable") & ";" ' CUENTA CONTABLE
       Glosa = Glosa & Trim(RS("Descripcion Cuenta")) & ";" ' DESCRIPCION CUENTA
'       glosa = glosa & Trim(RS("Ceco Correcto")) & ";" 'Ceco Correcto = 1; Error = 0
       
       If Trim(RS("Ceco Correcto")) = "0" Then
       
          EstCeco = False
       
       End If
       
       Print #9, Glosa
    
       MonTotal = MonTotal + RS("Monto Traspaso")
       
       RS.MoveNext
   
   Loop
   
   Glosa = ""
   Print #9, Glosa
    
   Glosa = ""
   Print #9, Glosa
    
   Glosa = ""
   Glosa = "" & ";" & "" & ";" & "Total Traspasos entre Profit Center" & ";" & "" & ";" & MonTotal
   Print #9, Glosa
   
   Close #9
   fg_descarga
   
   Set XL = CreateObject("Excel.application")
   'FileName = dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".txt"
   XL.Workbooks.OpenText dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".txt", , 1, 1, , , , , , , True, ";"
   
   XL.ActiveWorkbook.SaveAs FileName:=dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".xls", _
                                     FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                                     ReadOnlyRecommended:=False, CreateBackup:=False
   
   XL.Quit
   Set XL = Nothing
   
   NomHoja = FileName '"TraspasoSalidaAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora
   FileNameOslo = FileName & ".xls"
   FileName = dir_trabajo_Inf & "InformesAXFacturacionManual\" & FileName & ".xls"
   
'   If RevisarArchivoExcel(FileName, FileNameOslo, NomHoja) Then
    If EstCeco Then
   
      Call MsgBox("Se generó archivo para contabilización Traspaso Salida en OPTIMUM." & Chr(13) & "Por favor envíe el archivo mediante correo a su ejecutivo contable, éste se encuentra en carpeta" & Chr(13) & dir_trabajo_Inf & "InformesAXFacturacionManual\TraspasoSalidaAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".xls", vbInformation)
   
   Else
   
      Call MsgBox("Se generó archivo con problema la columna Profit Center Destino esta con error de código Ceco destino, modifique el archivo antes de ser enviado" & Chr(13) & "Por favor envíe el archivo mediante correo a su ejecutivo contable, éste se encuentra en carpeta" & Chr(13) & dir_trabajo_Inf & "InformesAXFacturacionManual\TraspasoSalidaAX_" & Ceco & "_" & periodommaa & "_" & Folio & "_" & hora & ".xls", vbInformation)
   
   End If
   
   If Dir(Mid((FileName), 1, Len((FileName)) - 3) & "txt") <> "" Then
   
      Kill Mid((FileName), 1, Len((FileName)) - 3) & "txt"
      
   End If
  

Else
   
   GenerarTraspasoSalidaAX = False
   fg_descarga

End If
RS.Close
Set RS = Nothing

Exit Function
error:
    fg_descarga
    GenerarTraspasoSalidaAX = False
    
    XL.Quit
    Set XL = Nothing

    MsgBox Err.Description, vbCritical

End Function

Function RevisarArchivoExcel(NomArchivo As String, SoloNomArchivo As String, SheetName As String) As Boolean

On Error GoTo error

Dim objArchivoXls As Object
Dim co1 As Long
Dim intUltimo As Long
Const xlDown As Integer = -4121

RevisarArchivoExcel = True

If Len(Dir(NomArchivo)) > 0 Then  'la ruta de tu libro el mio se llama movimiento

    Set objArchivoXls = GetObject(NomArchivo) ' una vez mas la ruta
    objArchivoXls.Worksheets(Mid(SheetName, 1, 31)).Activate  ' en este combo podes listar las hojas que tenga tu libro nota: el combo no lo llena el libro

    With objArchivoXls.ActiveSheet
        
        .Parent.Windows(SoloNomArchivo).Visible = True
        intUltimo = .Range("A1").End(xlDown).Row + 1
        
        For co1 = 1 To intUltimo + 0

            'Fin de archivo
            If .Cells(co1, 3).Value = "Total Traspasos entre Profit Center" Then
    
                Exit For
    
            End If
    
            'validar si existe errores codigo ceco
            If .Cells(co1, 8).Value = "0" Then
       
               .Cells(co1, 2).Interior.ColorIndex = 3
               RevisarArchivoExcel = False
            
            End If
            
            .Cells(co1, 8).Value = ""
            
        Next co1
    End With
    
    objArchivoXls.Save
    objArchivoXls.Parent.Quit
    Set objArchivoXls = Nothing

Else
    
    MsgBox "Archivo no existe"

End If

Exit Function
error:
    fg_descarga
    RevisarArchivoExcel = False
    MsgBox Err.Description, vbCritical
    objArchivoXls.Parent.Quit
    Set objArchivoXls = Nothing
End Function

Function GeneraCfcDigitado(Folio As Long, periodo As String, Inf_Tipo As String) As Boolean

On Error GoTo error

Dim RS           As New ADODB.Recordset
Dim RSAx         As New ADODB.Recordset
Dim Ceco         As String
Dim periodommaa  As String
Dim AnJes        As Object
Dim CencosAx     As String
Dim NomCencosAX  As String
Dim ContactoCeco As String
Dim CencosSgp    As String
Dim hora         As String
Dim Glosa        As String
Dim i            As Long
Dim FileName     As String
Dim XL           As Object 'Excel.Application
Dim TotalDocu    As Double
Dim TotalAlim    As Double
Dim TotalLim     As Double
Dim TotalOtr     As Double
Dim Tipo_Doc     As String

GeneraCfcDigitado = False
           
'--> Validar homologación Ceco y cuentas contable AX
Ceco = MuestraCasino(1)
Set RS = vg_db.Execute("SELECT csa.Cecos_AX, bc.cli_nombre, bc.cli_percon FROM Cecos_Sap_AX csa with (nolock) inner join b_clientes bc on bc.cli_codigo = csa.Cecos_Sap and bc.cli_socsap = csa.Sociedad_Sap and bc.cli_activo = '1' and bc.cli_tipo = 0 and bc.cli_codbod > 0 WHERE bc.cli_codigo = '" & Ceco & "'")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación Ceco con AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
CencosAx = Trim(RS!Cecos_AX)
NomCencosAX = Trim(RS!cli_nombre)
ContactoCeco = IIf(IsNull(RS!cli_percon), "", RS!cli_percon)

RS.Close
Set RS = Nothing

Set RS = vg_db.Execute("sgp_Sel_ValidarCuentasAx '" & Ceco & "'")
If Not RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe homologación cuentas contables AX..., Proceso Cancelado", vbCritical + vbOKOnly, MsgTitulo
   Exit Function

End If
RS.Close
Set RS = Nothing
           
GeneraCfcDigitado = True

Tipo_Doc = IIf(Inf_Tipo = "P", "CFCPortalElectronico_", "CFCManual_")

Set RS = vg_db.Execute("sgp_Sel_GeneraCfcDigitado_V01 '" & Ceco & "',  " & Folio & ", '" & periodo & "', '" & Inf_Tipo & "'")
If Not RS.EOF Then
    
    Set AnJes = CreateObject("scripting.filesystemobject")
    If Not AnJes.FolderExists(dir_trabajo_Inf & "InformesAXFacturacionManual") Then
       
       Call AnJes.CreateFolder(dir_trabajo_Inf & "InformesAXFacturacionManual")
    
    End If
    
    CencosSgp = Trim(Ceco)
    hora = Format(Time$, "HHMMSS")
    
    periodommaa = IIf(periodo = "", 0, Mid(periodo, 5, 2) & Mid(periodo, 1, 4))
    
    If (AnJes.FileExists(dir_trabajo_Inf & "InformesAXFacturacionManual\" & Tipo_Doc & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".txt")) Then
       
       Kill (dir_trabajo_Inf & "InformesAXFacturacionManual\" & Tipo_Doc & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".txt")
    
    End If
    
    FileName = dir_trabajo_Inf & "InformesAXFacturacionManual\" & Tipo_Doc & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".txt"
    
    Open dir_trabajo_Inf & "InformesAXFacturacionManual\" & Tipo_Doc & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".txt" For Append As #9
    Close #9
    Open dir_trabajo_Inf & "InformesAXFacturacionManual\" & Tipo_Doc & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".txt" For Append As #9
    
    If RSAx.State = 1 Then RSAx.Close

    Set RSAx = vg_db.Execute("sgp_Ins_GrabaLogFacAX '" & Tipo_Doc & "" & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".xls" & "', '" & CencosSgp & "', " & periodo & ",  " & Folio & ", '" & Inf_Tipo & "'")
    If Not RSAx.EOF Then

       If RSAx(0) > 0 Then

          fg_descarga
          MsgBox RSAx(0) & " " & RSAx(1), vbCritical + vbOKOnly, MsgTitulo
          RSAx.Close: Set RSAx = Nothing
          Exit Function

       End If

    End If
    RSAx.Close: Set RSAx = Nothing
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Glosa = Glosa & IIf(Inf_Tipo = "P", "CFC ELECTRONICO ", "CFC MANUAL ") & Meses("01/" & Mid(periodommaa, 1, 2) & "/" & Mid(periodommaa, 3, 4)) & " " & Mid(periodommaa, 3, 4) & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";"
    Glosa = Glosa & "Mes" & ";"
    Glosa = Glosa & Meses("01/" & Mid(periodommaa, 1, 2) & "/" & Mid(periodommaa, 3, 4)) & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Glosa = Glosa & "CFC N°" & ";"
    Glosa = Glosa & Folio & ";"
    Glosa = Glosa & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";"
    Glosa = Glosa & "Año" & ";"
    Glosa = Glosa & Mid(periodommaa, 3, 4) & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Glosa = Glosa & "Profit Center" & ";"
    Glosa = Glosa & CencosAx & ";"
    Glosa = Glosa & NomCencosAX & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Glosa = Glosa & "BRU" & ";"
    Glosa = Glosa & "098" & ";"
    Glosa = Glosa & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";"
    Glosa = Glosa & "Contacto en sitio" & ";"
    Glosa = Glosa & ContactoCeco & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "" & ";"
    Print #9, Glosa
    
    TotalDocu = 0
    TotalAlim = 0
    TotalLim = 0
    TotalOtr = 0
    
    Do While Not RS.EOF
    
       TotalDocu = TotalDocu + RS![Total Factura]
       TotalAlim = TotalAlim + RS![Alimentos 60111100]
       TotalLim = TotalLim + RS![Desechables 60261100]
       TotalOtr = TotalOtr + RS![Otros Gasto]
        
       RS.MoveNext
    Loop
    
    Glosa = ""
    Glosa = Glosa & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";" & "" & ";"
    Glosa = Glosa & IIf(TotalDocu = 0, "-", TotalDocu) & ";" & IIf(TotalAlim = 0, "-", TotalAlim) & ";" & IIf(TotalLim = 0, "-", TotalLim) & ";" & IIf(TotalOtr = 0, "-", TotalOtr) & ";"
    Print #9, Glosa
    
    Glosa = ""
    Glosa = Glosa & "RUT PROVEEDOR" & ";" & "RAZON SOCIAL PROVEEDOR" & ";" & "FOLIO FACTURA" & ";" & "FECHA EMISIÓN" & ";" & "IMPUESTO ADICIONAL ILA" & ";" & "IMPUESTO ADICIONAL CARNE" & ";" & "IMPUESTO ADICIONAL HARINA" & ";" & "TOTAL FACTURA" & ";" & "ALIMENTOS 60111100" & ";" & "DESECHABLES 60261100" & ";" & "OTRO GASTO" & ";" & "CUENTA CONTABLE OTRO GASTO" & ";" & "DESCRIPCION CUENTA" & ";" & "SOLICITUD ACTIVO FIJO" & ";" & "N° DE SOLICITUD INVERSION A.F" & ";" & "TIPO DOCUMENTO" & ";" & "CÓDIGO DOCUMENTO" & ";" & "OBSERVACIONES" & ";"
    Print #9, Glosa
    
    
    RS.MoveFirst
    Do While Not RS.EOF
    
       Glosa = ""
       Glosa = Glosa & Trim(RS("Rut Proveedor")) & ";" ' RUT PROVEEDOR
       Glosa = Glosa & RS("Razon Social Proveedor") & ";" ' RAZON SOCIAL PROVEEDOR
       Glosa = Glosa & (RS("Folio Factura")) & ";"  ' FOLIO FACTURA
       Glosa = Glosa & " " & Mid(Trim(RS("Fecha Emision")), 1, 2) & "-" & Mid(Trim(RS("Fecha Emision")), 3, 2) & "-" & Mid(Trim(RS("Fecha Emision")), 5, 4) & ";" ' FECHA EMISIÓN
       Glosa = Glosa & IIf(RS("Impuesto Adicional Ila") = 0, "", RS("Impuesto Adicional Ila")) & ";" ' IMPUESTO ADICIONAL ILA
       Glosa = Glosa & IIf(RS("Impuesto Adicional Carne") = 0, "", RS("Impuesto Adicional Carne")) & ";" ' IMPUESTO ADICIONAL CARNE
       Glosa = Glosa & IIf(RS("Impuesto Adicional Harina") = 0, "", RS("Impuesto Adicional Harina")) & ";" ' IMPUESTO ADICIONAL HARINA
       Glosa = Glosa & RS("Total Factura") & ";" ' TOTAL FACTURA
       Glosa = Glosa & IIf(RS("Alimentos 60111100") = 0, "", RS("Alimentos 60111100")) & ";" ' ALIMENTOS 60111100
       Glosa = Glosa & IIf(RS("Desechables 60261100") = 0, "", RS("Desechables 60261100")) & ";" ' DESECHABLES 60261100
       Glosa = Glosa & IIf(RS("Otros Gasto") = 0, "", RS("Otros Gasto")) & ";" ' OTRO GASTO
       Glosa = Glosa & RS("Cuenta Contable Otro Gasto") & ";" ' CUENTA CONTABLE OTRO GASTO
       Glosa = Glosa & RS("Descripcion Cuenta") & ";" ' DESCRIPCION CUENTA
       Glosa = Glosa & RS("Solicitud Activo Fijo") & ";" ' N° DE SOLICITUD INVERSION A.F
       Glosa = Glosa & RS("N° De Solicitud Inversion A.F") & ";" ' N° DE SOLICITUD INVERSION A.F
       Glosa = Glosa & RS("Tipo Documento") & ";" ' TIPO DOCUMENTO
       Glosa = Glosa & RS("Codigo Documento") & ";" ' CÓDIGO DOCUMENTO
       Glosa = Glosa & RS("OBSERVACIONES") & ";" ' OBSERVACIONES
       Print #9, Glosa
    
       RS.MoveNext
   
   Loop
   Close #9
   Set XL = CreateObject("Excel.application")
   XL.Workbooks.OpenText Mid((FileName), 1, Len((FileName)) - 3) & "txt", , 1, 1, , , , , , , True, ";"
   
   XL.Range("D11:D40").Select
   XL.Selection.NumberFormat = "m/d/yyyy"
    
    XL.Range("H9:K9").Select
    XL.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    XL.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With XL.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    XL.Range("A10:R40").Select
    XL.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    XL.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With XL.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With XL.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    XL.ActiveWindow.Zoom = 82
   
   XL.ActiveWorkbook.SaveAs FileName:=Mid((FileName), 1, Len((FileName)) - 3) & "xls", _
                                     FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
                                     ReadOnlyRecommended:=False, CreateBackup:=False
   
   XL.Quit
   Set XL = Nothing
   
   If Dir(Mid((FileName), 1, Len((FileName)) - 3) & "txt") <> "" Then
   
      Kill Mid((FileName), 1, Len((FileName)) - 3) & "txt"
      
   End If
   
   fg_descarga
   Call MsgBox("Se generó archivo para contabilización " & IIf(Inf_Tipo = "P", "CFC PORTAL ELECTRONICO", "CFC en MANUAL.") & Chr(13) & "Por favor envíe el archivo mediante correo a su ejecutivo contable, éste se encuentra en carpeta" & Chr(13) & dir_trabajo_Inf & "InformesAXFacturacionManual\" & Tipo_Doc & CencosAx & "_" & periodommaa & "_" & Folio & "_" & hora & ".xls", vbInformation)

Else
   
   GeneraCfcDigitado = False
   fg_descarga

End If
RS.Close
Set RS = Nothing

Exit Function
error:
    fg_descarga
    GeneraCfcDigitado = False
    
    XL.Quit
    Set XL = Nothing

    MsgBox Err.Description, vbCritical

End Function

Function ExplorarCarpeta(ByVal Ruta As String)

On Error GoTo error

Dim R As Long
Dim AnJes As Object

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(Ruta) Then
   Call AnJes.CreateFolder(Ruta)
End If
   
R = ShellExecute(0, "open", Ruta, 0, 0, 1)

Exit Function
error:
    fg_descarga
    MsgBox Err.Description, vbCritical

End Function

Function ValidaPCServidorVacio() As Boolean
    
On Error GoTo error

    ValidaPCServidorVacio = True
    Dim varNombreServidor As String
    
    varNombreServidor = Trim(GetParametro("SvrAppCont"))
    
    If Trim(varNombreServidor) = "" Then
        
        ValidaPCServidorVacio = False
    
    End If
        
Exit Function
error:
    fg_descarga
    MsgBox Err.Description, vbCritical

End Function

Function ValidaPCServidor() As Boolean
    
On Error GoTo error

    Dim varNombreEquipo, varNombreServidor As String
    Dim sEquipo As String * 255
    GetComputerName sEquipo, 255
    
    varNombreEquipo = Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1)
    varNombreServidor = Trim(GetParametro("SvrAppCont"))
    
    If varNombreEquipo = varNombreServidor Then
        
        ValidaPCServidor = True
    
    Else
        
        ValidaPCServidor = False
    
    End If
        
Exit Function
error:
    fg_descarga
    MsgBox Err.Description, vbCritical

End Function

Public Function ObtenerMACcomputadora() As String

On Error GoTo error

Dim colNetAdapters, objWMIService As Object
Dim strComputer As String
Dim objitem
strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colNetAdapters = objWMIService.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")

For Each objitem In colNetAdapters
    
    ObtenerMACcomputadora = objitem.MACAddress

Next

Exit Function
error:
    fg_descarga
    MsgBox Err.Description, vbCritical

End Function

Function ValidaTraspasodeSalida(Ceco As String, Bodega As Long, Folio As Long) As Boolean
    
On Error GoTo error

ValidaTraspasodeSalida = True
    
Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ResumenTraspasoSalidaAX '" & Ceco & "', " & Folio & ", " & Bodega & " ")
If RS.EOF Then
   
   ValidaTraspasodeSalida = False

End If
        
RS.Close
Set RS = Nothing

Exit Function
error:
    If RS.State = 1 Then RS.Close
    fg_descarga
    MsgBox Err.Description, vbCritical
    ValidaTraspasodeSalida = False

End Function

Function IsFormLoaded(FormToCheck As Form) As Integer

Dim y As Integer

For y = 0 To Forms.count - 1
    
    If Forms(y) Is FormToCheck Then
       
       IsFormLoaded = True
       Exit Function
    
    End If

Next

IsFormLoaded = False
 
End Function

Function fg_Archivo(RutaTrabajo As String, NombreArchivo As String)

Dim i As Long

i = 1

For i = 1 To 99999
    
    If Dir(RutaTrabajo & NombreArchivo & "_" & fg_pone_cero(Trim(Str(i)), 5)) = "" Then
        
        fg_Archivo = RutaTrabajo & NombreArchivo & "_" & fg_pone_cero(Trim(Str(i)), 5): Exit Function
    
    End If

Next i

End Function

Function Generar_ArchivoExcel(ParametroProced As String, RutaTrabajo As String, NombreArchivoExcel As String)

Dim RS       As New ADODB.Recordset
'Definición variables excel
Dim xlApp    As Object
Dim xlWb     As Object
Dim xlWs     As Object
Dim XL              As New Excel.Application 'Crea el objeto excel
'Dim NomArchivoExcel As String
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("Hoja1")

'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(ParametroProced)

'-------> Check version of Excel
Call encabezado(RS, xlWs)

xlWs.Cells(2, 1).CopyFromRecordset RS

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit

'xlApp.Columns("A:A").Select
'xlApp.Selection.Delete Shift:=xlToLeft

'NomArchivoExcel = fg_ArchivoXls("GuiaCDLogists_")
      
xlWb.Close True, RutaTrabajo & NombreArchivoExcel

XL.Workbooks.Open RutaTrabajo & NombreArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

'-- Cerrar Excel
xlApp.Quit

'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing

End Function

Function Generar_ArchivoExcelVta(ParametroProced As String, RutaTrabajo As String, NombreArchivoExcel As String)

Dim RS       As New ADODB.Recordset

'Definición variables excel
Dim xlApp    As Object
Dim xlWb     As Object
Dim xlWs     As Object
Dim XL              As New Excel.Application 'Crea el objeto excel
'Dim NomArchivoExcel As String
'-------> Create an instance of Excel and add a workbook
Set xlApp = CreateObject("Excel.Application")
Set xlWb = xlApp.Workbooks.Add
Set xlWs = xlWb.Worksheets("hoja1")
xlWb.Sheets("Hoja1").Name = "input ventas"

'-------> Display Excel and give user control of Excel's lifetime
xlApp.UserControl = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute(ParametroProced)

'xlApp.Columns("E:E").Select
'
'xlApp.Selection.HorizontalAlignment = xlRight
'
'xlApp.Selection.NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"

'-------> Check version of Excel
Call encabezadoVta(RS, xlWs)

xlWs.Cells(2, 1).CopyFromRecordset RS

'-------> Auto-fit the column widths and row heights
xlApp.Selection.CurrentRegion.Columns.AutoFit
xlApp.Selection.CurrentRegion.Rows.AutoFit

xlApp.Columns("E:E").Select
'
xlApp.Selection.HorizontalAlignment = xlRight
'
xlApp.Selection.NumberFormat = "[$-x-sysdate]dddd, mmmm dd, yyyy"
'xlApp.Selection.Delete Shift:=xlToLeft

'NomArchivoExcel = fg_ArchivoXls("GuiaCDLogists_")
      
xlWb.Close True, RutaTrabajo & NombreArchivoExcel

XL.Workbooks.Open RutaTrabajo & NombreArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
XL.Visible = True
XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

'-- Cerrar Excel
xlApp.Quit

'-------> Release Excel references
Set xlWs = Nothing
Set xlWb = Nothing
Set xlApp = Nothing

End Function

Function Unix2Dos(File) As Boolean

On Error GoTo Man_Error

Unix2Dos = True

    Dim fs As Object, Txt As String
    Set fs = CreateObject("Scripting.FileSystemObject")

    Txt = fs.OpenTextFile(File, 1).ReadAll
    Txt = Replace(Txt, vbLf, vbCrLf)
    fs.OpenTextFile(File, 2).Write Txt

Exit Function
Man_Error:
    
    Unix2Dos = False
    MsgBox Err.Description, vbCritical, MsgTitulo

End Function

Function SacarNombreArchivo(File As String, Caracter As String) As String

Dim Largo         As Long
Dim NombreArchivo As String
Dim i             As Long
Largo = Len(Trim(File))
NombreArchivo = ""

'Sacar nombre de archivo
For i = Largo To 1 Step -1
    
    If Mid$(Trim(File), i, 1) = Caracter Then Exit For
        
    NombreArchivo = Mid$(Trim(File), i, 1) & NombreArchivo
    
Next i

SacarNombreArchivo = NombreArchivo

End Function

Function fg_GrabaLogSistema(cNUsuario As String, cOpcion As Long, cOpcionSistema As String, cDatoNuevo As String, cDatoAnterior As String, CDetalleOperacion As String) As String

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Ins_logsistema_V01 '" & cNUsuario & "', " & cOpcion & ", '" & IIf(cOpcionSistema <> "MINSAN", Fg_Ponerpunto(cOpcionSistema), cOpcionSistema) & "', '" & cDatoAnterior & "', '" & cDatoNuevo & "', '" & CDetalleOperacion & "'")

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

Function fg_ValidaPassword(cUsr As String, cPass As String, cMsg As String) As Boolean

Dim RS_Dato2    As New ADODB.Recordset
Dim cPassLong   As Long
Dim cPassAnt    As Long
Dim cContPass   As Long
Dim i           As Long
Dim cCAr        As String
Dim cCantNum    As Integer
Dim cCantCar    As Integer
Dim cCantCarMay As Integer
Dim cCantEsp    As Integer
Dim cCantExep   As Integer
Dim cCarEsp     As String

fg_ValidaPassword = True

'Revisa caracteres
cCarEsp = GetParametro_Seguridad("pscara")
cCantNum = 0
cCantCar = 0
cCantEsp = 0
cCantExep = 0
cCantCarMay = 0

For i = 1 To Len(cPass)
    
    cCAr = Mid(cPass, i, 1)
    
    If (Asc(cCAr) >= 48 And Asc(cCAr) <= 57) Then
        
        cCantNum = cCantNum + 1
    
    End If
    
    If (Asc(cCAr) >= 65 And Asc(cCAr) <= 90) Then
    
       cCantCarMay = cCantCarMay + 1
    
    End If
    
    If (Asc(cCAr) >= 97 And Asc(cCAr) <= 122) Or Asc(cCAr) = 241 Or Asc(cCAr) = 209 Then
        
        cCantCar = cCantCar + 1
    
    End If
    
    If InStr(cCarEsp, cCAr) <> 0 Then
        
        cCantEsp = cCantEsp + 1
    
    End If
    
    If Not ((Asc(cCAr) >= 48 And Asc(cCAr) <= 57)) And Not ((Asc(cCAr) >= 97 And Asc(cCAr) <= 122) Or (Asc(cCAr) >= 65 And Asc(cCAr) <= 90) Or Asc(cCAr) = 241 Or Asc(cCAr) = 209) And Not (InStr(cCarEsp, cCAr) <> 0) Then
        
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
cPassLong = GetParametro_Seguridad("pslong")
If Len(cPass) < cPassLong Then
    
    MsgBox "La password debe tener una longitud mínima de " & cPassLong & " caracteres.", vbCritical + vbOKOnly, cMsg
    fg_ValidaPassword = False
    Exit Function

End If

If RS_Dato2.State = 1 Then RS_Dato2.Close
RS_Dato2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'Revisa password anteriores
cPassAnt = GetParametro_Seguridad("psante")
cContPass = 0
Set RS_Dato2 = vg_db.Execute("sgp_Sel_Log_CambiaPass 1, '" & cUsr & "', " & fg_TraeLogConcepto("vg_logsis_CambiaPass") & "")
If Not RS_Dato2.EOF Then
    
    Do While Not RS_Dato2.EOF
        
        If fg_Encripta(cPass) = RS_Dato2!dato_nuevo Then
            
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
RS_Dato2.Close
Set RS_Dato2 = Nothing

End Function

Function fg_TraeLogConcepto(referencia As String) As Long

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_TraeLogConcepto = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_TraeLogConcepto '" & referencia & "'")

If Not RS.EOF Then
   
   fg_TraeLogConcepto = TipoDato(RS!loc_Id, "")
   
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

Function fg_pone_rchar(ByVal cadena As String, ByVal cuanto As Integer, ByVal char As String) As String

'pone caracteres a la derecha
fg_pone_rchar = Trim(cadena) & String(cuanto - Len(Trim(cadena)), char)

End Function
