VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RutinaLectura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Segmento(ByVal opcion As Long, ByVal codseg As Long, ByVal descri As String) As String
Dim sql1 As String
Segmento = ""
Select Case opcion
Case 1
    Segmento = "SELECT seg_codigo FROM a_segmento ORDER BY seg_codigo DESC"
Case 2
    Segmento = "SELECT * FROM a_segmento WHERE (seg_codigo = " & codseg & " OR " & codseg & " = 0) ORDER BY seg_codigo"
Case 3
    Segmento = "SELECT seg_codigo, seg_nombre FROM a_segmento WHERE seg_codigo LIKE '%" & UCase(descri) & "%'"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(seg_nombre) ", " UPPER(seg_nombre) ")
    Segmento = "SELECT seg_codigo, seg_nombre FROM a_segmento WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Minutas(ByVal opcion As Long, ByVal codreg As Long, ByVal codser As Long, ByVal Fecha As Long, ByVal tipmin As String) As String
Dim sql1 As String
Minutas = ""
Select Case opcion
Case 1
    sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    Minutas = "SELECT DISTINCT a.min_cencos " & _
              "FROM  b_minuta a, b_minutadet b " & _
              "WHERE a.min_codigo = b.mid_codigo " & _
              "AND   a.min_cencos = '" & MuestraCasino(1) & "' " & _
              "AND   a.min_codreg = " & codreg & " " & _
              "AND   a.min_codser = " & codser & " " & _
              "AND " & sql1 & "   = " & Fecha & " " & _
              "AND   b.mid_tipmin = '" & tipmin & "'"
Case 2
    sql1 = IIf(vg_tipbase = "1", " val(mid(b.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) ")
    Minutas = "SELECT COUNT(a.mid_codigo) AS nreg FROM b_minutadet a, b_minuta b WHERE " & _
              "b.min_codigo = a.mid_codigo AND " & sql1 & " = " & Fecha & " AND a.mid_tipmin = '" & tipmin & "'"
Case 3
   sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
   Minutas = "SELECT DISTINCT b.mid_fecval FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = " & Fecha & " AND b.mid_fecval > 0"
Case 4
    sql1 = IIf(vg_tipbase = "1", " val(mid(b.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) ")
    Minutas = "SELECT b.min_fecmin, b.min_indblo " & _
              "FROM b_minuta b, b_minutadet c " & _
              "WHERE b.min_codigo = c.mid_codigo " & _
              "AND   b.min_cencos = '" & MuestraCasino(1) & "' AND b.min_codreg = " & codreg & " AND b.min_codser = " & codser & " " & _
              "AND   " & sql1 & " = " & Fecha & " AND c.mid_tipmin = '" & tipmin & "' ORDER BY b.min_fecmin"
Case 5
    sql1 = IIf(vg_tipbase = "1", " val(mid(b.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) ")
'    Minutas = "SELECT c.mid_tipmin, c.mid_numlin, c.mid_codrec, c.mid_descri, c.mid_cosrec, c.mid_cosdes, c.mid_rec5eta, b.min_fecmin, b.min_indblo, " & _
'              "       a.rec_nombre, c.mid_numrac, c.mid_estser, c.mid_tiprec, d.ess_nombre FROM b_receta a, b_minuta b, b_minutadet c, a_estservicio d " & _
'              "WHERE  b.min_codigo = c.mid_codigo AND c.mid_codrec = a.rec_codigo AND c.mid_estser = d.ess_codigo AND b.min_cencos = d.ess_cencos " & _
'              "AND    b.min_cencos = '" & MuestraCasino(1) & "' AND b.min_codreg = " & codreg & " AND b.min_codser = " & codser & " " & _
'              "AND    " & sql1 & " = " & Fecha & " AND c.mid_tipmin = '" & tipmin & "' ORDER BY c.mid_estser, b.min_fecmin, c.mid_numlin"
    Minutas = "sgp_Sel_TraerDetalleMinuta '" & MuestraCasino(1) & "', " & codreg & ", " & codser & ", " & Fecha & ", '" & tipmin & "'"
Case 6
    sql1 = IIf(vg_tipbase = "1", " val(mid(min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),min_fecmin),1,6)) ")
    Minutas = "SELECT min_racteo, min_racrea, min_fecmin FROM b_minuta " & _
              "WHERE  min_cencos = '" & MuestraCasino(1) & "' AND min_codreg = " & codreg & " AND min_codser = " & codser & " " & _
              "AND    " & sql1 & " = " & Fecha & " ORDER BY min_fecmin"
Case 7
     Minutas = "SELECT b.mid_fecval FROM  b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' " & _
               "AND a.min_codreg = " & codreg & " AND a.min_codser = " & codser & " AND a.min_fecmin = " & Fecha & " " & _
               "AND b.mid_tipmin = '" & tipmin & "' AND b.mid_fecval > 0"
Case 8
    sql1 = IIf(vg_tipbase = "1", " VAL(MID(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    Minutas = "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b " & _
              "WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' AND a.min_codreg = " & codreg & " AND a.min_codser = " & codser & " " & _
              "AND   " & sql1 & " = " & Fecha & " AND b.mid_tipmin = '" & tipmin & "'"
Case 9
    sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    Minutas = "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b " & _
              "WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' AND a.min_codreg = " & codreg & " AND a.min_codser = " & codser & " " & _
               "AND   " & sql1 & " = " & Fecha & " AND b.mid_tipmin = '" & tipmin & "'"
Case 10
    Minutas = "SELECT a.min_codigo FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & MuestraCasino(1) & "' " & _
              "AND a.min_codreg = " & codreg & " AND a.min_codser = " & codser & " AND a.min_fecmin = " & Fecha & " AND b.mid_tipmin = '" & tipmin & "'"
Case 11
    Minutas = "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC"
Case 12
'    Minutas = "SELECT DISTINCT b.reg_codigo, b.reg_nombre, d.reg_codigo activoreg, c.ser_codigo, d.ser_codigo activoser, c.ser_nombre FROM b_minuta a inner join a_regimen b on a.min_codreg = b.reg_codigo inner join a_servicio c on a.min_codser = c.ser_codigo left outer join a_pto_lectura_vales_servicio d on b.reg_codigo = d.reg_codigo and c.ser_codigo = d.ser_codigo WHERE a.min_cencos = '" & MuestraCasino(1) & "' ORDER BY b.reg_codigo, b.reg_nombre, c.ser_codigo, c.ser_nombre"
    Minutas = "SELECT DISTINCT " & _
              "b.reg_codigo " & _
              ", b.reg_nombre " & _
              ", d.reg_codigo activoreg " & _
              ", c.ser_codigo " & _
              ", d.ser_codigo activoser " & _
              ", c.ser_nombre " & _
              "FROM    b_minuta a " & _
              "INNER JOIN a_regimen b ON a.min_codreg = b.reg_codigo " & _
              "INNER JOIN a_servicio c ON a.min_codser = c.ser_codigo " & _
              "LEFT OUTER JOIN a_pto_lectura_vales_servicio d ON b.reg_codigo = d.reg_codigo " & _
                                                              " AND c.ser_codigo = d.ser_codigo " & _
              "LEFT OUTER JOIN a_pto_lectura_vales e ON d.lec_codlecvales = e.lec_codlecvales OR  e.lec_nombrepc like '" & tipmin & "' " & _
              "WHERE   a.min_cencos = '" & MuestraCasino(1) & "' and c.ser_activo = '1' " & _
              "ORDER BY b.reg_codigo " & _
              ", b.reg_nombre " & _
              ", c.ser_codigo " & _
              ", c.ser_nombre"
Case 13
20121029 Minutas = "SELECT DISTINCT " & _
          "b.reg_codigo, b.reg_nombre, b.reg_codigo activoreg, c.ser_codigo, " & _
          "c.ser_codigo activoser, c.ser_nombre, d.ate_codatencion " & _
          "FROM    b_minuta a " & _
          "        INNER JOIN a_regimen b ON a.min_codreg = b.reg_codigo " & _
          "        INNER JOIN a_servicio c ON a.min_codser = c.ser_codigo " & _
          "        INNER JOIN a_pto_lectura_vales_servicio d ON b.reg_codigo = d.reg_codigo " & _
          "                                                     AND c.ser_codigo = d.ser_codigo " & _
          "        INNER JOIN a_pto_atencion x ON x.ate_codatencion = d.ate_codatencion " & _
          "        INNER JOIN a_pto_lectura_vales e ON d.lec_codlecvales = e.lec_codlecvales " & _
          "                                            AND e.lec_nombrepc LIKE '" & tipmin & "' " & _
          "WHERE   a.min_cencos = '" & MuestraCasino(1) & "' " & _
          "        AND c.ser_activo = '1' " & _
          "        AND D.lec_codlecvales = " & codreg & " " & _
          "        AND d.ate_codatencion = " & codser & " " & _
          "        AND e.lec_activo = 1 "
Minutas = Minutas + "Union All "
Minutas = Minutas + "SELECT DISTINCT " & _
          "b.reg_codigo, b.reg_nombre, b.reg_codigo activoreg, c.ser_codigo, " & _
          "c.ser_codigo activoser, c.ser_nombre, 0 AS ate_codatencion " & _
          "FROM    b_minuta a " & _
          "        INNER JOIN a_regimen b ON a.min_codreg = b.reg_codigo " & _
          "        INNER JOIN a_servicio c ON a.min_codser = c.ser_codigo " & _
          "WHERE   a.min_cencos = '" & MuestraCasino(1) & "' " & _
          "        AND c.ser_activo = '1' " & _
          "        AND ( a.min_codreg NOT IN ( " & _
          "              SELECT DISTINCT " & _
          "                        aplvs.reg_codigo " & _
          "              FROM      dbo.a_pto_lectura_vales_servicio aplvs " & _
          "                        INNER JOIN a_pto_lectura_vales e ON e.lec_codlecvales = aplvs.lec_codlecvales " & _
          "                                                            AND e.lec_nombrepc LIKE '" & tipmin & "' AND e.lec_codlecvales = " & codreg & " AND aplvs.ate_codatencion = " & codser & " " & _
          "                                                            AND e.lec_activo = 1 ) " & _
          "              OR a.min_codser NOT IN ( " & _
          "              SELECT DISTINCT " & _
          "                        aplvs.ser_codigo " & _
          "              FROM      dbo.a_pto_lectura_vales_servicio aplvs " & _
          "                        INNER JOIN a_pto_lectura_vales e ON e.lec_codlecvales = aplvs.lec_codlecvales " & _
          "                                                            AND e.lec_nombrepc LIKE '" & tipmin & "' AND e.lec_codlecvales = " & codreg & " AND aplvs.ate_codatencion = " & codser & " " & _
          "                                                            AND e.lec_activo = 1 ) " & _
          "            ) " & _
          "ORDER BY b.reg_codigo, b.reg_nombre, c.ser_codigo, c.ser_nombre"
   
'20121029 Minutas = "SELECT DISTINCT " & _
'          "b.reg_codigo, b.reg_nombre, b.reg_codigo activoreg, c.ser_codigo, " & _
'          "c.ser_codigo activoser, c.ser_nombre, d.ate_codatencion " & _
'          "FROM    b_minuta a " & _
'          "        INNER JOIN a_regimen b ON a.min_codreg = b.reg_codigo " & _
'          "        INNER JOIN a_servicio c ON a.min_codser = c.ser_codigo " & _
'          "        INNER JOIN a_pto_lectura_vales_servicio d ON b.reg_codigo = d.reg_codigo " & _
'          "                                                     AND c.ser_codigo = d.ser_codigo " & _
'          "        INNER JOIN a_pto_atencion x ON x.ate_codatencion = d.ate_codatencion " & _
'          "        INNER JOIN a_pto_lectura_vales e ON d.lec_codlecvales = e.lec_codlecvales " & _
'          "                                            AND e.lec_nombrepc LIKE '" & tipmin & "' " & _
'          "WHERE   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'          "        AND c.ser_activo = '1' " & _
'          "        AND D.lec_codlecvales = " & codreg & " " & _
'          "        AND d.ate_codatencion = " & codser & " " & _
'          "        AND e.lec_activo = 1 "
'Minutas = Minutas + "Union All "
'Minutas = Minutas + "SELECT DISTINCT " & _
'          "b.reg_codigo, b.reg_nombre, b.reg_codigo activoreg, c.ser_codigo, " & _
'          "c.ser_codigo activoser, c.ser_nombre, 0 AS ate_codatencion " & _
'          "FROM    b_minuta a " & _
'          "        INNER JOIN a_regimen b ON a.min_codreg = b.reg_codigo " & _
'          "        INNER JOIN a_servicio c ON a.min_codser = c.ser_codigo " & _
'          "WHERE   a.min_cencos = '" & MuestraCasino(1) & "' " & _
'          "        AND c.ser_activo = '1' " & _
'          "        AND ( a.min_codreg NOT IN ( " & _
'          "              SELECT DISTINCT " & _
'          "                        aplvs.reg_codigo " & _
'          "              FROM      dbo.a_pto_lectura_vales_servicio aplvs " & _
'          "                        INNER JOIN a_pto_lectura_vales e ON e.lec_codlecvales = aplvs.lec_codlecvales " & _
'          "                                                            AND e.lec_nombrepc LIKE '" & tipmin & "' " & _
'          "                                                            AND e.lec_activo = 1 ) " & _
'          "              OR a.min_codser NOT IN ( " & _
'          "              SELECT DISTINCT " & _
'          "                        aplvs.ser_codigo " & _
'          "              FROM      dbo.a_pto_lectura_vales_servicio aplvs " & _
'          "                        INNER JOIN a_pto_lectura_vales e ON e.lec_codlecvales = aplvs.lec_codlecvales " & _
'          "                                                            AND e.lec_nombrepc LIKE '" & tipmin & "' " & _
'          "                                                            AND e.lec_activo = 1 ) " & _
'          "            ) " & _
'          "ORDER BY b.reg_codigo, b.reg_nombre, c.ser_codigo, c.ser_nombre"

    
'    Minutas = "SELECT DISTINCT " & _
'              "b.reg_codigo " & _
'              ", b.reg_nombre " & _
'              ", d.reg_codigo activoreg " & _
'              ", c.ser_codigo " & _
'              ", d.ser_codigo activoser " & _
'              ", c.ser_nombre " & _
'              ",(SELECT top 1 1 FROM a_pto_lectura_vales_pto_atencion f WHERE d.lec_codlecvales = f.lec_codlecvales AND f.lec_codlecvales = " & codreg & ") AS existe " & _
'              "FROM    b_minuta a " & _
'              "INNER JOIN a_regimen b ON a.min_codreg = b.reg_codigo " & _
'              "INNER JOIN a_servicio c ON a.min_codser = c.ser_codigo " & _
'              "LEFT OUTER JOIN a_pto_lectura_vales_servicio d ON b.reg_codigo = d.reg_codigo " & _
'                                                              " AND c.ser_codigo = d.ser_codigo " & _
'              "LEFT OUTER JOIN a_pto_lectura_vales e ON d.lec_codlecvales = e.lec_codlecvales AND  e.lec_nombrepc like '" & tipmin & "' " & _
'              "WHERE   a.min_cencos = '" & MuestraCasino(1) & "' and c.ser_activo = '1' " & _
'              "ORDER BY b.reg_codigo " & _
'              ", b.reg_nombre " & _
'              ", c.ser_codigo " & _
'              ", c.ser_nombre"
End Select
End Function

Public Function Producto(ByVal opcion As Long, ByVal codpro As String, ByVal descri As String) As String
Dim sql1 As String
Producto = ""
Select Case opcion
Case 1
    Producto = "SELECT DISTINCT pro_codtip FROM b_productos"
Case 2
    Producto = "SELECT DISTINCT a.pro_codigo, a.pro_nombre, (SELECT TOP 1 ppd_propon FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = a.pro_codigo AND ppd_fecdia <= " & Format(Date, "yyyymmdd") & " ORDER BY ppd_fecdia DESC) AS ppd_propon, pro_maepro, a.pro_codtip, CASE WHEN a.pro_maepro = 0 THEN 'Ambos' ELSE c.tis_nombre END AS tis_nombre " & _
                "FROM b_productos a, a_tiposervicio c, b_clientes d  " & _
                "WHERE (c.tis_codigo = d.cli_codtis OR a.pro_maepro < 1) " & _
                "AND    d.cli_codigo = '" & MuestraCasino(1) & "' " & _
                "AND   (c.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) " & _
                "AND   (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre"
End Select
End Function

Public Function ProductoIng(ByVal opcion As Long, ByVal codpri As String, ByVal descri As String) As String
Dim sql1 As String
ProductoIng = ""
Select Case opcion
Case 1
    ProductoIng = "SELECT ing.ing_codigo, ing.ing_nombre FROM b_ingrediente ing, b_productosing AS proing WHERE proing.pri_coding = ing.ing_codigo AND proing.pri_codpro = '" & codpri & "' ORDER BY ing.ing_nombre"
End Select
End Function

Public Function ProductoNut(ByVal opcion As Long, ByVal codpro As String, ByVal codpru As String, ByVal descri As String) As String
Dim sql1 As String
ProductoNut = ""
Select Case opcion
Case 1
    ProductoNut = "SELECT * FROM b_productonut WHERE pnu_codpro = '" & codpro & "' AND pnu_codapo = " & codpru
Case 2
    ProductoNut = "SELECT * FROM b_productonut WHERE pnu_codpro = '" & codpro & "' "
End Select
End Function

Public Function CurvaABC(ByVal opcion As Long, ByVal codcta As String, ByVal descri As String) As String
Dim sql1 As String
CurvaABC = ""
Select Case opcion
Case 1
    CurvaABC = "SELECT * FROM a_curvaabc WHERE (abc_codigo = '" & codcta & "' OR '" & codcta & "' = '') ORDER BY abc_codigo"
Case 2
    CurvaABC = "SELECT abc_codigo, abc_nombre, abc_porce FROM a_curvaabc WHERE abc_codigo LIKE '%" & UCase(descri) & "%'"
Case 3
   sql1 = IIf(vg_tipbase = "1", " UCASE(abc_nombre) ", " UPPER(abc_nombre) ")
   CurvaABC = "SELECT abc_codigo, abc_nombre, abc_porce FROM a_curvaabc WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Bodega(ByVal opcion As Long, ByVal codbod As Long, ByVal descri As String) As String
Dim sql1 As String
Bodega = ""
Select Case opcion
Case 1
    Bodega = "SELECT bod_nombre FROM a_bodega WHERE bod_codigo = " & codbod & ""
Case 2
    Bodega = "SELECT bod_codigo FROM a_bodega ORDER BY bod_codigo DESC"
Case 3
    Bodega = "SELECT * FROM a_bodega WHERE (bod_codigo = " & codbod & " OR " & codbod & " = 0) ORDER BY bod_codigo"
Case 4
    Bodega = "SELECT bod_codigo,bod_nombre,bod_ubicac FROM a_bodega WHERE bod_codigo LIKE '%" & UCase(descri) & "%'"
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(bod_nombre) ", " UPPER(bod_nombre) ")
    Bodega = "SELECT bod_codigo, bod_nombre, bod_ubicac FROM a_bodega WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Servicio(ByVal opcion As Long, ByVal codser As Long, ByVal descri As String) As String
Dim sql1 As String
Servicio = ""
Select Case opcion
Case 1
    Servicio = "SELECT * FROM  a_servicio WHERE ser_activo = '1' ORDER BY ser_orden, ser_codigo"
Case 2
    Servicio = "SELECT * FROM a_servicio WHERE ser_codigo = " & codser & " AND ser_activo = '1'"
Case 3
    Servicio = "SELECT ser_codigo FROM a_servicio WHERE ser_codigo < 10000 ORDER BY ser_codigo DESC"
Case 4
    Servicio = "SELECT * FROM a_servicio WHERE ser_codigo LIKE '%" & UCase(descri) & "%'"
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(ser_nombre) ", " UPPER(ser_nombre) ")
    Servicio = "SELECT * FROM a_servicio WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
Case 6
    Servicio = "SELECT * FROM a_servicio WHERE (ser_codigo = " & codser & " OR " & codser & " = 0)"
Case 7
    Servicio = "SELECT * FROM a_servicio ORDER BY ser_orden, ser_nombre"
Case 8
    Servicio = "SELECT DISTINCT * FROM a_servicio WHERE ser_codigo = " & codser & ""
End Select
End Function

Public Function Cliente(ByVal opcion As Long, ByVal cencos As String, ByVal descri As String) As String
Dim sql1 As String
Cliente = ""
Select Case opcion
Case 1
    Cliente = "SELECT * FROM b_clientes WHERE cli_codigo = '" & cencos & "' AND cli_tipo = 0"
Case 2
    Cliente = "SELECT cli_codigo, cli_nombre FROM b_clientes WHERE (cli_codigo = '" & cencos & "' OR '" & cencos & "' = '') AND cli_tipo = 1 AND cli_activo = '1' ORDER BY cli_codigo"
Case 3
    Cliente = "SELECT * FROM b_clientes WHERE cli_tipo = 0 OR cli_tipo = 2 ORDER BY cli_nombre"
Case 4
    Cliente = "SELECT * FROM b_clientes WHERE cli_codigo = '" & cencos & "'"
Case 5
    Cliente = "SELECT * FROM b_clientes WHERE cli_codigo = '" & cencos & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
Case 6
   sql1 = IIf(vg_tipbase = "1", " UCASE(cli_codigo) ", " UPPER(cli_codigo) ")
   Cliente = "SELECT cli_codigo, cli_nombre, cli_tipo FROM b_clientes WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%' AND (cli_tipo = 0 OR cli_tipo = 2) ORDER BY cli_codigo"
Case 7
   sql1 = IIf(vg_tipbase = "1", " UCASE(cli_nombre) ", " UPPER(cli_nombre) ")
   Cliente = "SELECT cli_codigo, cli_nombre, cli_tipo FROM b_clientes WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%' AND (cli_tipo = 0 OR cli_tipo = 2) ORDER BY cli_nombre"
Case 8
   sql1 = IIf(vg_tipbase = "1", " UCASE(cli_codigo) ", " UPPER(cli_codigo) ")
   Cliente = "SELECT COUNT(*) AS nreg FROM b_clientes WHERE " & sql1 & " LIKE '%" & UCase((descri)) & "%' AND (cli_tipo=0 OR cli_tipo=2)"
Case 9
   Cliente = "SELECT COUNT(*) AS nreg FROM b_clientes WHERE cli_tipo = 1 and cli_clisap = '1'"
Case 10
   Cliente = "SELECT cli_codigo FROM b_clientes WHERE cli_tipo = 1 and cli_clisap = '1'"
End Select
End Function

Public Function Regimen(ByVal opcion As Long, ByVal codreg As Long, ByVal descri As String) As String
Dim sql1 As String
Regimen = ""
Select Case opcion
Case 1
    Regimen = "SELECT * FROM  a_regimen ORDER BY reg_codigo"
Case 2
    Regimen = "SELECT * FROM a_regimen WHERE reg_codigo = " & codreg & ""
Case 3
    Regimen = "SELECT reg_codigo FROM a_regimen WHERE reg_codigo < 10000 ORDER BY reg_codigo DESC"
Case 4
    Regimen = "SELECT reg_codigo, reg_nombre FROM a_regimen WHERE reg_codigo LIKE '%" & UCase(descri) & "%'"
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(reg_nombre) ", " UPPER(reg_nombre) ")
    Regimen = "SELECT reg_codigo, reg_nombre FROM a_regimen WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Proveedor(ByVal opcion As Long, ByVal codpro As String, ByVal descri As String) As String
Dim sql1 As String
Proveedor = ""
Select Case opcion
Case 1
    Proveedor = "SELECT * FROM b_proveedor WHERE (prv_codigo = '" & codpro & "' OR '" & codpro & "' = '') ORDER BY prv_codigo"
Case 2
    Proveedor = "SELECT DISTINCT * FROM b_proveedor WHERE (prv_codigo = '" & codpro & "' OR '" & codpro & "' = '') AND prv_activo IN ('0','1') ORDER BY prv_nombre"
Case 3
    Proveedor = "SELECT *, (SELECT TOP 1 mun_nombre FROM a_municipio WHERE mun_codigo = b_proveedor.prv_codmun) AS mun_nombre FROM b_proveedor WHERE prv_codigo = '" & codpro & "' AND prv_activo IN ('0','1')"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(prv_codigo) ", " UPPER(prv_codigo) ")
    Proveedor = "SELECT prv_codigo, prv_nombre, prv_activo, prv_origen FROM b_proveedor WHERE prv_activo IN ('0','1') AND " & sql1 & " LIKE '%" & UCase(descri) & "%' ORDER BY prv_codigo"
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(prv_nombre) ", " UPPER(prv_nombre) ")
    Proveedor = "SELECT prv_codigo, prv_nombre, prv_activo, prv_origen FROM b_proveedor WHERE prv_activo IN ('0','1') AND " & sql1 & " LIKE '%" & UCase(descri) & "%' ORDER BY prv_nombre"
Case 6
    sql1 = IIf(vg_tipbase = "1", " UCASE(prv_codigo) ", " UPPER(prv_codigo) ")
    Proveedor = "SELECT COUNT(*) AS nreg FROM b_proveedor WHERE  prv_activo IN ('0','1') AND " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Nutriente(ByVal opcion As Long, ByVal codnut As Long, ByVal descri As String) As String
Dim sql1 As String
Nutriente = ""
Select Case opcion
Case 1
    Nutriente = "SELECT nut_codigo, nut_nombre, nut_indpri, nut_secnro FROM a_nutriente ORDER BY nut_secnro"
Case 2
    Nutriente = "SELECT nut_codigo FROM a_nutriente ORDER BY nut_codigo DESC"
Case 3
    Nutriente = "SELECT * FROM a_nutriente WHERE (nut_codigo = " & codnut & " OR " & codnut & " = 0) ORDER BY nut_secnro, nut_nombre"
Case 4
    Nutriente = "SELECT nut_codigo,nut_nombre,nut_nomuni,nut_indpri,nut_secnro FROM a_nutriente WHERE nut_codigo LIKE '%" & UCase(descri) & "%' ORDER BY nut_secnro, nut_nombre"
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(nut_nombre) ", " UPPER(nut_nombre) ")
    Nutriente = "SELECT nut_codigo, nut_nombre, nut_nomuni, nut_indpri, nut_secnro FROM a_nutriente WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%' ORDER BY nut_secnro, nut_nombre"
End Select
End Function

Public Function GrupoPaciente(ByVal opcion As Long) As String
GrupoPaciente = ""
Select Case opcion
Case 1
    GrupoPaciente = "SELECT grp_codigo, grp_nombre FROM a_grupopaciente ORDER BY grp_codigo"
End Select
End Function

Public Function Paciente(ByVal opcion As Long, ByVal codpac As String) As String
Paciente = ""
Select Case opcion
Case 1
    Paciente = "SELECT pac_nombre, pac_appaterno, pac_apmaterno FROM b_pacientes WHERE  pac_codigo = '" & codpac & "'"
End Select
End Function

Public Function Impuesto(ByVal opcion As Long, ByVal codimp As Long, ByVal descri As String) As String
Dim sql1 As String
Impuesto = ""
Select Case opcion
Case 1
    Impuesto = "SELECT imp_nombre, imp_codigo FROM a_impuesto ORDER BY imp_nombre"
Case 2
    Impuesto = "SELECT imp_codigo FROM a_impuesto ORDER BY imp_codigo DESC"
Case 3
    Impuesto = "SELECT * FROM a_impuesto WHERE (imp_codigo = " & codimp & " OR " & codimp & " = 0) ORDER BY imp_codigo"
Case 4
    Impuesto = "SELECT * FROM a_impuesto WHERE imp_codigo LIKE '%" & UCase(descri) & "%'"
Case 5
   sql1 = IIf(vg_tipbase = "1", " UCASE(imp_nombre) ", " UPPER(imp_nombre) ")
   Impuesto = "SELECT * FROM a_impuesto WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Ingrediente(ByVal opcion As Long, ByVal coding As String, ByVal descri As String) As String
Ingrediente = ""
Select Case opcion
Case 1
    Ingrediente = "SELECT DISTINCT a.ing_codigo, a.ing_nombre, a.ing_nomfan " & _
                  "FROM b_ingrediente a, b_productosing b, b_productos c, a_tiposervicio d, b_clientes e " & _
                  "WHERE a.ing_codigo = b.pri_coding " & _
                  "AND  (d.tis_codigo = e.cli_codtis OR c.pro_maepro < 1) AND e.cli_codigo = '" & MuestraCasino(1) & "' AND (d.tis_codigo = c.pro_maepro OR c.pro_maepro < 1) " & _
                  "AND   b.pri_codpro = c.pro_codigo ORDER BY a.ing_codigo"
Case 2
    Ingrediente = "SELECT ing_codigo FROM b_ingrediente WHERE ing_codigo = '" & coding & "'"
Case 3
'    Ingrediente = "SELECT ing.*, unm.unm_nombre, unm.unm_nomcor, c.cpi_precos, c.cpi_feccos FROM b_ingrediente ing, a_unidadmed unm, b_contlistpreing c " & _
'                  "WHERE ing_codigo = c.cpi_coding AND unm.unm_codigo = ing.ing_unimed AND ing.ing_codigo = '" & coding & "' AND c.cpi_cencos = '" & MuestraCasino(1) & "'"
    Ingrediente = "SELECT ing.*, unm.unm_nombre, unm.unm_nomcor, isnull(contlist.cpi_precos,0) as cpi_precos, isnull(contlist.cpi_feccos,0) as cpi_feccos FROM b_ingrediente ing " & _
                  "inner join a_unidadmed unm on unm.unm_codigo = ing.ing_unimed " & _
                  "left join (select c.cpi_coding, c.cpi_precos, c.cpi_feccos  from b_contlistpreing c where c.cpi_cencos = '" & MuestraCasino(1) & "') as contlist on ing_codigo = contlist.cpi_coding " & _
                  "WHERE  ing.ing_codigo = '" & coding & "' "
End Select
End Function

Public Function EstServicio(ByVal opcion As Long, ByVal codser As Long, ByVal codest As Long) As String
EstServicio = ""
Select Case opcion
Case 1
    EstServicio = "SELECT a_estservicio.* FROM a_estservicio WHERE ess_cencos = '" & MuestraCasino(1) & "' AND ess_codser = " & codser & " ORDER BY ess_orden"
Case 2
    EstServicio = "SELECT * FROM a_estservicio WHERE ess_cencos = '" & MuestraCasino(1) & "' AND ess_codser = " & codser & " AND ess_codigo = " & codest & ""
Case 3
    EstServicio = "SELECT ess_codigo FROM a_estservicio WHERE ess_cencos='" & MuestraCasino(1) & "' AND ess_codigo <= 9999 ORDER BY ess_codigo DESC"
End Select
End Function

Public Function CostoPatron(ByVal opcion As Long, ByVal codreg As Long, ByVal codser As Long, ByVal Fecha As Long) As String
CostoPatron = ""
Select Case opcion
Case 1
    CostoPatron = "SELECT cpa_descripcion, cpa_valor FROM b_costopatron WHERE cpa_cencos = '" & MuestraCasino(1) & "' AND cpa_codreg = " & codreg & " AND cpa_codser = " & codser & " AND cpa_anomes = " & Fecha & ""
Case 2
    CostoPatron = "SELECT * FROM b_costopatron WHERE cpa_cencos = '" & MuestraCasino(1) & "' AND cpa_codreg = " & codreg & " AND cpa_codser = " & codser & " AND cpa_anomes = " & Fecha & " AND cpa_descripcion = 'TECHO'"
Case 3
    CostoPatron = "SELECT * FROM b_costopatron WHERE cpa_cencos='" & MuestraCasino(1) & "' AND cpa_codreg=" & codreg & " AND cpa_codser =" & codser & " AND cpa_descripcion='TECHO' ORDER BY cpa_anomes"
End Select
End Function

Public Function ServicioRaciones(ByVal opcion As Long, ByVal codser As Long) As String
ServicioRaciones = ""
Select Case opcion
Case 1
    ServicioRaciones = "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos = '" & MuestraCasino(1) & "' AND sra_codser = " & codser & " GROUP BY sra_serdia ORDER BY sra_serdia"
Case 2
    ServicioRaciones = "SELECT * FROM a_serviciorac WHERE sra_cencos='" & MuestraCasino(1) & "' AND sra_codser = " & codser & " order by sra_coditem, sra_serdia"
End Select
End Function

Public Function EstructuraFijaDia(ByVal opcion As Long, ByVal codreg As Long, ByVal codser As Long, ByVal Fecha As Long, ByVal tipmin As String) As String
Dim sql1 As String
EstructuraFijaDia = ""
Select Case opcion
Case 1
    sql1 = IIf(vg_tipbase = "1", " val(mid(mfd_fecha,1,6)) ", " convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) ")
    EstructuraFijaDia = "SELECT DISTINCT mfd_cencos FROM b_minutafijadia " & _
            "WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
            "AND   mfd_codreg = " & codreg & " " & _
            "AND   mfd_codser = " & codser & " " & _
            "AND   " & sql1 & " = " & Fecha & " AND mfd_tipmin = '" & tipmin & "'"
End Select
End Function

Public Function MinutaFija(ByVal opcion As Long, ByVal codreg As Long, ByVal codser As Long, ByVal Fecha As Long, ByVal nrodia As Long, ByVal aAp As String, ByVal ctacon As String, ByVal codpro As String) As String
MinutaFija = ""
If Trim(ctacon) = "" Then ctacon = "b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
Select Case opcion
Case 1
    MinutaFija = "SELECT MAX(mif_fecval) AS fecval FROM b_minutafija WHERE mif_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   mif_codreg = " & codreg & " AND   mif_codser = " & codser & ""
Case 2
    MinutaFija = "SELECT SUM(c.ppd_propon*a.mif_canpro) AS cosesf " & _
                 "FROM  b_minutafija a, b_productos b, " & aAp & " c " & _
                 "WHERE a.mif_codpro = b.pro_codigo " & _
                 "AND   b.pro_codigo = c.ppd_codpro " & _
                 "AND   c.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   a.mif_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   a.mif_codreg = " & codreg & " " & _
                 "AND   a.mif_codser = " & codser & " " & _
                 "AND   a.mif_fecval = " & Fecha & " " & _
                 "AND   a.mif_dianro = " & nrodia & " " & _
                 "AND   " & ctacon & ""
Case 3
    MinutaFija = "SELECT SUM(c.ppd_propon*a.mif_canpro) AS cosesf " & _
                 "FROM  b_minutafija a, b_productos b, b_productospmpdia c " & _
                 "WHERE a.mif_codpro = b.pro_codigo " & _
                 "AND   b.pro_codigo = c.ppd_codpro " & _
                 "AND   c.ppd_cencos = '" & MuestraCasino(1) & "' AND c.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
                 "AND   a.mif_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   a.mif_codreg = " & codreg & " " & _
                 "AND   a.mif_codser = " & codser & " " & _
                 "AND   a.mif_fecval = " & Fecha & " " & _
                 "AND   a.mif_dianro = " & nrodia & " " & _
                 "AND   " & ctacon & ""
Case 4
    MinutaFija = "SELECT b.pro_codigo, b.pro_nombre, a.uni_nomcor, c.mif_canpro " & _
                 "FROM  a_unidad a, b_productos b, b_minutafija c " & _
                 "WHERE c.mif_codpro = b.pro_codigo " & _
                 "AND   b.pro_coduni = a.uni_codigo " & _
                 "AND   c.mif_cencos = '" & MuestraCasino(1) & "' " & _
                 "AND   c.mif_codreg = " & codreg & " " & _
                 "AND   c.mif_codser = " & codser & " " & _
                 "AND   c.mif_fecval = " & Fecha & " " & _
                 "AND   c.mif_dianro = " & nrodia & " " & _
                 "AND  (c.mif_codpro = '" & codpro & "' OR '" & codpro & "' = '')"
End Select
End Function

Public Function MinutaFijaDia(ByVal opcion As Long, ByVal codreg As Long, ByVal codser As Long, ByVal Fecha As Long, ByVal tipmin As String, ByVal ctacon As String) As String
Dim sql1 As String
MinutaFijaDia = ""
Select Case opcion
Case 1
    sql1 = IIf(vg_tipbase = "1", " val(mid(mfd_fecha,1,6)) ", " convert(int,substring(convert(varchar(8),mfd_fecha),1,6)) ")
    MinutaFijaDia = "SELECT DISTINCT mfd_cencos FROM b_minutafijadia WHERE mfd_cencos = '" & MuestraCasino(1) & "' " & _
                    "AND   mfd_codreg = " & codreg & " " & _
                    "AND   mfd_codser = " & codser & " " & _
                    "AND   " & sql1 & " = " & Fecha & " AND mfd_tipmin = '" & tipmin & "'"
Case 2
    If Trim(ctacon) = "" Then ctacon = "b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')"
    MinutaFijaDia = "SELECT SUM(a.mfd_canpro*a.mfd_cospro) AS cosesf " & _
                    "FROM b_minutafijadia a, b_productos b " & _
                    "WHERE a.mfd_codpro = b.pro_codigo " & _
                    "AND   a.mfd_cencos = '" & MuestraCasino(1) & "' " & _
                    "AND   a.mfd_codreg = " & codreg & " " & _
                    "AND   a.mfd_codser = " & codser & " " & _
                    "AND   a.mfd_fecha  = " & Fecha & " AND a.mfd_tipmin = '" & tipmin & "' " & _
                    "AND   " & ctacon & ""
'b.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "')
End Select
End Function

Public Function MinutaRaciones(ByVal opcion As Long, ByVal codreg As Long, ByVal codser As Long, ByVal rutcli As String, ByVal Fecha As Long) As String
MinutaRaciones = ""
Select Case opcion
Case 1
    MinutaRaciones = "SELECT SUM(mir_nrorac) AS mir_nrorac FROM b_minutaraciones " & _
                     "WHERE  mir_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND    mir_codreg = " & codreg & " " & _
                     "AND    mir_codser = " & codser & " " & _
                     "AND    mir_rutcli = '" & rutcli & "' " & _
                     "AND    mir_fecmin = " & Fecha & ""
End Select
End Function

Public Function UnidadMedida(ByVal opcion As Long, ByVal codume As Long, ByVal descri As String) As String
Dim sql1 As String
UnidadMedida = ""
Select Case opcion
Case 1
    UnidadMedida = "SELECT unm_codigo FROM a_unidadmed ORDER BY unm_codigo DESC"
Case 2
    UnidadMedida = "SELECT * FROM a_unidadmed WHERE (unm_codigo = " & codume & " OR " & codume & " = 0) ORDER BY unm_codigo"
Case 3
    UnidadMedida = "SELECT unm_codigo,unm_nombre,unm_nomcor FROM a_unidadmed WHERE unm_codigo LIKE '%" & UCase(descri) & "%'"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(unm_nombre) ", " UPPER(unm_nombre) ")
    UnidadMedida = "SELECT unm_codigo, unm_nombre, unm_nomcor FROM a_unidadmed WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Unidad(ByVal opcion As Long, ByVal coduni As Long, ByVal descri As String) As String
Dim sql1 As String
Unidad = ""
Select Case opcion
Case 1
    Unidad = "SELECT uni_codigo FROM a_unidad ORDER BY uni_codigo DESC"
Case 2
    Unidad = "SELECT * FROM a_unidad WHERE (uni_codigo = " & coduni & " OR " & coduni & " = 0) ORDER BY uni_codigo"
Case 3
    Unidad = "SELECT uni_codigo,uni_nombre,uni_nomcor FROM a_unidad WHERE uni_codigo LIKE '%" & UCase(descri) & "%'"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(uni_nombre) ", " UPPER(uni_nombre) ")
    Unidad = "SELECT uni_codigo, uni_nombre, uni_nomcor FROM a_unidad WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Embalaje(ByVal opcion As Long, ByVal codemb As Long, ByVal descri As String) As String
Dim sql1 As String
Embalaje = ""
Select Case opcion
Case 1
    Embalaje = "SELECT * FROM a_embalaje WHERE (emb_codigo = " & codemb & " OR " & codemb & " = 0) ORDER BY emb_codigo"
End Select
End Function

Public Function CategoriaDietetica(ByVal opcion As Long, ByVal codcdi As Long, ByVal descri As String) As String
Dim sql1 As String
CategoriaDietetica = ""
Select Case opcion
Case 1
    CategoriaDietetica = "SELECT * FROM a_recetacatdie WHERE car_previo = 0"
Case 2
    CategoriaDietetica = "SELECT DISTINCT * FROM a_recetacatdie WHERE car_previo = " & codcdi & " ORDER BY car_codigo"
Case 3
    CategoriaDietetica = "SELECT car_codigo FROM a_recetacatdie ORDER BY car_codigo DESC"
Case 4
    CategoriaDietetica = "SELECT * FROM a_recetacatdie WHERE car_codigo=" & codcdi
Case 5
    CategoriaDietetica = "SELECT COUNT(car_codigo) AS nreg FROM a_recetacatdie WHERE UCASE(car_nombre) LIKE '%" & UCase(descri) & "%' AND car_previo = " & codcdi & ""
Case 6
    CategoriaDietetica = "SELECT * FROM a_recetacatdie WHERE car_previo = " & codcdi & " order by car_nombre"
End Select
End Function

Public Function CuentaContable(ByVal opcion As Long, ByVal codcta As String, ByVal descri As String) As String
Dim sql1 As String
CuentaContable = ""
Select Case opcion
Case 1
    CuentaContable = "SELECT cta_codigo, cta_nombre FROM a_ctacontable WHERE (cta_codigo = '" & codcta & "' OR '" & codcta & "' = '') ORDER BY cta_codigo"
Case 2
    CuentaContable = "SELECT cta_codigo,cta_nombre FROM a_ctacontable WHERE cta_codigo LIKE '%" & UCase(descri) & "%'"
Case 3
    sql1 = IIf(vg_tipbase = "1", " ucase(cta_nombre) ", " UPPER(cta_nombre) ")
    CuentaContable = "SELECT cta_codigo, cta_nombre FROM a_ctacontable WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Sector(ByVal opcion As Long, ByVal codsec As Long, ByVal descri As String) As String
Dim sql1 As String
Sector = ""
Select Case opcion
Case 1
    Sector = "SELECT sec_codigo FROM a_sector ORDER BY sec_codigo DESC"
Case 2
    Sector = "SELECT * FROM a_sector WHERE (sec_codigo = " & codsec & " OR " & codsec & " = 0) ORDER BY sec_codigo"
Case 3
    Sector = "SELECT sec_codigo,sec_nombre,sec_orden FROM a_sector WHERE sec_codigo LIKE '%" & UCase(descri) & "%'"
Case 4
   sql1 = IIf(vg_tipbase = "1", " UCASE(sec_nombre) ", " UPPER(sec_nombre) ")
   Sector = "SELECT sec_codigo, sec_nombre, sec_orden FROM a_sector WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function TipoAjuste(ByVal opcion As Long, ByVal codtaj As Long, ByVal descri As String, ByVal tipaju As Long) As String
Dim sql1 As String
TipoAjuste = ""
Select Case opcion
Case 1
    TipoAjuste = "SELECT aju_codigo FROM a_tipoajuste ORDER BY aju_codigo DESC"
Case 2
    TipoAjuste = "SELECT * FROM a_tipoajuste WHERE (aju_codigo = " & codtaj & " OR " & codtaj & " = 0)  AND aju_tipaju = " & tipaju & "  and aju_codigo>9999 ORDER BY aju_codigo"
Case 3
    TipoAjuste = "SELECT aju_codigo, aju_nombre, aju_tipo FROM a_tipoajuste WHERE aju_codigo LIKE '%" & UCase(descri) & "%' AND aju_tipaju = " & tipaju & " and aju_codigo>9999"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(aju_nombre) ", " UPPER(aju_nombre) ")
    TipoAjuste = "SELECT aju_codigo, aju_nombre, aju_tipo FROM a_tipoajuste WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%' AND aju_tipaju = " & tipaju & "  and aju_codigo>9999"
End Select
End Function

Public Function TipoDocumento(ByVal opcion As Long, ByVal codtdo As String, ByVal descri As String) As String
Dim sql1 As String
TipoDocumento = ""
Select Case opcion
Case 1
    TipoDocumento = "SELECT * FROM a_tipodocumento WHERE (tdo_codigo = '" & codtdo & "' OR '" & codtdo & "' = '')"
Case 2
   sql1 = IIf(vg_tipbase = "1", " Ucase(tdo_codigo) ", " UPPER(tdo_codigo) ")
   TipoDocumento = "SELECT tdo_codigo, tdo_nombre, tdo_cladoc, tdo_orden FROM a_tipodocumento WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
Case 3
   sql1 = IIf(vg_tipbase = "1", " Ucase(tdo_nombre) ", " UPPER(tdo_nombre) ")
   TipoDocumento = "SELECT tdo_codigo, tdo_nombre, tdo_cladoc, tdo_orden FROM a_tipodocumento WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function Region(ByVal opcion As Long, ByVal codigo As String, ByVal descri As String) As String
Dim sql1 As String
Region = ""
Select Case opcion
Case 1
    Region = "SELECT reg_codigo, reg_nombre, " & _
             "(SELECT DISTINCT cds_cdosap FROM a_clasedocsap WHERE reg_codigo = cds_codreg AND  cds_coddoc = '" & codigo & "') AS cds_cdosap " & _
             "FROM a_region"
End Select
End Function

Public Function RecetaTipoPlato(ByVal opcion As Long, ByVal codigo As Long, ByVal descri As String) As String
Dim sql1 As String
RecetaTipoPlato = ""
Select Case opcion
Case 1
    RecetaTipoPlato = "SELECT * FROM a_recetatippla WHERE tip_previo = 0 ORDER BY tip_nombre"
Case 2
    RecetaTipoPlato = "SELECT DISTINCT tip_codigo FROM a_recetatippla WHERE tip_previo = " & codigo & " ORDER BY tip_codigo"
Case 3
    RecetaTipoPlato = "SELECT tip_codigo FROM a_recetatippla ORDER BY tip_codigo DESC"
Case 4
    RecetaTipoPlato = "SELECT DISTINCT * FROM a_recetatippla WHERE tip_codigo = " & codigo
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(tip_nombre) ", " UPPER(tip_nombre) ")
    RecetaTipoPlato = "SELECT COUNT(tip_codigo) AS nreg FROM a_recetatippla WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%' AND tip_previo = " & codigo
Case 6
    RecetaTipoPlato = "SELECT * FROM a_recetatippla WHERE tip_previo = " & codigo & " ORDER BY tip_codigo"
Case 7
    RecetaTipoPlato = "SELECT * FROM a_recetatippla WHERE tip_previo = " & codigo & " order by tip_nombre"
End Select
End Function

Public Function TipoServicio(ByVal opcion As Long, ByVal codigo As Long, ByVal descri As String) As String
Dim sql1 As String
TipoServicio = ""
Select Case opcion
Case 1
    TipoServicio = "SELECT tis_codigo FROM a_tiposervicio ORDER BY tis_codigo DESC"
Case 2
    TipoServicio = "SELECT * FROM a_tiposervicio WHERE (tis_codigo = " & codigo & " OR " & codigo & " = 0) ORDER BY tis_codigo"
Case 3
    TipoServicio = "SELECT tis_codigo, tis_nombre FROM a_tiposervicio WHERE tis_codigo LIKE '%" & UCase(descri) & "%'"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(tis_nombre) ", " UPPER(tis_nombre) ")
    TipoServicio = "SELECT tis_codigo, tis_nombre FROM a_tiposervicio WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
End Select
End Function

Public Function PuntoAtencion(ByVal opcion As Long, ByVal codigo As Long, ByVal descri As String) As String
Dim sql1 As String
PuntoAtencion = ""
Select Case opcion
Case 1
    PuntoAtencion = "SELECT ate_codatencion FROM a_pto_atencion ORDER BY ate_codatencion DESC"
Case 2
    PuntoAtencion = "SELECT * FROM a_pto_atencion WHERE (ate_codatencion = " & codigo & " OR " & codigo & " = 0) ORDER BY ate_codatencion"
Case 3
    PuntoAtencion = "SELECT ate_codatencion, ate_descripcion FROM a_pto_atencion WHERE ate_codatencion LIKE '%" & UCase(descri) & "%'"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(ate_descripcion) ", " UPPER(ate_descripcion) ")
    PuntoAtencion = "SELECT ate_codatencion, ate_descripcion FROM a_pto_atencion WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
Case 5
    PuntoAtencion = "SELECT  distinct a.* " & _
                 ", b.ate_codatencion activo " & _
                 "FROM    a_pto_atencion a " & _
                 "LEFT OUTER JOIN a_pto_lectura_vales_pto_atencion b ON a.ate_codatencion = b.ate_codatencion " & _
                 "LEFT OUTER JOIN a_pto_lectura_vales c ON b.lec_codlecvales = c.lec_codlecvales OR  c.lec_nombrepc like '" & descri & "' " & _
                 "ORDER BY a.ate_codatencion"
Case 6
    PuntoAtencion = "SELECT  distinct a.* " & _
                 ", b.ate_codatencion activo " & _
                 "FROM    a_pto_atencion a " & _
                 "inner join a_pto_lectura_vales_pto_atencion b ON a.ate_codatencion = b.ate_codatencion " & _
                 "inner join a_pto_lectura_vales c ON b.lec_codlecvales = c.lec_codlecvales WHERE  c.lec_nombrepc like '" & descri & "' and c.lec_activo = 1" & _
                 "ORDER BY a.ate_codatencion"
End Select
End Function

Public Function Personal(ByVal opcion As Long, ByVal codigo As String, ByVal descri As String) As String
Dim sql1 As String
Personal = ""
Select Case opcion
Case 1
    Personal = "SELECT per_rut FROM b_persona ORDER BY per_rut DESC"
Case 2
    Personal = "SELECT * FROM b_persona WHERE (per_rut = '" & codigo & "' OR '" & codigo & "' = '') ORDER BY per_rut"
Case 3
    Personal = "SELECT a.per_rut, a.per_nombre, a.cli_codigo, a.per_codbarra, b.cli_nombre FROM b_persona a, b_clientes b WHERE a.cli_codigo = b.cli_codigo and a.per_rut LIKE '%" & UCase(descri) & "%'"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(a.per_nombre) ", " UPPER(a.per_nombre) ")
    Personal = "SELECT a.per_rut, a.per_nombre, a.cli_codigo, a.per_codbarra, b.cli_nombre  FROM b_persona a, b_clientes b WHERE a.cli_codigo = b.cli_codigo and " & sql1 & " LIKE '%" & UCase(descri) & "%'"
Case 5
    Personal = "SELECT a.*, b.cli_nombre FROM b_persona a inner join b_clientes b on a.cli_codigo = b.cli_codigo WHERE (per_rut = '" & codigo & "' OR '" & codigo & "' = '') ORDER BY per_rut"
Case 6
    Personal = "SELECT a.cli_codigo FROM b_persona a inner join  b_clientes b on a.cli_codigo = b.cli_codigo WHERE a.per_codbarra = '" & codigo & "' and b.cli_tipo = '1'"
Case 7
    Personal = "SELECT per_rut, per_codbarra FROM b_persona WHERE per_codbarra = '" & codigo & "'"
End Select
End Function

Public Function PtoLecturaVales(ByVal opcion As Long, ByVal codigo As Long, ByVal descri As String) As String
Dim sql1 As String
PtoLecturaVales = ""
Select Case opcion
Case 1
    PtoLecturaVales = "SELECT TOP 1 lec_codlecvales, lec_ubicacion FROM a_pto_lectura_vales WHERE lec_nombrepc like '" & descri & "' and lec_activo = 1"
Case 2
    PtoLecturaVales = "SELECT lec_codlecvales FROM a_pto_lectura_vales where lec_activo = 1 ORDER BY lec_codlecvales DESC"
Case 3
    PtoLecturaVales = "SELECT ate_codatencion, ate_descripcion FROM a_pto_lectura_vales WHERE ate_codatencion LIKE '%" & UCase(descri) & "%' and lec_activo = 1"
Case 4
    sql1 = IIf(vg_tipbase = "1", " UCASE(ate_descripcion) ", " UPPER(ate_descripcion) ")
    PtoLecturaVales = "SELECT ate_codatencion, ate_descripcion FROM a_pto_atencion WHERE " & sql1 & " LIKE '%" & UCase(descri) & "%'"
Case 5
    sql1 = IIf(vg_tipbase = "1", " UCASE(e.lec_nombrepc) ", " UPPER(e.lec_nombrepc) ")
    PtoLecturaVales = "        SELECT DISTINCT " & _
                    "d.ate_codatencion, x.ate_descripcion, b.reg_codigo, b.reg_nombre, " & _
                    "c.ser_codigo , c.ser_nombre " & _
                    "FROM    a_pto_lectura_vales e " & _
            "        INNER JOIN dbo.a_pto_lectura_vales_pto_atencion AS aplvpa ON e.lec_codlecvales = aplvpa.lec_codlecvales " & _
            "        INNER JOIN a_pto_lectura_vales_servicio d ON aplvpa.lec_codlecvales = d.lec_codlecvales " & _
            "                                                     AND aplvpa.ate_codatencion = d.ate_codatencion " & _
            "        INNER JOIN a_regimen b ON b.reg_codigo = d.reg_codigo " & _
            "        INNER JOIN a_servicio c ON c.ser_codigo = d.ser_codigo " & _
            "        INNER JOIN a_pto_atencion x ON x.ate_codatencion = d.ate_codatencion " & _
            "WHERE   e.lec_nombrepc LIKE '%" & UCase(descri) & "%' " & _
            "        AND c.ser_activo = '1' " & _
            "        AND D.lec_codlecvales = " & codigo & " " & _
            "        AND e.lec_activo = 1 " & _
            "ORDER BY d.ate_codatencion, b.reg_codigo, c.ser_codigo"
Case 6
    PtoLecturaVales = "SELECT lec_codlecvales FROM a_pto_lectura_vales ORDER BY lec_codlecvales DESC"
End Select
End Function

Public Function PtoLecturaValesServicios(ByVal opcion As Long, ByVal codatencion As Long, ByVal descri As String, codreg As Long) As String
Dim sql1 As String
Dim RS As New ADODB.Recordset
Dim CodLecVales As Long
'Leer punto vales
RS.Open RutinaLectura.PtoLecturaVales(1, 0, descri), vg_db, adOpenStatic
If Not RS.EOF Then CodLecVales = RS!lec_codlecvales
RS.Close: Set RS = Nothing
    
PtoLecturaValesServicios = ""
Select Case opcion
Case 1
    PtoLecturaValesServicios = " SELECT DISTINCT " & _
                                "a.reg_codigo , a.reg_nombre " & _
                                "FROM    a_regimen a " & _
                                "        INNER JOIN a_pto_lectura_vales_servicio b ON a.reg_codigo = b.reg_codigo " & _
                                "        INNER JOIN a_pto_lectura_vales_pto_atencion c ON b.lec_codlecvales = c.lec_codlecvales " & _
                                "                                                         AND c.ate_codatencion = b.ate_codatencion " & _
                                "        INNER JOIN a_pto_lectura_vales d ON b.lec_codlecvales = d.lec_codlecvales " & _
                                "WHERE   d.lec_nombrepc LIKE '" & descri & "' " & _
                                "        AND b.lec_codlecvales = " & CodLecVales & " " & _
                                "        AND b.ate_codatencion = " & codatencion & " " & _
                                "        AND d.lec_activo = 1"
Case 2
    PtoLecturaValesServicios = "SELECT DISTINCT " & _
                                "a.ser_codigo , a.ser_nombre " & _
                                "FROM    a_servicio a " & _
                                "        INNER JOIN a_pto_lectura_vales_servicio b ON a.ser_codigo = b.ser_codigo " & _
                                "        INNER JOIN a_pto_lectura_vales_pto_atencion c ON b.lec_codlecvales = c.lec_codlecvales " & _
                                "                                                         AND c.ate_codatencion = b.ate_codatencion " & _
                                "        INNER JOIN a_pto_lectura_vales d ON b.lec_codlecvales = d.lec_codlecvales " & _
                                "WHERE   d.lec_nombrepc LIKE '" & descri & "' " & _
                                "        AND a.ser_activo = '1' " & _
                                "        AND d.lec_codlecvales = " & CodLecVales & " " & _
                                "        AND b.reg_codigo = " & codreg & " " & _
                                "        AND b.ate_codatencion = " & codatencion & " " & _
                                "        AND d.lec_activo = 1"
End Select
End Function

Public Function DetalleLectura(ByVal opcion As Long, ByVal cencos As String, ByVal codreg As Long, codser As Long, codate As Long, descri As String, fecbar As Long) As String
Dim sql1 As String
Dim Fecha As Long
Dim RS As New ADODB.Recordset
DetalleLectura = ""
Fecha = 0
If vg_tipbase = "1" Then
   RS.Open "select " & Format(Now, "yyyymmdd") & " as fecha", vg_db, adOpenStatic
   If Not RS.EOF Then
      Fecha = RS!Fecha
   End If
   RS.Close: Set RS = Nothing
Else
   RS.Open "select convert(varchar(8),getdate(),112) as fecha", vg_db, adOpenStatic
   If Not RS.EOF Then
      Fecha = RS!Fecha
   End If
   RS.Close: Set RS = Nothing
End If
sql1 = IIf(vg_tipbase = "1", " and format(a.fechahoraregistro,'yyyymmdd') = " & Fecha & " ", " and convert(varchar(8),a.fechahoraregistro,112) = " & Fecha & "")
Select Case opcion
Case 1
    DetalleLectura = "SELECT DISTINCT a.codigobarra, a.fechahoravale FROM b_detallelectura a inner join a_servicio b on a.ser_codigo = b.ser_codigo inner join a_regimen c on a.reg_codigo = c.reg_codigo inner join a_pto_atencion d on a.ate_codatencion = d.ate_codatencion WHERE a.cli_codigo = '" & cencos & "' and a.reg_codigo = " & codreg & " and a.ser_codigo = " & codser & " and a.ate_codatencion = " & codate & " " & sql1 & " order by a.fechahoravale"
Case 2
    DetalleLectura = "SELECT DISTINCT codigobarra FROM b_detallelectura a WHERE a.cli_codigo = '" & cencos & "' and a.reg_codigo = " & codreg & " and a.ser_codigo = " & codser & " and a.ate_codatencion = " & codate & " and a.codigobarra = '" & descri & "' " & sql1 & ""
Case 3
    sql1 = IIf(vg_tipbase = "1", " and format(fechahoraregistro,'yyyymmdd') = " & fecbar & " ", " and convert(varchar(8),fechahoraregistro,112) = " & fecbar & "")
    DetalleLectura = "SELECT DISTINCT codigobarra FROM b_detallelectura a WHERE a.cli_codigo = '" & cencos & "' and a.reg_codigo = " & codreg & " and a.ser_codigo = " & codser & " and a.ate_codatencion = " & codate & " and a.codigobarra = '" & descri & "' " & sql1 & ""
Case 4
    sql1 = IIf(vg_tipbase = "1", " and format(fechahoraregistro,'yyyymmdd') = " & fecbar & " ", " and convert(varchar(8),fechahoraregistro,112) = " & fecbar & "")
    DetalleLectura = "SELECT DISTINCT codigobarra FROM b_detallelectura a WHERE a.cli_codigo = '" & cencos & "' and a.codigobarra = '" & descri & "'" ' " & sql1 & ""
Case 5
    DetalleLectura = "SELECT DISTINCT codigobarra FROM b_detallelectura a WHERE a.cli_codigo = '" & cencos & "' a.codigobarra = '" & descri & "' " & sql1 & ""
End Select
End Function

Public Function ParametroCodBarra(ByVal opcion As Long, ByVal codigo As String, ByVal descri As String) As String
Dim sql1 As String
ParametroCodBarra = ""
Select Case opcion
Case 1
    sql1 = "(" & descri & ")"
    ParametroCodBarra = "SELECT cbar_atributo, cli_codigo, cbar_posinicial, cbar_largo FROM a_par_codigo_barra_cas WHERE cli_codigo = '" & codigo & "' and cbar_atributo in " & sql1 & ""
End Select
End Function

Public Function ParametroTipoVales(ByVal opcion As Long, cencos As String) As String
ParametroTipoVales = ""
Select Case opcion
Case 1
    ParametroTipoVales = "SELECT cbar_atributo, Largo, Posicion_inicial FROM a_par_tipo_vales WHERE cli_codigo = '" & cencos & "'"
End Select
End Function

Public Function ParametroCodigoBarra(ByVal opcion As Long, cencos As String, tipval As String) As String
ParametroCodigoBarra = ""
Select Case opcion
Case 1
    ParametroCodigoBarra = "SELECT * FROM a_par_codigo_barra_cas WHERE cli_codigo = '" & cencos & "' and id_tipo_vale = '" & tipval & "' order by cbar_posinicial"
Case 2
    ParametroCodigoBarra = "SELECT * FROM a_par_codigo_barra_cas WHERE cli_codigo = '" & cencos & "'"
End Select
End Function

Function Receta(ByVal opcion As Long, ByVal CodRec As Long, ByVal catdie As Long, ByVal tippla As Long, ByVal descri As String, ByVal tiprec As Long) As String

Dim sql1 As String
Receta = ""

Select Case opcion

Case 1
    
    Receta = "SELECT rec_codigo, rec_nombre FROM b_receta WHERE (rec_catdie = " & catdie & " OR " & catdie & " = 0) " & _
             "AND   (rec_fecvig > " & Format(Date, "yyyymmdd") & " OR rec_fecvig <= 0 OR (rec_fecvig) IS NULL) " & _
             "AND   (rec_tippla = " & tippla & " OR " & tippla & " = 0) ORDER BY rec_nombre"

Case 2
    
    Receta = "SELECT a.rec_codigo, a.rec_nombre, a.rec_nomfan, a.rec_catdie, a.rec_tippla, a.rec_basrac, b.car_nombre, c.tip_nombre " & _
             "FROM b_receta as a with (nolock) " & _
             "inner join a_recetacatdie as b with (nolock) on a.rec_catdie = b.car_codigo " & _
             "inner join a_recetatippla c with (nolock) on   " & _
             "a.rec_tippla = c.tip_codigo WHERE a.rec_codigo = " & CodRec & ""

Case 3
    
    Receta = "SELECT a.rec_codigo, a.rec_nombre, b.red_tiprec, SUM(b.red_canpro*d.cpi_precos) AS cosrec " & _
             "FROM   b_receta a, b_recetadet b, b_ingrediente c, b_contlistpreing d " & _
             "WHERE  b.red_codigo = a.rec_codigo AND ((b.red_tiprec <> 0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) " & _
             "AND    b.red_codpro = c.ing_codigo AND c.ing_codigo = d.cpi_coding AND d.cpi_cencos = '" & MuestraCasino(1) & "' " & _
             "AND   (a.rec_fecvig > " & Format(Date, "yyyymmdd") & " OR a.rec_fecvig <= 0 OR (a.rec_fecvig) IS NULL) " & _
             "AND   (a.rec_catdie = " & catdie & " OR " & catdie & " = 0) " & _
             "AND   (a.rec_tippla = " & tippla & " OR " & tippla & " = 0) " & _
             "GROUP BY a.rec_codigo, a.rec_nombre, b.red_tiprec ORDER BY a.rec_nombre, a.rec_codigo, b.red_tiprec"

Case 4 '-------> Calculo APorte Nutricionales
    
    Receta = "SELECT DISTINCT e.nut_nombre, e.nut_codigo, SUM((((b.red_pctnut/100)*(d.pnu_canapo*(b.red_canpro/a.rec_basrac)))/c.ing_facnut)) AS candiet, " & _
             "e.nut_secnro FROM b_receta a, b_recetadet b, b_ingrediente c, b_productonut d, a_nutriente e WHERE b.red_codigo = a.rec_codigo " & _
             "AND b.red_codpro = c.ing_codigo AND c.ing_codigo = d.pnu_codpro AND d.pnu_codapo = e.nut_codigo " & _
             "AND b.red_codigo = " & CodRec & " AND b.red_tiprec = " & tiprec & " AND b.red_cencos = '" & IIf(tiprec = 0, 0, MuestraCasino(1)) & "' AND c.ing_facnut > 0 GROUP BY e.nut_nombre, e.nut_codigo, e.nut_secnro ORDER BY e.nut_secnro"

Case 5
    
    Receta = "SELECT DISTINCT b.red_codigo, SUM(b.red_canpro) AS canpro, SUM((b.red_canpro*(b.red_pctcoc/100)) * (b.red_pctapr / 100)) AS canser, " & _
             "SUM((b.red_pctnut / 100) * b.red_canpro) AS cannet FROM b_receta a, b_recetadet b, b_ingrediente c WHERE b.red_codigo = a.rec_codigo " & _
             "AND b.red_codpro = c.ing_codigo AND b.red_codigo = " & CodRec & " AND b.red_tiprec = " & tiprec & " AND b.red_cencos = '" & IIf(tiprec = 0, 0, MuestraCasino(1)) & "' GROUP BY b.red_codigo"

End Select

End Function

Public Function Municipio(ByVal opcion As Long, ByVal codigo As Long, ByVal descri As String) As String
Dim sql1 As String
Municipio = ""
Select Case opcion
Case 1
    Municipio = "SELECT mun_codigo, mun_nombre, mun_retobl From a_municipio WHERE mun_codigo = " & codigo & ""
End Select
End Function

Public Function CierrePeriodo(ByVal opcion As Long, ByVal cencos As String, ByVal fecper As Long, ByVal fecini As Long, ByVal fecter As Long) As String
Dim sql1 As String
CierrePeriodo = ""
Select Case opcion
Case 1
    CierrePeriodo = "SELECT DISTINCT cie_fecini, cie_fecter, cie_periodo FROM b_cierreperiodo " & _
                    "WHERE cie_cencos = '" & cencos & "' " & _
                    "AND   cie_fecini <= " & fecini & " " & _
                    "AND   cie_fecter >= " & fecter & ""
End Select
End Function

Public Function TomaInventario(ByVal opcion As Long, ByVal codbod As Long, ByVal codpro As String, ByVal fectom As Long) As String
Dim sql1 As String
TomaInventario = ""
Select Case opcion
Case 1 '-------> Traer Inventario del mes anterior
    TomaInventario = "SELECT tin.tin_fectom, pro.pro_codigo, pro.pro_nombre, tin.tin_stofis, tin_propon " & _
                     "FROM b_tomainv tin, b_productos pro " & _
                     "WHERE tin.tin_codpro = pro.pro_codigo " & _
                     "AND   tin.tin_codbod = " & codbod & " " & _
                     "AND   tin.tin_codpro = '" & codpro & "' " & _
                     "AND   tin.tin_fectom <= " & fectom & " " & _
                     "ORDER BY tin.tin_fectom DESC"
End Select
End Function

Public Function AjusteInventario(ByVal opcion As Long, ByVal codbod As Long, ByVal codpro As String, ByVal fecini As Long, ByVal fecfin As Long) As String
Dim sql1 As String, sql2 As String
AjusteInventario = ""
Select Case opcion
Case 1 '-------> Traer Ajuste Inventario
    sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
    AjusteInventario = "SELECT tov.tov_fecemi, pro.pro_codigo, pro.pro_nombre, aju.aju_nombre, aju.aju_tipo, tov.tov_numdoc, dev.dev_canmin, dev.dev_precos " & _
                       "FROM b_totventas tov, b_detventas dev, b_productos pro, a_tipoajuste aju " & _
                       "WHERE tov.tov_rutcli = dev.dev_rutcli " & _
                       "AND   tov.tov_tipdoc = dev.dev_tipdoc " & _
                       "AND   tov.tov_numdoc = dev.dev_numdoc " & _
                       "AND   tov.tov_codser = aju.aju_codigo " & _
                       "AND   dev.dev_codmer = pro.pro_codigo " & _
                       "AND   tov.tov_codbod = " & codbod & " " & _
                       "AND   dev.dev_codmer = '" & codpro & "' " & _
                       "AND   tov.tov_fecemi >  " & sql1 & " " & _
                       "AND   tov.tov_fecemi <= " & sql2 & " " & _
                       "AND   tov.tov_tipdoc = 'AI' " & _
                       "AND   tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' " & _
                       "ORDER BY tov.tov_fecemi, pro.pro_nombre"

End Select
End Function

Public Function DocPro(ByVal opcion As Long, ByVal codbod As Long, ByVal codpro As String, ByVal fecini As Long, ByVal fecfin As Long, ByVal ctacon As String) As String

Dim sql1 As String, sql2 As String
DocPro = ""

Select Case opcion
    
    Case 1 '-------> Traer documento proveedor
        
        sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
        sql2 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
        
        DocPro = "SELECT toc.toc_fecdig, toc.toc_fecrem, pro.pro_codigo, pro.pro_nombre, de.dec_numdoc, " & _
                 "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
                 "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
                 "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
                 "WHERE toc.toc_rutpro = de.dec_rutpro AND toc.toc_tipdoc=de.dec_tipdoc " & _
                 "AND   toc.toc_numdoc = de.dec_numdoc AND de.dec_codmer=pro.pro_codigo " & _
                 "AND   toc.toc_codbod = " & codbod & " " & _
                 "AND   pro.pro_codigo = '" & codpro & "' " & _
                 "AND   de.dec_mueinv = 'S' AND toc.toc_tipdoc not in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND toc.toc_tipdoc NOT IN (select tdo_codigo from a_tipodocumento where tdo_IdCodigo in ('NC', 'CE')) AND de.dec_canrec > 0 " & _
                 "AND   toc.toc_fecrem > " & sql1 & " " & _
                 "AND   toc.toc_fecrem <= " & sql2 & " " & _
                 "ORDER BY toc.toc_fecdig, toc.toc_fecrem, pro.pro_nombre"

End Select

End Function

Public Function DocProImp(ByVal opcion As Long, ByVal rutpro As String, ByVal tipdoc As String, ByVal NumDoc As Long, ByVal numlin As Long, ByVal codpro As String) As String
Dim sql1 As String, sql2 As String
DocProImp = ""
Select Case opcion
Case 1 '-------> Traer documento proveedor impuesto
    DocProImp = "SELECT b.imp_inccos, SUM(a.imd_monimp) AS imd_monimp " & _
                "FROM  b_detcomprasimp a, a_impuesto b " & _
                "WHERE a.imd_rutdoc = '" & rutpro & "' " & _
                "AND   a.imd_tipdoc = '" & tipdoc & "' " & _
                "AND   a.imd_numdoc = " & NumDoc & " " & _
                "AND   a.imd_numlin = " & numlin & " " & _
                "AND   a.imd_codpro = '" & codpro & "' " & _
                "AND   a.imd_codimp = b.imp_codigo " & _
                "AND   b.imp_inccos = 1  GROUP BY b.imp_inccos"
Case 2 '-------> Traer documento proveedor impuesto calculo (cantidad recibida * precio documento)*()
    DocProImp = "SELECT b.imp_inccos, SUM(((( c.dec_precom * c.dec_canrec ) * ( a.imd_pctimp / 100 )))) AS imd_monimp " & _
                "FROM  b_detcomprasimp a, a_impuesto b, b_detcompras c " & _
                "WHERE a.imd_rutdoc = c.dec_rutpro " & _
                "AND   a.imd_tipdoc = c.dec_tipdoc " & _
                "AND   a.imd_numdoc = c.dec_numdoc " & _
                "AND   a.imd_numlin = c.dec_numlin " & _
                "AND   a.imd_codpro = c.dec_codmer " & _
                "AND   a.imd_rutdoc = '" & rutpro & "' " & _
                "AND   a.imd_tipdoc = '" & tipdoc & "' " & _
                "AND   a.imd_numdoc = " & NumDoc & " " & _
                "AND   a.imd_numlin = " & numlin & " " & _
                "AND   a.imd_codpro = '" & codpro & "' " & _
                "AND   a.imd_codimp = b.imp_codigo " & _
                "AND   b.imp_inccos = 1  GROUP BY b.imp_inccos"
End Select
End Function

Public Function DocTraspaso(ByVal opcion As Long, ByVal codbod As Long, ByVal codpro As String, ByVal fecini As Long, ByVal fecfin As Long) As String
Dim sql1 As String, sql2 As String
DocTraspaso = ""
Select Case opcion
Case 1 '-------> Traer documento traspaso entrada
    sql1 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecini) & "') ", " '" & Format(fg_Ctod1(fecini), "yyyymmdd") & "' ")
    sql2 = IIf(vg_tipbase = "1", " cdate('" & fg_Ctod1(fecfin) & "') ", " '" & Format(fg_Ctod1(fecfin), "yyyymmdd") & "' ")
    DocTraspaso = "SELECT tov.tov_fecemi, tov.tov_codser, pro.pro_codigo, pro.pro_nombre, tov.tov_numdoc, dev.dev_canmer, dev.dev_precos " & _
                  "FROM   b_totventas tov, b_detventas dev, b_productos pro " & _
                  "WHERE  tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc " & _
                  "AND    tov.tov_numdoc = dev.dev_numdoc AND dev.dev_codmer = pro.pro_codigo " & _
                  "AND    tov.tov_codbod = " & codbod & " AND pro.pro_codigo = '" & codpro & "' " & _
                  "AND    tov.tov_tipdoc = 'TR' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' " & _
                  "AND    tov.tov_fecemi > " & sql1 & " " & _
                  "AND    tov.tov_fecemi <= " & sql2 & " AND dev.dev_mueinv = 'S' " & _
                  "ORDER BY tov.tov_fecemi, pro.pro_nombre"
End Select
End Function

Public Function PrecioVta(ByVal opcion As Long, ByVal cencos As String, ByVal codreg As Long, ByVal codser As Long, ByVal Fecha As Long, ByVal codcli As String) As String
Dim sql1 As String, sql2 As String
PrecioVta = ""
Select Case opcion
Case 1 '-------> Precio Vta Cliente
    PrecioVta = "SELECT a.cli_codigo, isnull(a.cli_nombre,'') as cli_nombre, isnull(b.prv_preven,0) as prv_preven, isnull(b.prv_SPRS,'') as prv_SPRS " & _
                "FROM   b_clientes a, b_preciovta b " & _
                "WHERE  b.prv_rutcli = a.cli_codigo " & _
                "AND    b.prv_cencos = '" & cencos & "' " & _
                "AND    b.prv_codreg = " & codreg & " " & _
                "AND    b.prv_codser = " & codser & " " & _
                "AND    b.prv_fecvig = " & Fecha & " " & _
                "AND   (b.prv_rutcli = '" & codcli & "' OR '" & codcli & "' = '') " & _
                "AND    a.cli_tipo   = 1"
End Select
End Function

Public Function TipoVales(ByVal opcion As Long, ByVal codigo As Long, ByVal descri As String) As String
Dim sql1 As String
TipoVales = ""
Select Case opcion
Case 1
    TipoVales = "SELECT '  ' AS ID_Tipo_Vale, 'Seleccione...' AS Nombre, '00' AS Orden union SELECT ID_Tipo_Vale, Nombre, '01' AS Orden FROM a_par_tipo_vales WHERE cli_codigo = '" & descri & "' ORDER BY Orden ASC"
End Select
End Function

