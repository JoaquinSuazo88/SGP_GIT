VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_Guias 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   6120
   ClientTop       =   2445
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   4965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   4905
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2145
         Left            =   165
         TabIndex        =   2
         Top             =   300
         Width           =   4560
         _Version        =   393216
         _ExtentX        =   8043
         _ExtentY        =   3784
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   6
         MaxRows         =   1
         ScrollBars      =   2
         SpreadDesigner  =   "B_Guias.frx":0000
         ScrollBarTrack  =   1
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   360
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "B_Guias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset
Dim Impuestos() As Variant
Dim ctipdoc As String
Dim rutpro As String
Dim form1 As Object
Dim tipopc As String
Dim tpentg As Long

Public Function Cargar_DoctoGrilla(Form As Object, TipoDoc As String, TituloForm As String, rut As String, tipop As String, tpent As Long) As Boolean

On Error GoTo Man_Error

Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String

Cargar_DoctoGrilla = False

Set form1 = Form
ctipdoc = Trim(TipoDoc)
rutpro = rut
MsgTitulo = TituloForm
Caption = TituloForm
tipopc = tipop
tpentg = tpent
vaSpread1.MaxRows = 0

If TipoDoc = "SN" Then
   
   Toolbar1.Width = 6240
   Frame1.Width = 6230
   vaSpread1.Width = 5860
   sql1 = IIf(vg_tipbase = "1", " TRIM(toc_docaso) ", " lTRIM(toc_docaso) ")
   sql2 = IIf(vg_tipbase = "1", " TRIM(toc_numdoc) ", " lTRIM(toc_numdoc) ")
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
'   RS1.Open "SELECT toc_numdoc, toc_fecemi, toc_totdoc FROM b_totcompras WHERE " & sql1 & " IN (SELECT " & sql2 & " FROM b_totcompras WHERE (toc_tipdoc = 'FA' OR toc_tipdoc = 'FE') AND (toc_docsnc = '' OR toc_docsnc IS NULL) AND toc_codbod = " & vg_codbod & ") AND toc_codbod = " & vg_codbod & " AND toc_rutpro = '" & rut & "' AND toc_tipdoc = '" & Trim(TipoDoc) & "' AND (toc_docsnc = '' OR (toc_docsnc) IS NULL) ORDER BY toc_fecemi, toc_numdoc", vg_db, adOpenStatic
   Set RS1 = vg_db.Execute("sgp_Sel_TraerSolicitudNotaCreditoSinCerrar '" & rut & "', " & vg_codbod & ", '" & TipoDoc & "'")
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Function
   RS1.Close
   Set RS1 = Nothing
   
   Dim canali As Double, pctdes As Double, pctimp As Double
   Dim auxnumdoc As Long, auxdocfac As Long
   Dim auxfecha As Date
   
   sql1 = IIf(vg_tipbase = "1", " trim(a.toc_docsnc) ", " ltrim(a.toc_docsnc) ")
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
'   RS1.Open "SELECT a.toc_fecemi, a.toc_docaso, b.dec_numdoc, b.dec_codmer, b.dec_canmer, b.dec_precom, b.dec_ptotal, " & _
'            "b.dec_pctdes, b.dec_canrec, b.dec_prerec, b.dec_ptotrec, b.dec_prefle, b.dec_numlin  " & _
'            "FROM  b_totcompras a, b_detcompras b, b_productos c " & _
'            "WHERE a.toc_rutpro = b.dec_rutpro " & _
'            "AND   a.toc_tipdoc = b.dec_tipdoc " & _
'            "AND   a.toc_numdoc = b.dec_numdoc " & _
'            "AND   b.dec_codmer = c.pro_codigo " & _
'            "AND   a.toc_rutpro = '" & rut & "' " & _
'            "AND   a.toc_tipdoc = '" & TipoDoc & "' " & _
'            "AND   a.toc_codbod = " & vg_codbod & " " & _
''            "AND (" & sql1 & " = '' or (a.toc_docsnc) IS NULL) " & _
'            "ORDER BY a.toc_fecemi, b.dec_numdoc", vg_db, adOpenStatic
    Set RS1 = vg_db.Execute("sgp_Sel_TraerDetalleSolicitudNotaCreditoSinCerrar '" & rut & "', " & vg_codbod & ", '" & TipoDoc & "'")
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Function
    If Not RS1.EOF Then
       
       auxnumdoc = RS1!dec_numdoc
       auxfecha = RS1!toc_fecemi
       canali = 0
       auxdocfac = RS1!toc_docaso
       
       Do While Not RS1.EOF
          
          If RS1!dec_numdoc <> auxnumdoc Then
             
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.Row = vaSpread1.MaxRows
             If InStr(vg_Guias, auxnumdoc) <> 0 Then vaSpread1.Col = 1: vaSpread1.Value = 0
             
             vaSpread1.Col = 2
             vaSpread1.TypeHAlign = TypeHAlignRight
             vaSpread1.Value = auxnumdoc
             
             vaSpread1.Col = 3
             vaSpread1.ColHidden = False
             vaSpread1.TypeHAlign = TypeHAlignRight
             vaSpread1.Value = auxdocfac
             
             vaSpread1.Col = 4
             vaSpread1.TypeHAlign = TypeHAlignRight
             vaSpread1.Value = auxfecha
             
             vaSpread1.Col = 5
             vaSpread1.TypeHAlign = TypeHAlignRight
             vaSpread1.Value = Format(canali, fg_Pict(6, 2))
             
             Cargar_DoctoGrilla = True
             auxnumdoc = RS1!dec_numdoc
             auxfecha = RS1!toc_fecemi
             canali = 0
             auxdocfac = RS1!toc_docaso
          
          End If
          pctimp = 0
          pctdes = 0
          If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
          canali = Round(canali + (((RS1!dec_ptotal - RS1!dec_ptotrec)) - (((RS1!dec_ptotal - RS1!dec_ptotrec)) * (pctdes / 100)) + pctimp), vg_DPr)
          '------- Fin traer Impuesto adicionales
          RS1.MoveNext
       
       Loop
       
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
       If InStr(vg_Guias, auxnumdoc) <> 0 Then vaSpread1.Col = 1: vaSpread1.Value = 0
       
       vaSpread1.Col = 2
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.Value = auxnumdoc
       
       vaSpread1.Col = 3
       vaSpread1.ColHidden = False
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.Value = auxdocfac
       
       vaSpread1.Col = 4
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.Value = auxfecha
       
       vaSpread1.Col = 5
       vaSpread1.TypeHAlign = TypeHAlignRight
       vaSpread1.Value = Format(canali, fg_Pict(6, 2))
       
       Cargar_DoctoGrilla = True
    
    End If
    RS1.Close
    Set RS1 = Nothing

ElseIf TipoDoc = "GD" Then
   
'   RS1.Open "SELECT toc_numdoc, toc_fecemi, toc_totdoc FROM b_totcompras WHERE toc_rutpro = '" & rut & "' " & _
'            "AND (toc_docaso = '' OR toc_docaso IS NULL) AND toc_tipdoc = '" & Trim(TipoDoc) & "' AND toc_codbod = " & vg_codbod & " " & _
'            "ORDER BY toc_fecemi, toc_numdoc", vg_db, adOpenStatic
   Set RS1 = vg_db.Execute("sgp_Sel_TraerGuiaSinCerrar '" & rut & "', " & vg_codbod & ", '" & Trim(TipoDoc) & "'")
   Do While Not RS1.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      If InStr(vg_Guias, RS1!toc_numdoc) <> 0 Then vaSpread1.Col = 1: vaSpread1.Value = 0
      
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS1!toc_numdoc
      
      vaSpread1.Col = 3
      vaSpread1.ColHidden = True
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = ""
      
      vaSpread1.Col = 4
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS1!toc_fecemi
      
      vaSpread1.Col = 5
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = Format(RS1!toc_totdoc, fg_Pict(6, 2))
      
      vaSpread1.Col = 6
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS1!toc_tipdoc
      
      Cargar_DoctoGrilla = True
      RS1.MoveNext
   
   Loop
   RS1.Close
   Set RS1 = Nothing

ElseIf TipoDoc = "OC" Then
   
   Toolbar1.Width = 6240
   Frame1.Width = 6230
   vaSpread1.Width = 5860
   sql1 = IIf(vg_tipbase = "1", " val(format(solite_dtent, 'yyyymm')) ", " substring(CONVERT(varchar(10), solite_dtent,112),1,6) ")
   If tipopc = "docpro" Then
      
      sql2 = IIf(vg_tipbase = "1", " '" & Format(Form.Date1(0), "yyyymm") & "' ", " '" & Format(Form.Date1(0), "yyyymm") & "' ")
   
   Else
      
      sql2 = IIf(vg_tipbase = "1", " '" & Format(Form.fpDateTime1(0), "yyyymm") & "' ", " '" & Format(Form.fpDateTime1(0), "yyyymm") & "' ")
   
   End If
   sql3 = IIf(vg_tipbase = "1", " SUM(IIF(tipsol_idsol = 4,(-1 * pedite_qtcpa),pedite_qtcpa)) AS Cantidad ", " SUM(CASE WHEN tipsol_idsol = 4 THEN (-1 * pedite_qtcpa) ELSE pedite_qtcpa END) AS Cantidad ")
   sql4 = IIf(vg_tipbase = "1", " SUM(IIF(tipsol_idsol = 4,(-1 * pedite_qtcpa),pedite_qtcpa)) > 0 ", " SUM(CASE WHEN tipsol_idsol = 4 THEN (-1 * pedite_qtcpa) ELSE pedite_qtcpa END) > 0 ")
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   RS1.Open "SELECT DISTINCT solite_dtent, " & sql3 & " " & _
            "FROM b_ocsac " & _
            "WHERE cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
            "AND  (cadfor_nrcgc = '" & rut & "' OR '" & tipopc & "' = 'traspa') " & _
            "AND   " & sql1 & "   = " & sql2 & " AND pedite_flafo = " & tpent & " " & _
            "GROUP BY solite_dtent HAVING " & sql4 & "", vg_db, adOpenStatic
   
   Do While Not RS1.EOF
      
      sql3 = IIf(vg_tipbase = "1", " SUM(iif(a.tipsol_idsol = 4, (-1 * a.pedite_qtcpa), a.pedite_qtcpa) - b.ocr_cancom) AS difer ", " SUM(CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE a.pedite_qtcpa END - b.ocr_cancom) AS difer ")
      sql4 = IIf(vg_tipbase = "1", " cdate('" & RS1!solite_dtent & "') ", " '" & Format(RS1!solite_dtent, "yyyymmdd") & "' ")
      
      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      RS2.Open "SELECT " & sql3 & " " & _
               "FROM   b_ocsac a, b_ocsacrecibido b, b_formatocompras c, b_formatocomprassgp d, b_productos e " & _
               "WHERE  a.cadfor_nrcgc = b.ocr_rutpro " & _
               "AND    a.solite_dtent = b.ocr_fecoc  AND a.cpopro_cdpro = b.ocr_codprodsac AND a.cpopro_cdpro = c.foc_codsac AND c.foc_codsac = d.fcs_codsac and d.fcs_codsgp = e.pro_codigo " & _
               "AND    b.ocr_fecoc IS NOT NULL " & _
               "AND    a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
               "AND    " & sql1 & "   = " & sql2 & " " & _
               "AND   (a.cadfor_nrcgc = '" & rut & "' OR '" & tipopc & "' = 'traspa') AND a.solite_dtent = " & sql4 & " AND a.pedite_flafo = " & tpent & "", vg_db, adOpenStatic
      
      If RS2.EOF Or RS2!difer <> 0 Or IsNull(RS2!difer) Then
            
            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
  '         If InStr(vg_Guias, RS1!solfil_idsol) <> 0 Then vaSpread1.Col = 1: vaSpread1.Value = 0
            
            vaSpread1.Col = 2
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.text = ""
            vaSpread1.ColHidden = True
            
            vaSpread1.Col = 3
            vaSpread1.ColHidden = True
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.Value = ""
            vaSpread1.ColHidden = True
            
            vaSpread1.Col = 4
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.Value = Format(RS1!solite_dtent, "dd/mm/yyyy")
            
            vaSpread1.Col = 5
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.Value = ""
            vaSpread1.ColHidden = True
            
            Cargar_DoctoGrilla = True
      
      End If
      
      RS2.Close
      Set RS2 = Nothing
      RS1.MoveNext
   
   Loop
   RS1.Close: Set RS1 = Nothing

End If

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

'-------> Dimensiona el formulario a su estado de diseño
Me.Width = IIf(ctipdoc = "GD", 5055, 6500)
Me.Height = 3645
fg_centra Me
fg_carga ""
MsgTitulo = "Guías de Despacho"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Exit Sub
Man_Error:
MsgBox Err & ": " & Err.Number & "-" & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Error_Carga

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim i As Long, docaso As Long
Dim v_valdesc  As Double, v_valtot As Double, candoc As Double
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String
Dim tipdoc As String
Dim numdoc As Long

Select Case Button.Index

    Case 1
        
        '-------> En caso de que el documento ya exista
        form1.vaSpread1.Row = -1
        form1.vaSpread1.Col = 1
        
        If form1.vaSpread1.Lock = True And tipopc = "docpro" Then
            
            If MsgBox("Desea modificar selección...", vbInformation + vbOKCancel, MsgTitulo) = vbCancel Then Exit Sub
        
        End If
        
        '-------> Fin documento exista
        vg_Guias = ""
        vg_GuiasTipo = ""
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.Value = 1 Then
               
               form1.vaSpread1.MaxRows = 0
            
            End If
        
        Next i
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.Value = 1 Then
               
               If vg_FDC = "GD" Or vg_FDC = "SN" Then
                  
                  '-------> Carga Encabezado de Documento
                  If vg_FDC = "GD" Then
                  
                     vaSpread1.Col = 6
                     vg_GuiasTipo = vg_GuiasTipo & vaSpread1.Value & ";"
                  
                  End If
                  
                  vaSpread1.Col = 2
                  vg_Guias = vg_Guias & vaSpread1.Value & ";"
                  
                  
                  sql1 = IIf(vg_tipbase = "1", " trim(toc_rutpro) ", " ltrim(toc_rutpro) ")
                  
                  If RS1.State = 1 Then RS1.Close
                  RS1.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
'                  RS1.Open "SELECT toc_numinf, toc_tipinf, toc_docaso FROM b_totcompras WHERE " & sql1 & " = '" & vg_RDC & "' AND toc_numdoc = " & Val(vaSpread1.Value) & " AND toc_tipdoc = '" & vg_FDC & "' AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
                  If vg_FDC = "GD" Then
                     
                     Set RS1 = vg_db.Execute("sgp_Sel_ValidaGuiasPendientes '" & vg_RDC & "', " & Val(vaSpread1.Value) & ", " & vg_codbod & ", '" & vg_FDC & "'")
                     
                  ElseIf vg_FDC = "SN" Then
                  
                     Set RS1 = vg_db.Execute("sgp_Sel_ValidaSolicitudNPendientes '" & vg_RDC & "', " & Val(vaSpread1.Value) & ", " & vg_codbod & ", '" & vg_FDC & "'")
                  
                  End If
                  
                  docaso = IIf(IsNull(RS1!toc_docaso) Or Trim(RS1!toc_docaso) = "", 0, Val(RS1!toc_docaso))
                  If RS1!toc_tipinf = "C" Then
                     
                     form1.Option1(1).Value = True
                  
                  ElseIf RS1!toc_tipinf = "F" Then
                     
                     form1.Option1(0).Value = True
                  
                  End If
                  RS1.Close
                  Set RS1 = Nothing
                  
                  '-------> Detalle de Documento
                  If vg_pais = "CO" Then
                     
                     sql5 = IIf(vg_FDC = "SN", " IN ('FA','FE') ", " IN ('GD') ")
                     docaso = IIf(vg_FDC = "GD", Val(vaSpread1.Value), docaso)
                     
                     If RS2.State = 1 Then RS2.Close
                     RS2.CursorLocation = adUseClient
                     vg_db.CursorLocation = adUseClient
                     RS2.Open "SELECT DISTINCT ocr_rutpro FROM b_ocsacrecibido WHERE ocr_rutpro = '" & vg_RDC & "' AND ocr_tipdoc " & sql5 & " AND ocr_numdoc = " & docaso & "", vg_db, adOpenStatic
                     If Not RS2.EOF Then
                        
                        If RS1.State = 1 Then RS1.Close
                        RS1.CursorLocation = adUseClient
                        vg_db.CursorLocation = adUseClient
                        
                        RS1.Open "SELECT a.*, b.*, c.uni_nombre, (SELECT TOP 1 ocr_codprodsac FROM b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc " & sql5 & " AND ocr_numdoc = " & docaso & " AND ocr_codprodsgp = a.dec_codmer) AS ocr_codprodsac, " & _
                                 "(SELECT TOP 1 e.foc_nomsac FROM b_formatocompras e, b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc " & sql5 & " AND ocr_numdoc = " & docaso & " AND ocr_codprodsgp = a.dec_codmer AND ocr_codprodsac = e.foc_codsac) AS foc_nomsac, " & _
                                 "(SELECT TOP 1 e.foc_unisac FROM b_formatocompras e, b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc " & sql5 & " AND ocr_numdoc = " & docaso & " AND ocr_codprodsgp = a.dec_codmer AND ocr_codprodsac = e.foc_codsac) AS foc_unisac " & _
                                 "FROM b_detcompras a, b_productos b, a_unidad c WHERE a.dec_codmer = b.pro_codigo AND b.pro_coduni = c.uni_codigo AND a.dec_rutpro = '" & vg_RDC & "' AND a.dec_tipdoc = '" & vg_FDC & "' AND a.dec_numdoc = " & Val(vaSpread1.Value) & " ORDER BY a.dec_numlin", vg_db, adOpenStatic
                     
                     Else
                        
                        If RS1.State = 1 Then RS1.Close
                        RS1.CursorLocation = adUseClient
                        vg_db.CursorLocation = adUseClient

                        RS1.Open "SELECT a.*, b.*, c.uni_nombre, '' AS ocr_codprodsac, '' AS foc_nomsac, '' AS foc_unisac FROM b_detcompras a, b_productos b, a_unidad c WHERE a.dec_codmer = b.pro_codigo AND b.pro_coduni = c.uni_codigo AND a.dec_rutpro = '" & vg_RDC & "' AND a.dec_tipdoc = '" & vg_FDC & "' AND a.dec_numdoc = " & Val(vaSpread1.Value) & " ORDER BY a.dec_numlin", vg_db, adOpenStatic
                     
                     End If
                     RS2.Close
                     Set RS2 = Nothing
                  
                  Else
                     
                     If RS1.State = 1 Then RS1.Close
                     RS1.CursorLocation = adUseClient
                     vg_db.CursorLocation = adUseClient
                     
                     If vg_FDC = "GD" Then
                                          
'                        RS1.Open "SELECT a.*, b.*, c.uni_nombre, '' AS ocr_codprodsac, '' AS foc_nomsac, '' AS foc_unisac FROM b_detcompras a, b_productos b, a_unidad c WHERE a.dec_codmer = b.pro_codigo AND b.pro_coduni = c.uni_codigo AND a.dec_rutpro = '" & vg_RDC & "' AND a.dec_tipdoc = '" & vg_FDC & "' AND a.dec_numdoc = " & Val(vaSpread1.Value) & " ORDER BY a.dec_numlin", vg_db, adOpenStatic
                        numdoc = vaSpread1.text
                        
                        vaSpread1.Col = 6
                        tipdoc = IIf(Trim(vaSpread1.text) <> "", Trim(vaSpread1.text), "")
                        
                        Set RS1 = vg_db.Execute("sgp_Sel_DetalleComprasGuiasPendientes '" & vg_RDC & "', " & numdoc & ", '" & tipdoc & "'")
                  
                     ElseIf vg_FDC = "SN" Then
                     
                        Set RS1 = vg_db.Execute("sgp_Sel_DetalleComprasSNotaPendientes '" & vg_RDC & "', " & Val(vaSpread1.Value) & ", '" & vg_FDC & "' ")
                     
                     End If
                  End If
                  '----Deshabilitar Botones
                  form1.Frame3.Enabled = False
                  form1.Frame6.Enabled = False
                  form1.Frame5.Enabled = False
                  
                  For j = 1 To form1.vaSpread2.MaxRows
                      
                      form1.vaSpread2.Row = j
                      
                      form1.vaSpread2.Col = 1
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 2
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 3
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 4
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 5
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 6
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 7
                      form1.vaSpread2.Lock = True
                      
                      form1.vaSpread2.Col = 8
                      form1.vaSpread2.Lock = True
                      
                      If form1.vaSpread2.text = "1" Then form1.vaSpread2.Col = 5:  form1.vaSpread2.Lock = False
                  
                  Next j
                  '------->  Asigna código cfc o fifo segun corresponda
                  '-------> MSP : 16/08/2004
                  '-------> Detalle de Documento
                  Do While Not RS1.EOF
                     
                     form1.vaSpread1.MaxRows = form1.vaSpread1.MaxRows + 1
                     
                     form1.vaSpread1.Col = 1
                     form1.vaSpread1.Row = form1.vaSpread1.MaxRows
                     form1.vaSpread1.Lock = True
                     form1.vaSpread1.Value = IIf(vg_pais = "CO", RS1!ocr_codprodsac, RS1!dec_codmer)
                     
                     form1.vaSpread1.Col = 2
                     form1.vaSpread1.Lock = True
                     form1.vaSpread1.Value = IIf(vg_pais = "CO", Trim(RS1!foc_nomsac), Trim(RS1!pro_nombre))
                     
                     form1.vaSpread1.Col = 3
                     form1.vaSpread1.Lock = True
                     form1.vaSpread1.Value = IIf(vg_pais = "CO", Trim(RS1!foc_unisac), Trim(RS1!uni_nombre))
                     
                     If vg_FDC = "GD" Then
                         
                         form1.vaSpread1.Col = 4
                         form1.vaSpread1.Lock = False
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", RS1!dec_canmer, RS1!dec_cmefac)
                         
                         form1.vaSpread1.Col = 5
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", RS1!dec_precom, RS1!dec_pmefac)
                         
                         form1.vaSpread1.Col = 6
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = RS1!dec_pctdes
                         
                         form1.vaSpread1.Col = 7
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = RS1!dec_valdes
                         
                         form1.vaSpread1.Col = 8
                         form1.vaSpread1.Lock = False
                         form1.vaSpread1.Value = Round(RS1!dec_ptotal, 0)
                         
                         form1.vaSpread1.Col = 9
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", RS1!dec_canrec, RS1!dec_crefac)
                         
                         form1.vaSpread1.Col = 10
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", RS1!dec_prerec, RS1!dec_prefac)
                         
                         form1.vaSpread1.Col = 17
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", RS1!dec_prerec, RS1!dec_prefac)
                     
                     ElseIf vg_FDC = "SN" Then
                         
                         form1.vaSpread1.Col = 4
                         form1.vaSpread1.Lock = False
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", IIf(RS1!dec_canmer = RS1!dec_canrec, RS1!dec_canmer, (RS1!dec_canmer - RS1!dec_canrec)), IIf(RS1!dec_cmefac = RS1!dec_crefac, RS1!dec_cmefac, (RS1!dec_cmefac - RS1!dec_crefac)))
                         
                         form1.vaSpread1.Col = 5
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = IIf(vg_pais <> "CO", IIf(RS1!dec_precom = RS1!dec_prerec, RS1!dec_precom, (RS1!dec_precom - RS1!dec_prerec)), IIf(RS1!dec_pmefac = RS1!dec_prefac, RS1!dec_pmefac, (RS1!dec_pmefac - RS1!dec_prefac)))
                         
                         form1.vaSpread1.Col = 6
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = RS1!dec_pctdes
                         
                         If vg_pais <> "CO" Then
                            
                            v_valdesc = (IIf(RS1!dec_canmer = RS1!dec_canrec, RS1!dec_canmer, (RS1!dec_canmer - RS1!dec_canrec)) * IIf(RS1!dec_precom = RS1!dec_prerec, RS1!dec_precom, (RS1!dec_precom - RS1!dec_prerec))) * (RS1!dec_pctdes / 100)
                            v_valtot = (IIf(RS1!dec_canmer = RS1!dec_canrec, RS1!dec_canmer, (RS1!dec_canmer - RS1!dec_canrec)) * IIf(RS1!dec_precom = RS1!dec_prerec, RS1!dec_precom, (RS1!dec_precom - RS1!dec_prerec))) - v_valdesc
                         
                         Else
                            
                            v_valdesc = (IIf(RS1!dec_cmefac = RS1!dec_crefac, RS1!dec_cmefac, (RS1!dec_cmefac - RS1!dec_crefac)) * IIf(RS1!dec_pmefac = RS1!dec_prefac, RS1!dec_pmefac, (RS1!dec_pmefac - RS1!dec_prefac))) * (RS1!dec_pctdes / 100)
                            v_valtot = (IIf(RS1!dec_cmefac = RS1!dec_crefac, RS1!dec_cmefac, (RS1!dec_cmefac - RS1!dec_crefac)) * IIf(RS1!dec_pmefac = RS1!dec_prefac, RS1!dec_pmefac, (RS1!dec_pmefac - RS1!dec_prefac))) - v_valdesc
                         
                         End If
                         
                         form1.vaSpread1.Col = 7
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = v_valdesc
                         
                         form1.vaSpread1.Col = 8
                         form1.vaSpread1.Lock = False
                         form1.vaSpread1.Value = v_valtot
                         
                         form1.vaSpread1.Col = 9
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = 0
                         
                         form1.vaSpread1.Col = 10
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = 0
                         
                         form1.vaSpread1.Col = 17
                         form1.vaSpread1.Lock = True
                         form1.vaSpread1.Value = 0
                     
                     End If
                     
                     form1.vaSpread1.Col = 11
                     form1.vaSpread1.Value = RS1!dec_descri
                     
                     form1.vaSpread1.Col = 13
                     form1.vaSpread1.Value = RS1!dec_mueinv
                     
                     form1.vaSpread1.Col = 14
                     form1.vaSpread1.text = ""
                     
                     form1.vaSpread1.Col = 18
                     form1.vaSpread1.Value = IIf(vg_pais <> "CO", RS1!dec_precom, RS1!dec_pmefac)
                     
                     form1.vaSpread1.Col = 24
                     form1.vaSpread1.Lock = True
                     
                     form1.vaSpread1.Value = IIf(vg_pais = "CO", RS1!dec_codmer, IIf(IsNull(RS1!ocr_codprodsac), "", Trim(RS1!ocr_codprodsac)))
                     
                     form1.vaSpread1.Col = 25
                     form1.vaSpread1.Lock = True
                     form1.vaSpread1.Value = IIf(vg_pais = "CO", IIf(IsNull(RS1!pro_nombre), "No existe descripción SGP", Trim(RS1!pro_nombre)), IIf(IsNull(RS1!foc_nomsac), "No existe descripción SAC", Trim(RS1!foc_nomsac)))
                     
                     form1.vaSpread1.Col = 29
                     form1.vaSpread1.Value = IIf(IsNull(RS1!dec_faccon), 0, RS1!dec_faccon)
                     
                     Revisa RS1!dec_codmer, form1.vaSpread1.Row
                     
                     RS1.MoveNext
                 
                 Loop
                 RS1.Close
                 Set RS1 = Nothing
                 
                 If form1.vaSpread1.MaxRows > 0 Then
                    
                    form1.vaSpread1.Row = 1
                    form1.vaSpread1.Col = 24
                    
                    form1.Text2(0).text = Trim(form1.vaSpread1.text)
                    form1.vaSpread1.Row = 1
                    
                    form1.vaSpread1.Col = 25
                    form1.Text2(1).text = Trim(form1.vaSpread1.text)
                    
                    form1.vaSpread1.Row = 1
                    form1.vaSpread1.Col = 29
                    form1.Text2(2).text = Trim(form1.vaSpread1.text)
                 
                 End If
                 '-------> Fin detalle de Documento
               
               ElseIf vg_FDC = "OC" Then
                 
                 '-------> Bloquear columna de ordenes de compras
                 form1.vaSpread1.Visible = False
                 vaSpread1.Col = 4: vg_Guias = vg_Guias & vaSpread1.Value & ";"
                 If tipopc = "docpro" Then form1.Option1(1).Value = True
                 sql1 = IIf(vg_tipbase = "1", " SUM(IIF(a.tipsol_idsol = 4,(-1 * a.pedite_qtcpa), a.pedite_qtcpa)) AS CanEnt ", " SUM(CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE a.pedite_qtcpa END) AS CanEnt ")
                 sql2 = IIf(vg_tipbase = "1", " '" & Format(Trim(vaSpread1.text), "yyyymmdd") & "' ", " '" & Format(Trim(vaSpread1.text), "yyyymmdd") & "' ")
                 sql3 = IIf(vg_tipbase = "1", " ,(SELECT DISTINCT bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = b.pro_codigo) as bod_canmer ", " ,(SELECT DISTINCT bod_canmer FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_codpro = b.pro_codigo) as bod_canmer ")
                 sql4 = IIf(vg_tipbase = "1", " Format(a.solite_dtent, 'yyyymmdd') ", " convert(varchar(8),solite_dtent,112) ")
                 
                 If RS1.State = 1 Then RS1.Close
                 RS1.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
                 If vg_pais = "CO" Then
                    
                    RS1.Open "SELECT DISTINCT c.foc_codcat, a.cpopro_cdpro, a.cadfor_nrcgc, b.pro_codigo, b.pro_nombre, b.pro_ctrsto, b.pro_ctacon, b.pro_propon, " & _
                              "f.uni_nomcor, c.foc_nomsac, c.foc_faccon, c.foc_unisac, a.solite_dtent, a.pedite_vlpco, " & sql1 & " " & _
                              "" & sql3 & " " & _
                              "FROM b_ocsac a, b_productos b, b_formatocompras c, b_formatocomprassgp d, a_unidad f " & _
                              "Where c.foc_codsac   = d.fcs_codsac " & _
                              "AND   a.cpopro_cdpro = c.foc_codsac " & _
                              "AND   b.pro_codigo   = d.fcs_codsgp " & _
                              "AND   b.pro_coduni   = f.uni_codigo " & _
                              "AND   a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
                              "AND   " & sql4 & " = " & sql2 & " " & _
                              "AND  (a.cadfor_nrcgc = '" & vg_RDC & "' OR '" & vg_RDC & "' = '') " & _
                              "AND   a.pedite_flafo = " & tpentg & " " & _
                              "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven = 0) " & _
                               "GROUP BY c.foc_codcat, a.cpopro_cdpro, a.cadfor_nrcgc, b.pro_codigo, b.pro_nombre, b.pro_ctrsto, b.pro_ctacon, b.pro_propon, f.uni_nomcor, c.foc_nomsac, c.foc_faccon, c.foc_unisac, a.cpopro_cdpro, a.solite_dtent, a.pedite_vlpco ORDER BY c.foc_codcat, a.solite_dtent, a.cadfor_nrcgc, c.foc_nomsac", vg_db, adOpenStatic
                 
                 Else
                    
                    RS1.Open "SELECT DISTINCT c.foc_codcat, a.cpopro_cdpro, a.cadfor_nrcgc, b.pro_codigo, b.pro_nombre, b.pro_ctrsto, b.pro_ctacon, b.pro_propon, " & _
                             "f.uni_nomcor, c.foc_nomsac, c.foc_faccon, a.solite_dtent, a.pedite_vlpco, " & sql1 & " " & _
                             "" & sql3 & " " & _
                             "FROM b_ocsac a, b_productos b, b_formatocompras c, b_formatocomprassgp d, a_unidad f " & _
                             "Where c.foc_codsac   = d.fcs_codsac " & _
                             "AND   a.cpopro_cdpro = c.foc_codsac " & _
                             "AND   b.pro_codigo   = d.fcs_codsgp " & _
                             "AND   b.pro_coduni   = f.uni_codigo " & _
                             "AND   a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
                             "AND   " & sql4 & " = " & sql2 & " " & _
                             "AND  (a.cadfor_nrcgc = '" & vg_RDC & "' OR '" & vg_RDC & "' = '') " & _
                             "AND   a.pedite_flafo = " & tpentg & " " & _
                             "AND  (b.pro_fecven > " & Format(Date, "yyyymmdd") & " OR b.pro_fecven = 0) " & _
                             "GROUP BY c.foc_codcat, a.cpopro_cdpro, a.cadfor_nrcgc, b.pro_codigo, b.pro_nombre, b.pro_ctrsto, b.pro_ctacon, b.pro_propon, f.uni_nomcor, c.foc_nomsac, c.foc_faccon, a.cpopro_cdpro, a.solite_dtent, a.pedite_vlpco ORDER BY c.foc_codcat, a.solite_dtent, a.cadfor_nrcgc, c.foc_nomsac", vg_db, adOpenStatic
                  
                  End If
                  
                  '----Deshabilitar Botones
                  If tipopc = "docpro" Then
                     
                     form1.Frame3.Enabled = False
                     form1.Frame6.Enabled = False
                     form1.Frame5.Enabled = False
                     
                     For j = 1 To form1.vaSpread2.MaxRows
                         
                         form1.vaSpread2.Row = j
                         
                         form1.vaSpread2.Col = 1
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 2
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 3
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 4
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 5
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 6
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 7
                         form1.vaSpread2.Lock = True
                         
                         form1.vaSpread2.Col = 8
                         form1.vaSpread2.Lock = True
                         
                         If form1.vaSpread2.text = "1" Then form1.vaSpread2.Col = 5:  form1.vaSpread2.Lock = False
                     
                     Next j
                  
                  End If
                  
                  '------->  Asigna código cfc o fifo segun corresponda
                  '-------> MSP : 16/08/2004
                  '-------> Detalle de Documento
                  
                  If Not RS1.EOF Then
                     
                     sql1 = IIf(vg_tipbase = "1", " cdate('" & RS1!solite_dtent & "') ", " '" & Format(RS1!solite_dtent, "yyyymmdd") & "' ")
                  End If
                  
                  Do While Not RS1.EOF
                     
                     If RS2.State = 1 Then RS2.Close
                     RS2.CursorLocation = adUseClient
                     vg_db.CursorLocation = adUseClient
                     
                     RS2.Open "SELECT SUM(ocr_cancom) AS difer " & _
                              "FROM   b_ocsacrecibido " & _
                              "WHERE  ocr_rutpro = '" & rutpro & "' " & _
                              "AND    ocr_fecoc  = " & sql1 & " " & _
                              "AND    ocr_codprodsac = '" & Trim(RS1!cpopro_cdpro) & "' AND ocr_codprodsgp = '" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
                     If Not RS2.EOF And ((RS2!difer - RS1!canent) <> 0 Or IsNull(RS2!difer - RS1!canent)) And RS1!canent > 0 Then
                        
                        form1.vaSpread1.MaxRows = form1.vaSpread1.MaxRows + 1
                        
                        form1.vaSpread1.Col = 1
                        form1.vaSpread1.Row = form1.vaSpread1.MaxRows
                        form1.vaSpread1.Lock = True
                        
                        If vg_pais = "CO" Then
                           
                           form1.vaSpread1.Value = IIf(IsNull(RS1!cpopro_cdpro), "", Trim(RS1!cpopro_cdpro))
                           form1.vaSpread1.Col = 2
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = IIf(IsNull(RS1!foc_nomsac), "No existe producto sac, comuniquese departamento compras", Trim(RS1!foc_nomsac))
                           
                           form1.vaSpread1.Col = 3
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = IIf(IsNull(RS1!foc_unisac), "", Trim(RS1!foc_unisac))
                        
                        Else
                           
                           form1.vaSpread1.Value = IIf(IsNull(RS1!pro_codigo), "", Trim(RS1!pro_codigo))
                           
                           form1.vaSpread1.Col = 2
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = IIf(IsNull(RS1!pro_nombre), "No existe producto sac, comuniquese departamento compras", Trim(RS1!pro_nombre))
                           
                           form1.vaSpread1.Col = 3
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = IIf(IsNull(RS1!uni_nomcor), "", Trim(RS1!uni_nomcor))
                        
                        End If
                        
                        form1.vaSpread1.Col = 4: form1.vaSpread1.Lock = False
                        
                        If vg_pais = "CO" And IsNull(RS1!foc_faccon) Or RS1!foc_faccon < 0 Then
                           
                           form1.vaSpread1.Value = 0
                        
                        Else
                           
                           form1.vaSpread1.Value = IIf(IsNull(RS2!difer), IIf(IsNull(RS1!foc_faccon) Or RS1!foc_faccon < 0, 0, RS1!canent), IIf((RS1!canent - RS2!difer) < 0, RS1!canent, (RS1!canent - RS2!difer)))
                        
                        End If
                        
                        form1.vaSpread1.Col = 5
                        form1.vaSpread1.Lock = False
                        form1.vaSpread1.Value = RS1!pedite_vlpco
                        
                        form1.vaSpread1.Col = 6
                        form1.vaSpread1.Lock = IIf(tipopc <> "docpro", False, True): form1.vaSpread1.Value = IIf(tipopc <> "docpro", Round((RS1!pedite_vlpco * RS1!canent), 0), 0)
                        
                        form1.vaSpread1.Col = 7
                        form1.vaSpread1.Lock = IIf(tipopc <> "docpro", False, True): form1.vaSpread1.Value = IIf(tipopc <> "docpro", RS1!canent, 0)
                        
                        form1.vaSpread1.Col = 8
                        form1.vaSpread1.Lock = False:
                        
                        If vg_pais = "CO" And IsNull(RS1!foc_faccon) Or RS1!foc_faccon < 0 Then
                           
                           form1.vaSpread1.Value = 0
                           form1.vaSpread1.Col = 9
                           form1.vaSpread1.Lock = False
                           form1.vaSpread1.Value = 0
                        
                        Else
                           
                           form1.vaSpread1.Value = IIf(tipopc <> "docpro", "N", Round((RS1!pedite_vlpco * RS1!canent), 0))
                           
                           form1.vaSpread1.Col = 9
                           form1.vaSpread1.Lock = False
                           form1.vaSpread1.Value = 0
                           
                           form1.vaSpread1.Value = IIf(tipopc <> "docpro", IIf(IsNull(RS1!bod_canmer) Or RS1!bod_canmer = 0, 0, RS1!bod_canmer), RS1!canent)
                        
                        End If
                        
                        form1.vaSpread1.Col = 10
                        form1.vaSpread1.Lock = True
                        form1.vaSpread1.Value = IIf(tipopc <> "docpro", IIf(IsNull(RS1!pro_ctrsto), "N", IIf(RS1!pro_ctrsto = 1, "S", "N")), RS1!pedite_vlpco)
                        
                        If tipopc = "docpro" Then
                           
                           form1.vaSpread1.Col = 18
                           form1.vaSpread1.Value = RS1!pedite_vlpco
                           
                           form1.vaSpread1.Col = 20
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = ""
                           
                           form1.vaSpread1.Col = 21
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = Format(RS1!canent, fg_Pict(9, vg_DCa))
                           
                           form1.vaSpread1.Col = 22
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = Format(RS1!pedite_vlpco, fg_Pict(9, 2))
                           
                           form1.vaSpread1.Col = 23
                           form1.vaSpread1.Lock = True
                           form1.vaSpread1.Value = Format(RS1!solite_dtent, "dd/mm/yyyy")
                           
                           If vg_pais = "CO" Then
                              
                              form1.vaSpread1.Col = 24
                              form1.vaSpread1.Lock = True
                              form1.vaSpread1.Value = IIf(IsNull(RS1!pro_codigo), "", Trim(RS1!pro_codigo))
                              
                              form1.vaSpread1.Col = 25
                              form1.vaSpread1.Lock = True
                              form1.vaSpread1.Value = IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre))
                              
                              form1.vaSpread1.Col = 29
                              form1.vaSpread1.Lock = True
                              form1.vaSpread1.Value = IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
                           
                           Else
                              
                              form1.vaSpread1.Col = 24
                              form1.vaSpread1.Lock = True
                              form1.vaSpread1.Value = IIf(IsNull(RS1!cpopro_cdpro), "", Trim(RS1!cpopro_cdpro))
                              
                              form1.vaSpread1.Col = 25
                              form1.vaSpread1.Lock = True
                              form1.vaSpread1.Value = IIf(IsNull(RS1!foc_nomsac), "", Trim(RS1!foc_nomsac))
                              
                              form1.vaSpread1.Col = 29
                              form1.vaSpread1.Lock = True
                              form1.vaSpread1.Value = IIf(IsNull(RS1!foc_faccon), 0, RS1!foc_faccon)
                           
                           End If
                        
                        End If
                        
                        form1.vaSpread1.Col = 11
                        form1.vaSpread1.Value = IIf(tipopc <> "docpro", IIf(IsNull(RS1!pro_propon), 0, RS1!pro_propon), IIf(IsNull(RS1!pro_nombre), "No existe producto sac, comuniquese departamento compras", Trim(RS1!pro_nombre)))
                        
                        form1.vaSpread1.Col = 12
                        form1.vaSpread1.Value = IIf(tipopc <> "docpro", "N", IIf(IsNull(RS1!pro_ctacon), "", RS1!pro_ctacon))
                        
                        form1.vaSpread1.Col = 13
                        form1.vaSpread1.Value = IIf(tipopc <> "docpro", Format(RS1!canent, fg_Pict(9, vg_DCa)), IIf(IsNull(RS1!pro_ctrsto), "", IIf(RS1!pro_ctrsto = 1, "S", "N")))
                        
                        form1.vaSpread1.Col = 14
                        form1.vaSpread1.text = IIf(tipopc <> "docpro", Format(RS1!pedite_vlpco, fg_Pict(9, 2)), "")
                        
                        form1.vaSpread1.Col = 15
                        form1.vaSpread1.text = IIf(tipopc <> "docpro", Format(RS1!solite_dtent, "dd/mm/yyyy"), 0)
                        
                        form1.vaSpread1.Col = 16
                        form1.vaSpread1.text = IIf(tipopc <> "docpro", IIf(IsNull(RS1!cpopro_cdpro), "", Trim(RS1!cpopro_cdpro)), 0)
                        
                        form1.vaSpread1.Col = 17
                        form1.vaSpread1.Lock = True
                        form1.vaSpread1.Value = IIf(tipopc <> "docpro", IIf(IsNull(RS1!foc_nomsac), "", Trim(RS1!foc_nomsac)), RS1!pedite_vlpco)
                        
                        If tipopc = "docpro" Then
                           
                           Revisa RS1!pro_codigo, form1.vaSpread1.Row
                        
                        End If
                     
                     End If
                     RS2.Close
                     Set RS2 = Nothing
                     
                     RS1.MoveNext
                     
                 Loop
                 RS1.Close
                 Set RS1 = Nothing
                 
                 If form1.vaSpread1.MaxRows > 0 Then
                    
                    Dim codsgp As String
                    Dim codsac As String
                    
                    form1.vaSpread1.Row = 1
                    form1.vaSpread1.Col = 1
                    
                    codsgp = Trim(form1.vaSpread1.text)
                    
                    form1.vaSpread1.Row = 1
                    form1.vaSpread1.Col = IIf(tipopc <> "docpro", 16, 24)
                    
                    form1.Text2(0).text = Trim(form1.vaSpread1.text)
                    codsac = Trim(form1.vaSpread1.text)
                    form1.vaSpread1.Row = 1
                    form1.vaSpread1.Col = IIf(tipopc <> "docpro", 17, 25)
                    
                    form1.Text2(1).text = Trim(form1.vaSpread1.text)
                    form1.Image1(5).Visible = IIf(ValidarProductosSgpSac(Trim(codsac), Trim(codsgp)) And vg_pais = "CL", True, False)
                    
                    If tipopc = "docpro" Then
                       
                       form1.vaSpread1.Row = 1
                       form1.vaSpread1.Col = 29
                       form1.Text2(2).text = form1.vaSpread1.text
                    
                    End If
                 
                 End If
                 
                 '-------> Fin detalle de Documento
                 form1.vaSpread1.Visible = True
               
               End If
            
            End If
        
        Next i
        
End Select
Me.Hide
Unload Me

Exit Sub
Error_Carga:
MsgBox Err & ": " & Err.Number & " " & Err.Description, vbCritical, MsgTitulo
Resume Next

End Sub

Private Sub Revisa(codpro As String, Row As Long)

On Error GoTo Man_Error

Dim v_rut As String, estrut As Boolean
Dim regimp As String, autoret As String, cuohor As String
Dim codmun As Long
Dim RS2 As New ADODB.Recordset
Dim RS6 As New ADODB.Recordset

form1.vaSpread1.Col = 14
form1.vaSpread1.text = ""
v_rut = fg_DespintaRut(form1.fpText(0).text)
estrut = False
estrut = ValidarRetProveedor(v_rut)

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If estrut Then
   
   '-------> sacar datos de parametros
   RS2.Open "SELECT TOP 1 prv_regimp, prv_autret, prv_cuohor, prv_codmun " & _
            "FROM b_proveedor " & _
            "WHERE prv_codigo = '" & v_rut & "' AND prv_regimp IS NOT NULL", vg_db, adOpenForwardOnly
   
   If Not RS2.EOF Then
      
      regimp = IIf(IsNull(RS2!prv_regimp), "0", RS2!prv_regimp)
      autret = IIf(IsNull(RS2!prv_autret), "S", RS2!prv_autret)
      cuohor = IIf(IsNull(RS2!prv_cuohor), "N", RS2!prv_cuohor)
      codmun = IIf(IsNull(RS2!prv_codmun), 0, RS2!prv_codmun)
   
   End If
   RS2.Close
   Set RS2 = Nothing

End If

If RS6.State = 1 Then RS6.Close
RS6.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS6.Open "SELECT a.*, b.*, c.pro_codref, c.pro_codrei, c.pro_cuohor " & _
         "FROM   b_productosimp a, a_impuesto b, b_productos c " & _
         "WHERE  a.ipr_codimp = b.imp_codigo " & _
         "AND    a.ipr_codpro = c.pro_codigo " & _
         "AND    a.ipr_codpro = '" & codpro & "'", vg_db, adOpenForwardOnly
Do While Not RS6.EOF
   
   form1.vaSpread1.Col = form1.vaSpread1.ActiveCol
   form1.vaSpread1.Row = Row
         
   form1.vaSpread1.Col = 14
   form1.vaSpread1.text = form1.vaSpread1.text & Trim(Str(RS6!ipr_codimp)) & "&"
   
   form1.vaSpread1.Col = 14
   form1.vaSpread1.text = form1.vaSpread1.text & Trim(Str(IIf(IsNull(RS6!imp_pctimp), 0, RS6!imp_pctimp))) & "&"
   
   form1.vaSpread1.Col = 14
   form1.vaSpread1.text = form1.vaSpread1.text & Trim(Str(IIf(IsNull(RS6!imp_inccos), 0, RS6!imp_inccos))) & ";"
   
   '-------> Validar si aplica cuota hortofruticola
   If RS6!pro_cuohor = "S" And cuohor = "S" Then
      
      If CStr(RS6!ipr_codimp) = GetParametro("parrethorf") Then
         
         '-------> Colocar impuesto retención hortofruticola
         If form1.vaSpread2.SearchCol(1, 0, form1.vaSpread2.MaxRows, CStr(RS6!ipr_codimp), SearchFlagsNone) <> -1 Then
            
            form1.vaSpread2.Row = form1.vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS6!ipr_codimp), SearchFlagsNone)
            
            form1.vaSpread2.Col = 4
            form1.vaSpread2.text = IIf(IsNull(GetParametro("parhorfru")), Format(0, fg_Pict(3, 2)), Format(GetParametro("parhorfru"), fg_Pict(3, 2))) & " %"
            
            form1.vaSpread2.Col = 7
            form1.vaSpread2.text = IIf(IsNull(GetParametro("parhorfru")), 0, GetParametro("parhorfru"))
         
         End If
         
         form1.vaSpread1.Col = 14
         form1.vaSpread1.text = form1.vaSpread1.text & Trim(Str(RS6!ipr_codimp)) & "&"
         
         form1.vaSpread1.Col = 14
         form1.vaSpread1.text = form1.vaSpread1.text & Trim(Str(IIf(IsNull(GetParametro("parhorfru")), 0, GetParametro("parhorfru")))) & "&"
         
         form1.vaSpread1.Col = 14
         form1.vaSpread1.text = form1.vaSpread1.text & Trim(Str(IIf(IsNull(RS6!imp_inccos), 0, RS6!imp_inccos))) & ";"
      
      End If
   
   End If
   
   RS6.MoveNext

Loop
RS6.Close
Set RS6 = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarRetProveedor(rut As String) As Boolean

On Error GoTo Man_Error

Dim RS2 As New ADODB.Recordset
ValidarRetProveedor = False
'-------> Validar si proveedor posee impuestos adicionales
If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS2.Open "SELECT prv_regimp, prv_autret, prv_cuohor, prv_codmun FROM b_proveedor WHERE prv_codigo = '" & rut & "' AND prv_regimp IS NOT NULL", vg_db, adOpenForwardOnly
If Not RS2.EOF Then
   
   regimp = RS2!prv_regimp
   autret = RS2!prv_autret
   cuohor = RS2!prv_cuohor
   codmun = RS2!prv_codmun
   ValidarRetProveedor = True

End If
RS2.Close
Set RS2 = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function
