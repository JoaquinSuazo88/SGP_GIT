VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form B_HistPm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico Planificación Minutas"
   ClientHeight    =   3930
   ClientLeft      =   2250
   ClientTop       =   1905
   ClientWidth     =   7965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   7965
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   555
      Index           =   2
      Left            =   4320
      TabIndex        =   6
      Top             =   3240
      Width           =   1785
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   150
         Width           =   1560
      End
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Index           =   1
      Left            =   1920
      TabIndex        =   4
      Top             =   3240
      Width           =   2265
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   150
         Width           =   2040
      End
   End
   Begin VB.Frame Frame3 
      Height          =   555
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   1785
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   150
         Width           =   1560
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      _Version        =   393216
      _ExtentX        =   13044
      _ExtentY        =   5530
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
      MaxRows         =   30
      OperationMode   =   3
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_HistPm.frx":0000
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   3930
      Left            =   7425
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   6932
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_HistPm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS        As New ADODB.Recordset
Dim RS2       As New ADODB.Recordset
Dim op        As String
Dim Msgtitulo As String

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0
fg_descarga

End Sub

Sub LlenarHistPlan(tfor As String, cencos As String, tipmin As String, opcion As String)

Dim ancholcol As Double
Dim Titulo    As String, sql1 As String, sql2 As String
Dim i         As Long, codreg As Long, codser As Long, codbod As Long
Dim ValLcntH$
Dim tiptra    As Integer

Me.Caption = tfor
Msgtitulo = tfor
op = opcion

With vaSpread1
    
    Text1(1).Visible = False
    Text1(2).Visible = False
    Text1(3).Visible = False
    
    If opcion = "1" Then
       
       Me.Height = 3615
       Me.Width = 8025
       .Height = 3135
       fg_centra Me
       .MaxRows = 0: .MaxCols = 6: .Row = 0
       
       For i = 1 To .MaxCols
           
           If i = 1 Or i = 3 Then
              
              anchocol = 7.38
              If i = 1 Then Titulo = "C.Regimen"
              If i = 3 Then Titulo = "C.Servicio"
           
           End If
           
           If i = 2 Or i = 4 Then anchocol = 13.9: Titulo = "Descripción"
           If i = 5 Then anchocol = 8: Titulo = "Fecha"
           If i = 6 Then anchocol = 8: Titulo = "Estado"
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       sql1 = IIf(vg_tipbase = "1", "mid(d.min_fecmin,1,6) as fecha", "substring(convert(varchar(8),d.min_fecmin),1,6) as fecha")
       sql2 = IIf(vg_tipbase = "1", "mid(d.min_fecmin,1,6) DESC", "substring(convert(varchar(8),d.min_fecmin),1,6) DESC")
       RS.Open "SELECT DISTINCT d.min_cencos, a.cli_nombre, " & _
               "d.min_codreg, d.min_indblo, b.reg_nombre, " & _
               "d.min_codser, c.ser_nombre, " & sql1 & ", e.mid_tipmin " & _
               "FROM  b_clientes a, a_regimen b, a_servicio c, b_minuta d, b_minutadet e " & _
               "WHERE d.min_codigo = e.mid_codigo " & _
               "AND   d.min_cencos = a.cli_codigo " & _
               "AND   d.min_codreg = b.reg_codigo " & _
               "AND   d.min_codser = c.ser_codigo " & _
               "AND   d.min_cencos = '" & cencos & "' " & _
               "AND   e.mid_tipmin = '" & tipmin & "' " & _
               "AND   a.cli_tipo = 0 " & _
               "ORDER BY " & sql2 & ", d.min_codser", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    ElseIf opcion = "2" Then
       
       Me.Height = 3615
       Me.Width = 6200
       .Height = 3135
       .Width = 5550
       fg_centra Me
       .MaxRows = 0: .MaxCols = 4: .Row = 0
       
       For i = 1 To .MaxCols
           
           If i = 1 Then anchocol = 7.38: Titulo = "C.Regimen"
           If i = 2 Then anchocol = 20: Titulo = "Descripción"
           If i = 3 Then anchocol = 8: Titulo = "Fecha"
           If i = 4 Then anchocol = 8: Titulo = "Estado"
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       sql1 = IIf(vg_tipbase = "1", " mid(b.min_fecmin,1,6) ", " substring(convert(varchar(8),b.min_fecmin),1,6) ")
       RS.Open "SELECT DISTINCT b.min_codreg, a.reg_nombre, " & _
               "" & sql1 & " as fecha, c.mid_tipmin, b.min_indblo " & _
               "FROM a_regimen a, b_minuta b, b_minutadet c " & _
               "WHERE b.min_codigo = c.mid_codigo " & _
               "AND   b.min_codreg = a.reg_codigo " & _
               "AND   b.min_cencos = '" & cencos & "' " & _
               "AND   c.mid_tipmin = '" & tipmin & "' " & _
               "ORDER BY " & sql1 & " DESC, a.reg_nombre", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    ElseIf opcion = "3" Then
       
       Me.Height = 3615
       Me.Width = 8025
       .Height = 3135
       fg_centra Me
       .MaxRows = 0: .MaxCols = 6: .Row = 0
       
       For i = 1 To .MaxCols
           
           If i = 1 Or i = 3 Then
              
              anchocol = 7.38
              If i = 1 Then Titulo = "C.Regimen"
              If i = 3 Then Titulo = "C.Servicio"
           
           End If
           
           If i = 2 Or i = 4 Then anchocol = 13.9: Titulo = "Descripción"
           If i = 5 Then anchocol = 8: Titulo = "Fecha"
           If i = 6 Then anchocol = 8: Titulo = "Día"
           
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       ValLcntH = ""
       
       For i = 1 To Len(tipmin)
           
           If Asc(Mid(tipmin, i, 1)) = 124 Then
              
              If ValLcntH <> "" And codreg = 0 Then codreg = Val(ValLcntH): ValLcntH = ""
              If ValLcntH <> "" And codser = 0 Then codser = Val(ValLcntH): ValLcntH = "": Exit For
           
           Else
              
              ValLcntH = ValLcntH + Mid(tipmin, i, 1)
           
           End If
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open "SELECT DISTINCT a.mif_fecval, a.mif_dianro, a.mif_codreg, " & _
               "b.reg_nombre, a.mif_codser, c.ser_nombre " & _
               "FROM  b_minutafija a, a_regimen b, a_servicio c " & _
               "WHERE a.mif_codreg = b.reg_codigo " & _
               "AND   a.mif_codser = c.ser_codigo " & _
               "AND   a.mif_cencos = '" & cencos & "' " & _
               "AND   a.mif_codreg = " & codreg & " " & _
               "AND   a.mif_codser = " & codser & "", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    ElseIf opcion = "4" Then
       
       Text1(1).Visible = True
       Text1(2).Visible = True
       Text1(3).Visible = True
       
       If Trim(tfor) <> "" Then
          
          Me.Caption = Mid(tfor, 1, Len(tfor) - 1)
          Msgtitulo = Mid(tfor, 1, Len(tfor) - 1)
          
       End If
       
       Me.Height = 4305
       Me.Width = 7255
       .Height = 3135
       .Width = 6550
       fg_centra Me
       .MaxRows = 0: .MaxCols = 4: .Row = 0
       
       For i = 1 To .MaxCols
           
           If i = 1 Then anchocol = 15.38: Titulo = "Tipo"
           If i = 2 Then anchocol = 20: Titulo = "Folio"
           If i = 3 Then anchocol = 8: Titulo = "Fecha"
           If i = 4 Then anchocol = 8: Titulo = "Usuario"
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

       If IsNumeric(Right(tfor, 1)) And tipmin <> "IN ('C','P')" Then
          
          tiptra = Right(tfor, 1)
          Set RS = vg_db.Execute("sgp_Sel_HistoricoTraspaso '" & cencos & "', " & tiptra & "")
       
       Else
          If tipmin = "" Then
          
             tipmin = "IN ('C','P')"
          
          End If
          
          tipmin = Replace(Mid(tipmin, 1, Len(tipmin)), "'", """")
          Set RS = vg_db.Execute("sgp_Sel_HistoricoCfcFofi '" & cencos & "', '" & tipmin & "'")
          'RS.Open "SELECT inf_tipo, inf_numero, inf_feccie, inf_usuario FROM a_infcfcfofi WHERE inf_cencos = '" & cencos & "' AND inf_tipo " & tipmin & " ORDER BY inf_tipo, inf_numero desc", vg_db, adOpenStatic
        
       End If

       If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    ElseIf opcion = "5" Then
       
       Me.Height = 3615
       Me.Width = 7000
       .Height = 3135
       .Width = 6350
       fg_centra Me
       .MaxRows = 0
       .MaxCols = 5
       .Row = 0
       
       For i = 1 To .MaxCols
           
           If i = 1 Or i = 3 Then
              
              anchocol = 7.38
              If i = 1 Then Titulo = "C.Regimen"
              If i = 3 Then Titulo = "C.Servicio"
           
           End If
           
           If i = 2 Or i = 4 Then anchocol = 13.9: Titulo = "Descripción"
           If i = 5 Then anchocol = 8: Titulo = "Fecha"
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       ValLcntH = ""
       
       For i = 1 To Len(tipmin)
           
           If Asc(Mid(tipmin, i, 1)) = 124 Then
              
              If ValLcntH <> "" And codreg = 0 Then codreg = Val(ValLcntH): ValLcntH = ""
              If ValLcntH <> "" And codser = 0 Then codser = Val(ValLcntH): ValLcntH = "": Exit For
           
           Else
              
              ValLcntH = ValLcntH + Mid(tipmin, i, 1)
           
           End If
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open "SELECT DISTINCT c.prv_codreg, a.reg_nombre, " & _
               "c.prv_codser, b.ser_nombre, c.prv_fecvig " & _
               "FROM  a_regimen a, a_servicio b, b_preciovta c " & _
               "WHERE c.prv_codreg = a.reg_codigo " & _
               "AND   c.prv_codser = b.ser_codigo " & _
               "AND   c.prv_cencos = '" & cencos & "' " & _
               "AND   c.prv_codreg = " & codreg & " " & _
               "AND   c.prv_codser = " & codser & " ORDER BY c.prv_fecvig", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    ElseIf opcion = "6" Then
       
       Me.Height = 3635
       Me.Width = 5240
       .Height = 3135
       .Width = 4560
       fg_centra Me
       .MaxRows = 0: .MaxCols = 3: .Row = 0
       
       For i = 1 To .MaxCols
           
           If i = 1 Then anchocol = 7.38: Titulo = "C.Bodega"
           If i = 2 Then anchocol = 20: Titulo = "Descripción"
           If i = 3 Then anchocol = 8: Titulo = "Fecha Invetario"
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       codbod = IIf(Trim(cencos) = "", 0, cencos)
       RS.Open "SELECT DISTINCT b.bod_codigo, b.bod_nombre, a.tin_fectom " & _
               "FROM b_tomainv a, a_bodega b " & _
               "WHERE a.tin_codbod = b.bod_codigo " & _
               "AND   a.tin_codbod = " & vg_codbod & " ORDER BY tin_fectom DESC", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    ElseIf opcion = "7" Then
       
       Me.Caption = "Historico Salida Producción"
       Me.Height = 3615
       Me.Width = 5200
       .Height = 3135
       .Width = 4550
       fg_centra Me
       .MaxRows = 0: .MaxCols = 3: .Row = 0
       For i = 1 To .MaxCols
           
           If i = 1 Then anchocol = 7.38: Titulo = "C.Regimen"
           If i = 2 Then anchocol = 20: Titulo = "Descripción"
           If i = 3 Then anchocol = 8: Titulo = "Fecha"
    '       If i = 4 Then anchocol = 8: titulo = "Estado"
           .Col = i
           .ColWidth(i) = anchocol
           .text = Titulo
           .ColHidden = False
       
       Next i
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       sql1 = IIf(vg_tipbase = "1", "format(a.tov_fecpro,'yyyymm') ", "substring(CONVERT(varchar(10), a.tov_fecpro,112),1,6) ")
       RS.Open "SELECT DISTINCT a.tov_codreg AS min_codreg, b.reg_nombre, " & sql1 & " AS fecha, 2 AS mid_tipmin, 1 AS min_indblo " & _
               "FROM b_totventas a, a_regimen b " & _
               "WHERE a.tov_codreg = b.reg_codigo " & _
               "AND   a.tov_rutcli = '" & cencos & "' " & _
               "AND   a.tov_tipdoc = 'SP' AND a.tov_codbod = " & vg_codbod & " " & _
               "ORDER BY " & sql1 & " DESC, b.reg_nombre", vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub
    
    End If
    
    Do While Not RS.EOF
       
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
          
       If opcion = "1" Then
          
          .Col = 1
          .TypeHAlign = TypeHAlignRight
          .text = RS!min_codreg
          
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!reg_nombre)
          
          .Col = 3
          .TypeHAlign = TypeHAlignRight
          .text = RS!min_codser
          
          .Col = 4
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!ser_nombre)
         
          .Col = 5
          .TypeHAlign = TypeHAlignCenter
          .text = Mid(RS!Fecha, 5, 2) & "/" & Mid(RS!Fecha, 1, 4)
          
          .Col = 6
          .TypeHAlign = TypeHAlignCenter
          
          If (RS!min_indblo = 1 Or RS!min_indblo = 2) And RS!mid_tipmin = "1" Then
             
             .text = IIf(RS!min_indblo = 1, "Cerrado", "Bloqueado")
          
          ElseIf (RS!min_indblo = 0 Or IsNull(RS!min_indblo) Or RS!min_indblo = 11) And RS!mid_tipmin = "1" Then
             
             .text = "Abierto"
          
          ElseIf RS!Fecha < Val(Format(Date, "yyyymm")) And RS!mid_tipmin = "2" Then
             
             .text = "Cerrado"
          
          ElseIf RS!Fecha >= Val(Format(Date, "yyyymm")) And RS!mid_tipmin = "2" Then
             
             .text = "Abierto"
          
          End If
       
       ElseIf opcion = "2" Or opcion = "7" Then
          
          .Col = 1
          .TypeHAlign = TypeHAlignRight
          .text = RS!min_codreg
          
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!reg_nombre)
       
          .Col = 3
          .TypeHAlign = TypeHAlignCenter
          .text = Mid(RS!Fecha, 5, 2) & "/" & Mid(RS!Fecha, 1, 4)
      
          If opcion = "2" Then
             
             .Col = 4
             .TypeHAlign = TypeHAlignCenter
             
             If (RS!min_indblo = 1 Or IsNull(RS!min_indblo)) And RS!mid_tipmin = "1" Then
                
                .text = "Cerrado"
             
             ElseIf RS!min_indblo = 0 And RS!mid_tipmin = "1" Then
                
                .text = "Abierto"
             
             ElseIf RS!Fecha < Val(Format(Date, "yyyymm")) And RS!mid_tipmin = "2" Then
                
                .text = "Cerrado"
             
             ElseIf RS!Fecha >= Val(Format(Date, "yyyymm")) And RS!mid_tipmin = "2" Then
                
                .text = "Abierto"
             
             End If
          
          End If
       
       ElseIf opcion = "3" Then
          
          .Col = 1
          .TypeHAlign = TypeHAlignRight
          .text = RS!mif_codreg
          
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!reg_nombre)
          
          .Col = 3
          .TypeHAlign = TypeHAlignRight
          .text = RS!mif_codser
          
          .Col = 4
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!ser_nombre)
       
          .Col = 5
          .TypeHAlign = TypeHAlignCenter
          .text = Mid(RS!mif_fecval, 7, 2) & "/" & Mid(RS!mif_fecval, 5, 2) & "/" & Mid(RS!mif_fecval, 1, 4)
       
          .Col = 6
          .TypeHAlign = TypeHAlignCenter
          .text = fg_NomDia(RS!mif_dianro)
       
       ElseIf opcion = "4" Then
          
          .Col = 1
          .TypeHAlign = TypeHAlignCenter
          
          If RS!Inf_Tipo = "C" Then
             
             .text = "CFC"
          
          ElseIf RS!Inf_Tipo = "P" Then
             
             .text = "CFC - Portal Electronico"
          
          ElseIf RS!Inf_Tipo = "G" Then
             
             .text = "CTC - Portal Electronico"
          
          ElseIf RS!Inf_Tipo = "T" Then
             
             .text = "CTC"
          
          ElseIf RS!Inf_Tipo = "F" Then
             
             .text = "CFF"
          
          End If
          
          .Col = 2
          .TypeHAlign = TypeHAlignCenter
          .text = RS!inf_numero
          
          .Col = 3
          .TypeHAlign = TypeHAlignCenter
          .text = IIf(RS!inf_feccie = "0" Or IsNull(RS!inf_feccie), "", Mid(RS!inf_feccie, 7, 2) & "/" & Mid(RS!inf_feccie, 5, 2) & "/" & Mid(RS!inf_feccie, 1, 4))
       
          .Col = 4
          .TypeHAlign = TypeHAlignCenter
          .text = IIf(IsNull(RS!inf_usuario) Or Trim(RS!inf_usuario) = "", "No enviado", Trim(RS!inf_usuario))
       
       ElseIf opcion = "5" Then
          
          .Col = 1
          .TypeHAlign = TypeHAlignRight
          .text = RS!prv_codreg
          
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!reg_nombre)
          
          .Col = 3
          .TypeHAlign = TypeHAlignRight
          .text = RS!prv_codser
          
          .Col = 4
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!ser_nombre)
       
          .Col = 5
          .TypeHAlign = TypeHAlignCenter
          .text = Mid(RS!prv_fecvig, 7, 2) & "/" & Mid(RS!prv_fecvig, 5, 2) & "/" & Mid(RS!prv_fecvig, 1, 4)
       
       ElseIf opcion = "6" Then
          
          .Col = 1
          .TypeHAlign = TypeHAlignRight
          .text = RS!bod_codigo
          
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!bod_nombre)
          
          .Col = 3
          .TypeHAlign = TypeHAlignCenter
          .text = Mid(RS!tin_fectom, 7, 2) & "/" & Mid(RS!tin_fectom, 5, 2) & "/" & Mid(RS!tin_fectom, 1, 4)
       
       End If
       
       RS.MoveNext
    
    Loop
    RS.Close: Set RS = Nothing
    vg_codigo = ""

End With

End Sub

Private Sub Text1_Change(Index As Integer)

On Error GoTo error

Dim i As Long
Dim IndActivo As Integer

Select Case Index

Case 1

    Text1(2).text = ""
    Text1(3).text = ""

Case 2

    Text1(1).text = ""
    Text1(3).text = ""

Case 3

    Text1(1).text = ""
    Text1(2).text = ""

End Select

vaSpread1.Visible = False
If Trim(Text1(Index).text) <> "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           IndActivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = Index
           
           If IndActivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           Else
              
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index, 1
End If
'    vaSpread1_Click Index, 0
vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    
vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
If Trim(Text1(Index).text) = "" Then
       
   For i = 1 To vaSpread1.MaxRows
           
       vaSpread1.Row = i
       If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
   
   Next
   vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
   vaSpread1.SetActiveCell Index, 1

End If
vaSpread1.Visible = True

Exit Sub
error:
    MsgBox Err.Description, vbCritical

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Index

Case 1
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    MoverDatos

Case 3
    
    vg_codigo = ""
    Me.Hide
    Unload Me

End Select

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

MoverDatos

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode

Case 27
    
    Cerrar

End Select

End Sub

Private Sub MoverDatos()

With vaSpread1
    
    If .MaxRows < 1 Then Exit Sub
    .Row = .ActiveRow
    
    If op = "1" Then
       
       vg_codigo = "C"
       .Col = 1: vg_codregimen = Val(.text)
       .Col = 3: vg_codservicio = Val(.text)
       .Col = 5: vg_fecha = .text
    
    ElseIf op = "2" Or op = "7" Then
       
       .Col = 1: vg_codigo = .text
       .Col = 3: vg_auxfecha = .text
    
    ElseIf op = "3" Then
       
       vg_codigo = ""
       .Col = 5: vg_codigo = .text
       .Col = 6: vg_auxfecha = fg_NumDia(.text)
    
    ElseIf op = "4" Then
       
       vg_codigo = "": .Col = 2: vg_codigo = .text
       .Col = 1
       
       If Trim(.text) = "CFC" Then
          
          vg_codigo4 = "C"
       
       ElseIf Trim(.text) = "CFC - Portal Electronico" Then
          
          vg_codigo4 = "P"
       
       ElseIf Trim(.text) = "CTC" Then
          
          vg_codigo4 = "T"
       
       ElseIf Trim(.text) = "CTC - Portal Electronico" Then
          
          vg_codigo4 = "G"
       
       End If
    
    ElseIf op = "5" Then
       
       vg_codigo = ""
       .Col = 5: vg_codigo = .text
    
    ElseIf op = "6" Then
       
       vg_codigo = ""
       .Col = 1: vg_codigo = .text
       .Col = 3: vg_fecha = .text
    
    End If

End With

Cerrar

End Sub

Sub Cerrar()

Me.Hide
Unload Me

End Sub
