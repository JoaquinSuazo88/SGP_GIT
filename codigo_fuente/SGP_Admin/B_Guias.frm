VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form B_Guias 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3165
   ClientLeft      =   7665
   ClientTop       =   2595
   ClientWidth     =   2700
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   2595
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2145
         Left            =   315
         TabIndex        =   2
         Top             =   300
         Width           =   1920
         _Version        =   393216
         _ExtentX        =   3387
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
         MaxCols         =   2
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
Dim RS1 As New ADODB.Recordset, RS3 As New ADODB.Recordset
Dim Impuestos() As Variant

Public Sub Cargar_DoctoGrilla(TipoDoc As String, TituloForm As String, rut As String)
MsgTitulo = TituloForm
B_Guias.Caption = TituloForm
vaSpread1.MaxRows = 0
If TipoDoc = "SN" Then
   RS1.Open "select  toc_numdoc from b_totcompras where toc_docaso in (select trim(toc_numdoc)from b_totcompras where toc_tipdoc='FA' and (toc_docaso='' or toc_docaso is null)) and  toc_rutpro='" & rut & "' and toc_tipdoc='" & Trim(TipoDoc) & "' order by toc_numdoc", vg_db, adOpenStatic
   If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
   Do While Not RS1.EOF
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
         If InStr(vg_Guias, RS1!toc_numdoc) <> 0 Then vaSpread1.Col = 1: vaSpread1.Value = 0
         vaSpread1.Col = 2: vaSpread1.Value = RS1!toc_numdoc
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
Else
   RS1.Open "select toc_numdoc from b_totcompras where toc_rutpro='" & rut & "' and (toc_docaso='' or toc_docaso is null)   and toc_tipdoc='" & Trim(TipoDoc) & "' order by toc_numdoc", vg_db, adOpenStatic
   Do While Not RS1.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
      If InStr(vg_Guias, RS1!toc_numdoc) <> 0 Then vaSpread1.Col = 1: vaSpread1.Value = 0
      vaSpread1.Col = 2: vaSpread1.Value = RS1!toc_numdoc
      RS1.MoveNext
   Loop
   RS1.Close: Set RS1 = Nothing
End If
End Sub
Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error
'---Dimensiona el formulario a su estado de diseño
Me.Width = 2820
Me.Height = 3675
fg_centra Me
fg_carga ""
MsgTitulo = "Guías de Despacho"
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar "
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Exit Sub
Man_Error:
MsgBox Err & ": " & Err.Number & "-" & Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Error_Carga
Dim i As Long, TotE As Long, TotI As Long, TotImp, totnet, TotGen As Long
Dim RS1 As New ADODB.Recordset, RS As New ADODB.Recordset
Dim v_valdesc  As Double, v_valtot As Double
Select Case Button.Index
Case 1
    'En caso de que el documento ya exista ***
    M_DocPro.vaSpread1.Row = -1: M_DocPro.vaSpread1.Col = 1
    If M_DocPro.vaSpread1.Lock = True Then
        If MsgBox("Desea modificar selección...", vbInformation + vbOKCancel, MsgTitulo) = vbCancel Then Exit Sub
        'MsgBox "No puede utilizar guías en documentos ya existentes...", vbExclamation + vbOKOnly, MsgTitulo
        'Unload Me
        'Exit Sub
    End If
    '*** Fin documento exista ***
    TotE = 0: TotI = 0: TotImp = 0: totnet = 0: TotGen = 0: vg_Guias = ""
    
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.Value = 1 Then
            M_DocPro.vaSpread1.MaxRows = 0
        End If
    Next i
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.Value = 1 Then
           vaSpread1.Col = 2: vg_Guias = vg_Guias & Str(vaSpread1.Value) & ";"
           '---Carga Encabezado de Documento---
           RS1.Open "select toc_numinf,toc_tipinf from b_totcompras where trim(toc_rutpro) ='" & vg_RDC & "' and toc_numdoc =" & Val(vaSpread1.Value) & " and toc_tipdoc ='" & vg_FDC & "'", vg_db, adOpenStatic
           If RS1!toc_tipinf = "C" Then
                M_DocPro.Option1(1).Value = True
                M_DocPro.Double1(6).Value = Val(RS1!toc_numinf)
           ElseIf RS1!toc_tipinf = "F" Then
                M_DocPro.Option1(0).Value = True
                M_DocPro.Double1(6).Value = Val(RS1!toc_numinf)
           End If
           RS1.Close: Set RS1 = Nothing
           RS1.Open "SELECT a.*, b.*, c.uni_nombre from b_detcompras a, b_productos b, a_unidad c where a.dec_codmer=b.pro_codigo and b.pro_coduni=c.uni_codigo and a.dec_rutpro = '" & vg_RDC & "' and a.dec_tipdoc = '" & vg_FDC & "' and a.dec_numdoc = " & Val(vaSpread1.Value) & " order by a.dec_numlin", vg_db, adOpenStatic
           '----Deshabilitar Botones
           M_DocPro.Frame1.Enabled = False
           M_DocPro.Frame3.Enabled = False
           M_DocPro.Frame6.Enabled = False
           M_DocPro.vaSpread1.Row = -1: M_DocPro.vaSpread1.Col = -1: M_DocPro.vaSpread1.Lock = True
           '*** Asigna código cfc o fifo segun corresponda ***
           '*** MSP : 16/08/2004
           '******* Detalle de Documento
            Do While Not RS1.EOF
                M_DocPro.vaSpread1.MaxRows = M_DocPro.vaSpread1.MaxRows + 1
                M_DocPro.vaSpread1.Col = 1: M_DocPro.vaSpread1.Row = M_DocPro.vaSpread1.MaxRows: M_DocPro.vaSpread1.Value = RS1!dec_codmer
                M_DocPro.vaSpread1.Col = 2: M_DocPro.vaSpread1.Value = Trim(RS1!pro_nombre)
                M_DocPro.vaSpread1.Col = 3: M_DocPro.vaSpread1.Value = RS1!uni_nombre
                If vg_FDC = "GD" Then
                    M_DocPro.vaSpread1.Col = 4: M_DocPro.vaSpread1.Value = RS1!dec_canmer
                    M_DocPro.vaSpread1.Col = 5: M_DocPro.vaSpread1.Value = RS1!dec_precom
                    M_DocPro.vaSpread1.Col = 6: M_DocPro.vaSpread1.Value = RS1!dec_pctdes
                    M_DocPro.vaSpread1.Col = 7: M_DocPro.vaSpread1.Value = RS1!dec_valdes
                    M_DocPro.vaSpread1.Col = 8: M_DocPro.vaSpread1.Value = RS1!dec_ptotal
                    M_DocPro.vaSpread1.Col = 9: M_DocPro.vaSpread1.Value = RS1!dec_canrec
                    M_DocPro.vaSpread1.Col = 10: M_DocPro.vaSpread1.Value = RS1!dec_prerec
                ElseIf vg_FDC = "SN" Then
                    M_DocPro.vaSpread1.Col = 4: M_DocPro.vaSpread1.Value = (RS1!dec_canmer - RS1!dec_canrec)
                    M_DocPro.vaSpread1.Col = 5: M_DocPro.vaSpread1.Value = RS1!dec_precom
                    M_DocPro.vaSpread1.Col = 6: M_DocPro.vaSpread1.Value = RS1!dec_pctdes
                    v_valdesc = ((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) * (RS1!dec_pctdes / 100)
                    v_valtot = ((RS1!dec_canmer - RS1!dec_canrec) * RS1!dec_precom) - v_valdesc
                    M_DocPro.vaSpread1.Col = 7: M_DocPro.vaSpread1.Value = v_valdesc
                    M_DocPro.vaSpread1.Col = 8: M_DocPro.vaSpread1.Value = v_valtot
                    M_DocPro.vaSpread1.Col = 9: M_DocPro.vaSpread1.Value = 0
                    M_DocPro.vaSpread1.Col = 10: M_DocPro.vaSpread1.Value = 0
                End If
                M_DocPro.vaSpread1.Col = 11: M_DocPro.vaSpread1.Value = RS1!dec_descri
                M_DocPro.vaSpread1.Col = 13: M_DocPro.vaSpread1.Value = RS1!dec_mueinv
                M_DocPro.vaSpread1.Col = 14: M_DocPro.vaSpread1.Text = ""
                RS.Open "select a.*, b.* from b_productosimp a, a_impuesto b where a.ipr_codimp=b.imp_codigo and a.ipr_codpro='" & RS1!dec_codmer & "'", vg_db, adOpenStatic
                Do While Not RS.EOF
                    M_DocPro.vaSpread1.Row = M_DocPro.vaSpread1.MaxRows
                    M_DocPro.vaSpread1.Col = 14: M_DocPro.vaSpread1.Text = M_DocPro.vaSpread1.Text & Trim(Str(RS!ipr_codimp)) & "&"
                    M_DocPro.vaSpread1.Col = 14: M_DocPro.vaSpread1.Text = M_DocPro.vaSpread1.Text & Trim(Str(RS!imp_pctimp)) & "&"
                    M_DocPro.vaSpread1.Col = 14: M_DocPro.vaSpread1.Text = M_DocPro.vaSpread1.Text & Trim(Str(RS!imp_inccos)) & ";"
                    RS.MoveNext
                Loop
                RS.Close: Set RS = Nothing
                RS1.MoveNext
            Loop
            RS1.Close: Set RS1 = Nothing
            '******Fin detalle de Documento
        End If
    Next i
    M_DocPro.SumarTotales
    M_DocPro.vaSpread1.Enabled = False
    M_DocPro.Text1(0).Enabled = False
    M_DocPro.Double1(0).Enabled = False: M_DocPro.Double1(1).Enabled = False: M_DocPro.Double1(2).Enabled = False: M_DocPro.Double1(3).Enabled = False: M_DocPro.Double1(4).Enabled = False
End Select
Me.Hide
Unload Me
Exit Sub
Error_Carga:
MsgBox Err & ": " & Err.Number & " " & Err.Description, vbCritical, MsgTitulo
Resume Next
End Sub

