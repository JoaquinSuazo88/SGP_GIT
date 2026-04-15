VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Period 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención Periodo Cierre de Mes"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5805
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   5925
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5325
         Left            =   150
         TabIndex        =   2
         Top             =   360
         Width           =   5595
         _Version        =   393216
         _ExtentX        =   9869
         _ExtentY        =   9393
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   25
         SpreadDesigner  =   "M_Period.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Period"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim modo As String, fecha As String, MsgTitulo As String

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6840
Me.Width = 6090
fg_centra Me
modo = "M": MsgTitulo = "Periodo Cierre de Mes"
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 5, modo: Toolbar1.Buttons.Item(3).Visible = False: Toolbar1.Buttons.Item(4).Visible = True
LlenarDatos
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 7
    LlenarDatos
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    LlenarDatos
    modo = "": Gl_Ac_Botones Me, 1, 5, modo: Toolbar1.Buttons.Item(3).Visible = False: Toolbar1.Buttons.Item(4).Visible = True
Case 12
    Dim cieper As Long, fecini As Long, fecter As Long
    fg_carga ""
    vg_db.BeginTrans
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 4
        If vaSpread1.Text <> "0" Then
           vaSpread1.Col = 1: cieper = 0: cieper = Format(vaSpread1.Text, "yyyymm")
           vaSpread1.Col = 2: fecini = 0: fecini = Format(vaSpread1.Text, "yyyymmdd")
           vaSpread1.Col = 3: fecter = 0: fecter = Format(vaSpread1.Text, "yyyymmdd")
           vg_db.Execute "update b_periodo set cie_fecini=" & fecini & ", cie_fecter=" & fecter & " where cie_periodo=" & cieper & ""
        End If
    Next i
    vg_db.CommitTrans
    modo = "": Gl_Ac_Botones Me, 1, 5, modo: Toolbar1.Buttons.Item(3).Visible = False: Toolbar1.Buttons.Item(4).Visible = True
    fg_descarga
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    I_PeriodoCierre
Case 18
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub LlenarDatos()
vaSpread1.Visible = False: vaSpread1.MaxRows = 0
RS.Open "select * from b_periodo order by cie_periodo", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Text = Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText 'CellTypeDate
'      vaSpread1.TypeDateCentury = False
      vaSpread1.TypeHAlign = TypeHAlignCenter
'      vaSpread1.TypeSpin = False
'      vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
'      vaSpread1.TypeDateMax = Mid(RS!cie_fecter, 5, 2) & Mid(RS!cie_fecter, 7, 2) & Mid(RS!cie_fecter, 1, 4)
'      vaSpread1.TypeDateMin = Mid(RS!cie_fecini, 5, 2) & Mid(RS!cie_fecini, 7, 2) & Mid(RS!cie_fecini, 1, 4)
      vaSpread1.Text = Mid(RS!cie_fecini, 7, 2) & "/" & Mid(RS!cie_fecini, 5, 2) & "/" & Mid(RS!cie_fecini, 1, 4)
'      vaSpread1.TypeDateCentury = True
      vaSpread1.Lock = True
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeDate
      vaSpread1.TypeDateCentury = False
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.TypeSpin = False
      vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
      vaSpread1.TypeDateMin = "01011973":  vaSpread1.TypeDateMax = "31125000"
      vaSpread1.TypeDateMax = IIf(RS!cie_estado <> 1, Mid(RS!cie_fecter, 5, 2) & Mid(RS!cie_fecter, 7, 2) & Mid(RS!cie_fecter, 1, 4), Format(dEoM(Mid(RS!cie_fecter, 7, 2) & "/" & Mid(RS!cie_fecter, 5, 2) & "/" & Mid(RS!cie_fecter, 1, 4)), "mmddyyyy"))
      vaSpread1.TypeDateMin = Mid(RS!cie_fecini, 5, 2) & fg_pone_cero(Str(Val(Mid(RS!cie_fecini, 7, 2)) + 1), 2) & Mid(RS!cie_fecini, 1, 4)
      vaSpread1.Text = Mid(RS!cie_fecter, 7, 2) & "/" & Mid(RS!cie_fecter, 5, 2) & "/" & Mid(RS!cie_fecter, 3, 2)
      vaSpread1.TypeDateCentury = True
      vaSpread1.Lock = IIf(RS!cie_estado = 0 Or RS!cie_estado = 2, True, False)
      
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      If RS!cie_estado = 0 Then
         vaSpread1.Text = "Cerrado"
      ElseIf RS!cie_estado = 1 Then
         vaSpread1.Text = "Abierto"
      Else
         vaSpread1.Text = "Inhabilitado"
      End If
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
vaSpread1.Row = Row: vaSpread1.Col = Col
If ChangeMade = False Then fecha = vaSpread1.Text
If ChangeMade = True Then
   '------- Verificar si existen datos en esa fecha en ventas
   RS.Open "select count(*) as nreg from b_totventas where cdate(tov_fecemi)='" & vaSpread1.Text & "' and cdate(tov_fecpro)='" & vaSpread1.Text & "' and not isnull(tov_fecpro) and tov_estdoc<>'A'", vg_db, adOpenStatic
   If Not RS.EOF And RS!NReg > 0 And Not IsNull(RS!NReg) Then RS.Close: Set RS = Nothing: MsgBox "Existen documentos para esa fecha, no podra modificarse hasta el proximo periodo", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Text = fecha: Exit Sub
   RS.Close: Set RS = Nothing
   '------- Fin verificar si existen datos en esa fecha en ventas
   
   '------- Verificar si existen datos en esa fecha en documentos
   RS.Open "select count(*) as nreg from b_totcompras where cdate(toc_fecemi)='" & CDate(vaSpread1.Text) & "'", vg_db, adOpenStatic
   If Not RS.EOF And RS!NReg > 0 And Not IsNull(RS!NReg) Then RS.Close: Set RS = Nothing: MsgBox "Existen documentos para esa fecha, no podra modificarse hasta el proximo periodo", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Text = fecha: Exit Sub
   RS.Close: Set RS = Nothing
   '------- Fin verificar si existen datos en esa fecha en documentos

   Dim i As Long, fecini As Long, fecfin As Long
   fecini = 0: fecfin = 0
   For i = Row To vaSpread1.MaxRows
       vaSpread1.Row = i
       If i = Row Then
'          vaSpread1.Col = 3
'          vaSpread1.TypeDateMax = Mid(vaSpread1.Text, 4, 2) & Mid(vaSpread1.Text, 1, 2) & Mid(vaSpread1.Text, 7, 4)
          fecini = Mid(vaSpread1.Text, 7, 4) & Mid(vaSpread1.Text, 4, 2) & fg_pone_cero(Str(Val(Mid(vaSpread1.Text, 1, 2))), 2)
          fecfin = Mid(vaSpread1.Text, 7, 4) & Mid(vaSpread1.Text, 4, 2) & Mid(vaSpread1.Text, 1, 2)
       Else
          If Mid(fecfin, 7, 2) > 27 Then
             fecini = Format(dBoM(BEoM(Mid(fecini, 7, 2) & "/" & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4))), "yyyymmdd")
             fecfin = Format(dEoM(BEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4))), "yyyymmdd")
          Else
             fecini = Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymm") & fg_pone_cero(Str(Val(Mid(fecini, 7, 2))), 2)
             fecfin = Format(BEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymm") & Mid(fecfin, 7, 2)
          End If
          vaSpread1.Col = 3
          vaSpread1.TypeDateCentury = False
          vaSpread1.TypeHAlign = TypeHAlignCenter
          vaSpread1.TypeSpin = False
          vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
          a = "01011973": b = "31125000"
          vaSpread1.TypeDateMin = a: vaSpread1.TypeDateMax = b
          vaSpread1.TypeDateMin = Mid(fecini, 5, 2) & fg_pone_cero(Str(Val(Mid(fecini, 7, 2) + 1)), 2) & Mid(fecini, 1, 4)
          vaSpread1.TypeDateMax = Mid(fecfin, 5, 2) & Mid(fecfin, 7, 2) & Mid(fecfin, 1, 4)
          vaSpread1.Text = Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 3, 2)
          vaSpread1.TypeDateCentury = True
          
          vaSpread1.Col = 2
          vaSpread1.Text = IIf(Mid(fecfin, 7, 2) > 27, Mid(fecini, 7, 2), fg_pone_cero(Str(Val(Mid(fecini, 7, 2) + 1)), 2)) & "/" & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
       End If
   Next i
   If modo = "" Then modo = "M"
   Gl_Ac_Botones Me, 1, 0, modo
End If
End Sub

