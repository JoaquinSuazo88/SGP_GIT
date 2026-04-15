VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SsllIPA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Indice Precio a los alimentos por periodo"
   ClientHeight    =   7950
   ClientLeft      =   8415
   ClientTop       =   3075
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6255
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   4575
         _Version        =   393216
         _ExtentX        =   8070
         _ExtentY        =   11033
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   3
         MaxRows         =   50
         ScrollBars      =   2
         SpreadDesigner  =   "M_SsllIPA.frx":0000
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Período Actual"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   3
         Top             =   6600
         Width           =   1065
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0C0FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   1080
         Top             =   6630
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   2760
         Top             =   6630
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Períodos Anteriores"
         Height          =   195
         Index           =   0
         Left            =   3120
         TabIndex        =   2
         Top             =   6600
         Width           =   1395
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_SsllIPA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As ADODB.Recordset
Dim strSQL As String
Dim wvarMes As String
Dim FechaActual As String
Dim modo As String
Dim codigo, Msgtitulo, ModoGuardar As String
Dim est As Boolean
Dim wvarPeriodoUpdAnio, wvarPeriodoUpdMes, wvarPrecioUpd As String
Dim wvarOrdenGrd As String

Private Sub Form_Load()
    Me.HelpContextID = vg_OpcM
    Me.Height = 8475
    Me.Width = 6585
    fg_centra Me
    Msgtitulo = "Índice de Precios a los Alimentos por Período"
    
    modo = ""
    
    Gl_Mo_Botones Me, 1
    Gl_Ac_Botones Me, 1, 14, modo
    MoverDatosGrilla
    
End Sub

Sub MoverDatosGrilla()
    fg_carga ""
    Dim x As Boolean
    Dim i As Integer
    Dim varValCol1, varValCol2 As String
    vaSpread1.TextTip = 2
    vaSpread1.TextTipDelay = 250
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    vaSpread1.Lock = True
    
    strSQL = "SELECT ipa_period, ipa_valor " & _
             "FROM b_ssll_ipa " & _
             "ORDER BY ipa_period DESC"

    Set RS = vg_db.Execute(strSQL)
    
    FechaActual = Format(Date, "yyyymm")
    
    Do While Not RS.EOF
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1
        vaSpread1.text = Meses("01/" & Mid(RS!ipa_period, 5, 2) & "/" & Mid(RS!ipa_period, 1, 4)) & " / " & Mid(RS!ipa_period, 1, 4)
        vaSpread1.ForeColor = &H0& '&H808080
        vaSpread1.BackColorStyle = BackColorStyleUnderGrid: vaSpread1.BackColorStyle = BackColorStyleUnderGrid
        vaSpread1.BackColor = &H8000000F
        
        vaSpread1.Col = 2
        varValCol2 = CDbl(RS!ipa_valor)
        vaSpread1.text = varValCol2
        vaSpread1.ForeColor = &H0& '&H808080
        vaSpread1.BackColorStyle = BackColorStyleUnderGrid: vaSpread1.BackColorStyle = BackColorStyleUnderGrid
        vaSpread1.BackColor = &H8000000F
        
        vaSpread1.Col = 3
        vaSpread1.text = RS!ipa_period
        
        If (Val(RS!ipa_period) >= Val(FechaActual)) Then
            vaSpread1.Lock = False
            vaSpread1.Col = 1: vaSpread1.CellType = CellTypeEdit
            vaSpread1.ForeColor = &H0&: vaSpread1.BackColor = &HE0FEFE
            
            vaSpread1.Col = 2: vaSpread1.CellType = CellTypeEdit: vaSpread1.TypeHAlign = TypeHAlignRight
            
            If CDbl(vaSpread1.text) < 0 Then
                vaSpread1.ForeColor = &HFF&
            Else
                vaSpread1.ForeColor = &H0&
            End If
            
            vaSpread1.BackColor = &HE0FEFE
            
            If (Val(RS!ipa_period) = Val(FechaActual)) Then
                vaSpread1.Col = 1: vaSpread1.BackColorStyle = BackColorStyleUnderGrid: vaSpread1.BackColor = &HC0C0FF
                vaSpread1.Col = 2: vaSpread1.BackColorStyle = BackColorStyleUnderGrid: vaSpread1.BackColor = &HC0C0FF
            End If
            
        End If
                 
        RS.MoveNext
    Loop
    
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 3
        
        If (vaSpread1.text >= FechaActual) Then
            vaSpread1.Col = 2
            vaSpread1.Lock = False
            vaSpread1.CellType = CellTypeEdit
            vaSpread1.TypeHAlign = TypeHAlignRight
        End If
        
    Next i
        
    RS.Close: Set RS = Nothing
    
    vaSpread1.Visible = True
    
'    If (wvarOrdenGrd = "DESC") Then
'        vaSpread1.SortKey(1) = 3
'        vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending 'SortKeyOrderDescending
'        vaSpread1.Sort -1, -1, -1, -1, SortByRow
'    End If
    
    fg_descarga
End Sub

Function AsignaMesStr(mes)
    
    strmes = Right(Trim(mes), 2)
    
    If (strmes = "01") Then
        wvarMes = "Enero"
    ElseIf (strmes = "02") Then
        wvarMes = "Febrero"
    ElseIf (strmes = "03") Then
        wvarMes = "Marzo"
    ElseIf (strmes = "04") Then
        wvarMes = "Abril"
    ElseIf (strmes = "05") Then
        wvarMes = "Mayo"
    ElseIf (strmes = "06") Then
        wvarMes = "Junio"
    ElseIf (strmes = "07") Then
        wvarMes = "Julio"
    ElseIf (strmes = "08") Then
        wvarMes = "Agosto"
    ElseIf (strmes = "09") Then
        wvarMes = "Septiembre"
    ElseIf (strmes = "10") Then
        wvarMes = "Octubre"
    ElseIf (strmes = "11") Then
        wvarMes = "Noviembre"
    ElseIf (strmes = "12") Then
        wvarMes = "Diciembre"
    End If
    
    AsignaMesStr = wvarMes
    
End Function

Function AsignaMesNum(ByVal mes As String)
    Dim wvarStr, wvarMes, wvarNumMes, wvarAnio As String
    
    wvarStr = Trim(fg_Quitachar(mes, "/"))
    wvarAnio = Trim(Right(wvarStr, 4))
    wvarMes = Trim(Mid(wvarStr, 1, Len(wvarStr) - 4))
    
    If (wvarMes = "Enero") Then
        wvarNumMes = "01"
    ElseIf (wvarMes = "Febrero") Then
        wvarNumMes = "02"
    ElseIf (wvarMes = "Marzo") Then
        wvarNumMes = "03"
    ElseIf (wvarMes = "Abril") Then
        wvarNumMes = "04"
    ElseIf (wvarMes = "Mayo") Then
        wvarNumMes = "05"
    ElseIf (wvarMes = "Junio") Then
        wvarNumMes = "06"
    ElseIf (wvarMes = "Julio") Then
        wvarNumMes = "07"
    ElseIf (wvarMes = "Agosto") Then
        wvarNumMes = "08"
    ElseIf (wvarMes = "Septiembre") Then
        wvarNumMes = "09"
    ElseIf (wvarMes = "Octubre") Then
        wvarNumMes = "10"
    ElseIf (wvarMes = "Noviembre") Then
        wvarNumMes = "11"
    ElseIf (wvarMes = "Diciembre") Then
        wvarNumMes = "12"
    End If
    
    AsignaMesNum = wvarNumMes & "_" & wvarAnio
    
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Dim RS As ADODB.Recordset
Dim UltimoMes, UltimoAnio As String
Dim NuevoMes, NuevoAnio As String
Dim i       As Long
Dim strSQL  As String

Select Case Button.Index
Case 1 'NUEVO
    
    wvarOrdenGrd = "DESC"
    MoverDatosGrilla
    wvarOrdenGrd = ""

    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        
        vaSpread1.Col = 1
        vaSpread1.Lock = True
        vaSpread1.ForeColor = &H808080
        
        vaSpread1.Col = 2
        vaSpread1.Lock = True
        vaSpread1.ForeColor = &H808080
    Next i
    
    strSQL = "select count(*) as nreg, max(ipa_period) as ipa_period from b_ssll_ipa"
    Set RS = vg_db.Execute(strSQL)
    
    If Not RS.EOF Then
       If Not IsNull(RS!nReg) And RS!nReg >= 0 Then
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows 1, 1
          vaSpread1.Row = 1
'          vaSpread1.Row = vaSpread1.MaxRows
          vaSpread1.Col = 1
          If RS!nReg = 0 Then
             vaSpread1.text = Meses(BoM(Format(Date, "dd/mm/yyyy"))) & " / " & Format(BoM(Format(Date, "dd/mm/yyyy")), "yyyy")
          Else
             vaSpread1.text = Meses(EoM("27/" & Mid(RS!ipa_period, 5, 2) & "/" & Mid(RS!ipa_period, 1, 4))) & " / " & Format(EoM("27/" & Mid(RS!ipa_period, 5, 2) & "/" & Mid(RS!ipa_period, 1, 4)), "yyyy")
          End If
          
          vaSpread1.Col = 2
          vaSpread1.Lock = False
          vaSpread1.CellType = CellTypeNumber 'CellTypeEdit
          vaSpread1.TypeHAlign = TypeHAlignRight
       
          vaSpread1.SetActiveCell 2, 1: vaSpread1.SetFocus
       End If
    Else
    End If
    RS.Close: Set RS = Nothing
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    ModoGuardar = "Nuevo"
Case 3 'ALTERAR
Case 5 'BORRAR
Case 7 'ACTUALIZAR
    MoverDatosGrilla
Case 10 'CANCELAR-ACTIVO
    If MsgBox("Cancela...", vbQuestion + vbYesNo, "IPA por Período") = vbNo Then Exit Sub
    Cancela
    ModoGuardar = ""
Case 11 'CANCELAR
Case 12 'GUARDAR-ACTIVO
    If (ModoGuardar = "Nuevo") Then
        
        vaSpread1.Row = 1 'vaSpread1.MaxRows
        vaSpread1.Col = 1
        varNuevoStr = AsignaMesNum(vaSpread1.text)
        vaSpread1.Col = 2
        varvalor = vaSpread1.text
        
        If (varvalor = "") Then
            MsgBox "Debe Ingresar el Valor del IPA", vbExclamation + vbOKOnly, "IPA"
            Exit Sub
'        ElseIf (varvalor < 0) Then
'            MsgBox "El Valor del IPA Debe Ser Mayor a Cero", vbExclamation + vbOKOnly, "IPA"
'            Exit Sub
        End If
        
        If (varvalor <> "" And IsNumeric(varvalor)) Then
            
            varPeriodoAnio = Right(varNuevoStr, 4)
            varPeriodoMes = Left(varNuevoStr, 2)
            
            strSQL = "INSERT INTO b_ssll_ipa(ipa_period, ipa_valor) VALUES('" & varPeriodoAnio & varPeriodoMes & "', " & varvalor & ")"
            vg_db.Execute strSQL
            
        End If
        
    ElseIf (ModoGuardar = "Actualizar") Then
    
        If (wvarPrecioUpd = "") Then
            MsgBox "Debe Ingresar el Valor del IPA", vbExclamation + vbOKOnly, "IPA"
            Exit Sub
        End If
    
        If (wvarPeriodoUpdAnio <> "" And wvarPeriodoUpdMes <> "" And wvarPrecioUpd <> "") Then
            
            strSQL = "UPDATE b_ssll_ipa SET ipa_valor = " & wvarPrecioUpd & " " & _
                     "WHERE ipa_period = '" & wvarPeriodoUpdAnio & wvarPeriodoUpdMes & "'"
                     
            vg_db.Execute strSQL
    
            Cancela
            
            vaSpread1.Row = 1
            vaSpread1.Col = 1
            vaSpread1.SetFocus
        End If
    
    End If
    
    Cancela
    
    ModoGuardar = ""
    
Case 13 'GUARDAR
Case 15 'IMPRIMIR
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_SsllIPA
Case 18 'SALIR
    Me.Hide
    Unload Me
End Select
Exit Sub

Man_Error:
If Err = 3034 Then Exit Sub
If Err = 13 Then Exit Sub
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
    Dim varNuevoStr, varPeriodoAnio, varPeriodoMes As String

    If Toolbar1.Buttons(12).Visible = True Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    
    If (ModoGuardar <> "Nuevo") Then
    
        ModoGuardar = "Actualizar"
        
        vaSpread1.Row = Row
        vaSpread1.Col = 1
        varNuevoStr = AsignaMesNum(vaSpread1.text)
                
        wvarPeriodoUpdAnio = Right(varNuevoStr, 4)
        wvarPeriodoUpdMes = Left(varNuevoStr, 2)
        
        vaSpread1.Col = 2
        wvarPrecioUpd = vaSpread1.text
    
    End If
        
End Sub

Private Sub Cancela()
    OpGr = True
    vaSpread1.Row = vaSpread1.ActiveRow
    
    MoverDatosGrilla
    
    OpGr = False
    TipoOp = ""
    modo = ""
    Gl_Ac_Botones Me, 1, 14, modo
End Sub


