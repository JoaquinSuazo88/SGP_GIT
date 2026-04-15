VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_AjusteEstacionales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste Estacionales Receta"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   7455
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   11055
      Begin VB.CommandButton Command1 
         Caption         =   "Agr. Recetas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8880
         TabIndex        =   5
         Top             =   6840
         Width           =   1695
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   5640
         TabIndex        =   11
         Top             =   6720
         Width           =   2700
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   2595
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   4680
         TabIndex        =   10
         Top             =   6720
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   1680
         TabIndex        =   9
         Top             =   6720
         Width           =   2700
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   2
            Top             =   135
            Width           =   2595
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   720
         TabIndex        =   1
         Top             =   6720
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   795
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6255
         Left            =   240
         TabIndex        =   0
         Top             =   360
         Width           =   10575
         _Version        =   393216
         _ExtentX        =   18653
         _ExtentY        =   11033
         _StockProps     =   64
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_AjusteEstacionales.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_AjusteEstacionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo      As String
Dim OpGr      As Boolean
Public lc_Aux As String

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS               As New ADODB.Recordset
Dim RecetaOrigen     As Long
Dim AuxRecetaOrigen  As Long
Dim RecetaDestino    As Long
Dim AuxRecetaDestino As Long
Dim FechaIni         As String
Dim AuxFechaIni      As String
Dim FechaFin         As String
Dim AuxFechaFin      As String

Dim i             As Long

vaSpread1.Row = Fila
vaSpread1.Col = 1
RecetaOrigen = Val(vaSpread1.Value)
vaSpread1.Col = 3
RecetaDestino = Val(vaSpread1.Value)
vaSpread1.Col = 5
FechaIni = Replace(vaSpread1.Value, "", "")
vaSpread1.Col = 6
FechaFin = Replace(vaSpread1.Value, "", "")

'-------> Validar fecha Ini
If Mid(FechaIni, 3, 1) <> "/" Then

   vaSpread1.SetActiveCell 5, Fila: vaSpread1.SetFocus
   MsgBox "debe utilizar un separador " & "/ fecha fin", vbCritical, "Boton_Click"
   Exit Sub
    
End If

If Mid(FechaIni, 4, 2) = "01" Or Mid(FechaIni, 4, 2) = "03" Or Mid(FechaIni, 4, 2) = "05" Or Mid(FechaIni, 4, 2) = "07" Or Mid(FechaIni, 4, 2) = "08" Or Mid(FechaIni, 4, 2) = "10" Or Mid(FechaIni, 4, 2) = "12" Then

    If Val(Mid(FechaIni, 1, 2)) > 31 Or Val(Mid(FechaIni, 1, 2)) < 1 Then
    
       vaSpread1.SetActiveCell 5, Fila: vaSpread1.SetFocus
       MsgBox "el valor del día debe ser menor igual treinta y uno días...", vbCritical, "Boton_Click"
       Exit Sub
    
    End If


ElseIf Mid(FechaIni, 4, 2) = "02" Then
    
    If Val(Mid(FechaIni, 1, 2)) > 29 Or Val(Mid(FechaIni, 1, 2)) < 1 Then
    
       vaSpread1.SetActiveCell 5, Fila: vaSpread1.SetFocus
       MsgBox "el valor del día debe ser menor igual veinte nueve días...", vbCritical, "Boton_Click"
       Exit Sub
    
    End If

ElseIf Mid(FechaIni, 4, 2) = "04" Or Mid(FechaIni, 4, 2) = "06" Or Mid(FechaIni, 4, 2) = "09" Or Mid(FechaIni, 4, 2) = "11" Then

    If Val(Mid(FechaIni, 1, 2)) > 30 Or Val(Mid(FechaIni, 1, 2)) < 1 Then
    
       vaSpread1.SetActiveCell 5, Fila: vaSpread1.SetFocus
       MsgBox "el valor del día debe ser menor igual treinta días...", vbCritical, "Boton_Click"
       Exit Sub
                
    End If
                
ElseIf Val(Mid(FechaIni, 4, 2)) > 12 Then


    vaSpread1.SetActiveCell 5, Fila: vaSpread1.SetFocus
    MsgBox "Mes no corresponde...", vbCritical, "Boton_Click"
    Exit Sub

End If


'-------> Validar fecha Fin
If Mid(FechaFin, 3, 1) <> "/" Then

   vaSpread1.SetActiveCell 6, Fila: vaSpread1.SetFocus
   MsgBox "debe utilizar un separador " & "/ fecha fin", vbCritical, "Boton_Click"
   Exit Sub

End If

If Mid(FechaFin, 4, 2) = "01" Or Mid(FechaFin, 4, 2) = "03" Or Mid(FechaFin, 4, 2) = "05" Or Mid(FechaFin, 4, 2) = "07" Or Mid(FechaFin, 4, 2) = "08" Or Mid(FechaFin, 4, 2) = "10" Or Mid(FechaFin, 4, 2) = "12" Then

    If Val(Mid(FechaFin, 1, 2)) > 31 Or Val(Mid(FechaFin, 1, 2)) < 1 Then
    
       vaSpread1.SetActiveCell 6, Fila: vaSpread1.SetFocus
       MsgBox "el valor del día debe ser menor igual treinta y uno días...", vbCritical, "Boton_Click"
       Exit Sub
                
    End If


ElseIf Mid(FechaFin, 4, 2) = "02" Then
    
    If Val(Mid(FechaFin, 1, 2)) > 29 Or Val(Mid(FechaFin, 1, 2)) < 1 Then
    
       vaSpread1.SetActiveCell 6, Fila: vaSpread1.SetFocus
       MsgBox "el valor del día debe ser menor igual veinte nueve días...", vbCritical, "Boton_Click"
       Exit Sub
                
    End If

ElseIf Mid(FechaFin, 4, 2) = "04" Or Mid(FechaFin, 4, 2) = "06" Or Mid(FechaFin, 4, 2) = "09" Or Mid(FechaFin, 4, 2) = "11" Then

    If Val(Mid(FechaFin, 1, 2)) > 30 Or Val(Mid(FechaFin, 1, 2)) < 1 Then
    
       vaSpread1.SetActiveCell 6, Fila: vaSpread1.SetFocus
       MsgBox "el valor del día debe ser menor igual treinta días...", vbCritical, "Boton_Click"
       Exit Sub
                
    End If
                
ElseIf Val(Mid(FechaFin, 4, 2)) > 12 Then


    vaSpread1.SetActiveCell 6, Fila: vaSpread1.SetFocus
    MsgBox "Mes no corresponde...", vbCritical, "Boton_Click"
   Exit Sub

End If

'------> Validar iteración y traslape
For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 1
    AuxRecetaOrigen = Val(vaSpread1.Value)

    vaSpread1.Col = 3
    AuxRecetaDestino = Val(vaSpread1.Value)
    
    vaSpread1.Col = 5
    AuxFechaIni = Replace(vaSpread1.Value, "", "")

    vaSpread1.Col = 6
    AuxFechaFin = Replace(vaSpread1.Value, "", "")
    
    If i <> Fila Then
    
       If RecetaOrigen = AuxRecetaOrigen And RecetaDestino = AuxRecetaDestino And FechaIni = AuxFechaIni And FechaFin = AuxFechaFin Then
       
          MsgBox "El dato ya existe registrado...", vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
          
'       ElseIf RecetaOrigen = AuxRecetaOrigen And RecetaDestino = AuxRecetaDestino And (Val(AuxFechaIni) <= Val(FechaIni) And Val(AuxFechaFin) >= Val(FechaIni)) And (Val(AuxFechaIni) <= Val(FechaFin) And Val(AuxFechaFin) >= Val(FechaFin)) Then
       ElseIf RecetaOrigen = AuxRecetaOrigen And RecetaDestino = AuxRecetaDestino And (Format(FechaIni, "mmdd") >= Format(AuxFechaIni, "mmdd")) And (Format(FechaIni, "mmdd") <= Format(AuxFechaFin, "mmdd")) Then
          
          MsgBox "El dato esta traslapado...", vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
                    
       ElseIf RecetaOrigen = AuxRecetaOrigen And RecetaDestino = AuxRecetaDestino And (Format(FechaFin, "mmdd") >= Format(AuxFechaIni, "mmdd")) And (Format(FechaFin, "mmdd") <= Format(AuxFechaFin, "mmdd")) Then
          
          MsgBox "El dato esta traslapado...", vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
                    
       
       End If
    
    End If

Next i

OpGr = True

If Format(Trim(FechaFin), "mmdd") < Format(Trim(FechaIni), "mmdd") Then

   MsgBox "Fecha Final es menor fecha inicial...", vbCritical + vbOKOnly, MsgTitulo
   Exit Sub

End If

vaSpread1.Row = Fila
vaSpread1.Col = 1
RecetaOrigen = Val(vaSpread1.Value)
vaSpread1.Col = 3
RecetaDestino = Val(vaSpread1.Value)
vaSpread1.Col = 5
FechaIni = Replace(vaSpread1.Value, "/", "")
vaSpread1.Col = 6
FechaFin = Replace(vaSpread1.Value, "/", "")

If RecetaOrigen = 0 Then

   MsgBox "Falta información receta origen...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 1
   vaSpread1.SetActiveCell 1, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If RecetaDestino = 0 Then

   MsgBox "Falta información receta destino...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 3
   vaSpread1.SetActiveCell 3, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If Trim(FechaIni) = "" Then

   MsgBox "Falta información fecha inicio...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 5
   vaSpread1.SetActiveCell 5, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If Trim(FechaFin) = "" Then

   MsgBox "Falta información fecha final...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 6
   vaSpread1.SetActiveCell 6, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If modo = "A" Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_Ins_AjusteEstacional " & RecetaOrigen & ", " & RecetaDestino & ", '" & FechaIni & "', '" & FechaFin & "'")
   
   If Not RS.EOF Then
      
      If RS(0) > 0 Then
              
         RS.Close
         Set RS = Nothing
         
         MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
         Exit Sub
         
      Else
              
         MsgBox "Ajuste Estacionales [OK]", vbInformation + vbOKOnly, MsgTitulo
           
     End If

   
   End If
   RS.Close
   Set RS = Nothing
   
Else
  
  Dim testArray() As String
  '=
  vaSpread1.Row = Fila
  vaSpread1.Col = 7
  testArray = Split(vaSpread1.text, ";")
  AuxRecetaOrigen = testArray(0)
  AuxRecetaDestino = testArray(1)
  AuxFechaIni = testArray(2)
  AuxFechaFin = testArray(3)

  vg_db.Execute "DELETE FROM b_ajusteestacionales WHERE [Id_RecetaOrigen] = " & AuxRecetaOrigen & " and [Id_RecetaDestino] = " & AuxRecetaDestino & " and [FechaInicial] = '" & AuxFechaIni & "' and  [FechaFinal] = '" & AuxFechaFin & "'"

  Set RS = vg_db.Execute("sgpadm_Ins_AjusteEstacional " & RecetaOrigen & ", " & RecetaDestino & ", '" & FechaIni & "', '" & FechaFin & "'")

   If Not RS.EOF Then
      
      If RS(0) > 0 Then
              
         RS.Close
         Set RS = Nothing
         
         MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
         Exit Sub
         
      Else
              
         MsgBox "Ajuste Estacionales [OK]", vbInformation + vbOKOnly, MsgTitulo
           
     End If

   
   End If
   RS.Close
   Set RS = Nothing


End If
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

vg_nombre = ""
vg_codigo = ""
vg_left = Command1.Left + 550
B_TabEst.LlenaDatos "b_receta", "rec_", "Recetas", "AgregarRec"
B_TabEst.Show 1
If vg_codigo = "" Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol
vaSpread1.text = vg_codigo

vaSpread1.Col = vaSpread1.ActiveCol + 1
vaSpread1.text = vg_nombre

vaSpread1.SetActiveCell vaSpread1.ActiveCol + 2, vaSpread1.Row: vaSpread1.SetFocus

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
MsgTitulo = "Ajuste Estacionales Recetas"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo

MoverDatosGrillas
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 1 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""

ElseIf Index = 2 Then
   
   TextDet2(1).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""

ElseIf Index = 3 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(4).text = ""

ElseIf Index = 4 Then
   
   TextDet2(1).text = ""
   TextDet2(3).text = ""
   TextDet2(2).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 8
    vaSpread1.text = 0
    
Next

Select Case Index

Case 1, 2, 3, 4
    
    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 1 Or Index = 3, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 8
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 8
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 8
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 8
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RecetaOrigen  As Long
Dim RecetaDestino As Long
Dim FechaIni      As String
Dim FechaFin      As String

Select Case Button.Index

    Case 1
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1
        vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
        vaSpread1.SetFocus
    
    Case 3
        
        modo = "M"
        If vaSpread1.MaxRows < 1 Then Exit Sub
        Gl_Ac_Botones Me, 1, 0, modo
    
    Case 5
        
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        RecetaOrigen = Val(vaSpread1.Value)
        vaSpread1.Col = 3
        RecetaDestino = Val(vaSpread1.Value)
        vaSpread1.Col = 5
        FechaIni = Replace(vaSpread1.Value, "/", "")
        vaSpread1.Col = 6
        FechaFin = Replace(vaSpread1.Value, "/", "")
                
        vg_db.Execute "DELETE FROM b_ajusteestacionales WHERE [Id_RecetaOrigen] = " & RecetaOrigen & " and [Id_RecetaDestino] = " & RecetaDestino & " and [FechaInicial] = '" & FechaIni & "' and  [FechaFinal] = '" & FechaFin & "'"
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
    
    Case 7
        
        MoverDatosGrillas
    
    Case 10
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        If modo = "A" Then
            
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.DeleteRows vaSpread1.Row, 1
           vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        
        Else
            
           Cancela
        
        End If
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
    
    Case 12
        
        GrabaRegistro vaSpread1.ActiveRow
    
    Case 15
        
        If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        I_AjusteEstacionalReceta
    
    Case 18
        
        Me.Hide
        Unload Me

End Select
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then
   
   MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"

Else
    
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End If

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

'If modo = "" Then modo = "M"
'Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

OpGr = True
vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_AjusteEstacionales ")

Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 1
    vaSpread1.Value = RS!Id_RecetaOrigen
    
    vaSpread1.Col = 2
    vaSpread1.Value = Trim(RS!NombreOrigen)
    
    vaSpread1.Col = 3
    vaSpread1.Value = RS!Id_RecetaDestino
    
    vaSpread1.Col = 4
    vaSpread1.Value = Trim(RS!NombreDestino)
    
    vaSpread1.Col = 5
    vaSpread1.Value = fg_pone_cero(Mid(Trim(RS!FechaInicial), 1, 2), 2) & "/" & fg_pone_cero(Mid(Trim(RS!FechaInicial), 3, 2), 2)
    
    vaSpread1.Col = 6
    vaSpread1.Value = fg_pone_cero(Mid(Trim(RS!FechaFinal), 1, 2), 2) & "/" & fg_pone_cero(Mid(Trim(RS!FechaFinal), 3, 2), 2)
    
    vaSpread1.Col = 7
    vaSpread1.Value = RS!Id_RecetaOrigen & ";" & RS!Id_RecetaDestino & ";" & RS!FechaInicial & ";" & RS!FechaFinal
        
    vaSpread1.Col = 8
    vaSpread1.Value = 0
    
    RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing

Gl_Ac_Botones Me, 1, 1, modo
vaSpread1.Visible = True
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

Dim RS            As New ADODB.Recordset
Dim RecetaOrigen  As Long
Dim RecetaDestino As Long

'If modo = "" Then modo = "M"
Select Case Col

    Case 1
    
        If (modo = "M" Or modo = "A") And ChangeMade = True Then
            
            vaSpread1.Row = Row
            vaSpread1.Col = Col
            RecetaOrigen = vaSpread1.text
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_s_receta_V07 8, " & RecetaOrigen & ", '', 0, 0, 0, '" & vg_NUsr & "'")
            
            If Not RS.EOF Then
            
               vaSpread1.Col = 2
               vaSpread1.text = IIf(IsNull(RS!rec_nombre), "", RS!rec_nombre)
               vaSpread1.SetActiveCell 3, vaSpread1.Row: vaSpread1.SetFocus
               
            Else
            
               vaSpread1.Col = 2
               vaSpread1.text = ""
               vaSpread1.SetActiveCell 1, vaSpread1.Row: vaSpread1.SetFocus
               MsgBox "Receta origen no existe", vbCritical, "Boton_Click"
                
            End If
            RS.Close
            Set RS = Nothing
            
        End If
    
    Case 3

        If (modo = "M" Or modo = "A") And ChangeMade = True Then
        
            vaSpread1.Row = Row
            vaSpread1.Col = Col
            RecetaDestino = vaSpread1.text
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_s_receta_V07 8, " & RecetaDestino & ", '', 0, 0, 0, '" & vg_NUsr & "'")
            
            If Not RS.EOF Then
            
               vaSpread1.Col = 4
               vaSpread1.text = IIf(IsNull(RS!rec_nombre), "", RS!rec_nombre)
               vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus
               
            Else
            
               vaSpread1.Col = 4
               vaSpread1.text = ""
               vaSpread1.SetActiveCell 3, vaSpread1.Row: vaSpread1.SetFocus
               MsgBox "Receta destino no existe", vbCritical, "Boton_Click"
                
            End If
            RS.Close
            Set RS = Nothing
            
        End If
    
    Case 5
    
        If (modo = "M" Or modo = "A") And ChangeMade = True Then
        
            vaSpread1.Row = Row
            vaSpread1.Col = Col
            
            If Mid(vaSpread1.text, 3, 1) <> "/" Then
            
               vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus
               MsgBox "debe utilizar un separador " & "/ fecha inicio", vbCritical, "Boton_Click"
            
            End If
        
            If Mid(vaSpread1.text, 4, 2) = "01" Or Mid(vaSpread1.text, 4, 2) = "03" Or Mid(vaSpread1.text, 4, 2) = "05" Or Mid(vaSpread1.text, 4, 2) = "07" Or Mid(vaSpread1.text, 4, 2) = "08" Or Mid(vaSpread1.text, 4, 2) = "10" Or Mid(vaSpread1.text, 4, 2) = "12" Then
            
                If Val(Mid(vaSpread1.text, 1, 2)) > 31 Or Val(Mid(vaSpread1.text, 1, 2)) < 1 Then
                
                   vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus
                   MsgBox "el valor del día debe ser menor igual treinta y uno días...", vbCritical, "Boton_Click"
                            
                End If
            
            
            ElseIf Mid(vaSpread1.text, 4, 2) = "02" Then
                
                If Val(Mid(vaSpread1.text, 1, 2)) > 29 Or Val(Mid(vaSpread1.text, 1, 2)) < 1 Then
                
                   vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus
                   MsgBox "el valor del día debe ser menor igual veinte nueve días...", vbCritical, "Boton_Click"
                            
                End If
            
            ElseIf Mid(vaSpread1.text, 4, 2) = "04" Or Mid(vaSpread1.text, 4, 2) = "06" Or Mid(vaSpread1.text, 4, 2) = "09" Or Mid(vaSpread1.text, 4, 2) = "11" Then
            
                If Val(Mid(vaSpread1.text, 1, 2)) > 30 Or Val(Mid(vaSpread1.text, 1, 2)) < 1 Then
                
                   vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus
                   MsgBox "el valor del día debe ser menor igual treinta días...", vbCritical, "Boton_Click"
                            
                End If
                            
            ElseIf Val(Mid(vaSpread1.text, 4, 2)) > 12 Then
            
            
                vaSpread1.SetActiveCell 5, vaSpread1.Row: vaSpread1.SetFocus
                MsgBox "Mes no corresponde...", vbCritical, "Boton_Click"
            
            End If
                
        End If
'        vaSpread1.SetActiveCell 6, vaSpread1.Row: vaSpread1.SetFocus
        
    Case 6
    
        If (modo = "M" Or modo = "A") And ChangeMade = True Then

            vaSpread1.Row = Row
            vaSpread1.Col = Col
            
            If Mid(vaSpread1.text, 3, 1) <> "/" Then
            
               vaSpread1.SetActiveCell 6, vaSpread1.Row: vaSpread1.SetFocus
               MsgBox "debe utilizar un separador " & "/ fecha fin", vbCritical, "Boton_Click"
            
            End If
        
            If Mid(vaSpread1.text, 4, 2) = "01" Or Mid(vaSpread1.text, 4, 2) = "03" Or Mid(vaSpread1.text, 4, 2) = "05" Or Mid(vaSpread1.text, 4, 2) = "07" Or Mid(vaSpread1.text, 4, 2) = "08" Or Mid(vaSpread1.text, 4, 2) = "10" Or Mid(vaSpread1.text, 4, 2) = "12" Then
            
                If Val(Mid(vaSpread1.text, 1, 2)) > 31 Or Val(Mid(vaSpread1.text, 1, 2)) < 1 Then
                
                   vaSpread1.SetActiveCell 6, vaSpread1.Row: vaSpread1.SetFocus
                   MsgBox "el valor del día debe ser menor igual treinta y uno días...", vbCritical, "Boton_Click"
                            
                End If
            
            
            ElseIf Mid(vaSpread1.text, 4, 2) = "02" Then
                
                If Val(Mid(vaSpread1.text, 1, 2)) > 29 Or Val(Mid(vaSpread1.text, 1, 2)) < 1 Then
                
                   vaSpread1.SetActiveCell 6, vaSpread1.Row: vaSpread1.SetFocus
                   MsgBox "el valor del día debe ser menor igual veinte nueve días...", vbCritical, "Boton_Click"
                            
                End If
            
            ElseIf Mid(vaSpread1.text, 4, 2) = "04" Or Mid(vaSpread1.text, 4, 2) = "06" Or Mid(vaSpread1.text, 4, 2) = "09" Or Mid(vaSpread1.text, 4, 2) = "11" Then
            
                If Val(Mid(vaSpread1.text, 1, 2)) > 30 Or Val(Mid(vaSpread1.text, 1, 2)) < 1 Then
                
                   vaSpread1.SetActiveCell 6, vaSpread1.Row: vaSpread1.SetFocus
                   MsgBox "el valor del día debe ser menor igual treinta días...", vbCritical, "Boton_Click"
                            
                End If
                            
            ElseIf Val(Mid(vaSpread1.text, 4, 2)) > 12 Then
            
            
                vaSpread1.SetActiveCell 6, vaSpread1.Row: vaSpread1.SetFocus
                MsgBox "Mes no corresponde...", vbCritical, "Boton_Click"
            
            End If

        End If
        
End Select

'Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    
    GrabaRegistro Row

ElseIf Toolbar1.Buttons(12).Visible = False Then
    
    Cancela

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Cancela()

On Error GoTo Man_Error

MoverDatosGrillas

'Dim RS As New ADODB.Recordset
'OpGr = True
'vaSpread1.Row = vaSpread1.ActiveRow
'vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
'
'If RS.State = 1 Then RS.Close
'RS.CursorLocation = adUseClient
'vg_db.CursorLocation = adUseClient
'Set RS = vg_db.Execute("sgpadm_s_nutriente 5, 0,''")
'If Not RS.EOF Then
'
'    vaSpread1.Col = 2
'    vaSpread1.Value = Trim(RS!nut_nombre)
'
'    vaSpread1.Col = 3
'    vaSpread1.Value = Trim(RS!nut_nomuni)
'
'    vaSpread1.Col = 4
'    vaSpread1.CellType = 10
'    vaSpread1.TypeCheckText = ""
'    vaSpread1.TypeCheckCenter = True
'    vaSpread1.text = Trim(Str(RS!nut_indpri))
'
'    vaSpread1.Col = 5
'    vaSpread1.Value = Trim(RS!nut_secnro)
'
'End If
'RS.Close
'Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
