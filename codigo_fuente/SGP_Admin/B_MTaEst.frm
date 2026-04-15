VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_MTaEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5685
   ClientLeft      =   1815
   ClientTop       =   1830
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5685
   ScaleWidth      =   5505
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      ItemData        =   "B_MTaEst.frx":0000
      Left            =   120
      List            =   "B_MTaEst.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   8
      Top             =   3600
      Width           =   4635
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4935
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         TabIndex        =   3
         Top             =   555
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "B_MTaEst.frx":0004
         Left            =   1710
         List            =   "B_MTaEst.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Texto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   5
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Columna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   4
         Top             =   345
         Width           =   1440
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      Index           =   1
      ItemData        =   "B_MTaEst.frx":0022
      Left            =   720
      List            =   "B_MTaEst.frx":0024
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2220
      Left            =   0
      TabIndex        =   6
      Top             =   1080
      Width           =   4935
      _Version        =   393216
      _ExtentX        =   8705
      _ExtentY        =   3916
      _StockProps     =   64
      AllowMultiBlocks=   -1  'True
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
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
      MaxRows         =   13
      ScrollBars      =   2
      SpreadDesigner  =   "B_MTaEst.frx":0026
      StartingColNumber=   6
      ScrollBarTrack  =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5685
      Left            =   4965
      TabIndex        =   7
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   10028
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   0
      TabIndex        =   9
      Top             =   3360
      Width           =   4935
   End
End
Attribute VB_Name = "B_MTaEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim i As Long, j As Long, ibusca As Long, codigo As Long
Dim iCombo As Integer
Dim FindString  As String, SourceString As String
Dim prog As Object
Dim op As String

Private Sub Combo1_Click()

On Error GoTo Man_Error

If iCombo = 0 Then Text1.text = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    Cerrar

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
Me.Left = vg_left
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
iCombo = 1
Combo1.ListIndex = 1

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Sub LlenaDatos(titgen As String, programa As Object, subseg As Long, codReg As Long, codzon As Long, FecIni As Long, FecFin As Long, opcion As String, Indppr As String)

On Error GoTo Man_Error

Dim RS1       As New ADODB.Recordset
Dim i         As Long
Dim codCeco   As String
Dim seleccion As Integer
Dim Sql       As String

fg_carga ""
Me.Caption = titgen
Set prog = programa
op = opcion

With vaSpread1
    
    .MaxRows = 0
    List1(0).Clear: List1(1).Clear
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If opcion = "1" Then
        
        Set RS1 = vg_db.Execute("sgpadm_s_planifminuta 7, " & subseg & ", " & codReg & "," & codzon & ", 0, 0, " & FecIni & ", " & FecFin & "," & Indppr & "")
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fg_descarga: Exit Sub
    
    ElseIf opcion = "2" Then
       
       Set RS1 = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
    
    ElseIf opcion = "3" Then
       
       If titgen = "Regimen" Then
          
          RS1.Open "SELECT DISTINCT a.reg_codigo, a.reg_nombre FROM cas_a_regimen a, cas_log_regenviominutasitioremoto b WHERE a.reg_cecori = b.cecori AND a.reg_codigo = b.codreg AND a.reg_cecori = '" & vg_codigo & "' ORDER BY reg_codigo", vg_db, adOpenStatic
        
        Else
          
          RS1.Open "SELECT DISTINCT a.ser_codigo, a.ser_nombre, a.ser_orden FROM cas_a_servicio a, cas_log_regenviominutasitioremoto b WHERE  a.ser_cecori = b.cecori AND a.ser_codigo = b.codser AND a.ser_cecori = '" & vg_codigo & "' AND ser_activo = '1' ORDER BY ser_orden, ser_codigo", vg_db, adOpenStatic
        
        End If
    
    ElseIf opcion = "4" Then
        
        Set RS1 = vg_db.Execute("sgpadm_s_servicio 10, '', 0, 0")
    
    ElseIf opcion = "5" Then
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioMinutaBloque '" & vg_codigo & "', " & codReg & ", " & FecIni & ", " & FecFin & "")
    
    ElseIf opcion = "6" Then
    
            Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioMinutaBloqueMes '" & vg_codigo & "', " & codReg & ", " & FecIni & "")
        
    ElseIf opcion = "7" Then
    
        '--> Concatenar codigo ceco
        codCeco = ""

        For i = 1 To programa.vaSpread1.MaxRows
       
            programa.vaSpread1.Row = i
            programa.vaSpread1.Col = 1 'Seleccion
            seleccion = IIf(programa.vaSpread1.text = "", 0, programa.vaSpread1.text)
    
            If seleccion = 1 And programa.vaSpread1.RowHidden = False Then

               programa.vaSpread1.Col = 2
               codCeco = codCeco & "'" & programa.vaSpread1.text & "', "

            End If
  
        Next i

        Sql = ""
        If Trim(codCeco) <> "" Then
   
            Sql = Sql & Replace(Mid(codCeco, 1, Len(codCeco) - 2), "'", """")

        End If

        Set RS1 = vg_db.Execute("sgpadm_Sel_ServiciosClientes '" & Sql & "', '" & Format(programa.FpFecDesde.Value, "yyyymmdd") & "', '" & Format(programa.FpFecHasta.Value, "yyyymmdd") & "'")
    
    ElseIf opcion = "8" Then
    
        Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioMinutaBloqueOrder '" & Indppr & "', " & codReg & ", " & FecIni & ", " & FecFin & "")
        
    End If
    
    Do While Not RS1.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       
    If opcion = "7" Then
    
       programa.vaSpread3.Row = .MaxRows
       programa.vaSpread3.Col = 1
       If programa.vaSpread3.text = "1" Then
          .Col = 1
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = "1" ' checked
          If opcion = "1" Then
             List1(0).AddItem RS1!min_codser & " " & Trim(RS1!ser_nombre)
             List1(0).ItemData(List1(0).NewIndex) = RS1!min_codser
             List1(1).AddItem RS1!min_codser & " " & Trim(RS1!ser_nombre)
             List1(1).ItemData(List1(1).NewIndex) = RS1!min_codser
          ElseIf opcion = "2" Then
             List1(0).AddItem RS1!nut_codigo & " " & Trim(RS1!nut_nombre) & " " & Trim(RS1!nut_nomuni)
             List1(0).ItemData(List1(0).NewIndex) = RS1!nut_codigo
             List1(1).AddItem RS1!nut_codigo & " " & Trim(RS1!nut_nombre) & " " & Trim(RS1!nut_nomuni)
             List1(1).ItemData(List1(1).NewIndex) = RS1!nut_codigo
          ElseIf opcion = "7" Or opcion = "8" Then
             List1(0).AddItem RS1(0) & " " & Trim(RS1(1))
             List1(0).ItemData(List1(0).NewIndex) = RS1(0)
             List1(1).AddItem RS1(0) & " " & Trim(RS1(1))
             List1(1).ItemData(List1(1).NewIndex) = RS1(0)
          End If
       ElseIf programa.vaSpread3.text = "0" Or programa.vaSpread3.text = "" Then
          .Col = 1
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = "" ' checked
       End If
    
    Else
       programa.Row = .MaxRows
       programa.Col = 1
       If programa.text = "1" Then
          .Col = 1
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = "1" ' checked
          If opcion = "1" Then
             List1(0).AddItem RS1!min_codser & " " & Trim(RS1!ser_nombre)
             List1(0).ItemData(List1(0).NewIndex) = RS1!min_codser
             List1(1).AddItem RS1!min_codser & " " & Trim(RS1!ser_nombre)
             List1(1).ItemData(List1(1).NewIndex) = RS1!min_codser
          ElseIf opcion = "2" Then
             List1(0).AddItem RS1!nut_codigo & " " & Trim(RS1!nut_nombre) & " " & Trim(RS1!nut_nomuni)
             List1(0).ItemData(List1(0).NewIndex) = RS1!nut_codigo
             List1(1).AddItem RS1!nut_codigo & " " & Trim(RS1!nut_nombre) & " " & Trim(RS1!nut_nomuni)
             List1(1).ItemData(List1(1).NewIndex) = RS1!nut_codigo
          End If
       ElseIf programa.text = "" Then
          .Col = 1
          .CellType = 10
          .TypeCheckText = ""
          .TypeCheckCenter = True
          .text = "" ' checked
       End If
       
    End If
    
       If opcion = "1" Then
          
          .Col = 2
          .text = RS1!min_codser
          .Col = 3
          .text = Trim(RS1!ser_nombre)
       
       ElseIf opcion = "2" Then
          
          .Col = 2
          .text = RS1!nut_codigo
          .Col = 3
          .text = Trim(RS1!nut_nombre) & " " & Trim(RS1!nut_nomuni)
       
       ElseIf opcion = "3" Or opcion = "4" Or opcion = "5" Or opcion = "6" Or opcion = "7" Or opcion = "8" Then
          
          .Col = 2
          .text = RS1(0)
          .Col = 3
          .text = Trim(RS1(1))
       
       End If
       
       RS1.MoveNext
    
    Loop
    
    RS1.Close
    Set RS1 = Nothing
    fg_descarga
    
    iCombo = 0
    
End With

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub List1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    Cerrar

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Text1_Change()

On Error GoTo Man_Error

With vaSpread1

    If .MaxRows < 1 Then Exit Sub
    If LimpiaDato(Trim(Text1.text)) & Chr(KeyAscii) = "" Then Exit Sub
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       FindString = Text1.text
       For i = 1 To .MaxRows
           .Row = i
           .Col = 2
           SourceString = Trim(.text)
           indactivo = UCase(Trim(SourceString)) Like "*" & UCase(FindString) & "*"
           If indactivo = -1 Then
              If .RowHidden = True Then .RowHidden = False
           Else
              If .RowHidden = False Then .RowHidden = True
           End If
       Next i
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       FindString = Text1.text
       For i = 1 To .MaxRows
           .Row = i
           .Col = 3
           SourceString = Trim(.text)
           indactivo = UCase(Trim(SourceString)) Like "*" & UCase(FindString) & "*"
           If indactivo = -1 Then
              If .RowHidden = True Then .RowHidden = False
           Else
              If .RowHidden = False Then .RowHidden = True
           End If
       Next i
    End If
    
End With

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    Cerrar

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    MoverDatos

Case 3
    
    Cerrar

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Dim i As Long
Dim X As Long
    
Select Case BlockCol

'Case 1
'
'    vaSpread1.Col = 1
'
'    For i = BlockRow To BlockRow2
'
'        vaSpread1.Row = i
'
'        If vaSpread1.RowHidden = False Then
'
'           vaSpread1.Col = 1
'           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
'
'           If vaSpread1.Value = "1" Then
'
'              vaSpread1.Col = 2
'              codigo = Val(vaSpread1.text)
'              vaSpread1.Col = 3
'
'              List1(0).AddItem codigo & " " & Trim(vaSpread1.text)
'              List1(0).ItemData(List1(0).NewIndex) = codigo
'              List1(1).AddItem codigo & " " & Trim(vaSpread1.text)
'              List1(1).ItemData(List1(1).NewIndex) = codigo
'
'
'           Else
'
'              For x = 0 To List1(1).ListCount - 1
'
'                  List1(1).ListIndex = x
'                  vaSpread1.Col = 2
'                  codigo = Val(vaSpread1.text)
'
'                  If List1(1).ItemData(List1(1).ListIndex) = codigo Then
'
'                     List1(0).RemoveItem x
'                     List1(1).RemoveItem x
'                     Exit For
'
'                  End If
'
'              Next x
'
'           End If
'           iCombo = 0
'
'        End If
'
'    Next
'
'    For i = 1 To vaSpread1.MaxRows
'
'        vaSpread1.Row = i
'
'        If vaSpread1.RowHidden = False Then
'
'           vaSpread1.Col = 1
'           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
'
'           If vaSpread1.Value = "1" Then
'
'              vaSpread1.Col = 2
'              codigo = Val(vaSpread1.text)
'              vaSpread1.Col = 3
'
'              List1(0).AddItem codigo & " " & Trim(vaSpread1.text)
'              List1(0).ItemData(List1(0).NewIndex) = codigo
'              List1(1).AddItem codigo & " " & Trim(vaSpread1.text)
'              List1(1).ItemData(List1(1).NewIndex) = codigo
'
'
'           Else
'
'              For x = 0 To List1(1).ListCount - 1
'
'                  List1(1).ListIndex = x
'                  vaSpread1.Col = 2
'                  codigo = Val(vaSpread1.text)
'
'                  If List1(1).ItemData(List1(1).ListIndex) = codigo Then
'
'                     List1(0).RemoveItem x
'                     List1(1).RemoveItem x
'                     Exit For
'
'                  End If
'
'              Next x
'
'           End If
'           iCombo = 0
'
'        End If
'
'    Next
    
Case Is <> 1

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           iCombo = 1
           vaSpread1.Col = 1
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
           
           If vaSpread1.Value = "1" Then
           
              vaSpread1.Col = 2
              codigo = Val(vaSpread1.text)
              vaSpread1.Col = 3
           
              List1(0).AddItem codigo & " " & Trim(vaSpread1.text)
              List1(0).ItemData(List1(0).NewIndex) = codigo
              List1(1).AddItem codigo & " " & Trim(vaSpread1.text)
              List1(1).ItemData(List1(1).NewIndex) = codigo
           
           
           Else
           
              For X = 0 To List1(1).ListCount - 1
               
                  List1(1).ListIndex = X
                  vaSpread1.Col = 2
                  codigo = Val(vaSpread1.text)
               
                  If List1(1).ItemData(List1(1).ListIndex) = codigo Then
                  
                     List1(0).RemoveItem X
                     List1(1).RemoveItem X
                     Exit For
               
                  End If
           
              Next X
           
           End If
           iCombo = 0
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If iCombo = 1 Or Row < 1 Then Exit Sub

With vaSpread1
    
    Select Case Col
      
      Case 1
        
        .Row = .ActiveRow
        .Col = 1
        
        If .text = "1" Then ' checked
           
           .Col = 2
           codigo = Val(.text)
           .Col = 3
           List1(0).AddItem codigo & " " & Trim(.text)
           List1(0).ItemData(List1(0).NewIndex) = codigo
           List1(1).AddItem codigo & " " & Trim(.text)
           List1(1).ItemData(List1(1).NewIndex) = codigo
        
        ElseIf .text = "0" Then
           
           iCombo = 1
           .Col = 1
           .CellType = 10
           .TypeCheckText = ""
           .TypeCheckCenter = True
           .text = "0" ' checked
           iCombo = 0
           
           For i = 0 To List1(1).ListCount - 1
               
               List1(1).ListIndex = i
               .Col = 2
               codigo = Val(.text)
               
               If List1(1).ItemData(List1(1).ListIndex) = codigo Then
                  
                  List1(0).RemoveItem i
                  List1(1).RemoveItem i
                  Exit Sub
               
               End If
           
           Next i
        
        End If
    
    End Select
    
End With

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Col = 1 And Row = 0 Then
   
   Est = True
   iCombo = 1
   vaSpread1.Row = -1
   vaSpread1.Col = 1
   vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")
   
   Est = IIf(vaSpread1.Value = "1", True, False)
   
   If Not Est Then
      
      est1 = False
      List1(0).Clear: List1(1).Clear
   
   Else
      
      If est1 Then iCombo = 0: Exit Sub
      List1(0).Clear
      List1(1).Clear
      
      For i = 1 To vaSpread1.MaxRows
          
          vaSpread1.Row = i
          vaSpread1.Col = 2
          codigo = vaSpread1.text
          vaSpread1.Col = 3
          List1(0).AddItem codigo & Space(5) & Trim(vaSpread1.text)
          List1(0).ItemData(List1(0).NewIndex) = codigo
          List1(1).AddItem codigo & vaSpread1.text & Space(150) & "(" & fg_pone_espacio(CStr(codigo), 10) & ")" '& codigo
          List1(1).ItemData(List1(1).NewIndex) = codigo
      
      Next i
      
   End If
   iCombo = 0

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Or Row < 1 Then Exit Sub
MoverDatos

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 27
    
    Cerrar

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub MoverDatos()

On Error GoTo Man_Error

With vaSpread1

    If .MaxRows < 1 Then Exit Sub
    
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If .text = "1" And op = "7" Then
           
           prog.vaSpread3.Row = i
           prog.vaSpread3.Col = 1
           prog.vaSpread3.CellType = 10
           prog.vaSpread3.TypeCheckText = ""
           prog.vaSpread3.TypeCheckCenter = True
           prog.vaSpread3.text = "1" ' checked
        
        ElseIf .text = "1" And op <> "7" Then
           
           prog.Row = i
           prog.Col = 1
           prog.CellType = 10
           prog.TypeCheckText = ""
           prog.TypeCheckCenter = True
           prog.text = "1" ' checked
        
        ElseIf (.text = "" Or .text = "0") And op = "7" Then
           
           prog.vaSpread3.Row = i
           prog.vaSpread3.Col = 1
           prog.vaSpread3.CellType = 10
           prog.vaSpread3.TypeCheckText = ""
           prog.vaSpread3.TypeCheckCenter = True
           prog.vaSpread3.text = "" ' checked
        ElseIf (.text = "" Or .text = "0") And op <> "7" Then
           prog.Row = i
           prog.Col = 1
           prog.CellType = 10
           prog.TypeCheckText = ""
           prog.TypeCheckCenter = True
           prog.text = "" ' checked
        End If
    
    Next i
    
    .Row = .ActiveRow
    .Col = 1
    vg_codigo = Val(.text)

End With

Cerrar

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub Cerrar()

On Error GoTo Man_Error

Me.Hide
Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
