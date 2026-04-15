VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_Receta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Recetas"
   ClientHeight    =   7485
   ClientLeft      =   2205
   ClientTop       =   2670
   ClientWidth     =   9750
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7485
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   165
      Left            =   360
      TabIndex        =   13
      Top             =   1110
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   291
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6975
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   9495
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   3960
         TabIndex        =   23
         Top             =   5880
         Width           =   1140
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   24
            Top             =   135
            Width           =   1035
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   1320
         TabIndex        =   21
         Top             =   5880
         Width           =   2580
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   22
            Top             =   135
            Width           =   2475
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   480
         TabIndex        =   19
         Top             =   5880
         Width           =   780
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   20
            Top             =   135
            Width           =   675
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Metodo Preparación"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5040
         TabIndex        =   12
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Fantasia"
         Height          =   615
         Index           =   1
         Left            =   4800
         TabIndex        =   11
         Top             =   6240
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Receta"
         Height          =   615
         Index           =   0
         Left            =   960
         TabIndex        =   10
         Top             =   6240
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Selección Recetas"
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   5175
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3735
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   4935
            _Version        =   393216
            _ExtentX        =   8705
            _ExtentY        =   6588
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
            BackColorStyle  =   1
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
            MaxCols         =   13
            MaxRows         =   20
            SpreadDesigner  =   "I_Receta.frx":0000
            ScrollBarTrack  =   3
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "I_Receta.frx":07B3
         Left            =   960
         List            =   "I_Receta.frx":07CC
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   360
         Width           =   3645
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Seleccción Nutrientes"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4215
         Index           =   1
         Left            =   5400
         TabIndex        =   2
         Top             =   1560
         Width           =   4050
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   3345
            ItemData        =   "I_Receta.frx":089C
            Left            =   240
            List            =   "I_Receta.frx":089E
            MultiSelect     =   1  'Simple
            TabIndex        =   5
            Top             =   360
            Width           =   3570
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   " P%| G%| CHO%"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   4
            Top             =   3720
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   "Código Prod."
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   2280
            TabIndex        =   3
            Top             =   3720
            Visible         =   0   'False
            Width           =   1455
         End
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Todos"
         Height          =   255
         Index           =   9
         Left            =   2040
         TabIndex        =   18
         Top             =   1200
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Todos"
         Height          =   255
         Index           =   8
         Left            =   2040
         TabIndex        =   17
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Categoria Dietetica"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Plato"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   15
         Top             =   1215
         Width           =   885
      End
      Begin VB.Label lblREGSEL 
         Caption         =   "lblREGSEL"
         Height          =   270
         Left            =   4890
         TabIndex        =   14
         Top             =   390
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informes"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   420
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9750
      _ExtentX        =   17198
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "I_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer, iselecc As Integer, imarca As Integer
Dim iaporte As Long, Est As Boolean
Dim cuenta As Long

Private Sub CheckeaListaCompleta(Indicador As Integer)

On Error GoTo Man_Error

Select Case Indicador

    Case 1
      
      For i = 1 To vaSpread1.MaxRows
       
          vaSpread1.Col = 1:   vaSpread1.Row = i
          vaSpread1.Value = IIf(vaSpread1.Value = "1", "1", "1")
          lblREGSEL.Caption = i & "  recetas seleccionadas"
      
      Next i
    
    Case 2
      
      For i = 1 To vaSpread1.MaxRows
       
          vaSpread1.Col = 1:   vaSpread1.Row = i
          vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "0")
          lblREGSEL.Caption = " 0 recetas seleccionadas"
      
      Next i

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Combo1_Click()

On Error GoTo Man_Error

If Combo1.ItemData(Combo1.ListIndex) = 4 Or Combo1.ItemData(Combo1.ListIndex) = 5 Then
   
   Check2(1).Enabled = False
   Check2(1).Value = 0
   Check2(2).Enabled = False
   Check2(2).Value = 0
   List1.Enabled = False
   CheckeaListaCompleta (1)

ElseIf Combo1.ItemData(Combo1.ListIndex) = 3 Then
   
   Check2(1).Enabled = False
   Check2(1).Value = 0
   Check2(2).Enabled = False
   Check2(2).Value = 0
   List1.Enabled = False
   CheckeaListaCompleta (1)

ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
   
   Check1.Enabled = False
   Check1.Value = 0
   Check2(1).Enabled = True
   Check2(2).Enabled = True
   List1.Enabled = True
   Frame1(1).Enabled = True
   CheckeaListaCompleta (2)

Else
   
   Check2(1).Enabled = False
   Check2(1).Value = 0
   Check2(2).Enabled = False
   Check2(2).Value = 0
   List1.Enabled = False
   Frame1(1).Enabled = False
   
   If Combo1.ItemData(Combo1.ListIndex) = 1 Then
      
      Check1.Enabled = True
      Check1.Value = 0
   
   Else
      
      Check1.Enabled = False
      Check1.Value = 0
   
   End If
   CheckeaListaCompleta (2)

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim codTippla As Long
Dim nomTippla As String

fg_centra Me
fg_carga ""
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
cuenta = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_listaprecio 8, 0, 0, '" & vg_NUsr & "'")
If Not RS.EOF Then
   
   vg_codlpr = RS!lpr_codigo
   vg_fecha = RS!dlp_anomes

Else
   
   vg_fecha = 0

End If
RS.Close
Set RS = Nothing

lblREGSEL.Caption = cuenta & " recetas seleccionadas"
MsgTitulo = "Impresión de Recetas"

'-------> Mover recetas
vaSpread1.MaxRows = 0
imarca = 0
iselecc = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_receta_V07 21, 0, '" & IIf(M_Receta.Check2.Value = 1, "x", "") & "', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe maestro recetas", vbExclamation + vbOKOnly, MsgTitulo
   Me.Hide
   Unload Me
   
End If

Est = True
codTippla = 0
nomTippla = ""
Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 2
    vaSpread1.CellType = CellTypeStaticText
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = RS!rec_codigo
      
    vaSpread1.Col = 3
    vaSpread1.CellType = CellTypeStaticText
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = Trim(IIf(IsNull(RS!rec_nombre), "", RS!rec_nombre))
      
    vaSpread1.Col = 4
    vaSpread1.CellType = CellTypeStaticText
    vaSpread1.TypeHAlign = TypeHAlignRight
    vaSpread1.text = IIf(IsNull(RS!rec_nomfan), "", Trim(RS!rec_nomfan))
    
    vaSpread1.Col = 5
    vaSpread1.text = Trim(Mid(RS!rec_catdie1, 1, Len(RS!rec_catdie1) - 1))
    
    vaSpread1.Col = 6
    vaSpread1.text = Trim(Mid(RS!rec_tippla1, 1, Len(RS!rec_tippla1) - 1))
    
    vaSpread1.Col = 7
    vaSpread1.text = IIf(IsNull(RS!rec_basrac), "", RS!rec_basrac)
    
    vaSpread1.Col = 8
    vaSpread1.text = IIf(IsNull(RS!rec_fecvig), "", RS!rec_fecvig)
    
    vaSpread1.Col = 9
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(RS!rec_indppr = 1, "Real", "Propuesta")
    
    vaSpread1.Col = 10
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = IIf(IsNull(RS!rec_canser) = True, 0, RS!rec_canser)
    
    vaSpread1.Col = 11
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = RS!ofertas_asoc
    
    vaSpread1.Col = 12
    vaSpread1.TypeHAlign = TypeHAlignLeft
    vaSpread1.text = RS!Estacionalidad_asoc


    RS.MoveNext
Loop
Est = False
RS.Close
Set RS = Nothing

vaSpread1.SortKey(1) = 3
vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow

'-------> Llenar Tabla Nutrientes
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
If RS.EOF Then

   RS.Close
   Set RS = Nothing
   fg_descarga
   MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo
   Me.Hide
   Unload Me

End If

List1.Clear
iaporte = 0

Do While Not RS.EOF
   
   List1.AddItem Trim(RS!nut_nombre)
   List1.ItemData(List1.NewIndex) = RS!nut_codigo
   If RS!nut_indpri = 1 Then List1.Selected(iaporte) = True
   iaporte = iaporte + 1
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

Combo1.ListIndex = 0
fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"
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

If Index = 2 Then
   
   TextDet2(3).text = ""
   TextDet2(4).text = ""

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""
   TextDet2(4).text = ""

ElseIf Index = 4 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 13
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4
    
    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 13
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 13
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 13
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 13
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 13
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
           
           vaSpread1.Col = 13
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

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i      As Long
Dim codigo As String
Dim nompro As String
Dim nomfan As String
Dim aAp    As String
Dim RS     As New ADODB.Recordset
Dim spid   As Integer

Select Case Button.Index
  
  Case 1
    
    '-----Crea Manejo Spid-----
    vg_db.Execute "DELETE paso_servicio WHERE ser_spid=@@spid and ser_usr='" & vg_NUsr & "'"
    isel = 0
    
    '-------> Buscar spid
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then
    
       spid = RS!spid
       
    End If
    RS.Close
    Set RS = Nothing
    
    vg_db.Execute "INSERT INTO paso_servicio (ser_spid, ser_usr, ser_codigo) VALUES (" & spid & ", '" & vg_NUsr & "', 0)"
    '-----Crea tabla temporal-----
    
    If Combo1.ListIndex = 3 Or Combo1.ListIndex = 4 Then
       
       vg_db.Execute "DELETE paso_receta WHERE rec_usr='" & vg_NUsr & "'"
        j = 0
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Col = 1
            vaSpread1.Row = i
            
            If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
               
               vaSpread1.Col = 2
               codpro = codpro & "'" & Trim(vaSpread1.text) & "',"
               
               vaSpread1.Col = 3
               nompro = Trim(vaSpread1.text)
               
               vaSpread1.Col = 4
               nomfan = Trim(vaSpread1.text)
               
               j = j + 1
               
               If j > 50 Then
                  
                  If (Combo1.ListIndex = 3) Or (Combo1.ListIndex = 4) Then
                     
                     vg_db.Execute "INSERT INTO paso_receta (rec_spid,rec_usr, rec_codigo)  " & _
                                   "SELECT '" & spid & "','" & vg_NUsr & "', rec_codigo FROM b_receta WHERE rec_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
                  
                  End If
                  
                  codpro = ""
                  j = 0
               
               End If
            
            End If
        
        Next i
        
        If Trim(codpro <> "") Then
           
           If (Combo1.ListIndex = 3) Or (Combo1.ListIndex = 4) Or (Combo1.ListIndex = 7) Then
              
              vg_db.Execute "INSERT INTO paso_receta (rec_spid, rec_usr, rec_codigo)  " & _
                            "SELECT '" & spid & "','" & vg_NUsr & "', rec_codigo FROM b_receta WHERE rec_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
           
           End If
           
           codpro = ""
        
        End If
    
    End If
    '----------------------------------------
    iselecc = 0
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        
        If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
        
           iselecc = 1
           Exit For
        
        End If
        
    Next i
    
    If iselecc = 0 Then
    
       MsgBox "Debe Seleccionar A lo menor una Receta", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
       
    End If
    
    Frame1(0).Enabled = False
    Toolbar1.Enabled = False
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       
       I_NombreRecetas cuenta
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       
       I_TarjetaRecetas cuenta
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
       
       iselecc = 0
       For i = 0 To List1.ListCount - 1
           
           If List1.Selected(i) = True Then iselecc = 1: Exit For
       
       Next i
       If iselecc = 0 Then
          
          MsgBox "Debe Seleccionar A lo Menos Un Aporte Nutricional", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
          
       End If
'       I_AporteRecetas cuenta
       ExportarExcelRecetaAporte
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 3 Then
       
       I_RecetasConProdCostoCero spid, vg_NUsr
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 4 Then
       
       I_ProductosCostoCero spid, vg_NUsr
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 5 Then
       
       I_IngredienteSinProductos spid, vg_NUsr
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 6 Then
       
       ExportarExcelEncabezadoReceta
          
    End If
    Frame1(0).Enabled = True
    Toolbar1.Enabled = True
  
  Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ExportarExcelEncabezadoReceta()

On Error GoTo Man_Error

Dim i               As Long
Dim RS              As New ADODB.Recordset
Dim MyBuffer        As String
Dim seleccion       As String
Dim CodRec          As Long
Dim NomArchivoExcel As String
Dim Extension       As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel

  '-------> Rescata Ceco Seleccionado
  seleccion = 0
  fg_carga ""
  
  Let MyBuffer = ""
  Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  Let MyBuffer = MyBuffer & "<R>"
  
  For i = 1 To vaSpread1.MaxRows
  
      vaSpread1.Row = i
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
      If seleccion = 1 And vaSpread1.RowHidden = False Then
          
          CodRec = 0
          vaSpread1.Col = 2
          CodRec = vaSpread1.text
                   
          MyBuffer = MyBuffer & " <RD"
          MyBuffer = MyBuffer & " R = " & Chr(34) & CodRec & Chr(34)
          MyBuffer = MyBuffer & "/>"
      
      
      End If
  
      DoEvents
       
  Next i

  MyBuffer = MyBuffer & "</R>"
      
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient

  Set RS = vg_db.Execute("sgpadm_Sel_XmlExportarExcelEncabezadoReceta_V03 '" & MyBuffer & "'")

  If Not RS.EOF Then
             
     If RS.RecordCount > 1020000 Then
      
        RS.Close
        Set RS = Nothing
      
        MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Recetas", vbCritical
        Exit Sub
   
     End If
             

    '-------> Guardar nombre archivo excel
    NomArchivoExcel = ""
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.Filter = "Todos los archivos *.xls,*.xlsx"
    On Error Resume Next
    CD.ShowSave
               
    '-------> JPAZ Permite controlar Boton Cancelar
    If Err.Number = 32755 Then
       
       MsgBox "Proceso cancelado"
       Exit Sub
    
    End If
                
    If CD.FileName = "" Then
       
       MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
       Exit Sub
    
    Else
       
       Extension = ""
       Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
       
       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
          Exit Sub
       End If
       
       NomArchivoExcel = CD.FileName
    
    End If
       
       '-------> Create an instance of Excel and add a workbook
       Set xlApp = CreateObject("Excel.Application")
       Set xlWb = xlApp.Workbooks.Add
       Set xlWs = xlWb.Worksheets("Hoja1")
  
       '-------> Display Excel and give user control of Excel's lifetime
       xlApp.UserControl = True
    
       '-------> Check version of Excel
       Call encabezado(RS, xlWs)
        
       xlWs.Cells(2, 1).CopyFromRecordset RS

       '-------> Auto-fit the column widths and row heights
'       xlApp.Selection.CurrentRegion.Columns.AutoFit
'       xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'       xlApp.Columns("A:B").Select
'       xlApp.Selection.Delete Shift:=xlToLeft
  
'       NomArchivoExcel = fg_ArchivoXls("ExportarExcel_EncabezadoReceta")
                    
       xlWb.Close True, NomArchivoExcel

''       Dim XL As New excel.Application 'Crea el objeto excel
       XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
       XL.Visible = True
       XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

       '-- Cerrar Excel
       xlApp.Quit
      
       '-------> Release Excel references
       Set xlWs = Nothing
       Set xlWb = Nothing
       Set xlApp = Nothing
          
       MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
                                                    
  End If
  
  RS.Close
  Set RS = Nothing

  fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ExportarExcelRecetaAporte()

On Error GoTo Man_Error

Dim i               As Long
Dim RS              As New ADODB.Recordset
Dim MyBuffer        As String
Dim MyBufferAporte  As String
Dim seleccion       As String
Dim CodRec          As Long
Dim Nutriente       As Long
Dim NomArchivoExcel As String
Dim Extension       As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel

  '-------> Rescata Ceco Seleccionado
  seleccion = 0
  fg_carga ""
  
  Let MyBuffer = ""
  Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  Let MyBuffer = MyBuffer & "<Recetas>"
  
  For i = 1 To vaSpread1.MaxRows
  
      vaSpread1.Row = i
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
      If seleccion = 1 And vaSpread1.RowHidden = False Then
          
          CodRec = 0
          vaSpread1.Col = 2
          CodRec = vaSpread1.text
                   
          MyBuffer = MyBuffer & " <Rec"
          MyBuffer = MyBuffer & " R = " & Chr(34) & CodRec & Chr(34)
          MyBuffer = MyBuffer & "/>"
      
      
      End If
  
      DoEvents
       
  Next i

  MyBuffer = MyBuffer & "</Recetas>"
      
      
  Let MyBufferAporte = ""
  Let MyBufferAporte = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  Let MyBufferAporte = MyBufferAporte & "<Nutrientes>"
      
  For i = 0 To List1.ListCount - 1
           
      If List1.Selected(i) = True Then
      
          Nutriente = List1.ItemData(i)
          MyBufferAporte = MyBufferAporte & " <Nut"
          MyBufferAporte = MyBufferAporte & " Nutriente = " & Chr(34) & Nutriente & Chr(34)
          MyBufferAporte = MyBufferAporte & "/>"
      
      End If
      
       
  Next i

  MyBufferAporte = MyBufferAporte & "</Nutrientes>"

  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient

  Set RS = vg_db.Execute("sgpadm_Sel_ExcelDetalleRecetaAportes '" & MyBuffer & "', '" & MyBufferAporte & "'")

  If Not RS.EOF Then
             
     If RS.RecordCount > 1020000 Then
      
        RS.Close
        Set RS = Nothing
      
        MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Recetas", vbCritical
        Exit Sub
   
     End If
             

    '-------> Guardar nombre archivo excel
    NomArchivoExcel = ""
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.Filter = "Todos los archivos *.xls,*.xlsx"
    On Error Resume Next
    CD.ShowSave
               
    '-------> JPAZ Permite controlar Boton Cancelar
    If Err.Number = 32755 Then
       
       MsgBox "Proceso cancelado"
       Exit Sub
    
    End If
                
    If CD.FileName = "" Then
       
       MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
       Exit Sub
    
    Else
       
       Extension = ""
       Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
       
       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
          Exit Sub
       End If
       
       NomArchivoExcel = CD.FileName
    
    End If
       
       '-------> Create an instance of Excel and add a workbook
       Set xlApp = CreateObject("Excel.Application")
       Set xlWb = xlApp.Workbooks.Add
       Set xlWs = xlWb.Worksheets("Hoja1")
  
       '-------> Display Excel and give user control of Excel's lifetime
       xlApp.UserControl = True
    
       '-------> Check version of Excel
       Call encabezado(RS, xlWs)
        
       xlWs.Cells(2, 1).CopyFromRecordset RS

       '-------> Auto-fit the column widths and row heights
'       xlApp.Selection.CurrentRegion.Columns.AutoFit
'       xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'       xlApp.Columns("A:B").Select
'       xlApp.Selection.Delete Shift:=xlToLeft
  
'       NomArchivoExcel = fg_ArchivoXls("ExportarExcel_EncabezadoReceta")
                    
       xlWb.Close True, NomArchivoExcel

''       Dim XL As New excel.Application 'Crea el objeto excel
       XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
       XL.Visible = True
       XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

       '-- Cerrar Excel
       xlApp.Quit
      
       '-------> Release Excel references
       Set xlWs = Nothing
       Set xlWb = Nothing
       Set xlApp = Nothing
          
       MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
                                                    
  End If
  
  RS.Close
  Set RS = Nothing

  fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ExportarExcelTarjetaReceta(ByVal usuario As String, ByVal spid As Long)

On Error GoTo Man_Error

Dim i               As Long
Dim RS              As New ADODB.Recordset
Dim NomArchivoExcel As String
Dim Extension       As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel

  '-------> Rescata Ceco Seleccionado
  fg_carga ""
      
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
     
  Set RS = vg_db.Execute("sgpadm_Sel_CostoDetRecetaOrgComprasTarjeta_V01 '" & fg_codigocbo(M_Receta.Combo3, 0, 4, "") & "', 1, " & Format(M_Receta.FpFecDesde, "yyyymmdd") & ",'" & usuario & "', " & spid & "")

  If Not RS.EOF Then
             
     If RS.RecordCount > 1020000 Then
      
        RS.Close
        Set RS = Nothing
      
        MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Recetas", vbCritical
        Exit Sub
   
     End If
             

    '-------> Guardar nombre archivo excel
    NomArchivoExcel = ""
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.Filter = "Todos los archivos *.xls,*.xlsx"
    On Error Resume Next
    CD.ShowSave
               
    '-------> JPAZ Permite controlar Boton Cancelar
    If Err.Number = 32755 Then
       
       MsgBox "Proceso cancelado"
       Exit Sub
    
    End If
                
    If CD.FileName = "" Then
       
       MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
       Exit Sub
    
    Else
       
       Extension = ""
       Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
       
       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
          Exit Sub
       End If
       
       NomArchivoExcel = CD.FileName
    
    End If
       
       '-------> Create an instance of Excel and add a workbook
       Set xlApp = CreateObject("Excel.Application")
       Set xlWb = xlApp.Workbooks.Add
       Set xlWs = xlWb.Worksheets("Hoja1")
  
       '-------> Display Excel and give user control of Excel's lifetime
       xlApp.UserControl = True
    
       '-------> Check version of Excel
       Call encabezado(RS, xlWs)
        
       xlWs.Cells(2, 1).CopyFromRecordset RS

       '-------> Auto-fit the column widths and row heights
'       xlApp.Selection.CurrentRegion.Columns.AutoFit
'       xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'       xlApp.Columns("A:B").Select
'       xlApp.Selection.Delete Shift:=xlToLeft
  
'       NomArchivoExcel = fg_ArchivoXls("ExportarExcel_EncabezadoReceta")
                    
       xlWb.Close True, NomArchivoExcel

''       Dim XL As New excel.Application 'Crea el objeto excel
       XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
       XL.Visible = True
       XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

       '-- Cerrar Excel
       xlApp.Quit
      
       '-------> Release Excel references
       Set xlWs = Nothing
       Set xlWb = Nothing
       Set xlApp = Nothing
          
       MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
                                                    
  End If
  
  RS.Close
  Set RS = Nothing

  fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

If Est Then Exit Sub
Dim i As Long
vaSpread1.Col = 1
Est = True

For i = BlockRow To BlockRow2
    
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")

Next

lblREGSEL.Caption = BlockRow2 - IIf(BlockRow = 1, 0, 1) & " recetas seleccionadas"

Est = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

cuenta = 0
vaSpread1.Row = Row
vaSpread1.Col = Col

If Row = -1 And vaSpread1.text = 0 Then
    
    cuenta = 0

ElseIf Row = -1 And vaSpread1.text = 1 And vaSpread1.RowHidden = False Then
    
    cuenta = vaSpread1.MaxRows

Else
   
   If vaSpread1.text = 1 And vaSpread1.RowHidden = False Then cuenta = cuenta + 1 Else cuenta = cuenta - 0

   
End If
lblREGSEL.Caption = cuenta & " recetas seleccionadas"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
