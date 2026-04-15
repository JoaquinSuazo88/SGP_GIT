VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_Produc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Productos"
   ClientHeight    =   6165
   ClientLeft      =   1860
   ClientTop       =   2220
   ClientWidth     =   8910
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
   ScaleHeight     =   6165
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5775
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   8895
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Fantasia"
         Height          =   375
         Index           =   1
         Left            =   4800
         TabIndex        =   9
         Top             =   5280
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre Producto"
         Height          =   375
         Index           =   0
         Left            =   960
         TabIndex        =   8
         Top             =   5280
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Selección Productos"
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   135
         TabIndex        =   7
         Top             =   960
         Width           =   4455
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3735
            Left            =   135
            TabIndex        =   12
            Top             =   330
            Width           =   4215
            _Version        =   393216
            _ExtentX        =   7435
            _ExtentY        =   6588
            _StockProps     =   64
            AllowMultiBlocks=   -1  'True
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
            MaxCols         =   5
            MaxRows         =   20
            SpreadDesigner  =   "I_Produc.frx":0000
            ScrollBarTrack  =   3
         End
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "I_Produc.frx":0512
         Left            =   1080
         List            =   "I_Produc.frx":0514
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   3795
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         Caption         =   "Seleccción Nutrientes"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   4215
         Index           =   1
         Left            =   4695
         TabIndex        =   2
         Top             =   960
         Width           =   4050
         Begin VB.CheckBox Check2 
            Appearance      =   0  'Flat
            Caption         =   " P%| G%| CHO%"
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   3720
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.ListBox List1 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   3345
            Left            =   240
            MultiSelect     =   1  'Simple
            TabIndex        =   4
            Top             =   360
            Width           =   3570
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
      Begin VB.CheckBox Check1 
         Caption         =   "Metodo Preparación"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5580
         TabIndex        =   11
         Top             =   375
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblREGSEL 
         Caption         =   "lblREGSEL"
         Height          =   270
         Left            =   5190
         TabIndex        =   13
         Top             =   450
         Width           =   2565
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Informes"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   435
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_Produc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim i As Integer, iselecc As Integer
Dim Opx As String, cuenta As Long

Private Sub Combo1_Click()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If Combo1.ListIndex = 0 Or Combo1.ListIndex = 3 Then
    
    Check2(1).Enabled = False
    Check2(1).Value = 0
    Check2(2).Enabled = False
    Check2(2).Value = 0
    List1.Clear
    List1.Enabled = False
    Frame1(1).Enabled = False
    Frame1(1).Caption = ""
    Check1.Enabled = False
    Check1.Value = 0
    Option1(0).Visible = IIf(Opx = "P", False, True)
    Option1(1).Visible = IIf(Opx = "P", False, True)

ElseIf Combo1.ListIndex = 2 And Opx = "P" Then
    
    Dim impuesto As Long
    '------- Llenar Tabla Impuesto
    RS.Open "sgpadm_s_impuesto 2, 0, ''", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro impuestos", vbExclamation + vbOKOnly, MsgTitulo: Me.Hide: Unload Me
    List1.Clear: impuesto = 0
    
    Do While Not RS.EOF
        List1.AddItem Trim(RS!imp_nombre)
        List1.ItemData(List1.NewIndex) = RS!imp_codigo
        List1.Selected(impuesto) = True
        impuesto = impuesto + 1
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    
    '******************************** '
    Check2(1).Enabled = False
    Check2(1).Value = 0
    Check2(2).Enabled = False
    Check2(2).Value = 0
    List1.Enabled = True
    Frame1(1).Caption = "Selección Impuesto"
    Frame1(1).Enabled = True
    Check1.Enabled = True
    Check1.Value = 0

ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
    
    Dim iaporte As Long
    '------- Llenar Tabla Nutrientes
    RS.Open "sgpadm_s_nutriente 1, 0, ''", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fg_descarga: MsgBox "No existe maestro nutrientes", vbExclamation + vbOKOnly, MsgTitulo: Me.Hide: Unload Me
    List1.Clear: iaporte = 0
    Do While Not RS.EOF
        List1.AddItem Trim(RS!nut_nombre)
        List1.ItemData(List1.NewIndex) = RS!nut_codigo
        If RS!nut_indpri = 1 Then List1.Selected(iaporte) = True
        iaporte = iaporte + 1
        RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    '******************************** '
    Check1.Enabled = False
    Check1.Value = 0
    Check2(1).Enabled = True
    Check2(2).Enabled = True
    List1.Enabled = True
    Frame1(1).Enabled = True
    Frame1(1).Caption = "Selección Nutrientes"
    
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

fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
cuenta = 0
lblREGSEL.Caption = cuenta & " registros seleccionados"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Productos"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i As Long, codpro As String, nompro As String, nomfan As String, aAp As String
Dim RS2 As New ADODB.Recordset
Select Case Button.Index

Case 1
    
    fg_carga ""
    '-----Crea tabla temporal-----
    If Combo1.ListIndex = 0 Or Combo1.ListIndex = 2 Or Combo1.ListIndex = 3 Or (Combo1.ListIndex = 1 Or Opx = "I") Then
       
       vg_db.Execute "DELETE tmp_proding WHERE tem_usuario='" & vg_NUsr & "'"
        j = 0
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Col = 1: vaSpread1.Row = i
            If vaSpread1.text = "1" Then
               vaSpread1.Col = 2: codpro = codpro & "'" & Trim(vaSpread1.text) & "',"
               vaSpread1.Col = 3: nompro = Trim(vaSpread1.text)
               vaSpread1.Col = 4: nomfan = Trim(vaSpread1.text)
               j = j + 1
               If j > 50 Then
                  If (Combo1.ListIndex = 0 Or Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Or Combo1.ListIndex = 3) And Opx = "P" Then
                     vg_db.Execute "INSERT INTO tmp_proding (tem_usuario, tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                                   "SELECT '" & vg_NUsr & "', pro_codigo, pro_nombre, '" & nomfan & "', '0' FROM b_productos WHERE pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
                  Else
                     vg_db.Execute "INSERT INTO tmp_proding (tem_usuario, tem_codigo, tem_nombre,tem_nomfan, tem_codpat) " & _
                                   "SELECT '" & vg_NUsr & "', ing_codigo, ing_nombre, ing_nomfan, '0' FROM b_ingrediente WHERE ing_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
                  End If
                  If Combo1.ListIndex = 2 And Opx = "P" Then
                     vg_db.Execute "INSERT INTO tmp_proding  (tem_usuario, tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                                   "SELECT '" & vg_NUsr & "', ipr_codimp, 'x', '', ipr_codpro " & _
                                   "FROM b_productosimp WHERE ipr_codpro IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
                  ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
                     vg_db.Execute "INSERT INTO tmp_proding  (tem_usuario, tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                                   "SELECT '" & vg_NUsr & "', pnu_codapo, pnu_canapo, '', '" & pnu_codpro & "' " & _
                                   "FROM b_productonut WHERE pnu_codpro IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
                  End If
                  codpro = "": j = 0
               End If
            End If
        Next i
        If Trim(codpro <> "") Then
           
           If (Combo1.ListIndex = 0 Or Combo1.ListIndex = 1 Or Combo1.ListIndex = 2 Or Combo1.ListIndex = 3) And Opx = "P" Then
              vg_db.Execute "INSERT INTO tmp_proding (tem_usuario, tem_codigo, tem_nombre,tem_nomfan, tem_codpat) " & _
                            "SELECT '" & vg_NUsr & "', pro_codigo, pro_nombre, '" & nonfan & "', '0' FROM b_productos WHERE pro_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
           ElseIf Combo1.ListIndex = 2 Or (Combo1.ListIndex = 1 Or Opx = "I") Then
              vg_db.Execute "INSERT INTO tmp_proding (tem_usuario, tem_codigo, tem_nombre,tem_nomfan, tem_codpat) " & _
                            "SELECT '" & vg_NUsr & "', ing_codigo, ing_nombre, ing_nomfan, '0' FROM b_ingrediente WHERE ing_codigo IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
           End If
           If Combo1.ListIndex = 2 And Opx = "P" Then
              vg_db.Execute "INSERT INTO tmp_proding (tem_usuario, tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                            "SELECT '" & vg_NUsr & "', ipr_codimp, 'x', '', ipr_codpro " & _
                            "FROM b_productosimp WHERE ipr_codpro IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
           ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
              vg_db.Execute "INSERT INTO tmp_proding (tem_usuario, tem_codigo, tem_nombre, tem_nomfan, tem_codpat) " & _
                            "SELECT '" & vg_NUsr & "', pnu_codapo, pnu_canapo, '', '" & pnu_codpro & "' " & _
                            "FROM b_productonut WHERE pnu_codpro IN (" & Mid(codpro, 1, Len(codpro) - 1) & ")"
           End If
           codpro = ""
        
        End If
    
    End If
    '----------------------------------------
    iselecc = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then iselecc = 1: Exit For
    Next i
    If iselecc = 0 Then MsgBox "Debe seleccionar a lo menor un producto", vbExclamation + vbOKOnly, MsgTitulo: fg_descarga: Exit Sub
    If Combo1.ListIndex = 0 Then
       If Opx = "P" Then
          I_Productos1
       ElseIf Opx = "I" Then
          i_ingrediente_excel
       ElseIf Opx = "C" Then
          I_FormatoCompra
       End If
    ElseIf Combo1.ListIndex = 1 And Opx = "P" Then
       I_Productos2
    ElseIf Combo1.ListIndex = 2 And Opx = "P" Then
       iselecc = 0
       For i = 0 To List1.ListCount - 1
           If List1.Selected(i) = True Then iselecc = 1: Exit For
       Next i
       If iselecc = 0 Then MsgBox "Debe seleccionar a lo menos un Impuesto", vbExclamation + vbOKOnly, MsgTitulo: fg_descarga: Exit Sub
       I_ImpuestoProductos
    ElseIf Combo1.ListIndex = 3 And Opx = "P" Then
       I_IngredientesProductos
    ElseIf Combo1.ListIndex = 1 And Opx = "I" Then
       iselecc = 0
       For i = 0 To List1.ListCount - 1
           If List1.Selected(i) = True Then iselecc = 1: Exit For
       Next i
       If iselecc = 0 Then MsgBox "Debe seleccionar a lo menos un Aporte Nutricional", vbExclamation + vbOKOnly, MsgTitulo: fg_descarga: Exit Sub
       I_AporteProductos
    End If
    fg_descarga

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub i_ingrediente_excel()

Dim RS As New ADODB.Recordset
Dim NomArchivoExcel As String
Dim Extension       As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel


On Error GoTo Man_Error

  fg_carga ""
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient

  Set RS = vg_db.Execute("sgpadm_s_ingrediente_V02 18, 0, '" & vg_NUsr & "'")

  If Not RS.EOF Then
             
     If RS.RecordCount > 1020000 Then
      
        RS.Close
        Set RS = Nothing
      
        MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ingredientes", vbCritical
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
       xlWb.Close True, NomArchivoExcel

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

Dim i As Long
vaSpread1.Col = 1
For i = BlockRow To BlockRow2
    
    vaSpread1.Row = i
    vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")

Next

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

vaSpread1.Row = Row: vaSpread1.Col = Col
If Row = -1 And vaSpread1.text = 0 Then
    cuenta = 0
ElseIf Row = -1 And vaSpread1.text = 1 Then
    cuenta = vaSpread1.MaxRows
Else
    If vaSpread1.text = 1 Then cuenta = cuenta + 1 Else cuenta = cuenta - 1
End If
lblREGSEL.Caption = cuenta & " registros seleccionados"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
'If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.Value = IIf(vaSpread1.Value = "1", "", "1")
End Sub

Sub TraspasoGrilla(vaSrepadX As vaSpread, op As String)

On Error GoTo Man_Error

Dim i As Long, j As Long
Dim vecfamprod() As Variant
fg_carga ""
Opx = op

If Opx = "P" Then
    
    Combo1.Clear
    Combo1.AddItem "Productos Formato 1"
    Combo1.AddItem "Productos Formato 2"
    Combo1.AddItem "Impuesto Productos"
    Combo1.AddItem "Ingredientes & Productos"
    MsgTitulo = "Impresión de Productos"
    Me.Caption = "Imprimir Productos"
    '------- traer numero registro familia productos
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "sgpadm_Sel_productos 12, '', '', '" & vg_NUsr & "'", vg_db, adOpenForwardOnly
    If RS1.EOF Or RS1!nReg < 1 Then RS1.Close: Set RS1 = Nothing: Exit Sub
    ReDim vecfamprod(RS1!nReg, 3)
    RS1.Close: Set RS1 = Nothing
    RS.Open "sgpadm_p_filtrarfamproducto", vg_db, adOpenForwardOnly
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    i = 1
    Do While Not RS.EOF
       vecfamprod(i, 1) = Trim(RS!pro_codtip)
       If Trim(RS!tip_nombre) = "" Then
          
          vecfamprod(i, 2) = ""
       
       Else
          
          vecfamprod(i, 2) = Mid(RS!tip_nombre, 1, Len(RS!tip_nombre) - 1) 'Mid(RS(0), 1, Len(RS(0)) - 1) 'fg_BuscaenArbol(RS!pro_codtip, "a_tipopro", "tip_codigo")
          vecfamprod(i, 3) = IIf(RS!pro_indppr = 1, "Real", "Propuesta")
       
       End If
       RS.MoveNext:    i = i + 1
    Loop
    RS.Close: Set RS = Nothing
    
    vaSpread1.MaxRows = vaSrepadX.MaxRows
    For i = 1 To vaSrepadX.MaxRows
        
        vaSpread1.Row = i: vaSrepadX.Row = i
        vaSpread1.Col = 2: vaSrepadX.Col = 1
        vaSpread1.TypeHAlign = 1: vaSpread1.CellType = 5
        vaSpread1.Value = vaSrepadX.Value
        vaSpread1.Col = 3: vaSrepadX.Col = 2
        vaSpread1.TypeHAlign = 0: vaSpread1.CellType = 5
        vaSpread1.Value = vaSrepadX.Value
        vaSpread1.Col = 5: vaSrepadX.Col = 6 ' Se posiciona en el tipo Real o Propuesta
        vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.CellType = 5
        vaSpread1.Value = vaSrepadX.Value  ' Retorna valor
        vaSpread1.Col = 4: vaSrepadX.Col = 5
        For j = 1 To UBound(vecfamprod)
            
            If vecfamprod(j, 1) = vaSrepadX.text Then vaSpread1.text = vecfamprod(j, 2): Exit For ': vaSpread1.Col = 5: vaSpread1.text = vecfamprod(j, 3)
        
        Next j
    
    Next i

ElseIf Opx = "I" Then
    
    Combo1.Clear
    Combo1.AddItem "Ingredientes"
    Combo1.AddItem "Aportes Nutricionales Ingredientes"
    MsgTitulo = "Impresión de Ingredientes"
    Me.Caption = "Imprimir Ingredientes"
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "sgpadm_s_ingrediente_V02 8, '', ''", vg_db, adOpenForwardOnly ', adOpenStatic
    i = 1
    
    Do While Not RS1.EOF
        
        vaSpread1.MaxRows = i: vaSpread1.Row = i
        vaSpread1.Col = 2: vaSpread1.TypeHAlign = 1: vaSpread1.CellType = 5
        vaSpread1.Value = RS1!ing_codigo
        vaSpread1.Col = 3: vaSpread1.TypeHAlign = 0: vaSpread1.CellType = 5
        vaSpread1.Value = RS1!ing_nombre
        vaSpread1.Col = 4: vaSpread1.Value = RS1!ing_nomfan
        vaSpread1.Col = 5: vaSpread1.Value = IIf(RS1!ing_indppr = "1", "Real", "Propuesta")
        RS1.MoveNext: i = i + 1
    
    Loop
    RS1.Close: Set RS1 = Nothing

ElseIf Opx = "C" Then
    
    Combo1.Clear
    Combo1.AddItem "Productos Compras"
    MsgTitulo = "Impresión Productos Compras"
    Me.Caption = "Imprimir Productos Compras"
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT pco_codigo, pco_nombre FROM b_productocompra ORDER BY pco_codigo", vg_db, adOpenForwardOnly ', adOpenStatic
    i = 1
    Do While Not RS1.EOF
        
        vaSpread1.MaxRows = i: vaSpread1.Row = i
        vaSpread1.Col = 2: vaSpread1.TypeHAlign = 1: vaSpread1.CellType = 5
        vaSpread1.Value = RS1!pco_codigo
        vaSpread1.Col = 3: vaSpread1.TypeHAlign = 0: vaSpread1.CellType = 5
        vaSpread1.Value = RS1!pco_nombre
        vaSpread1.Col = 4: vaSpread1.Value = ""
        
        RS1.MoveNext: i = i + 1
    
    Loop
    RS1.Close: Set RS1 = Nothing

End If
Combo1.ListIndex = 0
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
