VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form B_BusVas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Recetas o Ingredientes en Planificación"
   ClientHeight    =   6465
   ClientLeft      =   6360
   ClientTop       =   2280
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar &Siguiente"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Criterio"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   3375
      Begin VB.Frame Frame2 
         Caption         =   "Por"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton Option2 
            Caption         =   "Ingrediente"
            Height          =   255
            Left            =   1560
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Receta"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin EditLib.fpText Text1 
         Height          =   315
         Left            =   840
         TabIndex        =   3
         Top             =   1080
         Width           =   2340
         _Version        =   196608
         _ExtentX        =   4128
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   0
         ThreeDOutsideHighlightColor=   16777215
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   1
         BorderColor     =   -2147483642
         BorderWidth     =   1
         ButtonDisable   =   0   'False
         ButtonHide      =   0   'False
         ButtonIncrement =   1
         ButtonMin       =   0
         ButtonMax       =   100
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   16777215
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   50
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label2 
         Caption         =   "Buscar :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2655
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   4815
      _Version        =   393216
      _ExtentX        =   8493
      _ExtentY        =   4683
      _StockProps     =   64
      ColsFrozen      =   1
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
      MaxRows         =   18
      SpreadDesigner  =   "B_BusVas.frx":0000
      VisibleCols     =   2
      VisibleRows     =   18
      ScrollBarTrack  =   3
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   1815
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   4815
      _Version        =   393216
      _ExtentX        =   8493
      _ExtentY        =   3201
      _StockProps     =   64
      ColsFrozen      =   1
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
      MaxRows         =   18
      SpreadDesigner  =   "B_BusVas.frx":04B3
      VisibleCols     =   2
      VisibleRows     =   18
      ScrollBarTrack  =   3
   End
End
Attribute VB_Name = "B_BusVas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srow As Long, spid As Long
Dim ret As Integer
Dim Form1 As Form
Dim Est As Boolean
Dim text As String, RecetaSelect As String
Dim VecRecet() As String

Private Sub Command1_Click()
    
    If Option1.Value = True Then
        
        BuscaReceta
    
    Else
        
        BuscaIngrediente
    
    End If

End Sub

Private Sub Command2_Click()

If ret < 0 Or ret = 0 Then
'Form1.mnusearchnext.Enabled = False
End If
Unload Me

End Sub

Private Sub Command3_Click()
    
    If Option1.Value = True Then
       
       BuscaSiguiente
'        BuscaReceta
    Else
       
       BuscaSiguienteIngrediente
    
    End If
    
End Sub

Private Sub Form_Load()

fg_centra Me
Est = True
Text1.text = ""
text = ""
Command3.Enabled = False
Est = False
Me.Height = 2145
Me.Caption = "Buscar Recetas en Planificación"
Option2.Visible = True

If VarSitioRemoto = True Then
   
   Option2.Visible = False

End If

End Sub

Private Sub Option1_Click()
    
    LoadBuscar

End Sub

Private Sub Option2_Click()
    
    LoadBuscar

End Sub

Private Sub Text1_Change()
'If est Then Exit Sub

text = Text1.text
Command1.Enabled = True
Command3.Enabled = False
srow = 1
ret = 0

If Option2.Value = True Then
    
    CargaGrilla (text)

End If

End Sub

Sub Partidas(Form As Form)

Set Form1 = Form
Est = False

End Sub

Private Sub Text1_Validate(Cancel As Boolean)

text = Text1.text

End Sub

Private Sub LoadBuscar()

Dim RS  As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

vg_db.Execute "DELETE paso_servicio WHERE ser_spid = @@spid and ser_usr = '" & vg_NUsr & "'"
'--isel = 0
'-------> Buscar spid

Set RS = vg_db.Execute("SELECT @@spid spid")

If Not RS.EOF Then
   
   spid = RS!spid
   vg_db.Execute "INSERT INTO paso_servicio (ser_spid, ser_usr, ser_codigo) VALUES (" & spid & ", '" & vg_NUsr & "', " & Val(vg_codservicio) & ")"

End If
RS.Close
Set RS = Nothing

If Option1.Value = True Then
        
   Me.Height = 2145
   Me.Caption = "Buscar Recetas en Planificación"

Else
        
   Me.Height = 6945
   Me.Caption = "Buscar Ingrediente en Planificación"
   fg_carga ""
   Set RS1 = vg_db.Execute("sgpadm_s_llenapasoingrediente " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(vg_fecha) & ", " & vg_codlpr & ",'" & vg_NUsr & "'," & spid & "," & vg_IndpprSelec & "")
   fg_descarga
        '--RS1.Close: Set RS1 = Nothing

End If

End Sub

Private Sub BuscaReceta()

ret = 0

With Form1.vaSpread1
     
     For srow = 1 To .MaxRows - 1
         
         ret = .SearchRow(srow, 0, .maxcols - 1, text, 2)
         
         If ret > -1 Then
            
            .SetActiveCell ret, srow
            srow1 = srow
            Est = True
            Command1.Enabled = False
            Command3.Enabled = True
            Command3.SetFocus
            Exit Sub
'       Form1.mnusearchnext.Enabled = True
         
         End If
     
     Next srow
     
     If ret = -1 Then MsgBox "Texto no fue encontrado.": Exit Sub

End With

End Sub

Private Sub BuscaSiguiente()

Dim ret2 As Integer

ret2 = 0
Est = False

With Form1.vaSpread1
     
     For srow1 = srow To .MaxRows - 1
         
         If ret > -1 Then
            
            ret2 = .SearchRow(srow1, ret, .maxcols - 1, text, 2)
            
            If ret2 > -1 Then
               
               .SetActiveCell ret2, srow1
               srow = srow1
               ret = ret2
               ret2 = -1
               Exit Sub
            
            Else
               
               ret = 0
            
            End If
         
         ElseIf srow1 <> (.MaxRows - 1) Then
            
            ret = 1
         
         ElseIf srow1 = (.MaxRows - 1) Then
            
            ret = ret2
            ret2 = -1
         
         End If
     
     Next srow1
     
     Command1.Enabled = True
     Command3.Enabled = False
     srow = 1
     ret = 0
     Command1.SetFocus

End With

End Sub

Private Sub BuscaIngrediente()

Dim IngSelec As String
Dim i        As Long
Dim j        As Long
Dim RS1      As New ADODB.Recordset

If vaSpread1.MaxRows < 1 Then Exit Sub

fg_carga ""
j = 0
'vaSpread1.MaxRows = 0
DoEvents
Screen.MousePointer = 11

RecetaSelect = ""
IngSelec = ""
    
    vaSpread1.Col = 1
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        If vaSpread1.text = "1" Then

'            vaSpread1.Col = 3
            vaSpread1.Col = 2
            IngSelec = vaSpread1.text
            Exit For
        
        End If
    
    Next i

If IngSelec = "" Then fg_descarga: Exit Sub

Set RS1 = vg_db.Execute("sgpadm_s_FiltroPasoIngrediente 2, '" & IngSelec & "'")

If Not (RS1.EOF And RS1.BOF) Then RS1.MoveFirst: ReDim VecRecet(RS1!recCount - 1)

Do While Not RS1.EOF

    VecRecet(j) = RS1!rec_nombre
    j = j + 1

    RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing
fg_descarga

j = 0
RecetaSelect = VecRecet(j)

ret = 0

With Form1.vaSpread1
     
     For srow = 1 To .MaxRows - 1
         
         ret = .SearchRow(srow, 0, .maxcols - 1, RecetaSelect, 8)
         
         If ret > -1 Then
            
            .SetActiveCell ret, srow
            srow1 = srow
            Est = True
            Command1.Enabled = False
            Command3.Enabled = True
            Command3.SetFocus
            Exit Sub
'       Form1.mnusearchnext.Enabled = True
         
         End If
     
     Next srow
     
     If ret = -1 Then MsgBox "Texto no fue encontrado.": Exit Sub

End With

RS1.Close
Set RS1 = Nothing
fg_descarga

End Sub

Private Sub BuscaSiguienteIngrediente()

Dim ret2  As Integer
Dim srow1 As Long
Dim j     As Integer

ret2 = 0
Est = False

With Form1.vaSpread1
     
     For srow1 = srow To .MaxRows - 1
         
         If ret > -1 Then
            ret2 = .SearchRow(srow1, ret, Form1.vaSpread1.maxcols - 1, RecetaSelect, 8)
            
            If ret2 > -1 Then
               
               .SetActiveCell ret2, srow1
               srow = srow1
               ret = ret2
               ret2 = -1
               Exit Sub
            
            Else
               
               ret = 0
            
            End If
         
         ElseIf srow1 <> (.MaxRows - 1) Then
            
            ret = 1
         
         ElseIf srow1 = (.MaxRows - 1) Then
            
            ret = ret2
            ret2 = -1
         
         End If
     
     Next srow1
     
     If j < UBound(VecRecet) Then
        
        j = j + 1
        RecetaSelect = VecRecet(j)
        srow = 1
     
     For srow1 = srow To .MaxRows - 1
         
         If ret > -1 Then
            
            ret2 = .SearchRow(srow1, ret, Form1.vaSpread1.maxcols - 1, RecetaSelect, 2)
            
            If ret2 > -1 Then
               
               .SetActiveCell ret2, srow1
               srow = srow1
               ret = ret2
               ret2 = -1
               Exit Sub
            
            Else
               
               ret = 0
            
            End If
         
         ElseIf srow1 <> (.MaxRows - 1) Then
            
            ret = 1
         
         ElseIf srow1 = (.MaxRows - 1) Then
            
            ret = ret2
            ret2 = -1
         
         End If
     
     Next srow1
     
     Else
        
        Command1.Enabled = True
        Command3.Enabled = False
        srow = 1
        ret = 0
        Command1.SetFocus
        RecetaSelect = ""
     
     End If
     
End With

End Sub

Private Sub CargaGrilla(text As String)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

fg_carga ""
vaSpread1.MaxRows = 0
DoEvents
Screen.MousePointer = 11

Set RS1 = vg_db.Execute("sgpadm_s_FiltroPasoIngrediente 1, '" & text & "'")

If Not (RS1.EOF And RS1.BOF) Then
   
   RS1.MoveFirst
   DoEvents
   Screen.MousePointer = 11
   
   Do While Not RS1.EOF
          
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.Row = vaSpread1.MaxRows
    
          vaSpread1.Col = 2
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.TypeHAlign = TypeHAlignLeft
          vaSpread1.text = RS1!ing_codigo
          
          vaSpread1.Col = 3
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.TypeHAlign = TypeHAlignLeft
          vaSpread1.text = RS1!ing_nombre
          
          vaSpread1.Col = 4
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.TypeHAlign = TypeHAlignLeft
          vaSpread1.text = IIf(RS1!ing_indppr = "1", "Real", "Propuesta")
          
          RS1.MoveNext
          
   Loop
       'Label1(0).Visible = True
       'Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
   RS1.Close
   Set RS1 = Nothing

Else
   
   RS1.Close
   Set RS1 = Nothing

End If

fg_descarga
Man_Error:

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

Dim RS1 As New ADODB.Recordset
Dim i   As Long

Select Case Col

Case 1
    
    vaSpread1.Col = 1
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        If vaSpread1.text = "1" And vaSpread1.Row <> Row Then
            
            vaSpread1.text = 0
        
        ElseIf vaSpread1.Row = Row Then
          
          vaSpread1.text = 1
          vaSpread1.Col = 2
          Set RS1 = vg_db.Execute("sgpadm_s_FiltroPasoIngrediente 3, '" & vaSpread1.text & "'")
          vaSpread1.Col = 1
          
          If Not (RS1.EOF And RS1.BOF) Then
             RS1.MoveFirst
             DoEvents
             Screen.MousePointer = 11
             vaSpread2.MaxRows = 0
             
             Do While Not RS1.EOF
                
                vaSpread2.MaxRows = vaSpread2.MaxRows + 1
                vaSpread2.Row = vaSpread2.MaxRows
    
                vaSpread2.Col = 1
                vaSpread2.CellType = CellTypeStaticText
                vaSpread2.TypeHAlign = TypeHAlignLeft
                vaSpread2.text = IIf(IsNull(RS1!api_codpro), "", RS1!api_codpro)
          
                vaSpread2.Col = 2
                vaSpread2.CellType = CellTypeStaticText
                vaSpread2.TypeHAlign = TypeHAlignLeft
                vaSpread2.text = IIf(IsNull(RS1!pro_nombre), "", RS1!pro_nombre)
          
                vaSpread2.Col = 3
                vaSpread2.CellType = CellTypeStaticText
                vaSpread2.TypeHAlign = TypeHAlignLeft
                vaSpread2.text = IIf(RS1!pro_indppr = "1", "Real", "Propuesta")

                RS1.MoveNext
          
             Loop
             'Label1(0).Visible = True
             'Label1(0).Caption = "Registro " & Format(vaSpread1.MaxRows, fg_Pict(6, 0))
             RS1.Close
             Set RS1 = Nothing
             fg_descarga
    
          End If
          
        End If
        
    Next i
   
End Select

End Sub
