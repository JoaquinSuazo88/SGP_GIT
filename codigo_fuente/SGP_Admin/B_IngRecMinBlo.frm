VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form B_IngRecMinBlo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Ingrediente & Receta Minuta Bloque"
   ClientHeight    =   6555
   ClientLeft      =   7140
   ClientTop       =   3165
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   1200
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar Criterio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   1
      Top             =   0
      Width           =   3375
      Begin VB.Frame Frame2 
         Caption         =   "Por"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   3015
         Begin VB.OptionButton Option1 
            Caption         =   "Receta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Ingrediente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   3
            Top             =   240
            Width           =   1335
         End
      End
      Begin EditLib.fpText Text1 
         Height          =   315
         Left            =   840
         TabIndex        =   5
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   615
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Buscar &Siguiente"
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
      Left            =   3720
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   2655
      Left            =   240
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
      SpreadDesigner  =   "B_IngRecMinBlo.frx":0000
      VisibleCols     =   2
      VisibleRows     =   18
      ScrollBarTrack  =   3
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   1815
      Left            =   240
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
      SpreadDesigner  =   "B_IngRecMinBlo.frx":04B3
      VisibleCols     =   2
      VisibleRows     =   18
      ScrollBarTrack  =   3
   End
End
Attribute VB_Name = "B_IngRecMinBlo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim srow As Long
Dim spid As Long
Dim ret As Integer
Dim j As Long
Dim Form1 As Form
Dim Est As Boolean
Dim text As String
Dim RecetaSelect As String
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

End If

Unload Me

End Sub

Private Sub Command3_Click()
    
    If Option1.Value = True Then
       
       BuscaSiguiente
    
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
Me.Caption = "Buscar Recetas Minuta Bloque"
Option2.Visible = True
If VarSitioRemoto = True Then Option2.Visible = False

End Sub

Private Sub Option1_Click()
    
    LoadBuscar

End Sub

Private Sub Option2_Click()

LoadBuscar

End Sub

'Private Sub Text1_Change()
'
'text = Text1.text
'Command1.Enabled = True
'Command3.Enabled = False
'srow = 1
'ret = 0
'
'If Option2.Value = True Then
'
'   CargaGrilla (text)
'
'End If
'
'End Sub

Sub Partidas(Form As Form)

Set Form1 = Form
Est = False

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub

text = Text1.text
Command1.Enabled = True
Command3.Enabled = False
srow = 1
ret = 0

If Option2.Value = True Then

   CargaGrilla (text)

End If

'SendKeys "{Tab}"

End Sub

Private Sub Text1_Validate(Cancel As Boolean)

text = Text1.text

End Sub

Private Sub LoadBuscar()

Dim RS           As New ADODB.Recordset
Dim i            As Long
Dim j            As Long
Dim CodigoReceta As Long
Dim TipoReceta   As Long
Dim StrRec       As String
Dim StrRecb      As String
Dim MyBuffer     As String

'------> formatear grilla 1 - 2
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
'-------> Buscar spid
Set RS = vg_db.Execute("SELECT @@spid spid")
If Not RS.EOF Then spid = RS!spid
RS.Close: Set RS = Nothing

If Option1.Value = True Then
   Me.Height = 2145
   Me.Caption = "Buscar Recetas en Planificación"
Else
   Me.Height = 6945
   Me.Caption = "Buscar Ingrediente en Planificación"
   fg_carga ""
   Let MyBuffer = ""
   Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
   Let MyBuffer = MyBuffer & "<CodigoReceta>"
   
   For i = 1 To Form1.vaSpread1.MaxRows
       
       Form1.vaSpread1.Row = i
       
       For j = 6 To (Form1.vaSpread1.maxcols - 4) Step 7
           
           DoEvents
           CodigoReceta = 0
           Form1.vaSpread1.Col = j + 5
           StrRec = Trim(Form1.vaSpread1.text)
           
           If Len(StrRec) <> 0 Then
              
              Do While InStr(StrRec, ";") <> 0
                 
                 StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                 StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                 CodigoReceta = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                 TipoReceta = Val(Mid(StrRecb, 1))
              
              Loop
           
           End If
           
           If CodigoReceta > 0 Then
              
              Let MyBuffer = MyBuffer & " <CodReceta"
              Let MyBuffer = MyBuffer & " CodigoReceta = " & Chr(34) & CodigoReceta & Chr(34)
              Let MyBuffer = MyBuffer & "/>"
           
           End If
       
       Next j
   
   Next i
   
   Let MyBuffer = MyBuffer & "</CodigoReceta>"
   Set RS = vg_db.Execute("sgpadm_Sel_XmlMoverIngReceta '" & MyBuffer & "', '" & vg_codcasino & "', " & vg_codregimen & ", '" & vg_NUsr & "', " & spid & "")
   Set RS = Nothing
   fg_descarga

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
ret2 = 0: Est = False
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
     Command1.Enabled = True: Command3.Enabled = False: srow = 1: ret = 0: Command1.SetFocus
End With

End Sub

Private Sub BuscaIngrediente()

Dim RS As New ADODB.Recordset
Dim IngSelec As String
Dim i As Integer

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

If IngSelec = "" Then
   fg_descarga
   MsgBox "Debe Seleccionar un ingrediente", vbExclamation + vbOKOnly, Me.Caption
   Exit Sub
End If

RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_FiltroRecetaPasoIngrediente '" & IngSelec & "', '" & vg_NUsr & "', " & spid & "")
If Not (RS.EOF And RS.BOF) Then
   RS.MoveFirst: ReDim VecRecet(RS.RecordCount)
   Do While Not RS.EOF
      VecRecet(j) = RS!rec_nombre
      j = j + 1
      RS.MoveNext
   Loop
End If
RS.Close
Set RS = Nothing
fg_descarga

j = 0
RecetaSelect = VecRecet(j)
ret = 0

With Form1.vaSpread1
     For srow = 1 To .MaxRows - 1
         ret = .SearchRow(srow, 0, .maxcols - 2, RecetaSelect, 8)
         If ret > -1 Then
            .SetActiveCell ret, srow
            srow1 = srow
            Est = True
            Command1.Enabled = False
            Command3.Enabled = True
            Command3.SetFocus
            Exit Sub
         End If
     Next srow
     If ret = -1 Then MsgBox "Texto no fue encontrado.": Exit Sub
End With

' RS1.Close: Set RS1 = Nothing: fg_descarga

End Sub

Private Sub BuscaSiguienteIngrediente()

Dim ret2 As Integer, srow1 As Long
ret2 = 0: Est = False
With Form1.vaSpread1
     For srow1 = srow To .MaxRows - 1
         If ret > -1 Then
            ret2 = .SearchRow(srow1, ret, Form1.vaSpread1.maxcols - 2, RecetaSelect, 8)
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
            ret2 = .SearchRow(srow1, ret, Form1.vaSpread1.maxcols - 2, RecetaSelect, 2)
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
        Command1.Enabled = True: Command3.Enabled = False: srow = 1: ret = 0: Command1.SetFocus: RecetaSelect = ""
     End If
End With

End Sub

Private Sub CargaGrilla(text As String)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
fg_carga ""
vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0
DoEvents
Screen.MousePointer = 11

Set RS = vg_db.Execute("sgpadm_Sel_BuscarPasoIngrediente '" & text & "', '" & vg_NUsr & "', " & spid & "")
If Not (RS.EOF And RS.BOF) Then
   RS.MoveFirst
   DoEvents
   Screen.MousePointer = 11
   Do While Not RS.EOF
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.Row = vaSpread1.MaxRows
    
          vaSpread1.Col = 2
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.TypeHAlign = TypeHAlignLeft
          vaSpread1.text = RS!ing_codigo
          
          vaSpread1.Col = 3
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.TypeHAlign = TypeHAlignLeft
          vaSpread1.text = RS!ing_nombre
          
          vaSpread1.Col = 4
          vaSpread1.CellType = CellTypeStaticText
          vaSpread1.TypeHAlign = TypeHAlignLeft
          vaSpread1.text = IIf(RS!ing_indppr = "1", "Real", "Propuesta")
          
          RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
Else
   RS.Close: Set RS = Nothing
End If
fg_descarga
Man_Error:

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

Dim RS As New ADODB.Recordset

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
           Set RS = vg_db.Execute("sgpadm_Sel_FiltroProdPasoIngrediente '" & vaSpread1.text & "', '" & vg_NUsr & "', " & spid & "")
           vaSpread1.Col = 1
          
           If Not (RS.EOF And RS.BOF) Then
              RS.MoveFirst
              DoEvents
              Screen.MousePointer = 11
              vaSpread2.MaxRows = 0
              Do While Not RS.EOF
                 vaSpread2.MaxRows = vaSpread2.MaxRows + 1
                 vaSpread2.Row = vaSpread2.MaxRows
    
                 vaSpread2.Col = 1
                 vaSpread2.CellType = CellTypeStaticText
                 vaSpread2.TypeHAlign = TypeHAlignLeft
                 vaSpread2.text = IIf(IsNull(RS!api_codpro), "", RS!api_codpro)
                  
                 vaSpread2.Col = 2
                 vaSpread2.CellType = CellTypeStaticText
                 vaSpread2.TypeHAlign = TypeHAlignLeft
                 vaSpread2.text = IIf(IsNull(RS!pro_nombre), "", RS!pro_nombre)
                  
                 vaSpread2.Col = 3
                 vaSpread2.CellType = CellTypeStaticText
                 vaSpread2.TypeHAlign = TypeHAlignLeft
                 vaSpread2.text = IIf(RS!pro_indppr = 1, "Real", "Propuesta")
        
                 RS.MoveNext
           
              Loop

           End If
           RS.Close: Set RS = Nothing: fg_descarga
        End If
    Next i
   
End Select

End Sub
