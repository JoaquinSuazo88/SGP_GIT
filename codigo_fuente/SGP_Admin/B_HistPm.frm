VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_HistPm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Histórico Planificación Minutas"
   ClientHeight    =   3060
   ClientLeft      =   1485
   ClientTop       =   1845
   ClientWidth     =   7935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   7935
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
      Height          =   3060
      Left            =   7395
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   5398
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
Option Explicit
Option Compare Text

Private RS          As New ADODB.Recordset
Private op          As String
Private MsgTitulo   As String
Private BtnX        As Variant

Private Sub Form_Activate()
    Call fg_descarga
End Sub

Private Sub Form_Load()
    Call fg_centra(Me)
    fg_carga ""
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    Call fg_descarga
End Sub

Sub LlenarHistPlan(tfor As String, subseg As Variant, TipMin As String, opcion As String, Optional codReg As Long, Optional CodIng As Long, Optional FilCatDie As Long, Optional FilTipPla As Long, Optional codser As Long)

Dim RS          As New ADODB.Recordset
Dim AnchoCol    As Double
Dim Titulo      As String
Dim i           As Long
Dim ValLcntH$


Me.Caption = tfor
MsgTitulo = tfor
op = opcion

If opcion = "1" Or opcion = "2" Or opcion = "3" Or opcion = "4" Or opcion = "5" Then
   
   Me.Height = 3615
   Me.Width = 8025
   vaSpread1.Height = 3135
   fg_centra Me
   vaSpread1.MaxRows = 0
   vaSpread1.maxcols = IIf(opcion = "1" Or "3" Or "5", 5, 4): vaSpread1.Row = 0
   
   For i = 1 To vaSpread1.maxcols
       
       If (i = 1 Or i = 3 Or i = 5) Then
          AnchoCol = 7.38
          If i = 1 Then Titulo = "C.Regimen"
          If i = 3 Or i = 5 Then Titulo = IIf(opcion = "1" Or opcion = "3" Or opcion = "5", "C.Servicio", "C.Ingred.")
       End If
       If (opcion = "1" Or opcion = "3" Or opcion = "5") And (i = 2 Or i = 4) Then AnchoCol = 17.9: Titulo = "Descripción"
       If (opcion = "2" Or opcion = "4") And i = 2 Then AnchoCol = 17.9: Titulo = "Descripción"
       If (opcion = "2" Or opcion = "4") And i = 4 Then AnchoCol = 26: Titulo = "Descripción"
       If i = 5 Then AnchoCol = 8: Titulo = "Fecha"
       vaSpread1.Col = i
       vaSpread1.ColWidth(i) = AnchoCol
       vaSpread1.text = Titulo
       vaSpread1.ColHidden = False
   
   Next i
   
   If opcion = "1" Then
        If VarSitioRemoto = False Then
            Set RS = vg_db.Execute("sgpadm_s_planifminuta 8, " & subseg & ", 0, 0," & 0 & ", 0, 0, 0," & vg_IndpprSelec & "")
        Else
            Set RS = vg_db.Execute("select cli_minsre, cli_blockmincontrato from  b_clientes where cli_codigo = '" & subseg & "' and cli_activo = '1'")
            If Not RS.EOF Then
               If RS!cli_minsre = "1" And RS!cli_blockmincontrato = "1" Then
                  Set RS = vg_db.Execute("sgpadm_s_planifminutaSitioRemoto 2, '" & subseg & "', 0, 0, 0, 0, 0")
               Else
                  Set RS = vg_db.Execute("sgpadm_s_planifminutaSitioRemoto 1, '" & subseg & "', 0, 0, 0, 0, 0")
               End If
            End If

        End If
   
   ElseIf opcion = "2" Then
       Set RS = vg_db.Execute("sgpadm_Sel_TablaGramajeHistoricoSubSegmento " & subseg & "")
   
   ElseIf opcion = "4" Then
       Set RS = vg_db.Execute("sgpadm_Sel_TablaGramajeHistoricoCeco '" & subseg & "'")
   
   ElseIf opcion = "3" Then
      Set RS = vg_db.Execute("sgpadm_s_minutarealcasino 3, '" & TipMin & "', 0, 0, 0, 0, 0, ''")
   
   ElseIf opcion = "5" Then
      Set RS = vg_db.Execute("sgpadm_Sel_HistoricoCecoMinutaBloque '" & TipMin & "'")
   
   End If
   
   If RS.EOF Then RS.Close: Set RS = Nothing: vg_codigo = "": Exit Sub

End If
Do While Not RS.EOF
   If opcion = "1" Or opcion = "3" Or opcion = "5" Then
      
      If Not IsNull(RS!reg_nombre) And Not IsNull(RS!ser_nombre) Then
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
      
         vaSpread1.Col = 1
         vaSpread1.TypeHAlign = TypeHAlignRight
         vaSpread1.text = RS!min_codreg
      
         vaSpread1.Col = 2
         vaSpread1.TypeHAlign = TypeHAlignLeft
         vaSpread1.text = IIf(IsNull(RS!reg_nombre), "No existe regimen", Trim(RS!reg_nombre))
      
         vaSpread1.Col = 3
         vaSpread1.TypeHAlign = TypeHAlignRight
         vaSpread1.text = RS!min_codser
      
         vaSpread1.Col = 4
         vaSpread1.TypeHAlign = TypeHAlignLeft
         vaSpread1.text = IIf(IsNull(RS!ser_nombre), "No existe servicio", Trim(RS!ser_nombre))
     
         vaSpread1.Col = 5
         vaSpread1.TypeHAlign = TypeHAlignCenter
         vaSpread1.text = Mid(RS!Fecha, 5, 2) & "/" & Mid(RS!Fecha, 1, 4)
      
      End If
   
   Else
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = RS(0)
      
      vaSpread1.Col = 2
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS(1))
      
      vaSpread1.Col = 3
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = RS(2)
      
      vaSpread1.Col = 4
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS(3))
   
   End If
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vg_codigo = ""

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

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vg_codigo = "C"
vaSpread1.Col = 1: vg_codregimen = Val(vaSpread1.text)
If op = "1" Or op = "3" Or op = "5" Then
   vaSpread1.Col = 3: vg_codservicio = Val(vaSpread1.text)
   vaSpread1.Col = 5: vg_fecha = vaSpread1.text
ElseIf op = "2" Or op = "4" Then
   vaSpread1.Col = 3: vg_codigo = Trim(vaSpread1.text)
End If
Cerrar

End Sub

Sub Cerrar()
Me.Hide
Unload Me
End Sub
