VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_Plami3 
   Appearance      =   0  'Flat
   Caption         =   "Cambiar Planificación Minuta De Propuesta a Real."
   ClientHeight    =   6960
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13005
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6960
   ScaleWidth      =   13005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   2640
      TabIndex        =   5
      Top             =   120
      Width           =   6375
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   2
         ItemData        =   "M_Plami3.frx":0000
         Left            =   1800
         List            =   "M_Plami3.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1200
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         ItemData        =   "M_Plami3.frx":0004
         Left            =   1800
         List            =   "M_Plami3.frx":0006
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "M_Plami3.frx":0008
         Left            =   1800
         List            =   "M_Plami3.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1845
         TabIndex        =   14
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   1845
         TabIndex        =   13
         Top             =   840
         Width           =   4095
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   1845
         TabIndex        =   12
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label1 
         Caption         =   "Servicio"
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
         Index           =   7
         Left            =   360
         TabIndex        =   11
         Top             =   1275
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Régimen"
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
         Index           =   6
         Left            =   360
         TabIndex        =   10
         Top             =   780
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Sub-Segmento"
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
         Index           =   5
         Left            =   360
         TabIndex        =   9
         Top             =   315
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      ScaleHeight     =   675
      ScaleWidth      =   6675
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ProgressBar gauge1 
         Height          =   210
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Datos"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   4
         Top             =   0
         Width           =   1125
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4320
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   12285
      _Version        =   393216
      _ExtentX        =   21669
      _ExtentY        =   7620
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
      MaxCols         =   12
      MaxRows         =   15
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_Plami3.frx":000C
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6960
      Left            =   12375
      TabIndex        =   1
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12277
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Plami3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset, RS3 As New ADODB.Recordset
Dim NomFor As String
Dim codSubseg, codreg, codser As Long

Private Sub Combo1_Click(Index As Integer)
    Select Case Index
    Case 0
      If Combo1(0).text <> "Todos" Then
      codSubseg = Val(fg_codigocbo(Combo1, 0, 2, ""))
        If codSubseg = 0 Then
          codSubseg = Val(fg_codigocbo(Combo1, 0, 1, ""))
        End If
      End If
      DetallePlantillaMinuta
    Case 1
      codreg = Val(fg_codigocbo(Combo1, 1, 5, ""))
      DetallePlantillaMinuta
    Case 2
      codser = Val(fg_codigocbo(Combo1, 2, 5, ""))
      DetallePlantillaMinuta
    End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Sub DetallePlantillaMinuta()
vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
RS1.Open "sgpadm_s_planifminuta 14, " & codSubseg & ", " & codreg & ", " & codser & ", 0, " & vg_fecha & ", 0,0,'2'", vg_db, adOpenForwardOnly  ', adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      DoEvents
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = Trim(RS1!min_subseg)
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS1!sub_nombre)
            
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS1!min_codreg)
      
      vaSpread1.Col = 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS1!reg_nombre)
      
      vaSpread1.Col = 6
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.text = Trim(RS1!min_codser)
      
      vaSpread1.Col = 7
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS1!ser_nombre)
      
      vaSpread1.Col = 8
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.text = Trim(RS1!fecha)
      
      vaSpread1.Col = 9
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = IIf(Trim(RS1!min_Indppr) = "2", "Propuesta", "Real")
      
      vaSpread1.TypeDateCentury = True
      RS1.MoveNext
   Loop
End If
RS1.Close: Set RS1 = Nothing
End Sub

Private Sub Form_Load()
Dim OpUsuario As String
codSubseg = 0: codreg = 0: codser = 0
fg_carga ""
Me.HelpContextID = vg_OpcM
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = True
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
fg_descarga

'Llenado de vg_fecha
Set RS = vg_db.Execute("sgpadm_s_listaprecio 8, 0, 0, '" & vg_NUsr & "'")
If Not RS.EOF Then
   vg_codlpr = RS!lpr_codigo
   vg_fecha = RS!dlp_anomes
Else
   vg_fecha = 0
End If
RS.Close: Set RS = Nothing

'Llenado primer combo Sub-Segmento
Set RS = vg_db.Execute("sgpadm_s_subsegmento 4, 0, '', ''")
Combo1(0).Clear
Combo1(0).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 1) & ")"
Do While Not RS.EOF
    Combo1(0).AddItem IIf(IsNull(RS!sub_codigo), "", RS!sub_codigo & " - " & RS!sub_nombre) & Space(150) & "(" & RS!sub_codigo & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(0).ListIndex = 0

'Llenado primer combo Regimen
Combo1(1).Clear
Set RS = vg_db.Execute("sgpadm_s_regimen 2, 0,''")
Combo1(1).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 1) & ")"
Do While Not RS.EOF
    Combo1(1).AddItem IIf(IsNull(RS!reg_nombre), "", RS!reg_codigo & " - " & RS!reg_nombre) & Space(150) & "(" & RS!reg_codigo & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(1).ListIndex = 0

'Llenado primer combo Servicio
Combo1(2).Clear
Set RS = vg_db.Execute("sgpadm_s_servicio 3,'',0,''")
Combo1(2).AddItem "Todos" & Space(150) & "(" & fg_pone_cero(Str(0), 1) & ")"
Do While Not RS.EOF
    Combo1(2).AddItem IIf(IsNull(RS!ser_nombre), "", RS!ser_codigo & " - " & RS!ser_nombre) & Space(150) & "(" & RS!ser_codigo & ")"
    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Combo1(2).ListIndex = 0

'Carga Grilla
DetallePlantillaMinuta
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim g, x, vcodsse As Long, vcodreg As Long, vcodser As Long, fecmin As Long
Dim est As Boolean
codsse = 0: codreg = 0: codser = 0: fecmin = 0: est = False
Select Case Button.Index
Case 2 '-------> Cambia Planificación minuta De Propuesta a Real
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then est = True: Exit Sub
    Next i
    If Not est Then fg_descarga: MsgBox "Debe seleccionar a lo menos un ítem", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    
    'Se obtienen valores
    Picture1.Visible = True: Label3.Visible = True: gauge1.Visible = True
    Picture1.Refresh: Label3.Refresh: gauge1.Refresh
    For g = 1 To vaSpread1.MaxRows
        gauge1.Value = Val((g / vaSpread1.MaxRows) * 100)
        vaSpread1.Col = 1: vaSpread1.Row = g
        If vaSpread1.Value = "1" Then 'Checked
           '------>  Se obtienen los valores por Línea
           vaSpread1.Col = 2: vcodsse = vaSpread1.text
           vaSpread1.Col = 4: vcodreg = vaSpread1.text
           vaSpread1.Col = 6: vcodser = vaSpread1.text
           vaSpread1.Col = 8: fecmin = vaSpread1.text
           
           vaSpread1.Col = 12
           If vaSpread1.Value = "1" Then 'Cambia Productos y Ingredientes
              vg_db.Execute "UPDATE b_productos SET b_productos.pro_indppr = '1' " & _
                            "FROM b_productos a, b_productosing b, b_ingrediente c, b_receta d, b_recetadet e, b_minuta f, b_minutadet g " & _
                            "WHERE  f.min_codigo = g.mid_codigo " & _
                            "AND    g.mid_codrec = d.rec_codigo " & _
                            "AND    d.rec_codigo = e.red_codigo " & _
                            "AND    e.red_codpro = c.ing_codigo " & _
                            "AND    c.ing_codigo = b.pri_coding " & _
                            "AND    b.pri_codpro = a.pro_codigo " & _
                            "AND    f.min_subseg = " & vcodsse & " " & _
                            "AND    f.min_codreg = " & vcodreg & " " & _
                            "AND    f.min_codser = " & vcodser & " " & _
                            "AND substring(convert(char(8),f.min_fecmin),1,6) = " & fecmin & " " & _
                            "AND    f.min_indppr = '2' " & _
                            "AND    a.pro_indppr = '2'"
              
              vg_db.Execute "UPDATE b_ingrediente SET b_ingrediente.ing_indppr = '1' " & _
                            "FROM b_ingrediente a, b_receta b, b_recetadet c, b_minuta d, b_minutadet e " & _
                            "WHERE  d.min_codigo = e.mid_codigo " & _
                            "AND    e.mid_codrec = b.rec_codigo " & _
                            "AND    b.rec_codigo = c.red_codigo " & _
                            "AND    c.red_codpro = a.ing_codigo " & _
                            "AND    b.min_subseg = " & vcodsse & " " & _
                            "AND    b.min_codreg = " & vcodreg & " " & _
                            "AND    b.min_codser = " & vcodser & " " & _
                            "AND substring(convert(char(8),b.min_fecmin),1,6) = " & fecmin & " " & _
                            "AND    b.min_indppr = '2' " & _
                            "AND    a.ing_indppr = '2'"
           End If
           
           vaSpread1.Col = 11
           If vaSpread1.Value = "1" Then 'Cambia recetas
              vg_db.Execute "UPDATE b_receta SET b_receta.rec_indppr = '1' " & _
                            "FROM b_receta a, b_minuta b, b_minutadet c " & _
                            "WHERE  b.min_codigo = c.mid_codigo " & _
                            "AND    c.mid_codrec = a.rec_codigo " & _
                            "AND    b.min_subseg = " & vcodsse & " " & _
                            "AND    b.min_codreg = " & vcodreg & " " & _
                            "AND    b.min_codser = " & vcodser & " " & _
                            "AND substring(convert(char(8),b.min_fecmin),1,6) = " & fecmin & " " & _
                            "AND    b.min_indppr = '2' " & _
                            "AND    a.rec_indppr = '2'"
           End If
                  
           vaSpread1.Col = 10
           If vaSpread1.Value = "1" Then 'Cambia planificación
               vg_db.Execute "UPDATE b_minuta SET b_minuta.min_indppr = '1' " & _
                             "WHERE min_subseg = " & vcodsse & " " & _
                             "AND   min_codreg = " & vcodreg & " " & _
                             "AND   min_codser = " & vcodser & " " & _
                             "AND   substring(convert(char(8),min_fecmin),1,6) = " & fecmin & " " & _
                             "AND   min_Indppr = '2'"
           End If
        End If
      Next g
      Picture1.Visible = False: gauge.Visible = False
      fg_descarga
      MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
      DetallePlantillaMinuta
    Case 4 '------- Salir
        Me.Hide
        Unload Me
    End Select
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    If vaSpread1.MaxRows > 0 And (Col = 1 Or Col = 10 Or Col = 11 Or Col = 12) And Row = 0 Then
          For i = 1 To vaSpread1.MaxRows
              vaSpread1.Col = Col: vaSpread1.Row = i
              If vaSpread1.text = "" Then
                vaSpread1.text = "1"
              Else
                vaSpread1.text = ""
              End If
          Next i
    End If
End Sub
