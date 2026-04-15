VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form C_AporteSansis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Nutricional Día"
   ClientHeight    =   6195
   ClientLeft      =   1305
   ClientTop       =   1905
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   11490
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle Nutricional x Plato Servicio"
      TabPicture(0)   =   "C_AporteSansis.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Resumen Nutricional x Servicio"
      TabPicture(1)   =   "C_AporteSansis.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(2)"
      Tab(1).Control(1)=   "Frame1(3)"
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   10335
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2175
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   10095
            _Version        =   393216
            _ExtentX        =   17806
            _ExtentY        =   3836
            _StockProps     =   64
            ColsFrozen      =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   501
            SpreadDesigner  =   "C_AporteSansis.frx":0038
            VisibleCols     =   8
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   10335
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2175
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   10095
            _Version        =   393216
            _ExtentX        =   17806
            _ExtentY        =   3836
            _StockProps     =   64
            ColsFrozen      =   8
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   501
            SpreadDesigner  =   "C_AporteSansis.frx":6EC6
            VisibleCols     =   8
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   2
         Left            =   -74880
         TabIndex        =   3
         Top             =   480
         Width           =   10335
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2175
            Index           =   2
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   10095
            _Version        =   393216
            _ExtentX        =   17806
            _ExtentY        =   3836
            _StockProps     =   64
            ColsFrozen      =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "C_AporteSansis.frx":D829
            UserResize      =   1
            VisibleCols     =   7
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   3
         Left            =   -74880
         TabIndex        =   1
         Top             =   3120
         Width           =   10335
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2175
            Index           =   3
            Left            =   120
            TabIndex        =   2
            Top             =   240
            Width           =   10095
            _Version        =   393216
            _ExtentX        =   17806
            _ExtentY        =   3836
            _StockProps     =   64
            ColsFrozen      =   7
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SpreadDesigner  =   "C_AporteSansis.frx":11E2D
            VisibleCols     =   7
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6195
      Left            =   10860
      TabIndex        =   9
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   10927
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "C_AporteSansis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim matrizaporte() As Double
Dim imat As Integer, proc As Integer, i As Integer, j As Integer, iaporte As Integer
Dim cantcalorias As Double, cantproteinas As Double, cantlipidos As Double, canthidratos As Double, cantporsolida As Double, cantacgrsat As Double
Dim cantsercalorias As Double, cantserproteinas As Double, cantserlipidos As Double, cantserhidratos As Double, cantserporsolida As Double, cantserporliquida As Double, cantseracgrsat As Double
Dim cantgrl1calorias As Double, cantgrl1proteinas As Double, cantgrl1lipidos As Double, cantgrl1hidratos As Double, cantgrl1porsolida As Double, cantgrl1porliquida As Double, cantgrl1acgrsat As Double
Dim cantgrl2calorias As Double, cantgrl2proteinas As Double, cantgrl2lipidos As Double, cantgrl2hidratos As Double, cantgrl2porsolida As Double, cantgrl2porliquida As Double, cantgrl2acgrsat As Double
Dim cantaporte As Double
Dim nomserv As String
Dim auxserv As Long

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Toolbar1.Buttons.Clear
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Sub LlenarApoPlan(ffor As Object, tfor As String, subseg As Variant, codReg As Long, codser As Long, Fecha As Long, TipMin As String, icol As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

vaSpread1(0).MaxRows = 0: vaSpread1(0).maxcols = 8
vaSpread1(1).MaxRows = 0: vaSpread1(1).maxcols = 8
vaSpread1(2).MaxRows = 0: vaSpread1(2).maxcols = 7
vaSpread1(3).MaxRows = 0: vaSpread1(3).maxcols = 7
SSTab1.Tab = 0

'-------> Llenar Nutrientes
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
iaporte = 0
ReDim Preserve matrizaporte(0)

If Not RS.EOF Then
   
   vaSpread1(0).MaxRows = 0
   
   Do While Not RS.EOF
      
      vaSpread1(0).Row = 0
      vaSpread1(0).maxcols = vaSpread1(0).maxcols + 1
      vaSpread1(0).Col = vaSpread1(0).maxcols
      vaSpread1(0).Value = Trim(RS!nut_nombre)

      vaSpread1(1).Row = 0
      vaSpread1(1).maxcols = vaSpread1(1).maxcols + 1
      vaSpread1(1).Col = vaSpread1(1).maxcols
      vaSpread1(1).Value = Trim(RS!nut_nombre)
      
      vaSpread1(2).Row = 0
      vaSpread1(2).maxcols = vaSpread1(2).maxcols + 1
      vaSpread1(2).Col = vaSpread1(2).maxcols
      vaSpread1(2).Value = Trim(RS!nut_nombre)
      
      vaSpread1(3).Row = 0
      vaSpread1(3).maxcols = vaSpread1(3).maxcols + 1
      vaSpread1(3).Col = vaSpread1(3).maxcols
      vaSpread1(3).Value = Trim(RS!nut_nombre)
      
      iaporte = iaporte + 1
      RS.MoveNext
   
   Loop
   
   ReDim Preserve matrizaporte(iaporte)

End If
RS.Close: Set RS = Nothing

C_AporteSansis.Caption = "Aporte Nutricional Día " & Format(Mid(Fecha, 7, 2) & "/" & Mid(Fecha, 5, 2) & "/" & Mid(Fecha, 1, 4))

cantsercalorias = 0
cantserproteinas = 0
cantserlipidos = 0
cantserhidratos = 0
proc = 0

cantserporsolida = 0
cantserporliquida = 0
cantseracgrsat = 0

cantgrl1calorias = 0
cantgrl1proteinas = 0
cantgrl1lipidos = 0
cantgrl1hidratos = 0
proc = 0

cantgrl1porsolida = 0
cantgrl1porliquida = 0
cantgrl1acgrsat = 0

cantgrl2calorias = 0
cantgrl2proteinas = 0
cantgrl2lipidos = 0
cantgrl2hidratos = 0
cantgrl2porsolida = 0
cantgrl2porliquida = 0
cantgrl2acgrsat = 0

auxserv = 0
For i = 1 To iaporte
    
    matrizaporte(i) = 0

Next i

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_DiaMinutaBloque '" & subseg & "', " & codReg & ", " & codser & ", " & Fecha & "")

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      If RS!Ser_codigo <> auxserv Then
         
         If proc = 1 Then
            
            MoverCtrServicio
         
         End If
         
         If RS!Ser_codigo = 2 Or RS!Ser_codigo = 23 Then
            
            imat = 1
         
         Else
            imat = 0
         
         End If
         
         vaSpread1(imat).MaxRows = vaSpread1(imat).MaxRows + 1
         vaSpread1(imat).Row = vaSpread1(imat).MaxRows
         vaSpread1(imat).Col = 2
         vaSpread1(imat).BackColor = &H80FF80
         vaSpread1(imat).Font.Bold = True
         vaSpread1(imat).Font.Size = 9
         vaSpread1(imat).Value = Trim(RS!ser_nombre)
         auxserv = RS!Ser_codigo
         nomserv = Trim(RS!ser_nombre)
         proc = 1
         MnitmRef = 0
         MnitmNo = 0
         SwTotal = 0
      
      End If
      
      vaSpread1(imat).MaxRows = vaSpread1(imat).MaxRows + 1
      vaSpread1(imat).Row = vaSpread1(imat).MaxRows
      
      vaSpread1(imat).Col = 1
      vaSpread1(imat).Value = 1
      
      vaSpread1(imat).Col = 2
      vaSpread1(imat).Value = RS!rec_nombre
      cantporsolida = 0
      
      vaSpread1(imat).Col = 3
      vaSpread1(imat).TypeHAlign = 1
      vaSpread1(imat).Value = Format(RS!rec_canser, fg_Pict(4, 2))
      vaSpread1(imat).ForeColor = &HFF0000
      
      cantporsolida = RS!rec_canser
      cantserporsolida = CCur(cantserporsolida + RS!rec_canser)
         
      If imat = 0 Then
            
         cantgrl1porsolida = CCur(cantgrl1porsolida + RS!rec_canser)
         
      Else
            
         cantgrl2porsolida = CCur(cantgrl2porsolida + RS!rec_canser)
       
      End If
      
      cantcalorias = 0
      cantproteinas = 0
      cantlipidos = 0
      canthidratos = 0
      cantacgrsat = 0
      j = 9
      i = 1
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS1 = vg_db.Execute("sgpadm_Sel_AporteMinutaBloqueDia_V02 " & RS!rec_codigo & ", " & codReg & ", '" & subseg & "'")
      
      If Not RS1.EOF Then
         
         Do While Not RS1.EOF
            
            vaSpread1(imat).Col = j
            vaSpread1(imat).CellType = 5
            vaSpread1(imat).TypeHAlign = 1
            vaSpread1(imat).Value = Format(RS1!valordietetico, fg_Pict(6, 2))
            vaSpread1(imat).ForeColor = &HFF0000
            
            If RS1!nut_codigo = 2 And RS1!valordietetico > 0 Then
               
               cantcalorias = RS1!valordietetico
               cantsercalorias = CCur(cantsercalorias + RS1!valordietetico)
               
               If imat = 0 Then
                  
                  cantgrl1calorias = CCur(cantgrl1calorias + RS1!valordietetico)
               
               Else
                  
                  cantgrl2calorias = CCur(cantgrl2calorias + RS1!valordietetico)
               
               End If
            
            ElseIf RS1!nut_codigo = 3 And RS1!valordietetico > 0 Then
               
               cantproteinas = RS1!valordietetico
               cantserproteinas = CCur(cantserproteinas + RS1!valordietetico)
               
               If imat = 0 Then
                  
                  cantgrl1proteinas = CCur(cantgrl1proteinas + RS1!valordietetico)
               
               Else
                  cantgrl2proteinas = CCur(cantgrl2proteinas + RS1!valordietetico)
               
               End If
            
            ElseIf RS1!nut_codigo = 6 And RS1!valordietetico > 0 Then
            
               cantlipidos = RS1!valordietetico
               cantserlipidos = CCur(cantserlipidos + RS1!valordietetico)
               
               If imat = 0 Then
                  
                  cantgrl1lipidos = CCur(cantgrl1lipidos + RS1!valordietetico)
               
               Else
                  
                  cantgrl2lipidos = CCur(cantgrl2lipidos + RS1!valordietetico)
               
               End If
            
            ElseIf RS1!nut_codigo = 4 And RS1!valordietetico > 0 Then
               
               canthidratos = RS1!valordietetico
               cantserhidratos = CCur(cantserhidratos + RS1!valordietetico)
               
               If imat = 0 Then
                  
                  cantgrl1hidratos = CCur(cantgrl1hidratos + RS1!valordietetico)
               
               Else
                  
                  cantgrl2hidratos = CCur(cantgrl2hidratos + RS1!valordietetico)
               
               End If
            
            ElseIf RS1!nut_codigo = 8 And RS1!valordietetico > 0 Then
            
               cantacgrsat = RS1!valordietetico
               cantseracgrsat = CCur(cantseracgrsat + RS1!valordietetico)
               
               If imat = 0 Then
                  
                  cantgrl1acgrsat = CCur(cantgrl1acgrsat + RS1!valordietetico)
               
               Else
                  
                  cantgrl2acgrsat = CCur(cantgrl2acgrsat + RS1!valordietetico)
               
               End If
            
            End If
            
            If i <= iaporte Then
               
               matrizaporte(i) = CCur(matrizaporte(i) + RS1!valordietetico)
            
            End If
            
            i = i + 1
            j = j + 1
            
            RS1.MoveNext
         
         Loop
      
      End If
      RS1.Close
      Set RS1 = Nothing
      
      vaSpread1(imat).Col = 5
      vaSpread1(imat).CellType = 5
      vaSpread1(imat).TypeHAlign = 1
      vaSpread1(imat).Value = Format(0, fg_Pict(6, 2))
      If cantcalorias > 0 And cantproteinas > 0 Then
         
         vaSpread1(imat).Value = Format(CCur(((cantproteinas * 4) / cantcalorias) * 100), fg_Pict(6, 2))
      
      End If
      vaSpread1(imat).ForeColor = &HFF0000
      
      vaSpread1(imat).Col = 6
      vaSpread1(imat).CellType = 5
      vaSpread1(imat).TypeHAlign = 1
      vaSpread1(imat).ForeColor = &HFF0000
      vaSpread1(imat).Value = Format(0, fg_Pict(6, 2))
      
      If cantcalorias > 0 And cantlipidos > 0 Then
      
         vaSpread1(imat).Value = Format(CCur(((cantlipidos * 9) / cantcalorias) * 100), fg_Pict(6, 2))
      
      End If
      
      
      vaSpread1(imat).Col = 7
      vaSpread1(imat).CellType = 5
      vaSpread1(imat).TypeHAlign = 1
      vaSpread1(imat).Value = Format(0, fg_Pict(6, 2))
      
      If canthidratos > 0 And cantcalorias > 0 Then
      
            vaSpread1(imat).Value = Format(CCur(((canthidratos * 4) / cantcalorias) * 100), fg_Pict(6, 2))
      
      End If

      vaSpread1(imat).ForeColor = &HFF0000
      
      vaSpread1(imat).Col = 8
      vaSpread1(imat).CellType = 5
      vaSpread1(imat).TypeHAlign = 1
      vaSpread1(imat).Value = Format(0, fg_Pict(6, 2))
      
      If cantcalorias > 0 And cantacgrsat > 0 Then
      
            vaSpread1(imat).Value = Format(CCur(((cantacgrsat * 9) / cantcalorias) * 100), fg_Pict(6, 2))
      
      End If

      vaSpread1(imat).ForeColor = &HFF0000
      
      RS.MoveNext
   
   Loop
   
   MoverCtrServicio
   
   If vaSpread1(0).MaxRows > 0 Then
      
      MoverGrlDiaVec1
   
   End If
   
   If vaSpread1(1).MaxRows > 0 Then
      
      MoverGrlDiaVec2
   
   End If

End If
RS.Close
Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub MoverCtrServicio()

On Error GoTo Man_Error

vaSpread1(imat).MaxRows = vaSpread1(imat).MaxRows + 1
vaSpread1(imat).Row = vaSpread1(imat).MaxRows
            
vaSpread1(imat).MaxRows = vaSpread1(imat).MaxRows + 1
vaSpread1(imat).Row = vaSpread1(imat).MaxRows

vaSpread1(imat).Col = 2
vaSpread1(imat).BackColor = &H80000016
vaSpread1(imat).Font.Bold = True
vaSpread1(imat).Font.Size = 9
vaSpread1(imat).Value = "Total " & nomserv
            
If imat = 0 Then
   
   vaSpread1(2).MaxRows = vaSpread1(2).MaxRows + 1
   vaSpread1(2).Row = vaSpread1(2).MaxRows
   vaSpread1(2).Col = 1
   vaSpread1(2).Font.Bold = True
   vaSpread1(2).Font.Size = 9
   vaSpread1(2).Value = nomserv
   vaSpread1(2).BackColor = &H80FF80

Else
   
   vaSpread1(3).MaxRows = vaSpread1(3).MaxRows + 1
   vaSpread1(3).Row = vaSpread1(3).MaxRows
   vaSpread1(3).Col = 1
   vaSpread1(3).Font.Bold = True
   vaSpread1(3).Font.Size = 9
   vaSpread1(3).Value = nomserv

End If

vaSpread1(imat).Col = 3
vaSpread1(imat).CellType = 5
vaSpread1(imat).TypeHAlign = 1
vaSpread1(imat).Value = IIf(cantserporsolida > 0, Format(cantserporsolida, fg_Pict(6, 2)), Format(0, fg_Pict(6, 2)))
vaSpread1(imat).ForeColor = &HFF&
   
If imat = 0 Then
      
   vaSpread1(2).Col = 2
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(cantserporsolida, fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF0000
   
Else
      
   vaSpread1(3).Col = 2
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(cantserporsolida, fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF0000
   
End If

vaSpread1(imat).Col = 4
vaSpread1(imat).CellType = 5
vaSpread1(imat).TypeHAlign = 1
vaSpread1(imat).Value = IIf(cantserporliquida > 0, Format(cantserporliquida, fg_Pict(6, 2)), Format(0, fg_Pict(6, 2)))
vaSpread1(imat).ForeColor = &HFF&
   
If imat = 0 Then
      
   vaSpread1(2).Col = 3
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(cantserporliquida, fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF0000
   
Else
      
   vaSpread1(3).Col = 3
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(cantserporliquida, fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF0000
   
End If

If cantsercalorias > 0 And cantserproteinas > 0 Then
   
   vaSpread1(imat).Col = 5
   vaSpread1(imat).CellType = 5
   vaSpread1(imat).TypeHAlign = 1
   vaSpread1(imat).Value = Format(CCur(((cantserproteinas * 4) / cantsercalorias) * 100), fg_Pict(6, 2))
   vaSpread1(imat).ForeColor = &HFF&
   
   If imat = 0 Then
      
      vaSpread1(2).Col = 4
      vaSpread1(2).CellType = 5
      vaSpread1(2).TypeHAlign = 1
      vaSpread1(2).Value = Format(CCur(((cantserproteinas * 4) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(2).ForeColor = &HFF0000
   
   Else
      
      vaSpread1(3).Col = 4
      vaSpread1(3).CellType = 5
      vaSpread1(3).TypeHAlign = 1
      vaSpread1(3).Value = Format(CCur(((cantserproteinas * 4) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(3).ForeColor = &HFF0000
   
   End If

End If

If cantsercalorias > 0 And cantserlipidos > 0 Then
   
   vaSpread1(imat).Col = 6
   vaSpread1(imat).CellType = 5
   vaSpread1(imat).TypeHAlign = 1
   vaSpread1(imat).Value = Format(CCur(((cantserlipidos * 9) / cantsercalorias) * 100), fg_Pict(6, 2))
   vaSpread1(imat).ForeColor = &HFF&
   
   If imat = 0 Then
      
      vaSpread1(2).Col = 5
      vaSpread1(2).CellType = 5
      vaSpread1(2).TypeHAlign = 1
      vaSpread1(2).Value = Format(CCur(((cantserlipidos * 9) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(2).ForeColor = &HFF0000
   
   Else
      
      vaSpread1(3).Col = 5
      vaSpread1(3).CellType = 5
      vaSpread1(3).TypeHAlign = 1
      vaSpread1(3).Value = Format(CCur(((cantserlipidos * 9) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(3).ForeColor = &HFF0000
   
   End If

End If

If cantserhidratos > 0 And cantsercalorias > 0 Then
   
   vaSpread1(imat).Col = 7
   vaSpread1(imat).CellType = 5
   vaSpread1(imat).TypeHAlign = 1
   vaSpread1(imat).Value = Format(CCur(((cantserhidratos * 4) / cantsercalorias) * 100), fg_Pict(6, 2))
   vaSpread1(imat).ForeColor = &HFF&
   
   If imat = 0 Then
      
      vaSpread1(2).Col = 6
      vaSpread1(2).CellType = 5
      vaSpread1(2).TypeHAlign = 1
      vaSpread1(2).Value = Format(CCur(((cantserhidratos * 4) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(2).ForeColor = &HFF0000
   
   Else
      
      vaSpread1(3).Col = 6
      vaSpread1(3).CellType = 5
      vaSpread1(3).TypeHAlign = 1
      vaSpread1(3).Value = Format(CCur(((cantserhidratos * 4) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(3).ForeColor = &HFF0000
   
   End If

End If

If cantseracgrsat > 0 And cantsercalorias > 0 Then
   
   vaSpread1(imat).Col = 8
   vaSpread1(imat).CellType = 5
   vaSpread1(imat).TypeHAlign = 1
   vaSpread1(imat).Value = Format(CCur(((cantseracgrsat * 9) / cantsercalorias) * 100), fg_Pict(6, 2))
   vaSpread1(imat).ForeColor = &HFF&
   
   If imat = 0 Then
      
      vaSpread1(2).Col = 7
      vaSpread1(2).CellType = 5
      vaSpread1(2).TypeHAlign = 1
      vaSpread1(2).Value = Format(CCur(((cantseracgrsat * 9) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(2).ForeColor = &HFF0000
   
   Else
      
      vaSpread1(3).Col = 7
      vaSpread1(3).CellType = 5
      vaSpread1(3).TypeHAlign = 1
      vaSpread1(3).Value = Format(CCur(((cantseracgrsat * 9) / cantsercalorias) * 100), fg_Pict(6, 2))
      vaSpread1(3).ForeColor = &HFF0000
   
   End If

End If

j = 9

For i = 1 To iaporte
    
    vaSpread1(imat).Col = j
    vaSpread1(imat).CellType = 5
    vaSpread1(imat).TypeHAlign = 1
    vaSpread1(imat).Value = Format(matrizaporte(i), fg_Pict(6, 2))
    vaSpread1(imat).ForeColor = &HFF&
    
    If imat = 0 Then
       vaSpread1(2).Col = j - 1
       vaSpread1(2).CellType = 5
       vaSpread1(2).TypeHAlign = 1
       vaSpread1(2).Value = Format(matrizaporte(i), fg_Pict(6, 2))
       vaSpread1(2).ForeColor = &HFF0000
    
    Else
       
       vaSpread1(3).Col = j - 1
       vaSpread1(3).CellType = 5
       vaSpread1(3).TypeHAlign = 1
       vaSpread1(3).Value = Format(matrizaporte(i), fg_Pict(6, 2))
       vaSpread1(3).ForeColor = &HFF0000
    
    End If
    
    j = j + 1

Next i

vaSpread1(imat).MaxRows = vaSpread1(imat).MaxRows + 1
vaSpread1(imat).Row = vaSpread1(imat).MaxRows
cantsercalorias = 0: cantserproteinas = 0: cantserlipidos = 0: cantserhidratos = 0: cantserporliquida = 0: cantserporsolida = 0: cantseracgrsat = 0

For i = 1 To iaporte
    
    matrizaporte(i) = 0

Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub MoverGrlDiaVec1()

On Error GoTo Man_Error

vaSpread1(0).MaxRows = vaSpread1(0).MaxRows + 1
vaSpread1(0).Row = vaSpread1(0).MaxRows
vaSpread1(0).Col = 2
vaSpread1(0).BackColor = &H80000016
vaSpread1(0).Font.Bold = True
vaSpread1(0).Font.Size = 9
vaSpread1(0).Value = "Total Día "
            
vaSpread1(2).MaxRows = vaSpread1(2).MaxRows + 1
vaSpread1(2).Row = vaSpread1(2).MaxRows

vaSpread1(2).MaxRows = vaSpread1(2).MaxRows + 1
vaSpread1(2).Row = vaSpread1(2).MaxRows
vaSpread1(2).Col = 1
vaSpread1(2).BackColor = &H80000016
vaSpread1(2).Font.Bold = True
vaSpread1(2).Font.Size = 9
vaSpread1(2).Value = "Total Día "

If cantgrl1porsolida > 0 Then
   
   vaSpread1(0).Col = 3
   vaSpread1(0).CellType = 5
   vaSpread1(0).TypeHAlign = 1
   vaSpread1(0).Value = Format(cantgrl1porsolida, fg_Pict(6, 2))
   vaSpread1(0).ForeColor = &HFF&

   vaSpread1(2).Col = 2
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(cantgrl1porsolida, fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF&

End If

If cantgrl1porliquida > 0 Then
   
   vaSpread1(0).Col = 4
   vaSpread1(0).CellType = 5
   vaSpread1(0).TypeHAlign = 1
   vaSpread1(0).Value = Format(cantgrl1porliquida, fg_Pict(6, 2))
   vaSpread1(0).ForeColor = &HFF&
   
   vaSpread1(2).Col = 3
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(cantgrl1porliquida, fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF&

End If

If cantgrl1calorias > 0 And cantgrl1proteinas > 0 Then
   
   vaSpread1(0).Col = 5
   vaSpread1(0).CellType = 5
   vaSpread1(0).TypeHAlign = 1
   vaSpread1(0).Value = Format(CCur(((cantgrl1proteinas * 4) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(0).ForeColor = &HFF&

   vaSpread1(2).Col = 4
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(CCur(((cantgrl1proteinas * 4) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF&

End If

If cantgrl1calorias > 0 And cantgrl1lipidos > 0 Then
   
   vaSpread1(0).Col = 6
   vaSpread1(0).CellType = 5
   vaSpread1(0).TypeHAlign = 1
   vaSpread1(0).Value = Format(CCur(((cantgrl1lipidos * 9) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(0).ForeColor = &HFF&

   vaSpread1(2).Col = 5
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(CCur(((cantgrl1lipidos * 9) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF&

End If

If cantgrl1hidratos > 0 And cantgrl1calorias > 0 Then
   
   vaSpread1(0).Col = 7
   vaSpread1(0).CellType = 5
   vaSpread1(0).TypeHAlign = 1
   vaSpread1(0).Value = Format(CCur(((cantgrl1hidratos * 4) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(0).ForeColor = &HFF&

   vaSpread1(2).Col = 6
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(CCur(((cantgrl1hidratos * 4) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF&

End If

If cantgrl1acgrsat > 0 And cantgrl1calorias > 0 Then
   
   vaSpread1(0).Col = 8
   vaSpread1(0).CellType = 5
   vaSpread1(0).TypeHAlign = 1
   vaSpread1(0).Value = Format(CCur(((cantgrl1acgrsat * 9) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(0).ForeColor = &HFF&

   vaSpread1(2).Col = 7
   vaSpread1(2).CellType = 5
   vaSpread1(2).TypeHAlign = 1
   vaSpread1(2).Value = Format(CCur(((cantgrl1acgrsat * 9) / cantgrl1calorias) * 100), fg_Pict(6, 2))
   vaSpread1(2).ForeColor = &HFF&

End If

j = 9
For i = 1 To iaporte
    
    vaSpread1(0).Col = j
    vaSpread1(0).CellType = 5
    vaSpread1(0).TypeHAlign = 1
    vaSpread1(0).Value = Format(0, fg_Pict(6, 2))
    vaSpread1(0).ForeColor = &HFF&
    vaSpread1(2).Col = j - 1
    vaSpread1(2).CellType = 5
    vaSpread1(2).TypeHAlign = 1
    vaSpread1(2).Value = Format(0, fg_Pict(6, 2))
    vaSpread1(2).ForeColor = &HFF&
    j = j + 1

Next i

cantaporte = 0
For i = 1 To vaSpread1(0).MaxRows - 1
    
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 1
    
    If Val(vaSpread1(0).Value) > 0 Then
       
       For j = 9 To vaSpread1(0).maxcols
           
           vaSpread1(0).Row = i
           vaSpread1(0).Col = j
           vaSpread1(2).Col = j - 1
           
           If Val(vaSpread1(0).Value) > 0 Then
              
              cantaporte = Val(vaSpread1(0).Value)
              vaSpread1(0).Row = vaSpread1(0).MaxRows
              vaSpread1(0).Value = Format(CCur(vaSpread1(0).Value + cantaporte), fg_Pict(6, 2))
              vaSpread1(2).Row = vaSpread1(2).MaxRows
              vaSpread1(2).Value = Format(CCur(vaSpread1(2).Value + cantaporte), fg_Pict(6, 2))
           
           End If
       
       Next j
    
    End If

Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub MoverGrlDiaVec2()

On Error GoTo Man_Error

vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
vaSpread1(1).Row = vaSpread1(1).MaxRows
vaSpread1(1).Col = 2
vaSpread1(1).BackColor = &H80000016
vaSpread1(1).Font.Bold = True
vaSpread1(1).Font.Size = 9
vaSpread1(1).Value = "Total Día "
            
vaSpread1(3).MaxRows = vaSpread1(3).MaxRows + 1
vaSpread1(3).Row = vaSpread1(3).MaxRows

vaSpread1(3).MaxRows = vaSpread1(3).MaxRows + 1
vaSpread1(3).Row = vaSpread1(3).MaxRows

vaSpread1(3).Col = 1
vaSpread1(3).BackColor = &H80000016
vaSpread1(3).Font.Bold = True
vaSpread1(3).Font.Size = 9
vaSpread1(3).Value = "Total Día "

If cantgrl2porsolida > 0 Then
   
   vaSpread1(1).Col = 3
   vaSpread1(1).CellType = 5
   vaSpread1(1).TypeHAlign = 1
   vaSpread1(1).Value = Format(cantgrl2porsolida, fg_Pict(6, 2))
   vaSpread1(1).ForeColor = &HFF&

   vaSpread1(3).Col = 2
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(cantgrl2porsolida, fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF&

End If

If cantgrl2porliquida > 0 Then
   
   vaSpread1(1).Col = 4
   vaSpread1(1).CellType = 5
   vaSpread1(1).TypeHAlign = 1
   vaSpread1(1).Value = Format(cantgrl2porliquida, fg_Pict(6, 2))
   vaSpread1(1).ForeColor = &HFF&
   
   vaSpread1(3).Col = 3
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(cantgrl2porliquida, fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF&

End If

If cantgrl2calorias > 0 And cantgrl2proteinas > 0 Then
   
   vaSpread1(1).Col = 5
   vaSpread1(1).CellType = 5
   vaSpread1(1).TypeHAlign = 1
   vaSpread1(1).Value = Format(CCur(((cantgrl2proteinas * 4) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(1).ForeColor = &HFF&

   vaSpread1(3).Col = 4
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(CCur(((cantgrl2proteinas * 4) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF&

End If

If cantgrl2calorias > 0 And cantgrl2lipidos > 0 Then
   
   vaSpread1(1).Col = 6
   vaSpread1(1).CellType = 5
   vaSpread1(1).TypeHAlign = 1
   vaSpread1(1).Value = Format(CCur(((cantgrl2lipidos * 9) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(1).ForeColor = &HFF&

   vaSpread1(3).Col = 5
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(CCur(((cantgrl2lipidos * 9) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF&

End If

If cantgrl2hidratos > 0 And cantgrl2calorias > 0 Then
   
   vaSpread1(1).Col = 7
   vaSpread1(1).CellType = 5
   vaSpread1(1).TypeHAlign = 1
   vaSpread1(1).Value = Format(CCur(((cantgrl2hidratos * 4) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(1).ForeColor = &HFF&

   vaSpread1(3).Col = 6
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(CCur(((cantgrl2hidratos * 4) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF&

End If

If cantgrl2acgrsat > 0 And cantgrl2calorias > 0 Then
   
   vaSpread1(1).Col = 8
   vaSpread1(1).CellType = 5
   vaSpread1(1).TypeHAlign = 1
   vaSpread1(1).Value = Format(CCur(((cantgrl2acgrsat * 9) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(1).ForeColor = &HFF&

   vaSpread1(3).Col = 7
   vaSpread1(3).CellType = 5
   vaSpread1(3).TypeHAlign = 1
   vaSpread1(3).Value = Format(CCur(((cantgrl2acgrsat * 9) / cantgrl2calorias) * 100), fg_Pict(6, 2))
   vaSpread1(3).ForeColor = &HFF&

End If

j = 9
For i = 1 To iaporte
    
    vaSpread1(1).Col = j
    vaSpread1(1).CellType = 5
    vaSpread1(1).TypeHAlign = 1
    vaSpread1(1).Value = Format(0, fg_Pict(6, 2))
    vaSpread1(1).ForeColor = &HFF&
    vaSpread1(3).Col = j - 1
    vaSpread1(3).CellType = 5
    vaSpread1(3).TypeHAlign = 1
    vaSpread1(3).Value = Format(0, fg_Pict(6, 2))
    vaSpread1(3).ForeColor = &HFF&
    j = j + 1

Next i

cantaporte = 0
For i = 1 To vaSpread1(1).MaxRows - 1
    
    vaSpread1(1).Row = i
    vaSpread1(1).Col = 1
    
    If Val(vaSpread1(1).Value) > 0 Then
       
       For j = 9 To vaSpread1(1).maxcols
           
           vaSpread1(1).Row = i
           vaSpread1(1).Col = j
           vaSpread1(3).Col = j - 1
           
           If Val(vaSpread1(1).Value) > 0 Then
              
              cantaporte = Val(vaSpread1(1).Value)
              vaSpread1(1).Row = vaSpread1(1).MaxRows
              vaSpread1(1).Value = Format(CCur(vaSpread1(1).Value + cantaporte), fg_Pict(6, 2))
              vaSpread1(3).Row = vaSpread1(3).MaxRows
              
              If Val(vaSpread1(3).Value) < 1 Then vaSpread1(3).Value = 0
              
              vaSpread1(3).Value = Format(CCur(vaSpread1(3).Value + cantaporte), fg_Pict(6, 2))
           
           End If
       
       Next j
    
    End If

Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 2
 
 If vaSpread1(0).MaxRows < 1 Then Exit Sub
 Call ExportarExcel
            
Case 4
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub ExportarExcel()

On Error GoTo Man_Error

Dim NashXl  As excel.Application

icol = IIf(SSTab1.Tab = 0, 0, 2)

    fg_carga ""
    Set NashXl = CreateObject("excel.application")
    Set NashXl = New excel.Application
    NashXl.SheetsInNewWorkbook = 1
    NashXl.Workbooks.Add
    vaSpread1(icol).AllowMultiBlocks = True
    vaSpread1(icol).SetSelection 1, -1, vaSpread1(icol).maxcols, vaSpread1(icol).MaxRows
    vaSpread1(icol).ClipboardCopy
    NashXl.ActiveSheet.Paste
    
    'Ajustar columna
    NashXl.Cells.Select
    NashXl.Cells.EntireColumn.AutoFit
    vaSpread1(icol).AllowMultiBlocks = False: vaSpread1(icol).SetSelection 1, 0, vaSpread1(icol).maxcols, vaSpread1(icol).MaxRows
    fg_descarga
    NashXl.Visible = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

