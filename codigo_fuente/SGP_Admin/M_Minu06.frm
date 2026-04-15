VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Minu06 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aporte Nutricional Día"
   ClientHeight    =   5925
   ClientLeft      =   1305
   ClientTop       =   1905
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   11220
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Detalle Nutricional x Plato Servicio"
      TabPicture(0)   =   "M_Minu06.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Resumen Nutricional x Servicio"
      TabPicture(1)   =   "M_Minu06.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame1(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame1(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   0
         Left            =   -74880
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
            SpreadDesigner  =   "M_Minu06.frx":0038
            VisibleCols     =   8
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   1
         Left            =   -74880
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
            SpreadDesigner  =   "M_Minu06.frx":6EC8
            VisibleCols     =   8
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   2
         Left            =   120
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
            SpreadDesigner  =   "M_Minu06.frx":D82D
            UserResize      =   1
            VisibleCols     =   7
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Index           =   3
         Left            =   120
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
            SpreadDesigner  =   "M_Minu06.frx":11E33
            VisibleCols     =   7
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5925
      Left            =   10590
      TabIndex        =   9
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   10451
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Minu06"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
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
fg_centra Me
fg_carga ""
Toolbar1.ImageList = partida.IL1
Toolbar1.Buttons.Clear
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

vaSpread1(0).MaxRows = 0: vaSpread1(0).MaxCols = 8
vaSpread1(1).MaxRows = 0: vaSpread1(1).MaxCols = 8
vaSpread1(2).MaxRows = 0: vaSpread1(2).MaxCols = 7
vaSpread1(3).MaxRows = 0: vaSpread1(3).MaxCols = 7
SSTab1.Tab = 0

'------- Llenar Nutrientes
Set RS = vg_db.Execute("min_s_nutriente 9, 0, ''")
iaporte = 0
ReDim Preserve matrizaporte(0)
If Not RS.EOF Then
   vaSpread1(0).MaxRows = 0
   Do While Not RS.EOF
      vaSpread1(0).Row = 0
      vaSpread1(0).MaxCols = vaSpread1(0).MaxCols + 1
      vaSpread1(0).Col = vaSpread1(0).MaxCols
      vaSpread1(0).Value = Trim(RS!Ntrnt_Name)

      vaSpread1(1).Row = 0
      vaSpread1(1).MaxCols = vaSpread1(1).MaxCols + 1
      vaSpread1(1).Col = vaSpread1(1).MaxCols
      vaSpread1(1).Value = Trim(RS!Ntrnt_Name)
      
      vaSpread1(2).Row = 0
      vaSpread1(2).MaxCols = vaSpread1(2).MaxCols + 1
      vaSpread1(2).Col = vaSpread1(2).MaxCols
      vaSpread1(2).Value = Trim(RS!Ntrnt_Name)
      
      vaSpread1(3).Row = 0
      vaSpread1(3).MaxCols = vaSpread1(3).MaxCols + 1
      vaSpread1(3).Col = vaSpread1(3).MaxCols
      vaSpread1(3).Value = Trim(RS!Ntrnt_Name)
      
      iaporte = iaporte + 1
      RS.MoveNext
   Loop
   ReDim Preserve matrizaporte(iaporte)
End If
RS.Close: Set RS = Nothing

M_Minu02.vaSpread1.Row = M_Minu02.vaSpread1.ActiveRow
M_Minu02.vaSpread1.Col = M_Minu02.vaSpread1.ActiveCol
Select Case M_Minu02.vaSpread1.Col
Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250
    icol = M_Minu02.vaSpread1.Col
Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251
    icol = M_Minu02.vaSpread1.Col - 1
Case 4, 12, 20, 28, 36, 44, 52, 60, 68, 76, 84, 92, 100, 108, 116, 124, 132, 140, 148, 156, 164, 172, 180, 188, 196, 204, 212, 220, 228, 236, 244
    icol = M_Minu02.vaSpread1.Col - 2
Case 8, 16, 24, 32, 40, 48, 56, 64, 72, 80, 88, 96, 104, 112, 120, 128, 136, 144, 152, 160, 168, 176, 184, 192, 200, 208, 216, 224, 232, 240, 248, 256
    icol = M_Minu02.vaSpread1.Col - 6
End Select
M_Minu02.vaSpread1.Col = icol + 7
If Val(M_Minu02.vaSpread1.Value) = 0 Then Exit Sub
'fecha = Val(M_Minu02.vaSpread1.Value)
M_Minu02.vaSpread1.Row = 0
M_Minu02.vaSpread1.Col = icol + 1: fecha = Format(vg_fecha, "yyyymm") & fg_pone_cero(Right(M_Minu02.vaSpread1.Text, 2), 2)
M_Minu06.Caption = "Aporte Nutricional Día " & Format(Mid(fecha, 7, 2) & "/" & Mid(fecha, 5, 2) & "/" & Mid(fecha, 1, 4))
cantsercalorias = 0: cantserproteinas = 0: cantserlipidos = 0: cantserhidratos = 0: proc = 0: cantserporsolida = 0: cantserporliquida = 0: cantseracgrsat = 0
cantgrl1calorias = 0: cantgrl1proteinas = 0: cantgrl1lipidos = 0: cantgrl1hidratos = 0: proc = 0: cantgrl1porsolida = 0: cantgrl1porliquida = 0: cantgrl1acgrsat = 0
cantgrl2calorias = 0: cantgrl2proteinas = 0: cantgrl2lipidos = 0: cantgrl2hidratos = 0: cantgrl2porsolida = 0: cantgrl2porliquida = 0: cantgrl2acgrsat = 0
auxserv = 0
For i = 1 To iaporte
    matrizaporte(i) = 0
Next i
Set RS = vg_db.Execute("min_s_diaminutaporte 1, '" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codsegmento & ", " & fecha & "")
If Not RS.EOF Then
   Do While Not RS.EOF
      If RS!codigo_servicio <> auxserv Then
         If proc = 1 Then
            MoverCtrServicio
         End If
         If RS!codigo_servicio = 2 Or RS!codigo_servicio = 23 Then
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
         vaSpread1(imat).Value = Trim(RS!Serv_Name)
         auxserv = RS!codigo_servicio
         nomserv = Trim(RS!Serv_Name)
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
      vaSpread1(imat).Value = RS!Rcpe_Desc
      cantporsolida = 0
      If RS!Rcpe_Uom_Code_No = 2 Then
         vaSpread1(imat).Col = 3
         vaSpread1(imat).TypeHAlign = 1
         vaSpread1(imat).Value = Format(RS!Rcpe_Prtn_Size_Val, fg_Pict(4, 2))
         vaSpread1(imat).ForeColor = &HFF0000
         cantporsolida = RS!Rcpe_Prtn_Size_Val
         cantserporsolida = CCur(cantserporsolida + RS!Rcpe_Prtn_Size_Val)
         If imat = 0 Then
            cantgrl1porsolida = CCur(cantgrl1porsolida + RS!Rcpe_Prtn_Size_Val)
         Else
            cantgrl2porsolida = CCur(cantgrl2porsolida + RS!Rcpe_Prtn_Size_Val)
         End If
      ElseIf RS!Rcpe_Uom_Code_No = 4 Then
         vaSpread1(imat).Col = 4
         vaSpread1(imat).TypeHAlign = 1
         vaSpread1(imat).Value = Format(RS!Rcpe_Prtn_Size_Val, fg_Pict(4, 2))
         vaSpread1(imat).ForeColor = &HFF0000
         cantserporliquida = CCur(cantserporliquida + RS!Rcpe_Prtn_Size_Val)
         If imat = 0 Then
            cantgrl1porliquida = CCur(cantgrl1porliquida + RS!Rcpe_Prtn_Size_Val)
         Else
            cantgrl2porliquida = CCur(cantgrl2porliquida + RS!Rcpe_Prtn_Size_Val)
         End If
      End If
      cantcalorias = 0: cantproteinas = 0: cantlipidos = 0: canthidratos = 0: cantacgrsat = 0
      j = 9: i = 1
      Set RS1 = vg_db.Execute("min_s_calaporterecetas " & RS!Rcpe_No & "")
      If Not RS1.EOF Then
         Do While Not RS1.EOF
            vaSpread1(imat).Col = j
            vaSpread1(imat).CellType = 5
            vaSpread1(imat).TypeHAlign = 1
            vaSpread1(imat).Value = Format(RS1!valordietetico, fg_Pict(6, 2))
            vaSpread1(imat).ForeColor = &HFF0000
            If RS1!Ntrnt_Code = 1 And RS1!valordietetico > 0 Then
               cantcalorias = RS1!valordietetico
               cantsercalorias = CCur(cantsercalorias + RS1!valordietetico)
               If imat = 0 Then
                  cantgrl1calorias = CCur(cantgrl1calorias + RS1!valordietetico)
               Else
                  cantgrl2calorias = CCur(cantgrl2calorias + RS1!valordietetico)
               End If
            ElseIf RS1!Ntrnt_Code = 3 And RS1!valordietetico > 0 Then
               cantproteinas = RS1!valordietetico
               cantserproteinas = CCur(cantserproteinas + RS1!valordietetico)
               If imat = 0 Then
                  cantgrl1proteinas = CCur(cantgrl1proteinas + RS1!valordietetico)
               Else
                  cantgrl2proteinas = CCur(cantgrl2proteinas + RS1!valordietetico)
               End If
            ElseIf RS1!Ntrnt_Code = 4 And RS1!valordietetico > 0 Then
               cantlipidos = RS1!valordietetico
               cantserlipidos = CCur(cantserlipidos + RS1!valordietetico)
               If imat = 0 Then
                  cantgrl1lipidos = CCur(cantgrl1lipidos + RS1!valordietetico)
               Else
                  cantgrl2lipidos = CCur(cantgrl2lipidos + RS1!valordietetico)
               End If
            ElseIf RS1!Ntrnt_Code = 5 And RS1!valordietetico > 0 Then
               canthidratos = RS1!valordietetico
               cantserhidratos = CCur(cantserhidratos + RS1!valordietetico)
               If imat = 0 Then
                  cantgrl1hidratos = CCur(cantgrl1hidratos + RS1!valordietetico)
               Else
                  cantgrl2hidratos = CCur(cantgrl2hidratos + RS1!valordietetico)
               End If
            ElseIf RS1!Ntrnt_Code = 18 And RS1!valordietetico > 0 Then
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
      RS1.Close: Set RS1 = Nothing
      If cantcalorias > 0 And cantproteinas > 0 Then
         vaSpread1(imat).Col = 5
         vaSpread1(imat).CellType = 5
         vaSpread1(imat).TypeHAlign = 1
         vaSpread1(imat).Value = Format(CCur(((cantproteinas * 4) / cantcalorias) * 100), fg_Pict(6, 2))
         vaSpread1(imat).ForeColor = &HFF0000
      End If
      If cantcalorias > 0 And cantlipidos > 0 Then
         vaSpread1(imat).Col = 6
         vaSpread1(imat).CellType = 5
         vaSpread1(imat).TypeHAlign = 1
         vaSpread1(imat).Value = Format(CCur(((cantlipidos * 9) / cantcalorias) * 100), fg_Pict(6, 2))
         vaSpread1(imat).ForeColor = &HFF0000
      End If
      If canthidratos > 0 And cantcalorias > 0 Then
         vaSpread1(imat).Col = 7
         vaSpread1(imat).CellType = 5
         vaSpread1(imat).TypeHAlign = 1
         vaSpread1(imat).Value = Format(CCur(((canthidratos * 4) / cantcalorias) * 100), fg_Pict(6, 2))
         vaSpread1(imat).ForeColor = &HFF0000
      End If
      If cantcalorias > 0 And cantacgrsat > 0 Then
         vaSpread1(imat).Col = 8
         vaSpread1(imat).CellType = 5
         vaSpread1(imat).TypeHAlign = 1
         vaSpread1(imat).Value = Format(CCur(((cantacgrsat * 9) / cantcalorias) * 100), fg_Pict(6, 2))
         vaSpread1(imat).ForeColor = &HFF0000
      End If

'      If cantporsolida > 0 And cantcalorias > 0 Then
'         vaSpread1(imat).Col = 8
'         vaSpread1(imat).CellType = 5
'         vaSpread1(imat).TypeHAlign = 1
'         vaSpread1(imat).Value = Format(CCur(cantcalorias / cantporsolida), fg_Pict(6, 2))
'         vaSpread1(imat).ForeColor = &HFF0000
'      End If
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
RS.Close: Set RS = Nothing: fg_descarga
End Sub

Private Sub MoverCtrServicio()
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
Else
   vaSpread1(3).MaxRows = vaSpread1(3).MaxRows + 1
   vaSpread1(3).Row = vaSpread1(3).MaxRows
   vaSpread1(3).Col = 1
   vaSpread1(3).Font.Bold = True
   vaSpread1(3).Font.Size = 9
   vaSpread1(3).Value = nomserv
End If
If cantserporsolida > 0 Then
   vaSpread1(imat).Col = 3
   vaSpread1(imat).CellType = 5
   vaSpread1(imat).TypeHAlign = 1
   vaSpread1(imat).Value = Format(cantserporsolida, fg_Pict(6, 2))
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
End If
If cantserporliquida > 0 Then
   vaSpread1(imat).Col = 4
   vaSpread1(imat).CellType = 5
   vaSpread1(imat).TypeHAlign = 1
   vaSpread1(imat).Value = Format(cantserporliquida, fg_Pict(6, 2))
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


'If cantserporsolida > 0 And cantsercalorias > 0 Then
'   vaSpread1(imat).Col = 8
'   vaSpread1(imat).CellType = 5
'   vaSpread1(imat).TypeHAlign = 1
'   vaSpread1(imat).Value = Format(CCur(cantsercalorias / cantserporsolida), fg_Pict(6, 2))
'   vaSpread1(imat).ForeColor = &HFF&
'   If imat = 0 Then
'      vaSpread1(2).Col = 7
'      vaSpread1(2).CellType = 5
'      vaSpread1(2).TypeHAlign = 1
'      vaSpread1(2).Value = Format(CCur(cantsercalorias / cantserporsolida), fg_Pict(6, 2))
'      vaSpread1(2).ForeColor = &HFF0000
'   Else
'      vaSpread1(3).Col = 7
'      vaSpread1(3).CellType = 5
'      vaSpread1(3).TypeHAlign = 1
'      vaSpread1(3).Value = Format(CCur(cantsercalorias / cantserporsolida), fg_Pict(6, 2))
'      vaSpread1(3).ForeColor = &HFF0000
'   End If
'End If
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
End Sub

Private Sub MoverGrlDiaVec1()
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


'If cantgrl1porsolida > 0 And cantgrl1calorias > 0 Then
'   vaSpread1(0).Col = 8
'   vaSpread1(0).CellType = 5
'   vaSpread1(0).TypeHAlign = 1
'   vaSpread1(0).Value = Format(CCur(cantgrl1calorias / cantgrl1porsolida), fg_Pict(6, 2))
'   vaSpread1(0).ForeColor = &HFF&

'   vaSpread1(2).Col = 7
'   vaSpread1(2).CellType = 5
'   vaSpread1(2).TypeHAlign = 1
'   vaSpread1(2).Value = Format(CCur(cantgrl1calorias / cantgrl1porsolida), fg_Pict(6, 2))
'   vaSpread1(2).ForeColor = &HFF&
'End If

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
       For j = 9 To vaSpread1(0).MaxCols
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
End Sub

Private Sub MoverGrlDiaVec2()
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

'If cantgrl2porsolida > 0 And cantgrl2calorias > 0 Then
'   vaSpread1(1).Col = 8
'   vaSpread1(1).CellType = 5
'   vaSpread1(1).TypeHAlign = 1
'   vaSpread1(1).Value = Format(CCur(cantgrl2calorias / cantgrl2porsolida), fg_Pict(6, 2))
'   vaSpread1(1).ForeColor = &HFF&

'   vaSpread1(3).Col = 7
'   vaSpread1(3).CellType = 5
'   vaSpread1(3).TypeHAlign = 1
'   vaSpread1(3).Value = Format(CCur(cantgrl2calorias / cantgrl2porsolida), fg_Pict(6, 2))
'   vaSpread1(3).ForeColor = &HFF&
'End If

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
       For j = 9 To vaSpread1(1).MaxCols
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
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    Me.Hide
    Unload Me
End Select
End Sub
