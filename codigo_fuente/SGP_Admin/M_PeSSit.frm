VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PeSSit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pegado Especial Sitio Remoto"
   ClientHeight    =   5310
   ClientLeft      =   4590
   ClientTop       =   1035
   ClientWidth     =   9405
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpreadMes1 
      Height          =   2325
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   270
      Visible         =   0   'False
      Width           =   4335
      _Version        =   393216
      _ExtentX        =   7646
      _ExtentY        =   4101
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   6
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_PeSSit.frx":0000
      UserResize      =   1
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5310
      Left            =   8865
      TabIndex        =   1
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   9366
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin FPSpread.vaSpread vaSpreadMes2 
      Height          =   2325
      Index           =   1
      Left            =   30
      TabIndex        =   2
      Top             =   2850
      Visible         =   0   'False
      Width           =   4335
      _Version        =   393216
      _ExtentX        =   7646
      _ExtentY        =   4101
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   6
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_PeSSit.frx":055C
      UserResize      =   1
   End
   Begin FPSpread.vaSpread vaSpreadMes3 
      Height          =   2325
      Index           =   0
      Left            =   4470
      TabIndex        =   3
      Top             =   270
      Visible         =   0   'False
      Width           =   4335
      _Version        =   393216
      _ExtentX        =   7646
      _ExtentY        =   4101
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   6
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_PeSSit.frx":0AB8
      UserResize      =   1
   End
   Begin FPSpread.vaSpread vaSpreadMes4 
      Height          =   2325
      Index           =   0
      Left            =   4470
      TabIndex        =   7
      Top             =   2880
      Visible         =   0   'False
      Width           =   4335
      _Version        =   393216
      _ExtentX        =   7646
      _ExtentY        =   4101
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   6
      ScrollBars      =   0
      SelectBlockOptions=   0
      SpreadDesigner  =   "M_PeSSit.frx":1014
      UserResize      =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Diciembre 2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   4470
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Diciembre 2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   6
      Top             =   30
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Diciembre 2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   30
      TabIndex        =   5
      Top             =   2610
      Visible         =   0   'False
      Width           =   4275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Diciembre 2004"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   4470
      TabIndex        =   4
      Top             =   30
      Visible         =   0   'False
      Width           =   4275
   End
End
Attribute VB_Name = "M_PeSSit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private BtnX    As Variant
Private Spread  As Long
Dim nrodia      As String
Dim VecDia()    As String
Dim NroMes      As String
Dim mes1        As Long
Dim mes2        As Long
Dim mes3        As Long
Dim mes4        As Long

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()
    
    Call fg_centra(Me)
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i As Long
Dim j As Long

    Select Case Button.Index
    
    Case 1
        
        Let vg_codigo = ""
        Let Vg_Codigo2 = ""
        Let Vg_Codigo3 = ""
        Let Vg_Codigo4 = ""
        
        For i = 1 To vaSpreadMes1(0).MaxRows
            
            vaSpreadMes1(0).Row = i
            
            For j = 1 To vaSpreadMes1(0).maxcols
                
                Let vaSpreadMes1(0).Col = j
                
                If vaSpreadMes1(0).CellType = CellTypeButton And vaSpreadMes1(0).TypeButtonColor = 13741485 Then
                    
                    Let vg_codigo = vg_codigo & Trim(vaSpreadMes1(0).TypeButtonText) & ";"
                
                End If
            
            Next j
        
        Next i
        
        For i = 1 To vaSpreadMes2(1).MaxRows
            
            vaSpreadMes2(1).Row = i
            
            For j = 1 To vaSpreadMes2(1).maxcols
                
                Let vaSpreadMes2(1).Col = j
                
                If vaSpreadMes2(1).CellType = CellTypeButton And vaSpreadMes2(1).TypeButtonColor = 13741485 Then
                    
                    Let Vg_Codigo2 = Vg_Codigo2 & Trim(vaSpreadMes2(1).TypeButtonText) & ";"
                
                End If
            
            Next j
        
        Next i
        
        For i = 1 To vaSpreadMes3(0).MaxRows
            
            vaSpreadMes3(0).Row = i
            
            For j = 1 To vaSpreadMes3(0).maxcols
                
                Let vaSpreadMes3(0).Col = j
                
                If vaSpreadMes3(0).CellType = CellTypeButton And vaSpreadMes3(0).TypeButtonColor = 13741485 Then
                    
                    Let Vg_Codigo3 = Vg_Codigo3 & Trim(vaSpreadMes3(0).TypeButtonText) & ";"
                
                End If
            
            Next j
        
        Next i
        
        For i = 1 To vaSpreadMes4(0).MaxRows
            
            vaSpreadMes4(0).Row = i
            
            For j = 1 To vaSpreadMes4(0).maxcols
                
                Let vaSpreadMes4(0).Col = j
                
                If vaSpreadMes4(0).CellType = CellTypeButton And vaSpreadMes4(0).TypeButtonColor = 13741485 Then
                    
                    Let Vg_Codigo4 = Vg_Codigo4 & Trim(vaSpreadMes4(0).TypeButtonText) & ";"
                
                End If
            
            Next j
        
        Next i
        
        If Trim(vg_codigo) = "" And Trim(Vg_Codigo2) = "" And Trim(Vg_Codigo3) = "" And Trim(Vg_Codigo4) = "" Then
            
            Call MsgBox("No existen días seleccionado", vbInformation + vbOKOnly, "Copiado Especial")
            Exit Sub
        
        End If
        
    Case 3
        
        Let vg_codigo = ""
    
    End Select
    
    Me.Hide
    Unload Me

End Sub

Sub Inicio(tfor As String, Op As String, fecpla As String, FecFin As String, tdia As String, tMes As String)

Dim Diferencia  As String
Dim X           As Long
Dim i           As Long
Dim diafin      As String
Dim nrosem      As Long
Dim ValLcntH    As Variant
Dim j           As Long
Dim FecInicial  As String

    Me.Caption = tfor
    MsgTitulo = tfor
    'fecmin = TipM
    nrodia = tdia
    Let NroMes = tMes

    Let Diferencia = DateDiff("m", "01/" & Mid(fecpla, 5, 2) & "/" & Mid(fecpla, 1, 4), "01/" & Mid(FecFin, 5, 2) & "/" & Mid(FecFin, 1, 4))
            
    For X = 1 To (Diferencia + 1)
        
        Let FecInicial = fecpla
        
        Select Case X
            
            Case 1
            
                '------- Armar calendario
                vaSpreadMes1(0).Visible = True
                Label1(0).Visible = True
                M_PeSSit.Width = 5010
                M_PeSSit.Height = 8160
                vaSpreadMes1(0).Row = -1: vaSpreadMes1(0).Col = -1:
                vaSpreadMes1(0).BackColor = Label1(0).BackColor
                Label1(0).Caption = Meses(fg_Ctod1(fecpla & "01")) & " " & Mid(fg_pone_cero(fecpla, 6), 1, 4)
                Vg_Mes1 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                diafin = fg_mes(Mid(fg_pone_cero(fecpla, 6), 5, 2) & Mid(fg_pone_cero(fecpla, 6), 1, 4))
                mes1 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                nrosem = 1
                
                For i = 1 To diafin
                    
                    Select Case fg_Dia(fecpla & fg_pone_cero(i, 2))
                    
                    Case 1
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 7
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeCheckType = TypeCheckTypeNormal
                '        vaSpread1.TypeButtonPicture = 0
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                        nrosem = nrosem + 1
                    
                    Case 2
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 1
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                    
                    Case 3
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 2
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                    
                    Case 4
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 3
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                    
                    Case 5
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 4
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                    
                    Case 6
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 5
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                    
                    Case 7
                        
                        vaSpreadMes1(0).Row = nrosem
                        vaSpreadMes1(0).Col = 6
                        vaSpreadMes1(0).CellType = CellTypeButton
                        vaSpreadMes1(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes1(0).TypeButtonTextColor = &H800000
                        vaSpreadMes1(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes1(0).TypeButtonText = CStr(i)
                    
                    End Select
                
                Next i
                
                vaSpreadMes1(0).RetainSelBlock = False
                Let fecpla = FecInicial
                
            Case 2
            
                fecpla = Format(DateAdd("m", 1, fg_Ctod1(fecpla & "01")), "yyyymmdd")
                Label1(1).Caption = Meses(fg_Ctod1(fecpla)) & " " & Mid(fg_pone_cero(fecpla, 6), 1, 4)
                Vg_Mes2 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                fecpla = Mid(fecpla, 1, 6)
                
                vaSpreadMes2(1).Visible = True
                Label1(1).Visible = True
                M_PeSSit.Width = 5010
                M_PeSSit.Height = 8160
                vaSpreadMes2(1).Row = -1: vaSpreadMes2(1).Col = -1:
                vaSpreadMes2(1).BackColor = Label1(0).BackColor
                
                mes2 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                
                diafin = fg_mes(Mid(fg_pone_cero(fecpla, 6), 5, 2) & Mid(fg_pone_cero(fecpla, 6), 1, 4))
                
                nrosem = 1
    
                For i = 1 To diafin
                    
                    Select Case fg_Dia(fecpla & fg_pone_cero(i, 2))
                    
                    Case 1
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 7
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeCheckType = TypeCheckTypeNormal
                '        vaSpread1.TypeButtonPicture = 0
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                        nrosem = nrosem + 1
                    
                    Case 2
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 1
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                    
                    Case 3
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 2
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                    
                    Case 4
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 3
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                    
                    Case 5
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 4
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                    
                    Case 6
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 5
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                    
                    Case 7
                        
                        vaSpreadMes2(1).Row = nrosem
                        vaSpreadMes2(1).Col = 6
                        vaSpreadMes2(1).CellType = CellTypeButton
                        vaSpreadMes2(1).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes2(1).TypeButtonTextColor = &H800000
                        vaSpreadMes2(1).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes2(1).TypeButtonText = CStr(i)
                    
                    End Select
                
                Next i
                
                vaSpreadMes2(1).RetainSelBlock = False
                Let fecpla = FecInicial
                
            Case 3
                
                fecpla = Format(DateAdd("m", 2, fg_Ctod1(fecpla & "01")), "yyyymmdd")
                Label1(2).Caption = Meses(fg_Ctod1(fecpla)) & " " & Mid(fg_pone_cero(fecpla, 6), 1, 4)
                Vg_Mes3 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                fecpla = Mid(fecpla, 1, 6)
            
                vaSpreadMes3(0).Visible = True
                Label1(2).Visible = True
                M_PeSSit.Width = 9495
                M_PeSSit.Height = 5745
                vaSpreadMes3(0).Row = -1: vaSpreadMes3(0).Col = -1:
                vaSpreadMes3(0).BackColor = Label1(0).BackColor
                mes3 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                diafin = fg_mes(Mid(fg_pone_cero(fecpla, 6), 5, 2) & Mid(fg_pone_cero(fecpla, 6), 1, 4))
                nrosem = 1
                
                For i = 1 To diafin
                    
                    Select Case fg_Dia(fecpla & fg_pone_cero(i, 2))
                    
                    Case 1
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 7
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeCheckType = TypeCheckTypeNormal
                '        vaSpread1.TypeButtonPicture = 0
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                        nrosem = nrosem + 1
                    
                    Case 2
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 1
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                    
                    Case 3
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 2
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                    
                    Case 4
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 3
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                    
                    Case 5
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 4
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                    
                    Case 6
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 5
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                    
                    Case 7
                        
                        vaSpreadMes3(0).Row = nrosem
                        vaSpreadMes3(0).Col = 6
                        vaSpreadMes3(0).CellType = CellTypeButton
                        vaSpreadMes3(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes3(0).TypeButtonTextColor = &H800000
                        vaSpreadMes3(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes3(0).TypeButtonText = CStr(i)
                    
                    End Select
                
                Next i
                
                vaSpreadMes3(0).RetainSelBlock = False
                Let fecpla = FecInicial
                
            Case 4
                
                fecpla = Format(DateAdd("m", 3, fg_Ctod1(fecpla & "01")), "yyyymmdd")
                Label1(3).Caption = Meses(fg_Ctod1(fecpla)) & " " & Mid(fg_pone_cero(fecpla, 6), 1, 4)
                Vg_Mes4 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                fecpla = Mid(fecpla, 1, 6)
            
                vaSpreadMes4(0).Visible = True
                Label1(3).Visible = True
                M_PeSSit.Width = 9495
                M_PeSSit.Height = 5745
                vaSpreadMes4(0).Row = -1: vaSpreadMes4(0).Col = -1:
                vaSpreadMes4(0).BackColor = Label1(0).BackColor
                mes4 = Mid(fg_pone_cero(fecpla, 6), 5, 2)
                diafin = fg_mes(Mid(fg_pone_cero(fecpla, 6), 5, 2) & Mid(fg_pone_cero(fecpla, 6), 1, 4))
                nrosem = 1
                
                For i = 1 To diafin
                    
                    Select Case fg_Dia(fecpla & fg_pone_cero(i, 2))
                    
                    Case 1
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 7
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeCheckType = TypeCheckTypeNormal
                '        vaSpread1.TypeButtonPicture = 0
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                        nrosem = nrosem + 1
                    
                    Case 2
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 1
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                    
                    Case 3
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 2
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                    
                    Case 4
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 3
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                    
                    Case 5
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 4
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                    
                    Case 6
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 5
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                    
                    Case 7
                        
                        vaSpreadMes4(0).Row = nrosem
                        vaSpreadMes4(0).Col = 6
                        vaSpreadMes4(0).CellType = CellTypeButton
                        vaSpreadMes4(0).TypeButtonType = TypeButtonTypeTwoState
                        vaSpreadMes4(0).TypeButtonTextColor = &H800000
                        vaSpreadMes4(0).TypeButtonAlign = TypeButtonAlignBottom
                        vaSpreadMes4(0).TypeButtonText = CStr(i)
                    
                    End Select
                
                Next i
                
                vaSpreadMes4(0).RetainSelBlock = False
                Let fecpla = FecInicial
                
        End Select
        
    Next X
    '------- mover días no permitidos
    ReDim Preserve VecDia(0)
    ValLcntH = ""
    i = 0
    
    For j = 1 To Len(tdia)
        
        If Asc(Mid(tdia, j, 1)) <> 59 Then
           
           ValLcntH = ValLcntH + Mid(tdia, j, 1)
        
        Else
           
           ReDim Preserve VecDia(i): VecDia(i) = ValLcntH: ValLcntH = "": i = i + 1
        
        End If
    
    Next j
    If Trim(ValLcntH) <> "" Then ReDim Preserve VecDia(i): VecDia(i) = ValLcntH

End Sub


Private Sub vaSpreadMes1_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Long

    vaSpreadMes1(0).Row = Row: vaSpreadMes1(0).Col = Col
    For i = 0 To UBound(VecDia)

'        If NroMes = 1 Then
            
            If fg_pone_cero(mes1, 2) & Trim(fg_pone_cero(vaSpreadMes1(0).TypeButtonText, 2)) = VecDia(i) Then
                
                vaSpreadMes1(0).TypeButtonType = TypeButtonTypeNormal
                Exit Sub
            
            End If

'        End If
    
    Next i
    
    If ButtonDown = 1 Then
       
       vaSpreadMes1(0).TypeButtonColor = &HD1ADAD
    
    ElseIf ButtonDown = 0 Then
       
       vaSpreadMes1(0).TypeButtonColor = Label1(0).BackColor
    
    End If

End Sub

Private Sub vaSpreadMes2_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim i As Long

    vaSpreadMes2(1).Row = Row: vaSpreadMes2(1).Col = Col
    
    For i = 0 To UBound(VecDia)
'        If NroMes = 2 Then
            
            If fg_pone_cero(mes2, 2) & Trim(fg_pone_cero(vaSpreadMes2(1).TypeButtonText, 2)) = VecDia(i) Then
                
                vaSpreadMes2(1).TypeButtonType = TypeButtonTypeNormal
                Exit Sub
            
            End If
'        End If
    
    Next i
    
    If ButtonDown = 1 Then
       
       vaSpreadMes2(1).TypeButtonColor = &HD1ADAD
    
    ElseIf ButtonDown = 0 Then
       
       vaSpreadMes2(1).TypeButtonColor = Label1(1).BackColor
    
    End If

End Sub

Private Sub vaSpreadMes3_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

Dim i As Long

    vaSpreadMes3(0).Row = Row: vaSpreadMes3(0).Col = Col
    
    For i = 0 To UBound(VecDia)

'        If NroMes = 3 Then
            
            If fg_pone_cero(mes3, 2) & Trim(fg_pone_cero(vaSpreadMes3(0).TypeButtonText, 2)) = VecDia(i) Then
                
                vaSpreadMes3(0).TypeButtonType = TypeButtonTypeNormal
                Exit Sub
            
            End If

'        End If
    
    Next i
    
    If ButtonDown = 1 Then
       
       vaSpreadMes3(0).TypeButtonColor = &HD1ADAD
    
    ElseIf ButtonDown = 0 Then
       
       vaSpreadMes3(0).TypeButtonColor = Label1(2).BackColor
    
    End If

End Sub

Private Sub vaSpreadMes4_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

Dim i As Long

    vaSpreadMes4(0).Row = Row: vaSpreadMes4(0).Col = Col
    
    For i = 0 To UBound(VecDia)

'        If NroMes = 3 Then
            
            If fg_pone_cero(mes4, 2) & Trim(fg_pone_cero(vaSpreadMes4(0).TypeButtonText, 2)) = VecDia(i) Then
                
                vaSpreadMes4(0).TypeButtonType = TypeButtonTypeNormal
                Exit Sub
            
            End If

'        End If
    
    Next i
    
    If ButtonDown = 1 Then
       
       vaSpreadMes4(0).TypeButtonColor = &HD1ADAD
    
    ElseIf ButtonDown = 0 Then
       
       vaSpreadMes4(0).TypeButtonColor = Label1(2).BackColor
    
    End If

End Sub

