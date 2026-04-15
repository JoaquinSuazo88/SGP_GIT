VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_ForComPrexCeCo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definir Excepción Formato de Compras por Centro de Costo"
   ClientHeight    =   9090
   ClientLeft      =   960
   ClientTop       =   1380
   ClientWidth     =   16905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   16905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   6135
      Left            =   135
      TabIndex        =   2
      Top             =   2805
      Width           =   16605
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1320
         TabIndex        =   15
         Top             =   5040
         Width           =   900
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   16
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   2280
         TabIndex        =   13
         Top             =   5040
         Width           =   3780
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   14
            Top             =   135
            Width           =   3675
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4665
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   16215
         _Version        =   393216
         _ExtentX        =   28601
         _ExtentY        =   8229
         _StockProps     =   64
         ButtonDrawMode  =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "M_ForComPrexCeCo.frx":0000
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Eliminar Ingrediiente"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   14625
         TabIndex        =   6
         Top             =   5400
         Width           =   1440
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Agregar Ingrediente"
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   12945
         TabIndex        =   5
         Top             =   5400
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   3360
      TabIndex        =   1
      Top             =   480
      Width           =   10455
      Begin VB.CommandButton Command1 
         Caption         =   "Limpiar"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   8640
         TabIndex        =   12
         Top             =   1440
         Width           =   1440
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   2850
         TabIndex        =   3
         Top             =   450
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
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
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
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
         MaxLength       =   10
         MultiLine       =   0   'False
         PasswordChar    =   ""
         IncHoriz        =   0.25
         BorderGrayAreaColor=   -2147483637
         NoPrefix        =   0   'False
         ScrollV         =   0   'False
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   2850
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   1425
         _Version        =   196608
         _ExtentX        =   2514
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
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         ThreeDInsideStyle=   0
         ThreeDInsideHighlightColor=   -2147483637
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
         ButtonStyle     =   3
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
         AutoBeep        =   0   'False
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
         OnFocusPosition =   1
         ControlType     =   0
         Text            =   ""
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
         DateMax         =   "00000000"
         DateMin         =   "00000000"
         TimeMax         =   "000000"
         TimeMin         =   "000000"
         TimeString1159  =   ""
         TimeString2359  =   ""
         DateDefault     =   "00000000"
         TimeDefault     =   "000000"
         TimeStyle       =   0
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         PopUpType       =   0
         DateCalcY2KSplit=   60
         CaretPosition   =   0
         IncYear         =   1
         IncMonth        =   1
         IncDay          =   0
         IncHour         =   0
         IncMinute       =   0
         IncSecond       =   0
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha (dd/mm/aaaa)"
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
         Index           =   2
         Left            =   960
         TabIndex        =   11
         Top             =   1005
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   4560
         TabIndex        =   8
         Top             =   450
         Width           =   5655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Left            =   960
         TabIndex        =   4
         Top             =   525
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   4080
         Picture         =   "M_ForComPrexCeCo.frx":1C2F
         Top             =   360
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   4590
         TabIndex        =   9
         Top             =   480
         Width           =   5655
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16905
      _ExtentX        =   29819
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
Attribute VB_Name = "M_ForComPrexCeCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo      As String
Dim codigo    As Long
Dim MsgTitulo As String
Dim Est       As Boolean
Dim estmod    As Boolean

Private Sub Command1_Click(Index As Integer)
    
On Error GoTo Man_Error
    
    Dim RS    As New ADODB.Recordset
    Dim Sql   As String
    Dim i     As Long
    Dim tiene As Long
    Dim msg, Response
    
    Dim lisnom  As String
    Dim liscod  As String
    Dim lisprov As String
    
    tiene = 0
    
Select Case Index
    
    Case 0
        
        Sql = Trim(LimpiaDato(fpText.text))
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS = vg_db.Execute("sgpadm_Sel_planif_excep '" & Sql & "'")
        If RS.EOF Then
           
           MsgBox "No Existen Planificaciones, para este Convenio.", vbInformation, "Ingredientes"
            RS.Close
            Set RS = Nothing

           Exit Sub
        
        End If
        
        RS.Close
        Set RS = Nothing
        
        vg_left = fpayuda.Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente Real", "IngRealCasino"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        'fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
        'valido que ingrediente no exista en la grilla
'        If vaSpread1.DataRowCnt = 0 Then
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 2
            
            If Trim(vaSpread1.text) = vg_codigo Then
               
               MsgBox "Ingrediente ya existe, seleccione otro", vbInformation, "Ingredientes"
               Exit Sub
            
            End If
        
        Next i
            
        Sql = Trim(LimpiaDato(fpText.text))
            
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgpadm_Sel_materialsap '" & Sql & "','" & vg_codigo & "'")
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           MsgBox "Ingrediente no tiene asociado material sap", vbInformation, "Ingredientes"
           Exit Sub
        
        End If
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 2
        vaSpread1.text = vg_codigo
        
        vaSpread1.Col = 3
        vaSpread1.text = vg_nombre
        
        i = 1
        
        lisnom = ""
        liscod = ""
        lisprov = ""
        
        Do While Not RS.EOF
            
            lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & LimpiaDato(RS!proveedor) & " - " & LimpiaDato(RS!fcs_CodMaterial) & " - " & RS!fcs_DenMaterial
            liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS!fcs_CodMaterial
            lisprov = lisprov & IIf(lisprov <> "", Chr$(9), "") & RS!proveedor
               
            vaSpread1.Col = 4
            vaSpread1.TypeComboBoxList = lisnom
            
            vaSpread1.Col = 5
            vaSpread1.TypeComboBoxList = lisprov
            
            vaSpread1.Col = 6
            vaSpread1.TypeComboBoxList = liscod

            If i = 1 Then
                
                vaSpread1.Col = 4
                vaSpread1.TypeComboBoxCurSel = 0
                vaSpread1.Col = 5 'proveedor
                vaSpread1.TypeComboBoxCurSel = 0
                vaSpread1.Col = 6
                vaSpread1.TypeComboBoxCurSel = 0
            
            End If
            
            RS.MoveNext: i = i + 1
        
        Loop
        
        vaSpread1.Col = 7
        vaSpread1.CellType = CellTypeDate
        vaSpread1.TypeSpin = False
        vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
        vaSpread1.TypeDateMin = "01011973"
        vaSpread1.TypeDateMax = "31125000"
        vaSpread1.text = ""
       
        vaSpread1.Col = 8
        vaSpread1.CellType = CellTypeDate
        vaSpread1.TypeSpin = False
        vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
        vaSpread1.TypeDateMin = "01011973"
        vaSpread1.TypeDateMax = "31125000"
        vaSpread1.text = ""
       
        vaSpread1.Col = 10
        vaSpread1.text = 0
        
        Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(4).Visible = True: Toolbar1.Buttons(5).Visible = False
        
        fpText.ControlType = ControlTypeStatic
        
    Case 1
        
        tiene = 0
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If Val(vaSpread1.text) = 1 Then
                
                tiene = 1
            
            End If
        
        Next
        
        If tiene = 0 Then
            
            MsgBox "No existen datos a eliminar", vbInformation, "Eliminar Ingredientes"
            Exit Sub
        
        End If
        
        If tiene = 1 Then
            
            msg = "¿Esta Seguro Que desea Eliminar?"
            Response = MsgBox(msg, 4 + 32, "Sistema Gestión")
            
            If Response = 6 Then
                
                For i = vaSpread1.MaxRows To 1 Step -1 'vaSpread1.MaxRows
                    
                    vaSpread1.Row = i
                    vaSpread1.Col = 1
                    
                    If Val(vaSpread1.text) = 1 Then
                        
                        vaSpread1.DeleteRows vaSpread1.Row, 1
                    
                    End If
                
                Next
            
            Else
                
                Exit Sub
            
            End If
        
        End If
        vaSpread1.MaxRows = vaSpread1.DataRowCnt
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
        fpText.ControlType = ControlTypeStatic
    
    Case 2 ''limpiar
        
        vaSpread1.MaxRows = 0
        Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(4).Visible = True
        Toolbar1.Buttons(5).Visible = False
        
        Command1(0).Enabled = False
        Command1(1).Enabled = False
        fpText.text = ""
        fpText.ControlType = ControlTypeNormal
        Call fpText_Change
    
    End Select

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

Me.HelpContextID = vg_OpcM
Me.Height = 9570
Me.Width = 16995
MsgTitulo = "Definir Excepcion Formato Compras"
fg_centra Me

modo = ""
Est = True
estmod = False
Gl_Mo_Botones Me, 20
Gl_Ac_Botones Me, 1, 17, modo
vaSpread1.MaxRows = 0
Command1(2).Enabled = False

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next

End Sub


Private Sub Form_Unload(Cancel As Integer)

'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
If fpText.text = "" Then fpayuda.Caption = "": Exit Sub
Sql = Trim(LimpiaDato(fpText.text))

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fpayuda.Caption = ""
   vaSpread1.MaxRows = 0
   
   If vaSpread1.MaxRows = 0 Then
       
       Command1(0).Enabled = True
   
   End If
   
   Command1(1).Enabled = False
   Command1(2).Enabled = False
   Toolbar1.Buttons.item(1).Visible = False
   Toolbar1.Buttons.item(2).Visible = True
   Exit Sub

End If

fpayuda.Caption = Trim(RS!Cli_nombre)
Command1(0).Enabled = True
Command1(2).Enabled = True
RS.Close: Set RS = Nothing
trae_formatocompra

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next

End Sub

Sub trae_formatocompra()
  
On Error GoTo Man_Error

  Dim RS As New ADODB.Recordset
  Dim Sql As String
  Dim lisnom As String
  Dim liscod As String
  Dim lisprov As String
  Dim Aux_ing_codigo As String
  vaSpread1.Visible = False
  vaSpread1.MaxRows = 0
  
  Sql = Trim(LimpiaDato(fpText.text))
  
  Aux_ing_codigo = ""
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  
  Set RS = vg_db.Execute("sgpadm_Sel_excepcionformato_V02'" & Sql & "'")
  If RS.EOF Then RS.Close: Set RS = Nothing:  vaSpread1.Visible = True: Exit Sub
  
  Do While Not RS.EOF
    
    If Trim(Aux_ing_codigo) <> Trim(RS!ing_codigo) Then
       
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
    
       vaSpread1.Col = 2
       vaSpread1.text = RS!ing_codigo
    
       vaSpread1.Col = 3
       vaSpread1.text = RS!ing_nombre
    
       lisnom = ""
       liscod = ""
       lisprov = ""
       
       lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & LimpiaDato(Trim(RS!proveedor)) & " - " & LimpiaDato(Trim(RS!fcs_CodMaterial)) & " - " & Trim(RS!fcs_DenMaterial)
       liscod = liscod & IIf(liscod <> "", Chr$(9), "") & Trim(RS!fcs_CodMaterial)
       lisprov = lisprov & IIf(lisprov <> "", Chr$(9), "") & Trim(RS!proveedor)
    
       vaSpread1.Col = 4
       vaSpread1.TypeComboBoxList = lisnom
       
       vaSpread1.Col = 4
       vaSpread1.TypeComboBoxCurSel = 0
       
       vaSpread1.Col = 5
       vaSpread1.TypeComboBoxList = lisprov
       
       vaSpread1.Col = 5
       vaSpread1.TypeComboBoxCurSel = 0
       
       vaSpread1.Col = 6
       vaSpread1.TypeComboBoxList = liscod
       
       vaSpread1.Col = 6
       vaSpread1.TypeComboBoxCurSel = 0
    
       vaSpread1.Col = 7
       vaSpread1.CellType = CellTypeDate
       vaSpread1.TypeSpin = False
       vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
       vaSpread1.TypeDateMin = "01011973"
       vaSpread1.TypeDateMax = "31125000"
       vaSpread1.text = IIf(IsNull(RS!Fecha_Inicio), "", RS!Fecha_Inicio)
       
       vaSpread1.Col = 8
       vaSpread1.CellType = CellTypeDate
       vaSpread1.TypeSpin = False
       vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY
       vaSpread1.TypeDateMin = "01011973"
       vaSpread1.TypeDateMax = "31125000"
       vaSpread1.text = IIf(IsNull(RS!Fecha_Termino), "", RS!Fecha_Termino)
       
       vaSpread1.Col = 10
       vaSpread1.text = 0
       
       Aux_ing_codigo = Trim(RS!ing_codigo)
    
    Else
    
       lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & LimpiaDato(Trim(RS!proveedor)) & " - " & LimpiaDato(RS!fcs_CodMaterial) & " - " & RS!fcs_DenMaterial
       liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS!fcs_CodMaterial
       lisprov = lisprov & IIf(lisprov <> "", Chr$(9), "") & RS!proveedor

       vaSpread1.Col = 4
       vaSpread1.TypeComboBoxList = lisnom
       
       vaSpread1.Col = 5
       vaSpread1.TypeComboBoxList = lisprov
       
       vaSpread1.Col = 6
       vaSpread1.TypeComboBoxList = liscod
    
    End If
    
        
    RS.MoveNext
  
  Loop

salir:
  
  vaSpread1.Visible = True

   If vaSpread1.MaxRows > 0 Then
      
      Toolbar1.Buttons(4).Visible = True: Toolbar1.Buttons(5).Visible = False
    
   End If
   
   RS.Close
   Set RS = Nothing

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next

End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

Case 120
    
    Image1_Click

End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next

End Sub

Private Sub Image1_Click()

On Error GoTo Man_Error

    vg_left = fpayuda.Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo: fpayuda.Caption = vg_nombre
   
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next
   
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

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 10
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3

    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 10
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 10
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 10
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 10
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 10
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
           
           vaSpread1.Col = 10
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

Dim RS       As New ADODB.Recordset
Dim MyBuffer As String
Dim Sql      As String
Dim i        As Long
Dim FecIni   As Date
Dim FecTer   As Date

Screen.MousePointer = 11
Select Case Button.Index

Case 1 'grabar

    'validar fechas inicio y termino
    For i = 1 To vaSpread1.MaxRows
    
        vaSpread1.Row = i
        
        FecIni = "00:00:00"
        FecTer = "00:00:00"
        
        vaSpread1.Col = 7
        
        If Trim(vaSpread1.text) <> "" Then
           
           If Not IsDate(Format(vaSpread1.text, "dd/mm/yyyy")) Then
           
              MsgBox "Fecha inicio no corresponde formato fecha...", vbExclamation + vbOKOnly, MsgTitulo
              Exit Sub
           
           End If
           
           FecIni = Format(vaSpread1.text, "dd/mm/yyyy")
        
        End If
              
        vaSpread1.Col = 8
        If Trim(vaSpread1.text) <> "" Then
                
           If Not IsDate(Format(vaSpread1.text, "dd/mm/yyyy")) Then
           
              MsgBox "Fecha termino no corresponde formato fecha...", vbExclamation + vbOKOnly, MsgTitulo
              Exit Sub
           
           End If
           
           FecTer = Format(vaSpread1.text, "dd/mm/yyyy")
        
        End If
        
        If FecTer < FecIni And FecIni <> "0:00:00" Then
        
           MsgBox "Fecha inicio es mayor fecha termino...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
         
        End If
        
        If FecIni <> "0:00:00" And FecTer = "0:00:00" Then
        
           MsgBox "Debe ingresar fecha termino...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If FecTer <> "0:00:00" And FecIni = "0:00:00" Then
        
           MsgBox "Debe ingresar fecha inicio...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
    Next i
    
    'registrar Log sistema
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Actualizar"), CStr(Me.HelpContextID), "", "", "")
    
    Let MyBuffer = ""
    Let Sql = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaFormato>"
    
    For i = 1 To vaSpread1.MaxRows
        
        MyBuffer = MyBuffer & " <Formato"
        vaSpread1.Row = i
        vaSpread1.Col = 2
        MyBuffer = MyBuffer & " CI = " & Chr(34) & Trim(LimpiaDato(vaSpread1.text)) & Chr(34)
        vaSpread1.Col = 5
        MyBuffer = MyBuffer & " CP = " & Chr(34) & Trim(LimpiaDato(vaSpread1.text)) & Chr(34)
        vaSpread1.Col = 6
        MyBuffer = MyBuffer & " CM = " & Chr(34) & Trim(LimpiaDato(vaSpread1.text)) & Chr(34)
        vaSpread1.Col = 7
        
        If IsDate(Trim(LimpiaDato(vaSpread1.text))) Then
        
           MyBuffer = MyBuffer & " FI = " & Chr(34) & Format(Trim(LimpiaDato(vaSpread1.text)), "yyyymmdd") & Chr(34)
           
        Else
        
           MyBuffer = MyBuffer & " FI = " & Chr(34) & Trim(LimpiaDato(vaSpread1.text)) & Chr(34)
        
        End If
        
        vaSpread1.Col = 8
        
        If IsDate(Trim(LimpiaDato(vaSpread1.text))) Then
        
           MyBuffer = MyBuffer & " FT = " & Chr(34) & Format(Trim(LimpiaDato(vaSpread1.text)), "yyyymmdd") & Chr(34)
        
        Else
        
           MyBuffer = MyBuffer & " FT = " & Chr(34) & Trim(LimpiaDato(vaSpread1.text)) & Chr(34)
        
        End If
        
        Let MyBuffer = MyBuffer & "/>"
    
    Next
    
    Let MyBuffer = MyBuffer & "</GrabaFormato>"
    Sql = Trim(LimpiaDato(Me.fpText))
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_ins_excepcionformato_V02 '" & MyBuffer & "', '" & Sql & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
    Set RS = Nothing
    
    MsgBox "Registros grabados con exito", vbInformation, "Grabar"
    Screen.MousePointer = 0
    Exit Sub

Case 4 'imprimir
    
'    If fpDateTime1.text = "" Then
'
'           MsgBox "Debe Ingresar Fecha a Imprimir", vbExclamation + vbOKOnly, MsgTitulo
'           Screen.MousePointer = 0
'           Exit Sub
'
'    End If
'
'    If vaSpread1.MaxRows < 1 Then
'
'       MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Screen.MousePointer = 0
'       Exit Sub
'
'    End If
'
    'registrar Log sistema
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), CStr(Me.HelpContextID), "", "", "")
    
    E_ExcepcionFormatoCompras.Show 1

    'I_Formatocompras_CC
    fg_descarga

'Case 7 ' Copiar

'    'registrar Log sistema
'    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso_Copiar"), CStr(Me.HelpContextID), "", "", "")
'    M_CopiarExcepcionFormato.Show 1
    
Case 9 'salir
    
    Me.Hide
    Unload Me

End Select
Screen.MousePointer = 0

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Exit Sub
Resume Next

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo Man_Error

Select Case ButtonMenu

    Case "Copiar Formato de Compras"
        
        'registrar Log sistema
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso_Copiar"), CStr(Me.HelpContextID), "", "", "")
        M_CopiarExcepcionFormato.Show 1
    
    Case "Bach - Input Ingresar", "Bach - Input Modificar", "Bach - Input Eliminar"
    
        
        'Abrimos el Commondialog con ShowOpen
        CD.DialogTitle = "Seleccione un archivo excel"
        CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
        CD.DefaultExt = "*.xls|*.xlsx"
        CD.FilterIndex = 2
        CD.Flags = cdlOFNFileMustExist
        CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
        CD.FileName = ""
        CD.ShowOpen
    
        'Si seleccionamos un archivo mostramos la excepcion
        If CD.FileName <> "" Then
           
           Dim msg As String
           msg = IIf(ButtonMenu = "Bach - Input Ingresar", "Esta Seguro Agregar Excepción Formato", IIf(ButtonMenu = "Bach - Input Modificar", "Esta Seguro Modificar Excepción Formato", "Esta Seguro Eliminar Excepción Formato"))
           If MsgBox(msg, vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
             
              Exit Sub
       
           End If

           ValidarPlantillaExcel CD.FileName, IIf(ButtonMenu = "Bach - Input Ingresar", 1, IIf(ButtonMenu = "Bach - Input Modificar", 2, 3))
                
        Else
            'Si no mostramos un texto de advertencia de que no se seleccionó _
            ninguno, ya que FileName devuelve una cadena vacía
            MsgBox "No seleccionó ningún archivo"
    
        End If
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub ValidarPlantillaExcel(NombreArchivo As String, OpMenu As Integer)

On Error GoTo Man_Error

Dim PathXls         As String
Dim File_Ext        As String
Dim NomHoja         As String
Dim dbexcel         As Database
Dim cn              As ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim MyBuffer        As String

Dim Ceco            As String
Dim CodIng          As String
Dim CodMat          As String
Dim codpro          As String
Dim FecIni          As Date
Dim FecTer          As Date

Dim EstPro          As Boolean
Dim NomArchivoExcel As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

PathXls = Trim(NombreArchivo)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))
NomHoja = "Hoja1$"
With cn
     
     Select Case File_Ext
        
        ' Excel 97/2003
        Case "XLS"
          
          .Provider = "Microsoft.Jet.OLEDB.4.0"
          .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
          
        ' Excel 2010
        Case "XLSX"

          .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
          .ConnectionString = "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
     
     End Select
     
     .Open

End With

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<ExcepFor>"
EstPro = True

RsExcel.Open ("SELECT * FROM [" & NomHoja & "]"), cn

If RsExcel.EOF Then Exit Sub

RsExcel.MoveFirst

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Or IsNull(RsExcel.Fields(0).Value) Then Exit Do
           
      Ceco = ""
      CodIng = ""
      CodMat = ""
      codpro = ""
      FecIni = "00:00:00"
      FecTer = "00:00:00"
       
      If Trim(RsExcel.Fields(0).Name) <> "Ceco" Or Trim(RsExcel.Fields(1).Name) <> "Ing# Ingrediente" Or _
         Trim(RsExcel.Fields(2).Name) <> "Material SAP" Or Trim(RsExcel.Fields(3).Name) <> "Proveedor" Or _
         Trim(RsExcel.Fields(4).Name) <> "Fecha Inicio" Or Trim(RsExcel.Fields(5).Name) <> "Fecha Termino" Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo
         EstPro = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor codigo ceco esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo
         EstPro = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código ingrediente esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo
         EstPro = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(2).Value) Or Trim(RsExcel.Fields(2).Value) = "" Or Not IsNumeric(RsExcel.Fields(2).Value) Then
   
         MsgBox "Valor código material SAP esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         EstPro = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(3).Value) Or Trim(RsExcel.Fields(3).Value) = "" Then
   
         MsgBox "Valor código proveedor esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         EstPro = False
         Exit Do
      
      End If
      
'      If IsNull(RsExcel.Fields(4).Value) Or Trim(RsExcel.Fields(4).Value) = "" Then
'
'         MsgBox "Valor fecha inicio esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
'         ValidarPlantillaExcel = False
'         Exit Do
'
'      End If
'
'      If IsNull(RsExcel.Fields(5).Value) Or Trim(RsExcel.Fields(5).Value) = "" Or Not IsNumeric(RsExcel.Fields(5).Value) Then
'
'         MsgBox "Valor fecha termino esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
'         ValidarPlantillaExcel = False
'         Exit Do
'
'      End If
                
      Ceco = RsExcel.Fields(0).Value
      CodIng = RsExcel.Fields(1).Value
      CodMat = RsExcel.Fields(2).Value
      codpro = RsExcel.Fields(3).Value
            
      If IsDate(RsExcel.Fields(4).Value) Then
         
         FecIni = RsExcel.Fields(4).Value
      
      End If
      
      If IsDate(RsExcel.Fields(5).Value) Then
      
         FecTer = RsExcel.Fields(5).Value
          
      End If
      
      If OpMenu = 1 Or OpMenu = 2 Then
      
         If FecTer < FecIni And FecIni <> "0:00:00" Then
        
            MsgBox "Debe fecha inicio es menor fecha termino...", vbExclamation + vbOKOnly, MsgTitulo
            EstPro = False
            Exit Do
        
         End If
        
         If FecIni <> "0:00:00" And FecTer = "0:00:00" Then
        
            MsgBox "Debe ingresar fecha termino...", vbExclamation + vbOKOnly, MsgTitulo
            EstPro = False
            Exit Do
        
         End If
        
         If FecTer <> "0:00:00" And FecIni = "0:00:00" Then
        
            MsgBox "Debe ingresar fecha inicio...", vbExclamation + vbOKOnly, MsgTitulo
            EstPro = False
            Exit Do
        
        End If
        
      End If
                
      Ceco = Replace(Trim(Ceco), Chr(34), "&quot;")
      Ceco = Replace(Trim(Ceco), Chr(38), "&amp;")
      Ceco = Replace(Trim(Ceco), Chr(39), "&apos;")
      Ceco = Replace(Trim(Ceco), Chr(60), "&lt;")
      Ceco = Replace(Trim(Ceco), Chr(62), "&gt;")
    
      CodIng = Replace(Trim(CodIng), Chr(34), "&quot;")
      CodIng = Replace(Trim(CodIng), Chr(38), "&amp;")
      CodIng = Replace(Trim(CodIng), Chr(39), "&apos;")
      CodIng = Replace(Trim(CodIng), Chr(60), "&lt;")
      CodIng = Replace(Trim(CodIng), Chr(62), "&gt;")
    
      CodMat = Replace(Trim(CodMat), Chr(34), "&quot;")
      CodMat = Replace(Trim(CodMat), Chr(38), "&amp;")
      CodMat = Replace(Trim(CodMat), Chr(39), "&apos;")
      CodMat = Replace(Trim(CodMat), Chr(60), "&lt;")
      CodMat = Replace(Trim(CodMat), Chr(62), "&gt;")
    
      codpro = Replace(Trim(codpro), Chr(34), "&quot;")
      codpro = Replace(Trim(codpro), Chr(38), "&amp;")
      codpro = Replace(Trim(codpro), Chr(39), "&apos;")
      codpro = Replace(Trim(codpro), Chr(60), "&lt;")
      codpro = Replace(Trim(codpro), Chr(62), "&gt;")
    
      FecIni = Replace(Trim(FecIni), Chr(34), "&quot;")
      FecIni = Replace(Trim(FecIni), Chr(38), "&amp;")
      FecIni = Replace(Trim(FecIni), Chr(39), "&apos;")
      FecIni = Replace(Trim(FecIni), Chr(60), "&lt;")
      FecIni = Replace(Trim(FecIni), Chr(62), "&gt;")
    
      FecTer = Replace(Trim(FecTer), Chr(34), "&quot;")
      FecTer = Replace(Trim(FecTer), Chr(38), "&amp;")
      FecTer = Replace(Trim(FecTer), Chr(39), "&apos;")
      FecTer = Replace(Trim(FecTer), Chr(60), "&lt;")
      FecTer = Replace(Trim(FecTer), Chr(62), "&gt;")
     
      MyBuffer = MyBuffer & " <ExF"
      MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
      MyBuffer = MyBuffer & " CI = " & Chr(34) & CodIng & Chr(34)
      MyBuffer = MyBuffer & " CM = " & Chr(34) & CodMat & Chr(34)
      MyBuffer = MyBuffer & " CP = " & Chr(34) & codpro & Chr(34)
      
      If IsDate(FecIni) Then
      
         MyBuffer = MyBuffer & " FI = " & Chr(34) & Format(FecIni, "yyyymmdd") & Chr(34)
      
      Else
      
         MyBuffer = MyBuffer & " FI = " & Chr(34) & FecIni & Chr(34)
    
      End If
      
      If IsDate(FecTer) Then
      
         MyBuffer = MyBuffer & " FT = " & Chr(34) & Format(FecTer, "yyyymmdd") & Chr(34)
      
      Else
      
         MyBuffer = MyBuffer & " FT = " & Chr(34) & FecTer & Chr(34)
         
      End If
              
      MyBuffer = MyBuffer & "/>"
   
   'End If
   
   DoEvents
           
   RsExcel.MoveNext
   
Loop
        
MyBuffer = MyBuffer & "</ExcepFor>"

RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing

If EstPro Then

      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
                            
      'registrar Log sistema
      If OpMenu = 1 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Agregar"), Me.HelpContextID, "", "", "")
        
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcepcionFormatoCompras_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
      
      ElseIf OpMenu = 2 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
         
         Set RS = vg_db.Execute("sgpadm_Upd_XmlExcepcionFormatoCompras_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
      
      Else
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), Me.HelpContextID, "", "", "")
      
         Set RS = vg_db.Execute("sgpadm_Del_XmlExcepcionFormatoCompras_V01 '" & MyBuffer & "', '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "'")
      
      End If

      If Not RS.EOF Then

         If RS(0) > 0 Then

            MsgBox RS(1)
            fg_descarga

            RS.Close
            Set RS = Nothing
   
            'registrar Log sistema eliminar & Agregado
      
            If OpMenu = 1 Then
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")
               
            Else
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregado"), Me.HelpContextID, "", "", "")
               
            End If
            
            Exit Sub
      
         Else
   
            MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo
            
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
            xlApp.Selection.CurrentRegion.Columns.AutoFit
            xlApp.Selection.CurrentRegion.Rows.AutoFit
    
            xlApp.Columns("A:A").Select
            xlApp.Selection.Delete Shift:=xlToLeft
  
            NomArchivoExcel = fg_ArchivoXls(IIf(OpMenu = 1, "ReporteError_RutasDespachosEliminar", "ReporteError_RutasDespachos_Agregar"))
                    
            xlWb.Close True, NomArchivoExcel

            Dim XL As New excel.Application 'Crea el objeto excel
            XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
            XL.Visible = True
            XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
            '-- Cerrar Excel
            xlApp.Quit
            '-------> Release Excel references
            Set xlWs = Nothing
            Set xlWb = Nothing
            Set xlApp = Nothing
   
            'registrar Log sistema eliminar & Agregado
      
            If OpMenu = 1 Then
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
               
            Else
            
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), Me.HelpContextID, "", "", "")
               
            End If
         
         End If
      
      End If
      RS.Close
      Set RS = Nothing

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    
On Error GoTo Man_Error

    Dim RS As New ADODB.Recordset
    Dim tiene, i As Integer
    If vaSpread1.ActiveCol = 1 Then
        
        vaSpread1.Col = 1
        vaSpread1.Row = vaSpread1.ActiveRow
        
        If vaSpread1.text = "1" Then
            
            vaSpread1.text = 0
        
        Else
            
            vaSpread1.text = 1
        
        End If
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If Val(vaSpread1.text) = 1 Then
                
                tiene = 1
                Me.Command1(1).Enabled = True
                Exit For
            
            End If
        
        Next
    
    End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
fpText.ControlType = ControlTypeStatic

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
    
On Error GoTo Man_Error

    Select Case Col
        
        Case 4
            
            Dim indice As Long
            Dim CodVal As String
            vaSpread1.Row = Row
            vaSpread1.Col = 4: indice = vaSpread1.TypeComboBoxCurSel
            vaSpread1.Col = 5: vaSpread1.TypeComboBoxCurSel = indice
            vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = indice
            
            Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(4).Visible = True: Toolbar1.Buttons(5).Visible = False

        End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error
            
Select Case Col

    Case 7, 8
    
            Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(4).Visible = True: Toolbar1.Buttons(5).Visible = False

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

Select Case KeyCode

Case 46
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    If vaSpread1.Col = 7 Or vaSpread1.Col = 8 Then
       
       vaSpread1.text = ""
    
       Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False
       Toolbar1.Buttons(4).Visible = True: Toolbar1.Buttons(5).Visible = False
       
    End If
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub








