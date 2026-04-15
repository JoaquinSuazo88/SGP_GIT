VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PlantillaMenuI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plantilla Menu"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   11370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   10455
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4335
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   9975
         _Version        =   393216
         _ExtentX        =   17595
         _ExtentY        =   7646
         _StockProps     =   64
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
         SpreadDesigner  =   "M_PlatillaMenuI.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1515
         TabIndex        =   3
         Top             =   675
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   2
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2565
         Picture         =   "M_PlatillaMenuI.frx":1866
         Top             =   600
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3015
         TabIndex        =   5
         Top             =   675
         Width           =   7215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   780
         Width           =   705
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3060
         TabIndex        =   6
         Top             =   720
         Width           =   7215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6675
      Left            =   10740
      TabIndex        =   7
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   11774
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_PlantillaMenuI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim Row  As Long

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
    Call fg_carga("")
    Me.HelpContextID = vg_OpcM
    Call fg_centra(Me)
    Let Me.Height = 7050
    Let Me.Width = 11460
    modo = ""
    Gl_Mo_Botones Me, 19
    Gl_Ac_Botones Me, 1, 16, modo
    
    Call FormatearDatos
    Call fg_descarga
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    
On Error GoTo Man_Error

    Select Case KeyCode
        
        Case 120
            
            If Index = 0 Then Image1_Click 2
    
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim RS       As New ADODB.Recordset

Select Case Index
    
     Case 0
        
        Set RS = vg_db.Execute("sgpadm_Sel_ServicioBloque " & IIf(Val(fpLongInteger1(0).Value) = 0, -1, Val(fpLongInteger1(0).Value)) & "")
        
        If RS.EOF Then
            
            RS.Close
            Set RS = Nothing
            fpayuda(2).Caption = ""
            Exit Sub
        
        End If
        fpayuda(2).Caption = Trim(RS!ser_nombre)
        RS.Close: Set RS = Nothing

End Select
    
    Call MoverGrilla(fpLongInteger1(0).text)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Call MoverGrilla(fpLongInteger1(0).text)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub FormatearDatos()

On Error GoTo Man_Error

    Let vaSpread1.MaxRows = 0
    Let fpLongInteger1(0).Value = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub MoverGrilla(servicio)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i  As Long

Set RS = vg_db.Execute("sgpadm_Sel_ServicioPlatillaMenu " & Val(fpLongInteger1(0).Value) & "")
        
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
i = 1

Do While Not RS.EOF
            
    vaSpread1.Row = i
   
   vaSpread1.Col = 2
   vaSpread1.text = RS!id
   
   vaSpread1.Col = 3
   vaSpread1.text = RS!descripcion
   
   i = i + 1
   
   RS.MoveNext

Loop
        
RS.Close
Set RS = Nothing

vaSpread1.Visible = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
    
    Select Case Index
        
        Case 2
            
            Let vg_left = fpayuda(2).Left + 2300
            Let vg_nombre = ""
            Let vg_codigo = ""
            Call B_TabEst.LlenaDatos("a_servicio", "", "Servicio", "SerBlo")
            Call B_TabEst.Show(1)
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpLongInteger1(0).Value = Val(vg_codigo)
            fpLongInteger1(0).SetFocus
            fpayuda(2).Caption = vg_nombre
    
    End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS       As New ADODB.Recordset
Dim Sql      As String
Dim IdBloque As Long

Select Case Button.Index
        
    Case 1, 3
            
        If ValidaDatos = False Then Exit Sub
    
        vg_codservicio = Val(fpLongInteger1(0).Value)
        vg_nombre = fpayuda(2).Caption
        
        If Button.Index = 1 Then
               
           Vg_PlaSer = "1"
               
        ElseIf Button.Index = 3 Then
            
           Vg_PlaSer = "2"
                
           If Not ValidaDatosGrilla Then Exit Sub
            
        End If
            
        Unload M_PlantillaMenuII
            
        Call M_PlantillaMenuII.Show(1)
            
    Case 6 'Eliminar registro plantilla menu
        
        If Not ValidaDatosGrilla Then Exit Sub
        
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
                   
        'registrar Log sistema eliminación
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), CStr(Me.HelpContextID), "", "", "")
                   
        Sql = ""
        Sql = " & IdBloque & "
        Set RS = vg_db.Execute("sgpadm_Del_MinutaBloque " & Sql & "")
                   
        If Not RS.EOF Then
                      
           If UCase(RS(0)) = "OK" Then
                         
               'registrar Log sistema Eliminar
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")
                         
               MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, MsgTitulo
               EstSel = True
               vaSpread1.DeleteRows Row, 1
               vaSpread1.MaxRows = vaSpread1.MaxRows - 1
                      
           Else
                         
               MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
                      
               'registrar Log sistema error Eliminacion
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
                      
           End If
                   
        End If
                
    Case 9 'Salir
            
        Me.Hide
        Unload Me
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    'registrar Log sistema error Eliminacion
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
    
End Sub

Private Function ValidaDatos() As Boolean

On Error GoTo Man_Error

Let ValidaDatos = True
     
If Len(fpLongInteger1(0).text) = 0 Then
       
   Call MsgBox("Debe Ingresar Servicio", vbInformation, Me.Caption)
   Call fpLongInteger1(0).SetFocus
   Let ValidaDatos = False
   Exit Function
    
End If
    
Exit Function
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

Private Function ValidaDatosGrilla() As Boolean

'validar que haya un datos seleccionado en la grilla

Dim i As Long

ValidaDatosGrilla = False
Row = 0
vg_IDBloque = 0

For i = 1 To vaSpread1.MaxRows
                   
    vaSpread1.Row = i
    vaSpread1.Col = 1
                   
    If vaSpread1.text = "1" Then
                      
       ValidaDatosGrilla = True
       vaSpread1.Col = 2
       vg_IDBloque = Val(vaSpread1.text)
       Row = vaSpread1.Row
                   
    End If
               
Next i
               
If Not ValidaDatosGrilla Then
                  
   MsgBox "Seleccione un bloque del detalle de la grilla", vbExclamation + vbOKOnly, Me.Caption
   Exit Function
               
End If

End Function
