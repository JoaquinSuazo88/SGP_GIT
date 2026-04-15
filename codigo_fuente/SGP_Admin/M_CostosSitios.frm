VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CostosSitios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Costos Sitios"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13920
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   13920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   1905
      TabIndex        =   2
      Top             =   435
      Width           =   9630
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   675
         Width           =   1335
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   1050
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
         ButtonStyle     =   1
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
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
         Text            =   "2018"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "yyyy"
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   330
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha"
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
         Left            =   840
         TabIndex        =   7
         Top             =   1125
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Org. Compras"
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
         Index           =   6
         Left            =   840
         TabIndex        =   6
         Top             =   405
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Costo"
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
         Left            =   840
         TabIndex        =   5
         Top             =   780
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   13635
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5295
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   13365
         _Version        =   393216
         _ExtentX        =   23574
         _ExtentY        =   9340
         _StockProps     =   64
         ColsFrozen      =   6
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
         EditModeReplace =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   19
         MaxRows         =   10
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_CostosSitios.frx":0000
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   13920
      _ExtentX        =   24553
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_CostosSitios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo      As String
Dim Est       As Boolean
Dim Msgtitulo As String
Public lc_Aux As String

Private Sub Combo1_Click(Index As Integer)

On Error GoTo Man_Error

MoverDatosGrilla

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Msgtitulo = "Parametro Costos Sitios"
fg_centra Me
fpDateTime1.text = Format(Date, "yyyy")
modo = ""
Est = True
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo

Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True

Toolbar1.Buttons(5).Visible = False
Toolbar1.Buttons(6).Visible = True

Combo1(0).Clear

If lc_Aux = "CosCom" Then
    
    Msgtitulo = "Parametro Costo Patrón Comercial"
    Me.Caption = "Parametro Costo Patrón Comercial"
    Combo1(0).AddItem "COMERCIAL" & Space(150) & "(0)"

'Else
'
'    Msgtitulo = "Parametro Costo Patrón Techo"
'    Me.Caption = "Parametro Costo Patrón Techo"
'    Combo1(0).AddItem "TECHO" & Space(150) & "(1)"

End If

Combo1(0).ListIndex = 0
vaSpread1.MaxRows = 0
Est = False
'vaSpread1.Col = -1: vaSpread1.Row = -1
'vaSpread1.BackColor = Shape1(1).FillColor
'vaSpread1.Lock = False
'vaSpread1.Row = -1
'vaSpread1.Col = 1: vaSpread1.BackColor = Shape1(2).FillColor: vaSpread1.Col = 2: vaSpread1.BackColor = Shape1(2).FillColor
'vaSpread1.Col = 1: vaSpread1.text = "": vaSpread1.ColHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", True, False)
'vaSpread1.Col = 2: vaSpread1.text = "": vaSpread1.ColHidden = IIf(fg_codigocbo(Combo1, 0, 1, "") = "0", True, False)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Sub MoverDatosGrilla()

On Error GoTo Man_Error

If Est Then Exit Sub

Dim i        As Long
Dim anomes   As Long
Dim RS       As New ADODB.Recordset
Dim Ceco     As String
Dim Regimen  As Long
Dim servicio As Long

Frame1.Enabled = False
    
vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpamd_Sel_ParametroCostoComercial '" & Trim(fpText.text) & "', " & Val(Format(fpDateTime1.Value, "yyyy")) & ", '" & Trim(Mid(Combo1(0).text, 1, 150)) & "'")
Ceco = ""
Regimen = 0
servicio = 0

fg_carga ""

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      DoEvents
        
      If RS!cli_codigo <> Ceco Or RS!reg_codigo <> Regimen Or RS!ser_codigo <> servicio Then
         
         vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         vaSpread1.Row = vaSpread1.MaxRows
         Ceco = RS!cli_codigo
         Regimen = RS!reg_codigo
         servicio = RS!ser_codigo
      
         vaSpread1.Col = 1
         vaSpread1.text = Trim(RS!cli_codigo)
      
         vaSpread1.Col = 2
         vaSpread1.text = Trim(RS!cli_nombre)
      
         vaSpread1.Col = 3
         vaSpread1.text = Trim(RS!reg_codigo)
         
         vaSpread1.Col = 4
         vaSpread1.text = Trim(RS!reg_nombre)
      
         vaSpread1.Col = 5
         vaSpread1.text = Trim(RS!ser_codigo)
      
         vaSpread1.Col = 6
         vaSpread1.text = Trim(RS!ser_nombre)
      
      End If
      
      vaSpread1.Col = (Val(Mid(RS!pcp_anomes, 5, 2)) + 6)
      vaSpread1.text = IIf(RS!pcp_valor > 0, RS!pcp_valor, "")
         
      vaSpread1.Col = 19
      vaSpread1.text = 0
      
      RS.MoveNext
   
   Loop
   
   vaSpread1.Col = -1: vaSpread1.Row = -1
   vaSpread1.Lock = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", False, True)
   vaSpread1.SetActiveCell 3, 1
'   vaSpread1.SetFocus
   
End If
RS.Close
Set RS = Nothing
vaSpread1.Visible = True

Frame1.Enabled = True
    
fg_descarga

Exit Sub
Man_Error:
fg_descarga
Frame1.Enabled = True
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub fpDateTime1_Change()

On Error GoTo Man_Error

MoverDatosGrilla

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub fpDateTime1_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i  As Long
Dim j  As Long

Select Case Button.Index

Case 3
    
    If Trim(fpText.text) = "" Or Trim(fpDateTime1.text) = "" Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True

    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    
    Frame1.Enabled = False

Case 5

Case 7
    
    MoverDatosGrilla

Case 10
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDatosGrilla
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True

    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    
    Frame1.Enabled = True

Case 12
    
    Dim Ceco        As String
    Dim descripcion As String
    Dim Regimen     As Long
    Dim servicio    As Long
    Dim valor       As Double
    Dim IndDia      As Long
    Dim MyBuffer    As String
    fg_carga ""
    
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaComercial>"
            
    For i = 1 To vaSpread1.MaxRows
        
        DoEvents
        vaSpread1.Row = i
        
        vaSpread1.Col = 1
        Ceco = vaSpread1.text
        
        vaSpread1.Col = 3
        Regimen = vaSpread1.text
        
        vaSpread1.Col = 5
        servicio = vaSpread1.text
        
        indmes = 1
         
         For j = 7 To vaSpread1.maxcols
            
            vaSpread1.Col = j
            valor = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
            DoEvents
            
            If vaSpread1.ForeColor = &HFF0000 Then
               
               MyBuffer = MyBuffer & " <Comercial"
               MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
               MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
               MyBuffer = MyBuffer & " Ser = " & Chr(34) & servicio & Chr(34)
               MyBuffer = MyBuffer & " Mes = " & Chr(34) & fg_pone_cero(indmes, 2) & Chr(34)
               MyBuffer = MyBuffer & " Val = " & Chr(34) & valor & Chr(34)
               
               MyBuffer = MyBuffer & "/>"
                            
                            
            
            End If
            
            indmes = indmes + 1
        
        Next j
    
    Next i
    
    MyBuffer = MyBuffer & "</GrabaComercial>"
    
    Set RS = vg_db.Execute("sgpadm_Ins_XmlParametroComercial '" & MyBuffer & "', 'COMERCIAL', " & Val(Format(fpDateTime1.Value, "yyyy")) & "")
    
    If Not RS.EOF Then
    
    
       If RS(0) > 0 Then
                  
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Msgtitulo
               
       Else
       
          MsgBox "Proceso Finalizo Correctamente ", vbInformation + vbOKOnly, Msgtitulo
       
       End If

    
    End If
    
    RS.Close
    Set RS = Nothing
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True

    Toolbar1.Buttons(5).Visible = False
    Toolbar1.Buttons(6).Visible = True
    
    Frame1.Enabled = True
    fg_descarga

Case 15
    

Case 18
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
If Err = -2147467259 Or 2147217900 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub

If modo = "" Then modo = "M"
'
    If ChangeMade = True And modo = "M" Then
          
       vaSpread1.Row = Row
       vaSpread1.Col = Col
       vaSpread1.ForeColor = &HFF0000

       vaSpread1.Col = 19
       vaSpread1.text = 1
       
   End If

   Gl_Ac_Botones Me, 1, 0, modo
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True

   Toolbar1.Buttons(5).Visible = False
   Toolbar1.Buttons(6).Visible = True
   
   Frame1.Enabled = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub
