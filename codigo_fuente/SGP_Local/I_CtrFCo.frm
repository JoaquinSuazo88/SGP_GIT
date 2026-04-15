VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-47E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail.dll"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_CtrFCo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control Facturas Compras"
   ClientHeight    =   2010
   ClientLeft      =   3810
   ClientTop       =   3810
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   8460
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   5535
      Left            =   30
      TabIndex        =   10
      Top             =   2040
      Visible         =   0   'False
      Width           =   8295
      Begin VB.TextBox Text1 
         Height          =   5055
         Index           =   0
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   11
         Top             =   240
         Visible         =   0   'False
         Width           =   7935
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   5055
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   7935
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   390
      Width           =   8235
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   4095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   5730
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   2265
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   390
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   720
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Lugar Fisico"
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
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
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
         Left            =   2520
         TabIndex        =   13
         Top             =   810
         Visible         =   0   'False
         Width           =   75
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Traspaso"
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
         Index           =   1
         Left            =   4260
         TabIndex        =   6
         Top             =   810
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Folio"
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
         Left            =   240
         TabIndex        =   5
         Top             =   810
         Width           =   690
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2685
         Picture         =   "I_CtrFCo.frx":0000
         Top             =   300
         Width           =   480
      End
      Begin VB.Label label2 
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
         Index           =   7
         Left            =   240
         TabIndex        =   3
         Top             =   465
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3150
         TabIndex        =   8
         Top             =   390
         Width           =   4800
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3180
         TabIndex        =   9
         Top             =   435
         Width           =   4800
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8460
      _ExtentX        =   14923
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATMAILLibCtl.ChilkatMailMan oMail 
      Left            =   720
      OleObjectBlob   =   "I_CtrFCo.frx":030A
      Top             =   3960
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   0
      OleObjectBlob   =   "I_CtrFCo.frx":0408
      Top             =   3840
   End
End
Attribute VB_Name = "I_CtrFCo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS        As New ADODB.Recordset
Dim RS1       As New ADODB.Recordset
Dim RS2       As New ADODB.Recordset
Dim RS3       As New ADODB.Recordset
Dim RS4       As New ADODB.Recordset
Dim RS5       As New ADODB.Recordset
Dim RS6       As New ADODB.Recordset
Dim i         As Integer
Dim isel      As Integer
Dim tiperr    As Long
Dim numenv    As Long
Dim totenv    As Long
Dim corenv    As Long
Dim MsgTitulo As String
Dim tipinf    As String
Dim est       As Boolean
Dim OpEnvioAx As Boolean

Private Sub Combo1_Click(Index As Integer)

On Error GoTo error

If est Then Exit Sub

Dim RS As New ADODB.Recordset
Dim LugFis As String

Select Case Index

Case 0

    If Not OpEnvioAx And tipinf <> "T" Then

       Label2(2).Caption = IIf(fg_codigocbo(Combo1, 0, 1, "") = "C", "Doc. SGP", "Doc. P. Electronica")
       
    End If

Case 1
    
    LugFis = fg_codigocbo(Combo1, 1, 4, 1)

    Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'ParLugFis'")
    If Not RS.EOF Then
       
       est = True
       vg_db.Execute ("sgp_Upd_Param 1, '" & MuestraCasino(1) & "', 'ParLugFis', '', '', '" & LugFis & "'")
       est = False
    
    Else
       
       vg_db.Execute ("sgp_Ins_Param 'ParLugFis','Parametro Lugar Fisico','C', '" & LugFis & "', '" & MuestraCasino(1) & "'")
    
    End If
    RS.Close
    Set RS = Nothing

End Select

Exit Sub
error:
fg_descarga
MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Activate()

On Error GoTo error

fg_descarga

Exit Sub
error:
fg_descarga
MsgBox Err.Description, vbCritical

End Sub

Private Sub Form_Load()

On Error GoTo error

Dim RS As New ADODB.Recordset
Me.Height = 2430
Me.Width = 8550

OpEnvioAx = True
'-------> Validar si el contrato tiene opción de envio sap
Set RS = vg_db.Execute("SELECT * FROM b_casinointerfaz with (nolock) WHERE cai_cencos = '" & MuestraCasino(1) & "'")

If Not RS.EOF And (tipinf = "C" Or tipinf = "T") Then
   
   Do While Not RS.EOF
   
      If RS!cai_codtii = 1 Then
      
         Me.Height = 8175
         Exit Do
   
      ElseIf RS!cai_codtii = 6 Then
    
         OpEnvioAx = False
         
         If tipinf = "C" Then
            
            Combo1(1).Visible = False
            Label1(1).Visible = False
            Combo1(0).Visible = True
            Label2(1).Visible = True
            Label2(1).Caption = "Tipo Dcoumento"
            Combo1(0).Clear
            Combo1(0).AddItem "Cfc Manual" & Space(1) & "(C)"
            Combo1(0).AddItem "Cfc Portal Electronico" & Space(1) & "(P)"
            Combo1(0).ListIndex = 0
         
         ElseIf tipinf = "T" Then
         
            Combo1(1).Visible = False
            Label1(1).Visible = False
          
            Combo1(0).Clear
            Combo1(0).AddItem "Entrada" & Space(150) & "(1)"
            Combo1(0).AddItem "Salida" & Space(150) & "(0)"
         
         End If
         
         Exit Do
      
      End If

      RS.MoveNext
    Loop
    
End If
RS.Close: Set RS = Nothing

fg_centra Me

MsgTitulo = "Control Facturas Compras"
Me.HelpContextID = vg_OpcM

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   ")
BtnX.Visible = True
BtnX.ToolTipText = "Vista Previa"
BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0)
BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar")
BtnX.Visible = True
BtnX.Enabled = False
BtnX.ToolTipText = IIf(Me.Height = 2115, "Generar Folio", "Enviar Documento CFC"): BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0)
BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico")
BtnX.Visible = True
BtnX.ToolTipText = IIf(tipinf = "C", "Historico CFC", IIf(tipinf = "T", "Historico TEC", "Historio FOFI"))
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0)
BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_GenerarArchivo", , tbrDefault, "A_GenerarArchivo")
BtnX.Visible = True
BtnX.ToolTipText = IIf(tipinf = "C", "Generar Facturación MANUAL", IIf(tipinf = "T", "Generar Traspaso de Salida", "FOFI"))
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0)
BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_ExploradorWindows", , tbrDefault, "A_ExploradorWindows")
BtnX.Visible = True
BtnX.ToolTipText = "Ver Carpeta"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0)
BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    ")
BtnX.Visible = True
BtnX.ToolTipText = "Salir"

fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)

If OpEnvioAx Then
   
   Combo1(0).Clear
   Combo1(0).AddItem "Entrada" & Space(150) & "(1)"
   Combo1(0).AddItem "Salida" & Space(150) & "(0)"

End If
Text1(0).Visible = False

est = True
'-------> Cargar Lugar Fisico
CargarDatoCombo Combo1, 1, "LugarFisico_AX", "cli_", "LugFis", "A"

'-------> buscar a_param lugar fisico
Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'ParLugFis'")
If Not RS.EOF Then
   
   Combo1(1).ListIndex = fg_buscacbostring(Combo1, 1, 4, (RS!par_valor))

End If
RS.Close
Set RS = Nothing
est = False

Exit Sub
error:
fg_descarga
MsgBox Err.Description, vbCritical

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT * FROM b_clientes with (nolock) WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0")
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyUp(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

Select Case KeyCode

    Case 120
    
    '    Image1_Click 0

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
        
        vg_left = fpayuda(0).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpText.text = vg_codigo
        fpayuda(0).Caption = vg_nombre

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim RS1         As New ADODB.Recordset
Dim RS5         As New ADODB.Recordset
Dim RS6         As New ADODB.Recordset
Dim LugarFisico As String
Dim periodo     As String
Dim Sql         As String

Select Case Button.Index

Case 1
    
    Dim fecemi As Long
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT * FROM b_clientes with (nolock) WHERE cli_codigo = '" & fpText.text & "' AND cli_tipo = 0")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close
    Set RS = Nothing
    
    If tipinf = "T" And Combo1(0).ListIndex = -1 Then
       
       MsgBox "Debe selecionar tipo traspaso", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    ElseIf tipinf = "G" Then
       
       Combo1(0).ListIndex = 0
    
    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    If tipinf = "C" Then
       
       Set RS = vg_db.Execute("SELECT MAX(toc_fecrem) AS fecemi FROM b_totcompras with (nolock) WHERE toc_codbod = " & vg_codbod & " AND toc_tipdoc in (select at.tdo_IdCodigo from a_tipodocumento as at with (nolock) where at.tdo_IdCodigo in ('FA' ,'FE' ,'NC' ,'CE' ,'ND', 'DE' )) AND toc_numinf = " & Val(fpLongInteger1(0).Value) & " AND toc_tipinf IN ('C')")
    
    ElseIf tipinf = "P" Then
       
'       Set RS = vg_db.Execute("SELECT MAX(toc_fecrem) AS fecemi FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND (toc_tipdoc = 'FA' OR toc_tipdoc = 'FE' OR toc_tipdoc = 'NC' OR toc_tipdoc = 'CE' or toc_tipdoc = 'ND' or toc_tipdoc = 'DE') AND toc_numinf = " & Val(fpLongInteger1(0).Value) & " AND toc_tipinf IN ('P')")
       Set RS = vg_db.Execute("SELECT MAX(toc_fecrem) AS fecemi FROM b_totcompras with (nolock) WHERE toc_codbod = " & vg_codbod & " AND toc_tipdoc in (select at.tdo_IdCodigo from a_tipodocumento as at with (nolock) where at.tdo_IdCodigo in ('FA' ,'FE' ,'NC' ,'CE' ,'ND', 'DE' )) AND toc_numinf = " & Val(fpLongInteger1(0).Value) & " AND toc_tipinf IN ('P')")
    ElseIf tipinf = "T" Then
       
       Set RS = vg_db.Execute("SELECT MAX(tov_fecemi) AS fecemi FROM b_totventas with (nolock) WHERE tov_codbod = " & vg_codbod & " AND tov_tipdoc = 'TR' AND tov_numinf = " & Val(fpLongInteger1(0).Value) & "")
    
    ElseIf tipinf = "G" Then
       
       Set RS = vg_db.Execute("SELECT MAX(tov_fecemi) AS fecemi FROM b_totventas with (nolock) WHERE tov_codbod = " & vg_codbod & " AND tov_tipdoc = 'TR' AND tov_numinf = " & Val(fpLongInteger1(0).Value) & "")
    
    ElseIf tipinf = "F" Then
       
       Set RS = vg_db.Execute("SELECT MAX(toc_fecemi) AS fecemi FROM b_totcompras with (nolock) WHERE toc_codbod = " & vg_codbod & " AND toc_numinf = " & Val(fpLongInteger1(0).Value) & " AND toc_tipinf = 'F'")
    
    End If
    
    If RS.State = 1 Then
       
       If RS.EOF Or IsNull(RS!fecemi) Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       fecemi = Format(RS!fecemi, "ddmmyyyy")
       RS.Close
       Set RS = Nothing
    
    End If
    
    If tipinf = "C" Or tipinf = "P" Then
       
       If fg_codigocbo(Combo1, 0, 1, "") <> 0 Then
       
          tipinf = fg_codigocbo(Combo1, 0, 1, "")
          
       End If
       I_CFC fpText.text, fecemi, Val(fpLongInteger1(0).Value), tipinf
    
    End If
    
    If tipinf = "T" Or tipinf = "G" Then
    
       I_CTC fpText.text, fecemi, Val(fpLongInteger1(0).Value), Val(fg_codigocbo(Combo1, 0, 1, "")), tipinf
       
    End If
    
    If tipinf = "F" Then
       
       I_NewFoFi fpText.text, fecemi, Val(fpLongInteger1(0).Value)
       
    End If

Case 3
    
    Dim numero As Long
    
'Solicitado x Rene Molina 20130213    If tipinf = "P" Then Exit Sub
    Toolbar1.Enabled = False: Frame1(0).Enabled = False
    
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT * FROM b_clientes with (nolock) WHERE cli_codigo = '" & Trim(fpText.text) & "' AND cli_tipo = 0")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close
    Set RS = Nothing
    
    numero = fpLongInteger1(0).Value
    If tipinf = "C" Or tipinf = "P" Then
       
'       tipinf = fg_codigocbo(Combo1, 0, 1, "")
       '-------> Validar si folio fue enviado por correlativo
       If ValidarEnvioCorrelativo(MuestraCasino(1), tipinf, numero) Then
          
          MsgBox "Existe Numero CFC anterior a este folio que no ha sido enviado (" & TraerFolioCFCPosterioesNoEnvio(MuestraCasino(1), tipinf) & ")", vbCritical + vbOKOnly, MsgTitulo
          Toolbar1.Enabled = True
          Frame1(0).Enabled = True
          Exit Sub
       
       End If
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       est = True
       Set RS = vg_db.Execute("SELECT DISTINCT toc_tipinf FROM b_totcompras with (nolock) WHERE toc_codbod = " & vg_codbod & " AND toc_numinf = " & numero & " AND toc_tipinf in ('P', 'C')")
       If Not RS.EOF Then
          
          '-------> Validar si existen documento que sean diferente a factura - nota credito - nota debito
          Sql = ""
          Sql = Sql + "SELECT DISTINCT a.toc_numinf FROM b_totcompras a, b_detcompras b, a_tipodocumento c " & _
                   "WHERE a.toc_rutpro = b.dec_rutpro AND   a.toc_tipdoc = b.dec_tipdoc " & _
                   "AND   a.toc_numdoc = b.dec_numdoc AND   a.toc_tipdoc = c.tdo_codigo " & _
                   "AND   a.toc_codbod = " & vg_codbod & " AND   a.toc_tipinf = '" & tipinf & "' "
                   
          If tipinf = "C" Then
          
             Sql = Sql + "AND (a.toc_envsap='0' OR (a.toc_envsap) IS NULL) "
          
          End If
          
          Sql = Sql + "AND   (c.tdo_cladoc) IS NOT NULL AND c.tdo_cladoc <> '' AND a.toc_numinf = " & numero & ""
          
          If RS1.State = 1 Then RS1.Close
          RS1.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          Set RS1 = vg_db.Execute(" " & Sql & " ")
          If RS1.EOF Then
             RS1.Close: Set RS1 = Nothing
             
             If RS1.State = 1 Then RS1.Close
             RS1.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS1 = vg_db.Execute("SELECT DISTINCT a.toc_numinf " & _
                      "FROM b_totcompras a, b_detcompras b, a_tipodocumento c " & _
                      "WHERE a.toc_rutpro = b.dec_rutpro " & _
                      "AND   a.toc_tipdoc = b.dec_tipdoc " & _
                      "AND   a.toc_numdoc = b.dec_numdoc " & _
                      "AND   a.toc_tipdoc = c.tdo_codigo " & _
                      "AND   a.toc_codbod = " & vg_codbod & " " & _
                      "AND   a.toc_tipinf = '" & tipinf & "' " & _
                      "AND   ((c.tdo_cladoc) IS NULL OR c.tdo_cladoc = '') " & _
                      "AND   a.toc_numinf = " & numero & "")
             
             If RS1.EOF Then
                RS1.Close: Set RS1 = Nothing
                
                If RS5.State = 1 Then RS5.Close
                RS5.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS5 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                
                If RS5.EOF Then
                   
                   vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                   vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                   fpLongInteger1(0).Value = TraerFolioDocumento(tipinf)
                   corenv = fpLongInteger1(0).Value
                
                End If
                RS5.Close: Set RS5 = Nothing
                
                If RS6.State = 1 Then RS6.Close
                RS6.CursorLocation = adUseClient
                vg_db.CursorLocation = adUseClient
                
                Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero) & " AND (inf_feccie = 0 OR (inf_feccie) IS NULL)")
                
                If Not RS6.EOF Then
                   
                   RS6.Close: Set RS6 = Nothing
                   vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                   Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                   If RS6.EOF Then
                      
                      vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                   
                   End If
                   RS6.Close: Set RS6 = Nothing
                
                Else
                   
                   RS6.Close: Set RS6 = Nothing
                
                End If
                MsgBox "No existe información, para enviar", vbExclamation + vbOKOnly, MsgTitulo
                Toolbar1.Enabled = True
                Frame1(0).Enabled = True
                RS.Close
                Set RS = Nothing
                Exit Sub
             
             Else
                
                est = False
             
             End If
          
          End If
          RS1.Close: Set RS1 = Nothing
          If Not est And (ValidarOpEnvio(MuestraCasino(1), 1) Or ValidarOpEnvio(MuestraCasino(1), 5) Or ValidarOpEnvio(MuestraCasino(1), 6)) Then
             
             '-------> Validar si existe guias despachos
             RS.Close: Set RS = Nothing
             Set RS5 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
             
             If RS5.EOF Then
                
                vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                fpLongInteger1(0).Value = TraerFolioDocumento(tipinf)
                corenv = fpLongInteger1(0).Value
             
             Else
                
                Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero) & " AND (inf_feccie = 0 OR (inf_feccie) IS NULL)")
                If Not RS6.EOF Then
                   
                   RS6.Close: Set RS6 = Nothing
                   vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                   Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                   
                   If RS6.EOF Then
                      
                      vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                   
                   End If
                   
                   RS6.Close
                   Set RS6 = Nothing
                
                Else
                   
                   RS6.Close
                   Set RS6 = Nothing
                
                End If
             
             End If
             
             RS5.Close
             Set RS5 = Nothing
             MsgBox "Folio fue enviado en su totalidad", vbExclamation + vbOKOnly, MsgTitulo
             Toolbar1.Enabled = True
             Frame1(0).Enabled = True
             Exit Sub
          
          End If
          '-------> Validar si tienes acceso envio sap
          If est And (ValidarOpEnvio(MuestraCasino(1), 1) Or ValidarOpEnvio(MuestraCasino(1), 5) Or ValidarOpEnvio(MuestraCasino(1), 6)) Then
             
             If Combo1(1).ListIndex = -1 And OpEnvioAx Then
                
                MsgBox "Debe selecionar lugar fisico", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub
             
             End If
             
             If ValidarOpEnvio(MuestraCasino(1), 1) Then
                
                If Not isInternetConnected(False, False, False) Then
                   
                   RS.Close
                   Set RS = Nothing
                   MsgBox "No hay conexión a internet, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
                   Toolbar1.Enabled = True
                   Frame1(0).Enabled = True
                   Exit Sub
                
                End If
             
             End If
             
             If Not GenerarArcSap(numero, tipinf) Then
                
                RS.Close: Set RS = Nothing
                
                If ValidarOpEnvio(MuestraCasino(1), 1) Then
                   
                   Text1(0).text = Text1(0).text & FechaHora & IIf(tiperr = 1, "Generación envío finalizado sin problema", IIf(tiperr = 2, "Generación envío finalizo con errores, " & (numenv) & " documentos no fueron enviados", IIf(tiperr = 3, "Generación envío finalizo con errores, " & (numenv) & " documentos no fueron enviados. Debe Informar a su monitor", "Generación envío finalizo con problema. Debe informar a su monitor"))) & VgLinea
                
                End If
                
                If corenv = fpLongInteger1(0).Value Then
                   
                   vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                   Set RS5 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                   
                   If RS5.EOF Then
                      
                      vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                   
                   End If
                   RS5.Close: Set RS5 = Nothing
                
                Else
                   vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                   
                   Set RS5 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                   If RS5.EOF Then
                      vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                   End If
                   RS5.Close: Set RS5 = Nothing
                
                End If
                
                If ValidarOpEnvio(MuestraCasino(1), 1) Then
                   
                   I_EnvioSap "1"
                
                End If
                
                '-------> Traer periodo
                Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
                If Not RS.EOF Then
                   
                   periodo = RS!cie_periodo
                
                End If
                RS.Close: Set RS = Nothing
                
                '--> Lugar Fisico
                If ValidarOpEnvio(MuestraCasino(1), 5) Then
                   
                   If (tipinf = "C" Or tipinf = "P") And Not OpEnvioAx Then
                   
                      If Not GeneraCfcDigitado(fpLongInteger1(0).Value, periodo, tipinf) Then
                      
                         Call MsgBox("No genero correctamente archivos MANUAL, trate de generar por envio Facturación MANUAL", vbInformation)
                   
                      End If
                   
                   
                   Else
                   
                      LugarFisico = fg_codigocbo(Combo1, 1, 4, 1)
                   
                      If Not GeneraCfcAX(fpLongInteger1(0).Value, periodo, LugarFisico) Then
                      
                         Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Facturación OPTIMUM", vbInformation)
                   
                      End If
                   
                   End If
                   
                End If
                
                fpLongInteger1(0).Value = TraerFolioDocumento(tipinf)
                corenv = fpLongInteger1(0).Value
                Toolbar1.Enabled = True
                Frame1(0).Enabled = True
                
                Exit Sub
             
             Else
                
                If ValidarOpEnvio(MuestraCasino(1), 1) Then
                   
                   Text1(0).text = Text1(0).text & FechaHora & IIf(tiperr = 1, "Generación envío finalizado sin problema", IIf(tiperr = 2, "Generación envío finalizo con errores, " & (numenv) & " documentos no fueron enviados", IIf(tiperr = 3, "Generación envío finalizo con errores, " & (numenv) & " documentos no fueron enviados. Debe informar a su monitor", "Generación envío finalizo con problema, debe informar a su monitor"))) & VgLinea
                
                End If
                
                RS.Close: Set RS = Nothing
                If corenv = fpLongInteger1(0).Value Then
                   
                   vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                   Set RS5 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                   If RS5.EOF Then
                      
                      vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                   
                   End If
                   RS5.Close: Set RS5 = Nothing
                
                Else
                   Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero) & " AND (inf_feccie = 0 OR (inf_feccie) IS NULL)")
                   If Not RS6.EOF Then
                      
                      RS6.Close
                      Set RS6 = Nothing
                      vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
                      
                      Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
                      If RS6.EOF Then
                         
                         vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
                      
                      End If
                      RS6.Close
                      Set RS6 = Nothing
                   
                   Else
                      
                      RS6.Close
                      Set RS6 = Nothing
                    
                    End If
                End If
                
                If ValidarOpEnvio(MuestraCasino(1), 1) Then
                   
                   I_EnvioSap "1"
                
                End If
                
                '-------> Traer periodo
                Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
                If Not RS.EOF Then
                   
                   periodo = RS!cie_periodo
                
                End If
                RS.Close
                Set RS = Nothing
                
                If ValidarOpEnvio(MuestraCasino(1), 5) Or ValidarOpEnvio(MuestraCasino(1), 6) Then
                   
                   If (tipinf = "C" Or tipinf = "P") And Not OpEnvioAx Then
                   
                      If Not GeneraCfcDigitado(fpLongInteger1(0).Value, periodo, tipinf) Then
                      
                          If tipinf = "P" Then
                             
                             Call MsgBox("Folio procesado, como el folio no corresponde al periodo actual. Debe generar el archivo excel por la vía de envio FACTURACION MANUAL", vbInformation)
                          
                          Else
                          
                            Call MsgBox("No genero correctamente archivos MANUAL, trate de generar por envio Facturación MANUAL", vbInformation)
                   
                          End If
                      End If
                   
                   Else
                      '--> Lugar Fisico
                      LugarFisico = fg_codigocbo(Combo1, 1, 4, 1)
                      
                      If Not GeneraCfcAX(fpLongInteger1(0).Value, periodo, LugarFisico) Then
                      
                         Call MsgBox("No genero correctamente archivos OPTIMUM, trate de generar por envio Facturación OPTIMUM", vbInformation)
                   
                      End If
                   
                   End If
                   
                End If
                
                fpLongInteger1(0).Value = TraerFolioDocumento(tipinf)
                corenv = fpLongInteger1(0).Value
                Toolbar1.Enabled = True
                Frame1(0).Enabled = True
                Exit Sub
             
             End If
          
          End If
       
       End If
    
    ElseIf tipinf = "T" Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("SELECT DISTINCT tov_tipdoc FROM b_totventas WHERE  tov_codbod = " & vg_codbod & " AND tov_tipdoc = 'TR' AND tov_numinf = " & numero & "")
    
    ElseIf tipinf = "G" Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("SELECT DISTINCT tov_tipdoc FROM b_totventas WHERE  tov_codbod = " & vg_codbod & " AND tov_tipdoc = 'TR' AND tov_numinf = " & numero & "")
    
    ElseIf tipinf = "F" Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("SELECT DISTINCT toc_tipinf FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND toc_numinf = " & numero & " AND toc_tipinf = 'F'")
    
    ElseIf tipinf = "P" Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("SELECT DISTINCT toc_tipinf FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND toc_numinf = " & numero & " AND toc_tipinf = 'P'")
    
    End If
    
    If RS.EOF Then
       
       RS.Close
       Set RS = Nothing
       '-------> Validar si folio fue enviado por correlativo
       
       If Not ValidarEnvioCorrelativo(MuestraCasino(1), tipinf, numero) And numero < TraerFolioDocumento(tipinf) Then
          
          If RS6.State = 1 Then RS6.Close
          RS6.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          
          Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero) & " AND (inf_feccie = 0 OR (inf_feccie) IS NULL)")
          
          If Not RS6.EOF Then
             
             RS6.Close
             Set RS6 = Nothing
             vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
             
             If RS6.State = 1 Then RS6.Close
             RS6.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS6 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
             If RS6.EOF Then
                
                vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
             
             End If
             RS6.Close
             Set RS6 = Nothing
             MsgBox "Folio fue cerrado, sin información", vbExclamation + vbOKOnly, MsgTitulo
          
          Else
             
             RS6.Close
             Set RS6 = Nothing
          
          End If
       
       Else
          
          MsgBox "No existe información, para enviar", vbExclamation + vbOKOnly, MsgTitulo
       
       End If
       Toolbar1.Enabled = True
       Frame1(0).Enabled = True
       Exit Sub
    
    ElseIf tipinf = "T" Or tipinf = "G" Or tipinf = "F" Or tipinf = "P" Or Not ValidarOpEnvio(MuestraCasino(1), 1) Then
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT inf_feccie FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(fpText.text) & "' AND inf_tipo = '" & tipinf & "' AND inf_numero = " & (fpLongInteger1(0).Value) & "")
        If Not RS1.EOF Then
           
           If RS1!inf_feccie > 0 Then
           
              RS.Close
              Set RS = Nothing
              RS1.Close
              Set RS1 = Nothing
              MsgBox "Nº Documento ya fue generado", vbExclamation + vbOKOnly, MsgTitulo
              Toolbar1.Enabled = True
              Frame1(0).Enabled = True
              Exit Sub
           
           End If
              
           RS1.Close
           Set RS1 = Nothing
        
           If MsgBox("El Folio Nº" & numero & " sera cerrado y se generara un nuevo folio, para los sgtes documentos " & Chr(13) & Space(50) & "¿ Desea continuar... ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
           
              RS.Close
              Set RS = Nothing
              Toolbar1.Enabled = True
              Frame1(0).Enabled = True
              Exit Sub
    
           End If
            
        End If
        
    End If
    RS.Close
    Set RS = Nothing
    
    vg_db.Execute "UPDATE a_infcfcfofi SET inf_feccie = " & Format(Date, "yyyymmdd") & ", inf_usuario = '" & vg_NUsr & "' WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & numero & ""
    
    If RS5.State = 1 Then RS5.Close
    RS5.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS5 = vg_db.Execute("SELECT DISTINCT inf_cencos FROM a_infcfcfofi WHERE inf_tipo = '" & tipinf & "' AND inf_cencos = '" & Trim(fpText.text) & "' AND inf_numero = " & (numero + 1) & "")
    If RS5.EOF Then
       
       vg_db.Execute "INSERT INTO a_infcfcfofi (inf_cencos, inf_tipo, inf_numero, inf_feccie) VALUES ('" & fpText.text & "', '" & tipinf & "', " & (numero + 1) & ", 0)"
    
    End If
    RS5.Close
    Set RS5 = Nothing
    
    If tipinf = "T" And ValidaTraspasodeSalida(Trim(fpText.text), vg_codbod, fpLongInteger1(0).Value) Then
    
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
    
       periodo = Format(Date, "yyyymmdd")
       Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
       If Not RS.EOF Then

          periodo = RS!cie_periodo

       End If
       RS.Close
       Set RS = Nothing
       
       If Not GenerarTraspasoSalidaAX(fpLongInteger1(0).Value, periodo) Then
             
          Call MsgBox("No genero correctamente archivos Traspaso salida manual, trate de generar por envio Facturación Traspaso salida manual", vbInformation)
             
       End If
           
    End If
    
    If tipinf = "T" Or tipinf = "G" Or tipinf = "F" Or tipinf = "P" Or Not ValidarOpEnvio(MuestraCasino(1), 1) Then
       
       MsgBox "Generación envió Finalizado Sin Problema", vbExclamation + vbOKOnly, MsgTitulo
    
    End If
    fpLongInteger1(0).Value = TraerFolioDocumento(tipinf)
    corenv = fpLongInteger1(0).Value
    Toolbar1.Enabled = True
    Frame1(0).Enabled = True

Case 5
   
    vg_codigo = ""
    vg_codigo4 = ""
    Dim titform As String
    Dim auxtipinf As String
    titform = ""
    If tipinf = "C" Or tipinf = "P" Then
       
       titform = "Histórico Control Facturas Compras"
       auxtipinf = "IN ('C','P')"
    
    ElseIf tipinf = "T" Or tipinf = "G" Then
       
        If Combo1(0).ListIndex = -1 Then
       
           MsgBox "Debe selecionar tipo traspaso", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
           
        End If
       
       titform = "Histórico Control Traspasos Contratos;" & fg_codigocbo(Combo1, 0, 1, "")
       auxtipinf = "IN ('T','G')"
    
    ElseIf tipinf = "F" Then
       
       titform = "Histórico Control Fondo Fijo (Fofi)"
       auxtipinf = "IN ('F')"
    
    End If
    
    B_HistPm.LlenarHistPlan titform, fpText.text, auxtipinf, 4
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    
    If tipinf = "C" Or tipinf = "P" Then
       
       tipinf = vg_codigo4
       Label2(2).Caption = IIf(tipinf = "C", "Doc. SGP", "Doc. P. Electronica")
    
       est = True
       
       Combo1(0).ListIndex = fg_buscacbostring(Combo1, 0, 1, (tipinf))
       
       est = False
       
    ElseIf tipinf = "T" Or tipinf = "G" Then
       
       tipinf = vg_codigo4
       Label2(2).Caption = IIf(tipinf = "T", "Doc. SGP", "Doc. P. Electronica")
    
    End If

Case 7 '-------> Generar facturación OPTIMUM
    
    OpEnvioAx = True
    
    If tipinf = "T" And (fg_codigocbo(Combo1, 0, 1, "") = 1 Or Combo1(0).ListIndex = -1) Then
       
       MsgBox "Para acceder a esta opción, solo tiene que seleccionar tipo traspaso de salida ", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    '-------> Validar si el contrato tiene opción de envio sap
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
   
    Set RS = vg_db.Execute("SELECT * FROM b_casinointerfaz WHERE cai_cencos = '" & MuestraCasino(1) & "'")

    If Not RS.EOF And (tipinf = "C" Or tipinf = "P" Or (tipinf = "T" And fg_codigocbo(Combo1, 0, 1, "") = 0)) Then
   
       Do While Not RS.EOF
  
          If RS!cai_codtii = 6 Then
    
             OpEnvioAx = False
             Exit Do
      
          End If

          RS.MoveNext
       Loop
    
    End If
    RS.Close
    Set RS = Nothing
    
    If (tipinf = "C" Or tipinf = "P") And Not OpEnvioAx Then
    
        P_GenCfcAx.Inicio 1
    
    ElseIf (tipinf = "C" Or tipinf = "P") And OpEnvioAx Then
    
        P_GenCfcAx.Inicio 2
        
    ElseIf (tipinf = "T") And Not OpEnvioAx Then
    
        P_GenCfcAx.Inicio 3
            
    Else
    
       Exit Sub
       
    End If

    P_GenCfcAx.Show 1, Partida

Case 9 '-------> Explorar carpeta windows
    
    OpEnvioAx = True
    '-------> Validar si el contrato tiene opción de envio sap
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT * FROM b_casinointerfaz WHERE cai_cencos = '" & MuestraCasino(1) & "'")

    If Not RS.EOF And (tipinf = "C" Or tipinf = "P" Or tipinf = "T") Then
   
       Do While Not RS.EOF
  
          If RS!cai_codtii = 6 Then
    
             OpEnvioAx = False
             Exit Do
      
          End If

          RS.MoveNext
       Loop
    
    End If
    RS.Close
    Set RS = Nothing

    If (tipinf = "C" Or tipinf = "P" Or tipinf = "T") And Not OpEnvioAx Then
    
       ExplorarCarpeta dir_trabajo_Inf & "InformesAXFacturacionManual"
       
    Else
       
       ExplorarCarpeta dir_trabajo_Inf & "InformesAXFacturacion"
    
    End If

Case 11
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = True
Frame1(0).Enabled = True
If RS.State = 1 Then RS.Close: Set RS = Nothing
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End Sub

Sub Inicio(tfor As String, tf As String)

tipinf = tf
Me.Caption = tfor
MsgTitulo = tfor
corenv = 0
If tf = "T" Then Label2(1).Visible = True: Combo1(0).Visible = True
If tf = "C" Or tf = "P" Or tf = "T" Then
   
   Label2(2).Caption = "Doc. SGP"
   Label2(2).Visible = True

Else
   
   Label2(2).Visible = False

End If

'-------> Buscar numero folio
fpLongInteger1(0).Value = TraerFolioDocumento(tipinf)
corenv = TraerFolioDocumento(tipinf)
'------->

End Sub


Function GenerarArcSap(numfol As Long, tipinf As String) As Boolean

Dim RS          As New ADODB.Recordset
Dim RS1         As New ADODB.Recordset
Dim RS2         As New ADODB.Recordset
Dim RS3         As New ADODB.Recordset
Dim RS4         As New ADODB.Recordset
Dim clacon      As String
Dim ctacon      As String
Dim fecdoc      As String
Dim NumDoc      As String
Dim rutpro      As String
Dim valcue      As String
Dim indimp      As String
Dim cencos      As String
Dim cuemay      As String
Dim nomarch     As String
Dim nomarchzip  As String
Dim mdir        As String
Dim tipdoc      As String
Dim valivai     As String
Dim i           As Long
Dim j           As Long
Dim valiva      As Long
Dim valimp      As Long
Dim flete       As Long
Dim excimp      As Boolean
Dim estsap      As Boolean
Dim totdoc      As Long
Dim EstEnc      As Boolean
Dim exiimp      As Boolean
Dim corr        As Long
Dim numlin      As Long
Dim candec      As Double
Dim candec2     As Double
Dim n_a         As String
Dim bkpf_bukrs  As String
Dim bkpf_blart  As String
Dim bkpf_budat  As String
Dim bkpf_bldat  As String
Dim bkpf_xblnr  As String
Dim bkpf_bktxt  As String
Dim bkpf_waers  As String
Dim bseg_newbs  As String
Dim bseg_newko  As String
Dim bseg_wrbtr  As String
Dim bseg_zuonr  As String
Dim bseg_kostl  As String
Dim n_acodimpto As String
Dim n_actaimpto As String
Dim n_amonimp   As String
Dim n_aimprecu  As String
Dim n_aotrimp   As String
Dim fecenv      As String
Dim numero      As Long
Dim parametro1  As String
Dim parametro2  As String
Dim parametro3  As String
Dim Segundos    As Byte
Dim StrMensaje  As String
Dim StrMensaje1 As String
Dim nommen      As String
Dim vRet        As Variant
Dim estado      As Boolean
Dim valfac      As Long
Dim esttim      As Boolean
Dim Sql         As String

On Error GoTo Man_EnvioCfc

DoEvents
n_a = ""
bkpf_bukrs = ""
bkpf_blart = ""
bkpf_budat = ""
bkpf_bldat = ""
bkpf_xblnr = ""
bkpf_bktxt = ""
bkpf_waers = ""
bseg_newbs = ""
bseg_newko = ""
bseg_wrbtr = ""
bseg_zuonr = ""
bseg_kostl = ""
n_acodimpto = ""
n_actaimpto = ""
n_amonimp = ""
n_aimprecu = ""
n_aotrimp = ""
excimp = True
estsap = True
esttim = True

'-------> Validar si existen documento del portal

If tipinf = "P" Then

   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute("sgp_Sel_ValidarFacturasDigitadaPortal_V01 '" & MuestraCasino(1) & "', " & numfol & ", '" & tipinf & "'")
   If Not RS.EOF Then
      
      GenerarArcSap = True
      RS.Close
      Set RS = Nothing
      Exit Function
      
   End If
   RS.Close
   Set RS = Nothing

End If

If ValidarOpEnvio(MuestraCasino(1), 1) Then
   
   '-------> Abrir mensaje de text
   Me.Height = 7725
   tiperr = 0
   fg_centra Me
   Frame2.Visible = True
   Text1(0).Visible = True
   Text1(0).Enabled = True 'False
   GenerarArcSap = True
   Text1(0).text = FechaHora & "CENCO : " & MuestraCasino(1) & " - " & MuestraCasino(2) & VgLinea
   Text1(0).text = Text1(0).text & FechaHora & "USUARIO : " & Environ("USERNAME") & VgLinea
   Text1(0).text = Text1(0).text & FechaHora & "Inicio del Proceso. Control Factura Compra : " & numfol & VgLinea

   '-------> Validar si existe usuario sap
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'sapusu'")
   If RS.EOF Then
      
      RS.Close
      Set RS = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "No tiene creado usuario, para Web Service" & VgLinea
      tiperr = 4
      GenerarArcInvSap = False
      Exit Function
   
   ElseIf IsNull(RS!par_valor) Or Trim(RS!par_valor) = "" Then
      
      RS.Close
      Set RS = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "usuario fue borrado, para Web Service" & VgLinea
      tiperr = 4
      GenerarArcInvSap = False
      Exit Function
   
   End If
   RS.Close
   Set RS = Nothing

   '-------> Validar si existe password sap
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'sappas'")
   If RS.EOF Then
      
      RS.Close
      Set RS = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "No tiene creado password, para Web Service" & VgLinea
      tiperr = 4
      GenerarArcInvSap = False
      Exit Function
   
   ElseIf IsNull(RS!par_valor) Or Trim(RS!par_valor) = "" Then
      
      RS.Close
      Set RS = Nothing
      Text1(0).text = Text1(0).text & FechaHora & "Password fue borrada, para Web Service" & VgLinea
      tiperr = 4
      GenerarArcInvSap = False
      Exit Function
   
   End If
   RS.Close
   Set RS = Nothing

End If

'-------> Traer sociedad del contrato
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT * FROM b_clientes WHERE cli_codigo = '" & MuestraCasino(1) & "' AND cli_tipo = 0")
If RS.EOF Or IsNull(RS!cli_socsap) Or Trim(RS!cli_socsap) = "" Then
   
   RS.Close
   Set RS = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No tiene asignado la sociedad de SAP, en contrato." & VgLinea
   tiperr = 4
   GenerarArcInvSap = False
   Exit Function

End If
bkpf_bukrs = Trim(RS!cli_socsap)
bkpf_bktxt = Trim(RS!cli_codigo) & "-" & numfol
RS.Close
Set RS = Nothing

'-------> Traer variable impuesto iva
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
If vg_tipbase = "1" Then
   
   Set RS = vg_db.Execute("SELECT imp_codsap FROM a_impuesto WHERE imp_adicional = 0 AND ((imp_codsap) IS NOT NULL OR trim(imp_codsap) <> '')")

Else
   
   Set RS = vg_db.Execute("SELECT imp_codsap FROM a_impuesto WHERE imp_adicional = 0 AND ((imp_codsap) IS NOT NULL OR ltrim(imp_codsap) <> '')")

End If

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No existe código SAP asignado impuesto iva, Comuniquesen con departamento de informatica." & VgLinea
   tiperr = 4
   GenerarArcSap = False
   Exit Function

End If
vg_csapiva = Trim(RS!imp_codsap)
RS.Close
Set RS = Nothing

'-------> Traer variables impuestos como Harina, Carne, Zona Franca, Disiel
vg_csapotros = ""
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
If vg_tipbase = "1" Then
   
   Set RS = vg_db.Execute("SELECT imp_codsap FROM a_impuesto WHERE imp_adicional = 1 AND ((imp_codsap) IS NOT NULL OR trim(imp_codsap) <> '')")

Else
   
   Set RS = vg_db.Execute("SELECT imp_codsap FROM a_impuesto WHERE imp_adicional = 1 AND ((imp_codsap) IS NOT NULL OR ltrim(imp_codsap) <> '')")

End If

If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No existe código SAP asignado impuesto Harina, Carne, Etc. Comuniquesen con departamento de informatica." & VgLinea
   tiperr = 4
   GenerarArcSap = False
   Exit Function

End If

Do While Not RS.EOF
    
    vg_csapotros = vg_csapotros & LCase(Trim(RS!imp_codsap)) & ";"
    RS.MoveNext

Loop

If Trim(vg_csapotros) <> "" Then
   
   vg_csapotros = Mid(vg_csapotros, 1, Len(vg_csapotros) - 1)

End If
RS.Close
Set RS = Nothing
'-------> Validar encabezado clave contable
'Activar cuando este ok If vg_claencsap = 0 Then MsgBox "No existe clave contabilización encabezado sap, comuniquese con departamento de informatica." & VgLinea & Space(50) & "Proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: GenerarArcSap = False: Exit Function
'-------> Validar detalle clave contable
'Activar cuando este ok If vg_cladetsap = 0 Then MsgBox "No existe clave contabilización sap, comuniquese con departamento de informatica." & VgLinea & Space(50) & "Proceso cancelado", vbCritical + vbOKOnly, Msgtitulo: GenerarArcSap = False: Exit Function
'-------> Validar documento exento impuesto
If Trim(vg_docexento) = "" Then
   
   Text1(0).text = Text1(0).text & FechaHora & "No existe clave exento sap, comuniquese con departamento de informatica." & VgLinea
   tiperr = 4: GenerarArcSap = False
   Exit Function

End If

'-------> Validar documento exento impuesto
If Trim(vg_docafecto) = "" Then
   
   Text1(0).text = Text1(0).text & FechaHora & "No existe clave afecto sap, comuniquese con departamento de informatica." & VgLinea
   tiperr = 4: GenerarArcSap = False
   Exit Function

End If

numlin = 1
corr = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
If vg_tipbase = "1" Then
   
   Sql = ""
   Sql = Sql + "SELECT DISTINCT a.toc_fecdig, a.toc_fecemi, a.toc_fecrem, a.toc_totdoc, b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, " & _
         "a.toc_ivadoc, a.toc_otrimp, a.toc_exedoc, SUM(b.dec_canmer* b.dec_precom+b.dec_prefle) AS valfac " & _
         "FROM b_totcompras a, b_detcompras b, a_tipodocumento c " & _
         "WHERE a.toc_rutpro = b.dec_rutpro " & _
         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
         "AND   a.toc_numdoc = b.dec_numdoc " & _
         "AND   a.toc_tipdoc = c.tdo_codigo " & _
         "AND   a.toc_codbod = " & vg_codbod & " AND a.toc_tipinf = '" & tipinf & "' "
            
   If tipinf = "C" Then
   
      Sql = Sql + "AND (a.toc_envsap = '0' OR ISNULL(a.toc_envsap)) "
   
   End If
   
   Sql = Sql + "AND   NOT ISNULL(c.tdo_cladoc) AND c.tdo_cladoc <> '' " & _
         "AND   a.toc_numinf = " & numfol & " GROUP BY a.toc_fecdig, a.toc_fecemi, a.toc_fecrem, a.toc_totdoc, b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, a.toc_ivadoc, a.toc_otrimp, a.toc_exedoc " & _
         "ORDER BY a.toc_fecdig, a.toc_fecemi, b.dec_rutpro, b.dec_numdoc"
   
   Set RS = vg_db.Execute("" & Sql & "")
   
Else
   
   Sql = ""
   Sql = Sql + "SELECT DISTINCT a.toc_fecdig, a.toc_fecemi, a.toc_fecrem, a.toc_totdoc, b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, " & _
         "a.toc_ivadoc, a.toc_otrimp, a.toc_exedoc, SUM(b.dec_canmer* b.dec_precom+b.dec_prefle) AS valfac " & _
         "FROM b_totcompras a, b_detcompras b, a_tipodocumento c " & _
         "WHERE a.toc_rutpro = b.dec_rutpro " & _
         "AND   a.toc_tipdoc = b.dec_tipdoc " & _
         "AND   a.toc_numdoc = b.dec_numdoc " & _
         "AND   a.toc_tipdoc = c.tdo_codigo " & _
         "AND   a.toc_codbod = " & vg_codbod & " AND   a.toc_tipinf = '" & tipinf & "' "
       
   If tipinf = "C" Then
   
      Sql = Sql + "AND (a.toc_envsap = '0' OR (a.toc_envsap) IS NULL) "
   
   End If
   
   Sql = Sql + "AND   (c.tdo_cladoc) IS NOT NULL AND c.tdo_cladoc <> '' " & _
         "AND   a.toc_numinf = " & numfol & " GROUP BY a.toc_fecdig, a.toc_fecemi, a.toc_fecrem, a.toc_totdoc, b.dec_rutpro, b.dec_tipdoc, b.dec_numdoc, a.toc_ivadoc, a.toc_otrimp, a.toc_exedoc " & _
         "ORDER BY a.toc_fecdig, a.toc_fecemi, b.dec_rutpro, b.dec_numdoc"
   Set RS = vg_db.Execute("" & Sql & "")
   
End If

If RS.EOF Then
   
   fg_descarga
   RS.Close
   Set RS = Nothing
   Text1(0).text = Text1(0).text & FechaHora & "No existe informaciòn a procesar, comuniquese con departamento de informatica." & VgLinea
   tiperr = 4
   GenerarArcSap = False
   Exit Function

End If
numenv = 0
totenv = 0

'-------> traer numero de documento

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = ""
Sql = Sql + "SELECT COUNT(a.toc_numdoc) AS nreg FROM b_totcompras a, a_tipodocumento b " & _
      "WHERE a.toc_tipdoc = b.tdo_codigo AND a.toc_codbod = " & vg_codbod & " " & _
      "AND  a.toc_numinf = " & numfol & " " & _
      "AND  a.toc_tipinf = '" & tipinf & "' "

If tipinf = "C" Then

   Sql = Sql + "AND (a.toc_envsap = '0' OR (a.toc_envsap) IS NULL) "

End If

Sql = Sql + "AND (b.tdo_cladoc) IS NOT NULL AND b.tdo_cladoc <> '' "
Set RS1 = vg_db.Execute("" & Sql & "")

If Not RS1.EOF Then totenv = IIf(IsNull(RS1!nreg), 0, RS1!nreg)
RS1.Close
Set RS1 = Nothing
Do While Not RS.EOF
   
   DoEvents
   '-------> Traer Impuesto IVA
   excimp = True
   valiva = Round((RS!valfac + RS!toc_ivadoc + RS!toc_otrimp), 0)
   If valiva <> RS!toc_totdoc Then valiva = RS!toc_totdoc
   candec = 0
   candec2 = 0
   totdoc = 0
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   '-------> Traer documento con impuestos
   Set RS1 = vg_db.Execute("SELECT DISTINCT dec_numdoc, " & _
            "(SELECT TOP 1 imd_rutdoc FROM b_detcomprasimp WHERE imd_rutdoc = dec_rutpro AND imd_tipdoc = dec_tipdoc AND imd_numdoc = dec_numdoc AND imd_codpro = dec_codmer) as imd_rutdoc " & _
            "FROM   b_detcompras " & _
            "WHERE  dec_rutpro = '" & RS!dec_rutpro & "' " & _
            "AND    dec_tipdoc = '" & RS!dec_tipdoc & "' " & _
            "AND    dec_numdoc = " & RS!dec_numdoc & "")
   
   EstEnc = True
   ctacon = ""
   i = 0
   EstEnc = True
   ctacon = ""
   i = 0
   
   Do While Not RS1.EOF
      
      '-------> Validar si documento es excepto ó biene con impuesto
      If IsNull(RS1!imd_rutdoc) Then
         
         excimp = False
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         If RS2.State = 1 Then RS2.Close
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         Set RS2 = vg_db.Execute("SELECT DISTINCT b.pro_tippro, b.pro_ctacon, c.cta_nombre, 0 AS imp_codigo, 0 AS imp_pctimp, 0 AS imp_inccos, '' AS imp_codsap, " & _
                  "(SELECT TOP 1 par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'codsapexe1') AS imp_cimsap1, " & _
                  "(SELECT TOP 1 par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'codsapexe2') AS imp_cimsap2, " & _
                  "(SELECT TOP 1 par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'codsapexe3') AS imp_cimsap3, " & _
                  "(SELECT TOP 1 par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'codsapexe4') AS imp_cimsap4, " & _
                  "(SELECT TOP 1 par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'tipope') AS cli_tipope, " & _
                  "0 AS imp_adicional, 0 AS valimp " & _
                  "FROM b_detcompras a, b_productos b, a_ctacontable c " & _
                  "WHERE a.dec_rutpro = '" & RS!dec_rutpro & "' " & _
                  "AND   a.dec_tipdoc = '" & RS!dec_tipdoc & "' " & _
                  "AND   a.dec_numdoc = " & RS!dec_numdoc & " " & _
                  "AND   a.dec_codmer = b.pro_codigo " & _
                  "AND   b.pro_ctacon = c.cta_codigo GROUP BY b.pro_tippro, b.pro_ctacon, c.cta_nombre ORDER BY b.pro_tippro, b.pro_ctacon")
      
      Else
         
         '-------> Fin traer impuesto IVA
         If RS2.State = 1 Then RS2.Close
         RS2.CursorLocation = adUseClient
         vg_db.CursorLocation = adUseClient
         
         Set RS2 = vg_db.Execute("SELECT c.pro_tippro, b.imp_codigo, b.imp_pctimp, b.imp_inccos, b.imp_codsap, b.imp_adicional, b.imp_cimsap1, b.imp_cimsap2, b.imp_cimsap3, b.imp_cimsap4, c.pro_ctacon, d.cta_nombre, ROUND(SUM(a.imd_monimp),0) AS valimp, " & _
                  "(SELECT TOP 1 par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'tipope') AS cli_tipope " & _
                  "FROM b_detcomprasimp a, a_impuesto b, b_productos c, a_ctacontable d " & _
                  "WHERE a.imd_rutdoc = '" & RS!dec_rutpro & "' " & _
                  "AND   a.imd_tipdoc = '" & RS!dec_tipdoc & "' " & _
                  "AND   a.imd_numdoc = " & RS!dec_numdoc & " " & _
                  "AND   a.imd_codpro = c.pro_codigo " & _
                  "AND   a.imd_codimp = b.imp_codigo " & _
                  "AND   c.pro_ctacon = d.cta_codigo AND a.imd_pctimp <> 0 " & _
                  "GROUP BY c.pro_tippro, b.imp_codigo, b.imp_pctimp, b.imp_inccos, b.imp_codsap, b.imp_adicional, b.imp_cimsap1, b.imp_cimsap2, b.imp_cimsap3, b.imp_cimsap4, c.pro_ctacon, d.cta_nombre ORDER BY  c.pro_ctacon, b.imp_codigo, b.imp_codsap ")
      
      End If
      j = 1
      
      If Not RS2.EOF Then
         
         Do While Not RS2.EOF
            
            '-------> Traer clase documento
            
            If vg_pais = "CO" Then
               
               If RS3.State = 1 Then RS3.Close
               RS3.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               Set RS3 = vg_db.Execute("SELECT DISTINCT b.cds_cdosap FROM a_tipodocumento a, a_clasedocsap b, b_clientes c " & _
                        "WHERE a.tdo_codigo = b.cds_coddoc " & _
                        "AND   b.cds_codreg = c.cli_codreg " & _
                        "AND   c.cli_codigo = '" & MuestraCasino(1) & "' " & _
                        "AND   c.cli_tipo   = 0 " & _
                        "AND   a.tdo_codigo = '" & RS!dec_tipdoc & "'")
               
               If RS3.EOF Then
                  
                  fg_descarga
                  Text1(0).text = Text1(0).text & FechaHora & "No existe clave documento SAP, en sistema SGP." & VgLinea
               
               Else
                  
                  bkpf_blart = Trim(RS3!cds_cdosap)
               
               End If
            Else
               
               If RS3.State = 1 Then RS3.Close
               RS3.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               Set RS3 = vg_db.Execute("SELECT * FROM a_tipodocumento WHERE tdo_codigo = '" & RS!dec_tipdoc & "'")
               If RS3.EOF Then
                  
                  bkpf_blart = ""
               
               Else
                  
                  bkpf_blart = Trim(RS3!tdo_cladoc)
               
               End If
            
            End If
            RS3.Close
            Set RS3 = Nothing
      
            '-------> Encabezado documento
            n_a = "X"
            bkpf_budat = Format(RS!toc_fecemi, "ddmmyyyy")
            bkpf_bldat = Format(RS!toc_fecrem, "ddmmyyyy")
            
            bkpf_xblnr = RS!dec_numdoc
            bkpf_waers = vg_tipmonsap '"CLP"
            bseg_newbs = IIf(RS!dec_tipdoc = "FA" Or RS!dec_tipdoc = "ND", "31", IIf(RS!dec_tipdoc = "NC", "21", IIf(RS!dec_tipdoc = "CE", "23", IIf(RS!dec_tipdoc = "FE", "33", "30"))))
            bseg_newko = RS!dec_rutpro
            bseg_wrbtr = IIf(IsNull(RS2!imp_codsap), Round(IIf(Trim(RS!dec_tipdoc) = "NC" Or Trim(RS!dec_tipdoc) = "CE", RS2!valimp, RS2!valimp), 0), IIf(Trim(RS!dec_tipdoc) = "NC" Or Trim(RS!dec_tipdoc) = "CE", valiva, valiva))
            bseg_zuonr = MuestraCasino(1)
            bseg_sgtxt = RS1!dec_numdoc
            bseg_kostl = MuestraCasino(1) '""
            n_acodimpto = ""
            n_actaimpto = ""
            n_amonimp = ""
            n_aimprecu = ""
            n_aotrimp = ""
            If (Trim(RS2!imp_codsap) = TraerCuentaIva(RS2!imp_codigo) Or IsNull(RS2!imp_codsap) Or Not excimp) And EstEnc = True Then
               
               codigo = 0
               numlin = 1
               
               If RS4.State = 1 Then RS4.Close
               RS4.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient

               Set RS4 = vg_db.Execute("SELECT cfc_codigo FROM sap_cfc ORDER BY cfc_codigo DESC")
               If Not RS4.EOF Then RS4.MoveFirst: codigo = RS4!cfc_codigo + 1 Else codigo = 1
               RS4.Close
               Set RS4 = Nothing
            
               vg_db.Execute "INSERT INTO sap_cfc VALUES (" & codigo & ", " & numlin & ", '" & n_a & "', '" & bkpf_bukrs & "', '" & bkpf_blart & "', '" & bkpf_budat & "', " & _
                             "'" & bkpf_bldat & "', '" & bkpf_xblnr & "', '" & bkpf_bktxt & "', '" & bkpf_waers & "', '" & bseg_newbs & "', " & _
                             "'" & bseg_newko & "', '" & bseg_wrbtr & "', '" & bseg_zuonr & "', '" & bseg_sgtxt & "', '" & bseg_kostl & "', " & _
                             "'" & n_acodimpto & "', '" & n_actaimpto & "', '" & n_amonimp & "', '" & n_aimprecu & "', '" & n_aotrimp & "')"
               numlin = numlin + 1
               If Trim(RS2!imp_codsap) = vg_csapiva Or Not excimp Then EstEnc = False
            
            End If
            RS2.MoveNext
            If Not RS2.EOF Then
               
               If Trim(RS2!imp_codsap) <> "" Or Not IsNull(RS2!imp_codsap) Then ctacon = IIf(Trim(RS2!imp_codsap) = vg_csapiva, Trim(RS2!pro_ctacon), Trim(RS2!imp_codsap))
            
            End If
            RS2.MovePrevious
            valfac = 0
            flete = 0
            If EstEnc = False And Not IsNull(RS1!imd_rutdoc) Then
               
               If RS3.State = 1 Then RS3.Close
               RS3.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               Set RS3 = vg_db.Execute("SELECT a.toc_fecemi, SUM(b.dec_canmer* b.dec_precom+b.dec_prefle) AS valfac, SUM(b.dec_prefle) AS dec_prefle, SUM(b.dec_ptotal+b.dec_prefle) AS valnet " & _
                        "FROM b_totcompras a, b_detcompras b, b_productos c, a_tipodocumento d " & _
                        "WHERE a.toc_rutpro = b.dec_rutpro " & _
                        "AND   a.toc_tipdoc = b.dec_tipdoc " & _
                        "AND   a.toc_numdoc = b.dec_numdoc " & _
                        "AND   b.dec_codmer IN (SELECT DISTINCT e.imd_codpro FROM b_detcomprasimp e WHERE b.dec_rutpro = e.imd_rutdoc AND b.dec_tipdoc = e.imd_tipdoc AND b.dec_numdoc = e.imd_numdoc AND e.imd_codimp = " & RS2!imp_codigo & ") " & _
                        "AND   b.dec_codmer = c.pro_codigo " & _
                        "AND   a.toc_tipdoc = d.tdo_codigo " & _
                        "AND   c.pro_ctacon = '" & RS2!pro_ctacon & "' " & _
                        "AND   a.toc_rutpro = '" & RS!dec_rutpro & "' " & _
                        "AND   a.toc_tipdoc = '" & RS!dec_tipdoc & "' " & _
                        "AND   a.toc_numdoc = " & RS!dec_numdoc & " " & _
                        "AND   a.toc_tipinf = '" & tipinf & "' " & _
                        "AND   (d.tdo_cladoc) IS NOT NULL AND d.tdo_cladoc <> '' " & _
                        "AND   a.toc_numinf = " & numfol & " AND a.toc_codbod = " & vg_codbod & " GROUP BY a.toc_fecemi ORDER BY a.toc_fecemi")
              If Not RS3.EOF Then valfac = RS3!valnet: flete = RS3!dec_prefle
              RS3.Close
              Set RS3 = Nothing
              If RS!toc_exedoc > 0 Then valfac = Round(valfac + RS!toc_otrimp, 0)
            
            ElseIf EstEnc = False And IsNull(RS1!imd_rutdoc) Then
               
               If RS3.State = 1 Then RS3.Close
               RS3.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               Set RS3 = vg_db.Execute("SELECT a.toc_fecemi, SUM(b.dec_canmer* b.dec_precom+b.dec_prefle) AS valfac, SUM(b.dec_prefle) AS dec_prefle, SUM(b.dec_ptotal+b.dec_prefle) AS valnet " & _
                        "FROM b_totcompras a, b_detcompras b, b_productos c, a_tipodocumento d " & _
                        "WHERE a.toc_rutpro = b.dec_rutpro " & _
                        "AND   a.toc_tipdoc = b.dec_tipdoc " & _
                        "AND   a.toc_numdoc = b.dec_numdoc " & _
                        "AND   b.dec_codmer = c.pro_codigo " & _
                        "AND   a.toc_tipdoc = d.tdo_codigo " & _
                        "AND   c.pro_ctacon = '" & RS2!pro_ctacon & "' " & _
                        "AND   a.toc_rutpro = '" & RS!dec_rutpro & "' " & _
                        "AND   a.toc_tipdoc = '" & RS!dec_tipdoc & "' " & _
                        "AND   a.toc_numdoc = " & RS!dec_numdoc & " " & _
                        "AND   a.toc_tipinf = '" & tipinf & "' " & _
                        "AND   (d.tdo_cladoc) IS NOT NULL AND d.tdo_cladoc <> '' " & _
                        "AND   a.toc_numinf = " & numfol & " AND a.toc_codbod = " & vg_codbod & " GROUP BY a.toc_fecemi ORDER BY a.toc_fecemi")
              If Not RS3.EOF Then valfac = RS3!valnet: flete = RS3!dec_prefle
              RS3.Close
              Set RS3 = Nothing
              If RS!toc_exedoc > 0 Then valfac = Round(valfac + RS!toc_otrimp, 0)
            
            End If
            '-------> Ver si existe más de un impuesto
            exiimp = False
            
            If RS3.State = 1 Then RS3.Close
            RS3.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            If vg_tipbase = "1" Then
               
               Set RS3 = vg_db.Execute("SELECT DISTINCT (SELECT DISTINCT COUNT(a.imd_codimp) AS nidm FROM a_impuesto WHERE imp_codigo = a.imd_codimp AND imp_adicional <> 0) AS nreg " & _
                        "FROM b_detcomprasimp a, a_impuesto b, b_productos c " & _
                        "WHERE a.imd_rutdoc = '" & RS!dec_rutpro & "' " & _
                        "AND   a.imd_tipdoc = '" & RS!dec_tipdoc & "' " & _
                        "AND   a.imd_numdoc = " & RS!dec_numdoc & " " & _
                        "AND   c.pro_ctacon = '" & RS2!pro_ctacon & "' " & _
                        "AND   a.imd_codpro = c.pro_codigo " & _
                        "AND   a.imd_codimp = b.imp_codigo AND b.imp_adicional")
            
            Else
               
               Set RS3 = vg_db.Execute("SELECT DISTINCT (SELECT DISTINCT COUNT(a_impuesto.imp_codigo) AS nidm FROM a_impuesto WHERE imp_codigo = a.imd_codimp AND imp_adicional <> 0) AS nreg " & _
                        "FROM b_detcomprasimp a, a_impuesto b, b_productos c " & _
                        "WHERE a.imd_rutdoc = '" & RS!dec_rutpro & "' " & _
                        "AND   a.imd_tipdoc = '" & RS!dec_tipdoc & "' " & _
                        "AND   a.imd_numdoc = " & RS!dec_numdoc & " " & _
                        "AND   c.pro_ctacon = '" & RS2!pro_ctacon & "' " & _
                        "AND   a.imd_codpro = c.pro_codigo " & _
                        "AND   a.imd_codimp = b.imp_codigo AND b.imp_adicional <> 0")
            
            End If
            If Not RS3.EOF Then
               
               If RS3!nreg > i Then exiimp = True
            
            End If
            RS3.Close
            Set RS3 = Nothing
            '-------> Mover costo exento documento
            If IsNull(RS1!imd_rutdoc) And valfac = 0 Then valfac = RS!toc_exedoc
            
            n_a = ""
            bkpf_budat = Format(RS!toc_fecemi, "ddmmyyyy")
            bkpf_bldat = Format(RS!toc_fecrem, "ddmmyyyy")
            
            bkpf_xblnr = RS!dec_numdoc
            bkpf_waers = vg_tipmonsap
            'pendiente cuenta 50 haber
            bseg_newbs = IIf(RS2!imp_adicional = 0 Or RS!dec_tipdoc = "NC" Or RS!dec_tipdoc = "CE", IIf(RS!dec_tipdoc = "FA" Or RS!dec_tipdoc = "FE" Or RS!dec_tipdoc = "ND" Or RS!dec_tipdoc = "DE", "40", "50"), "40")
            bseg_newko = IIf(RS2!imp_adicional = 0, Trim(RS2!pro_ctacon), "")
            bseg_wrbtr = IIf(RS2!imp_adicional = 0, Round(IIf(valfac > 0, IIf(RS!dec_tipdoc = "NC" Or RS!dec_tipdoc = "CE", valfac, valfac), IIf(RS!dec_tipdoc = "NC" Or RS!dec_tipdoc = "CE", RS!valfac, RS!valfac))), "")
            bseg_zuonr = IIf(RS2!imp_adicional = 0, Trim(RS!dec_rutpro), "")
            bseg_sgtxt = IIf(RS2!imp_adicional = 0, Trim(RS2!cta_nombre), "")
            bseg_kostl = IIf(RS2!imp_adicional = 0, Trim(MuestraCasino(1)), Trim(MuestraCasino(1)))
            If vg_pais = "CL" Then
               
               n_acodimpto = IIf(IsNull(RS2!imp_cimsap1), "", Trim(RS2!imp_cimsap1))
            
            ElseIf vg_pais = "CO" Then
               
               n_acodimpto = ""
               
               If RS2!pro_tippro = "0" And RS2!cli_tipope = "0" Then
                  
                  n_acodimpto = Trim(IIf(IsNull(RS2!imp_cimsap1), "", RS2!imp_cimsap1))
               
               ElseIf RS2!pro_tippro = "0" And RS2!cli_tipope = "1" Then
                  
                  n_acodimpto = Trim(IIf(IsNull(RS2!imp_cimsap2), "", RS2!imp_cimsap2))
               
               ElseIf RS2!pro_tippro = "1" And RS2!cli_tipope = "0" Then
                  
                  n_acodimpto = Trim(IIf(IsNull(RS2!imp_cimsap3), "", RS2!imp_cimsap3))
               
               ElseIf RS2!pro_tippro = "1" And RS2!cli_tipope = "1" Then
                  
                  n_acodimpto = Trim(IIf(IsNull(RS2!imp_cimsap4), "", RS2!imp_cimsap4))
               
               End If
            
            End If
            n_actaimpto = IIf(IsNull(RS2!imp_codsap) Or Trim(RS2!imp_codsap) = "", Trim(RS2!pro_ctacon), Trim(RS2!imp_codsap))
            If RS2!imp_adicional = "0" Then
               
               valimp = Format((RS2!valimp + (flete * (RS2!imp_pctimp / 100))), fg_Pict(9, 0))
               totdoc = (totdoc + valfac)
               candec = candec + ((RS2!valimp + (flete * (RS2!imp_pctimp / 100))) - Round(RS2!valimp + (flete * (RS2!imp_pctimp / 100)), 0)) 'Right(RS2!valimp, 2)
            
            Else
               
               valimp = Format(RS2!valimp, fg_Pict(9, 0))
               candec = candec + (RS2!valimp - Round(RS2!valimp, 0))
            
            End If
            totdoc = totdoc + valimp
            
            If j = RS2.RecordCount Then
               
               If totdoc > RS!toc_totdoc Then
                  
                  valimp = valimp + (RS!toc_totdoc - totdoc)
               
               ElseIf (RS!toc_totdoc - totdoc) = 1 Then
                  
                  valimp = valimp + 1
               
               ElseIf (RS!toc_totdoc - totdoc) = 2 Then
                  
                  valimp = valimp + 2
               
               ElseIf RS!toc_totdoc < totdoc Then
                  
                  valimp = valimp + Round(candec)
               
               End If
            
            End If
         
            n_amonimp = IIf(Not IsNull(RS2!valimp) And RS2!valimp > 0, Format(valimp), IIf(excimp, "", "0"))
            n_aimprecu = IIf(RS2!imp_adicional = "0", "NO", "SI")
            n_aotrimp = IIf(exiimp And excimp, "SI", "NO")
            vg_db.Execute "INSERT INTO sap_cfc VALUES (" & codigo & ", " & numlin & ", '" & n_a & "', '" & bkpf_bukrs & "', '" & bkpf_blart & "', '" & bkpf_budat & "', " & _
                          "'" & bkpf_bldat & "', '" & bkpf_xblnr & "', '" & bkpf_bktxt & "', '" & bkpf_waers & "', '" & bseg_newbs & "', " & _
                          "'" & bseg_newko & "', '" & bseg_wrbtr & "', '" & bseg_zuonr & "', '" & bseg_sgtxt & "', '" & bseg_kostl & "', " & _
                          "'" & n_acodimpto & "', '" & n_actaimpto & "', '" & n_amonimp & "', '" & n_aimprecu & "', '" & n_aotrimp & "')"
            numlin = numlin + 1
            RS2.MoveNext
            i = i + 1
            j = j + 1
            
         Loop
      
      End If
      RS2.Close
      Set RS2 = Nothing
      '-------> Fin traer Impuesto adicionales
      RS1.MoveNext
   
   Loop
   RS1.Close
   Set RS1 = Nothing
   
   '-------> Grabar log proceso
   fecenv = IIf(vg_tipbase = "1", Format(Date, "dd-mm-yyyy") & " " & Format(Time, "h:m:s"), Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s"))
   numero = 0
   
   If RS2.State = 1 Then RS2.Close
   RS2.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS2 = vg_db.Execute("SELECT numero FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' ORDER BY numero DESC")
   If Not RS2.EOF Then RS2.MoveFirst: numero = RS2!numero + 1 Else numero = 1
   RS2.Close
   Set RS2 = Nothing
   
   vg_db.Execute "INSERT INTO log_procesos (cencos, numero, fecha, tipo_proceso, rut, tipo_documento, num_documento, num_cfc, estado, mensaje, envio) " & _
                 "VALUES ('" & MuestraCasino(1) & "', " & numero & ", '" & fecenv & "', '1', '" & RS!dec_rutpro & "', '" & RS!dec_tipdoc & "' , '" & RS!dec_numdoc & "',  " & numfol & ", '0', '', " & codigo & ")"
  
   '-------> Mover parametro Web Service
   parametro1 = "1"
   parametro2 = codigo
   parametro3 = MuestraCasino(1)

   '-------> Proceso envio Web Service
   DoEvents
   
   If ValidarOpEnvio(MuestraCasino(1), 1) Then
      
      If vg_tipbase = "1" Then
         
         vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & LCase(App.Path) & "\" & "|" & "" & "|" & "" & "|" & "" & "|" & "" & "|")
      
      Else
         
         vRet = Shell(Trim(dir_trabajo) & "WsSapPortal.exe " & Trim(parametro1) & "|" & Trim(parametro2) & "|" & Trim(parametro3) & "|" & LCase(App.Path) & "\" & "|" & vg_SqlNSvr & "|" & vg_SqlBase & "|" & vg_SqlNUsr & "|" & vg_SqlPass & "|")
      
      End If
      
      If vRet = 0 Then
         
         RS1.Close
         Set RS1 = Nothing
         Text1(0).text = Text1(0).text & "Proceso cancelado, no hay comunicación con Web Service"
         GenerarArcSap = False
         Exit Function
      
      End If
      
      DoEvents
      '-------> Dejar en espera el proceso
      
      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS2 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '1' AND estado = '0'")
      
      Do While Not RS2.EOF
         
         DoEvents
         RS2.Close
         Set RS2 = Nothing

         Set RS2 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '1' AND estado = '0'")
         
      Loop
      
      RS2.Close
      Set RS2 = Nothing
   
   End If
   
   If ValidarOpEnvio(MuestraCasino(1), 5) Or ValidarOpEnvio(MuestraCasino(1), 6) Then
      
      vg_db.Execute ("update log_procesos set estado = 1 WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '1'")
   
   End If
   
   '-------> Proceso de estado de envio
   
   If ValidarOpEnvio(MuestraCasino(1), 1) Then
      
      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS2 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '1'")
      
      estado = True
      
      If Not RS2.EOF Then
         
         StrMensaje = Trim(RS2!mensaje)
         
         If Len(StrMensaje) <> 0 Then
            
            If estsap = True And RS2!estado = "1" Then
               
               Text1(0).text = Text1(0).text & FechaHora & "Mensaje SAP : " & VgLinea
               Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
               estsap = False
            
            ElseIf RS2!estado = "2" Or RS2!estado = "0" Or RS2!estado = "3" Then
               
               Text1(0).text = Text1(0).text & VgLinea
               Text1(0).text = Text1(0).text & FechaHora & IIf(RS2!estado = "3", "Mensaje Error : ", "Mensaje SAP : ") & VgLinea
               Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
            
            End If
            
            Do While InStr(StrMensaje, ";") <> 0 And InStr(StrMensaje, ";") <> 1
               
               If StrMensaje <> "" Then
                  
                  nommen = Mid(StrMensaje, 1, InStr(StrMensaje, "|") - 1)
                  StrMensaje = Mid(StrMensaje, InStr(StrMensaje, "|") + 1)
                  Text1(0).text = Text1(0).text & FechaHora & Trim(nommen) & " Doc. SGP - " & Trim(RS!dec_tipdoc) & " - " & RS!dec_numdoc & VgLinea
                  If InStr(nommen, "timed out") <> 0 Or InStr(nommen, "No esta conectado a la internet") <> 0 Then esttim = False
               
               End If
            
            Loop
            
            If RS2!estado = "2" Or RS2!estado = "0" Or RS2!estado = "3" Then
               
               estado = False ': GenerarArcSap = False
               numenv = numenv + 1
            
            End If
         
         End If
      
      End If
      RS2.Close
      Set RS2 = Nothing
   
   ElseIf ValidarOpEnvio(MuestraCasino(1), 5) Or ValidarOpEnvio(MuestraCasino(1), 6) Then
      
      If RS2.State = 1 Then RS2.Close
      RS2.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      Set RS2 = vg_db.Execute("SELECT * FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND numero = " & numero & " AND tipo_proceso = '1'")
      estado = True
      If Not RS2.EOF Then
         
         StrMensaje = Trim(RS2!mensaje)
      
      End If
      RS2.Close
      Set RS2 = Nothing
   
   End If
   
   '-------> Grabar tabla b_totcompras se genero sin problema
   If estado Then
      
      vg_db.Execute "UPDATE b_totcompras SET toc_envsap = '1' WHERE toc_rutpro = '" & RS!dec_rutpro & "' AND toc_tipdoc = '" & RS!dec_tipdoc & "' AND toc_numdoc = " & RS!dec_numdoc & " AND toc_codbod = " & vg_codbod & " AND toc_numinf = " & numfol & ""
   
   End If
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

If ValidarOpEnvio(MuestraCasino(1), 1) Then
   
   Text1(0).text = Text1(0).text & FechaHora & "---------------------------------------------------------" & VgLinea
   Text1(0).text = Text1(0).text & VgLinea
   GenerarArcSap = IIf(numenv > 2 Or Not esttim, False, True)
   tiperr = IIf(numenv > 2, 3, IIf(numenv = 0, 1, 2))

ElseIf ValidarOpEnvio(MuestraCasino(1), 5) Or ValidarOpEnvio(MuestraCasino(1), 6) Then
   
   GenerarArcSap = True
   tiperr = 3

End If

Exit Function
Man_EnvioCfc:
If RS.State = 1 Then
   
   RS.Close

End If

If RS1.State = 1 Then
   
   RS1.Close

End If

If RS2.State = 1 Then
   
   RS2.Close

End If

If RS3.State = 1 Then
   
   RS3.Close

End If

If RS4.State = 1 Then
   
   RS4.Close

End If

If Err = 53 Or Err = 5 Then
   
   Text1(0).text = Text1(0).text & FechaHora & "No existe Ejecutable de envio..." & VgLinea
   vg_db.Execute "UPDATE log_procesos SET estado = '3', mensaje = 'No existe ejecutable, para procesar Web Service' WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso = '1' AND num_cfc = " & numfol & " AND numero = " & numero & ""
   GenerarArcSap = False
   Exit Function

End If
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End Function

