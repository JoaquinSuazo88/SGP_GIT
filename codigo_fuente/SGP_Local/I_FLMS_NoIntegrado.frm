VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form I_FLMS_NoIntegrado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos No Integrados (FLMStoSGP)"
   ClientHeight    =   3210
   ClientLeft      =   3360
   ClientTop       =   3060
   ClientWidth     =   8385
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8385
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2475
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   8025
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1080
         Width           =   2475
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar Reporte"
         Height          =   735
         Left            =   5880
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   4755
         TabIndex        =   1
         Top             =   315
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
         Text            =   "04/01/2005"
         DateCalcMethod  =   3
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   8
         Top             =   360
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
         Text            =   "04/01/2005"
         DateCalcMethod  =   3
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   225
         TabIndex        =   3
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   3510
         TabIndex        =   2
         Top             =   360
         Width           =   1005
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   8385
      _ExtentX        =   14790
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_FLMS_NoIntegrado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim MsgTitulo As String
Public lc_Aux As String
Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim TipoDocumento   As String
'Dim Sql             As String
'Dim NomArchivoExcel As String
'Dim Extension       As String

'If Not ValidarDatos Then Exit Sub

fg_carga ""

'Exportar Excel
Dim NombreArchivoExcel As String
       
'-------> Crear directorio guias Logistico
If Dir(dir_trabajo_Inf & "" & "FLMStoSGP\", vbDirectory) = "" Then
       
   MkDir dir_trabajo_Inf & "" & "FLMStoSGP\"
          
End If
'-------> Fin crear directorio guias Logistico

'Validar si existe información
       
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If fg_codigocbo(Combo1, 0, 1, "") = 1 Then
    TipoDocumento = ""
ElseIf fg_codigocbo(Combo1, 0, 1, "") = 2 Then
    TipoDocumento = "SP"
ElseIf fg_codigocbo(Combo1, 0, 1, "") = 3 Then
    TipoDocumento = "DP"
ElseIf fg_codigocbo(Combo1, 0, 1, "") = 4 Then
    TipoDocumento = "MS"
ElseIf fg_codigocbo(Combo1, 0, 1, "") = 5 Then
    TipoDocumento = "ME"
End If

Set RS = vg_db.Execute("SGP_S_ReporteDatosNoIntegrados_FLMStoSGP " & Format(fpDateTime1(0).Value, "yyyymmdd") & ", " & Format(fpDateTime1(1).Value, "yyyymmdd") & ", '" & TipoDocumento & "'")
       
If RS.EOF Then
                
   fg_descarga
          
   MsgBox "No existe información, con los parametros indicados.. ", vbInformation + vbOKOnly, MsgTitulo
             
   RS.Close
   Set RS = Nothing
   Exit Sub
          
End If
RS.Close
Set RS = Nothing

NombreArchivoExcel = "ReporteDatosNoIntegrados_" & Format(Date, "yyyymmdd") & "_" & Format(Time, "HHMMSS")
Generar_ArchivoExcel "SGP_S_ReporteDatosNoIntegrados_FLMStoSGP " & Format(fpDateTime1(0).Value, "yyyymmdd") & ", " & Format(fpDateTime1(1).Value, "yyyymmdd") & ", '" & TipoDocumento & "'", dir_trabajo_Inf & "FLMStoSGP\", NombreArchivoExcel

fg_descarga
MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
Exit Sub
Man_Error:
    
    If RS.State = 1 Then RS.Close
    
    Call fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Height = 3500
Me.Width = 8500
Me.HelpContextID = vg_OpcM
fg_centra Me
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'fpDateTime1(0).DateTimeFormat = UserDefined
'fpDateTime1(0).UserDefinedFormat = "mm/yyyy"
'fpDateTime1(0).text = Format(Date, "mm/yyyy")

fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
fpDateTime1(1).text = Format(Date, "dd/mm/yyyy")


Combo1(0).Clear
Combo1(0).AddItem "Todos" & Space(150) & "(1)"
Combo1(0).AddItem "Salidas de Bodega" & Space(150) & "(2)"
Combo1(0).AddItem "Devoluciones a Bodega" & Space(150) & "(3)"
Combo1(0).AddItem "Raciones No Vendidas" & Space(150) & "(4)"
Combo1(0).AddItem "Mermas de Bodega" & Space(150) & "(5)"
Combo1(0).ListIndex = 0


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If Trim(fpDateTime1(0).text) = "" Or Trim(fpDateTime1(1).text) = "" Then Exit Sub
If Not IsDate(fpDateTime1(0).text) Or Not IsDate(fpDateTime1(1).text) Then Exit Sub

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

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
    
        Image1_Click 0
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_LostFocus()

On Error GoTo Man_Error



Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
        
        fpDateTime1(0).SetFocus

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo ManError

Select Case Button.Index
    
    Case 1
        
        fg_carga ""
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        
        fg_descarga
    
    Case 3
        
        Me.Hide
        Unload Me

End Select

Exit Sub
ManError:
fg_descarga
End Sub
