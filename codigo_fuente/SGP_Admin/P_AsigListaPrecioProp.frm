VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form P_AsigListaPrecioProp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Asignar Lista Convenios Ceco Propuesta"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   10965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
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
      Left            =   8880
      TabIndex        =   10
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   6720
      TabIndex        =   9
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin VB.OptionButton Option1 
         Caption         =   "Eliminar Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   12
         Top             =   1200
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Asignar Contrato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   11
         Top             =   1200
         Value           =   -1  'True
         Width           =   1695
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1755
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
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1755
         TabIndex        =   2
         Top             =   720
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3030
         Picture         =   "P_AsigListaPrecioProp.frx":0000
         Top             =   660
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
         Index           =   1
         Left            =   3480
         TabIndex        =   6
         Top             =   735
         Width           =   6735
      End
      Begin VB.Label Label2 
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
         Index           =   3
         Left            =   465
         TabIndex        =   5
         Top             =   780
         Width           =   1155
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3045
         Picture         =   "P_AsigListaPrecioProp.frx":030A
         Top             =   240
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
         Index           =   0
         Left            =   3495
         TabIndex        =   4
         Top             =   315
         Width           =   6735
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
         Left            =   480
         TabIndex        =   3
         Top             =   420
         Width           =   735
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
         Index           =   0
         Left            =   3540
         TabIndex        =   8
         Top             =   360
         Width           =   6735
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
         Index           =   1
         Left            =   3525
         TabIndex        =   7
         Top             =   780
         Width           =   6735
      End
   End
End
Attribute VB_Name = "P_AsigListaPrecioProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim NombreCeco As String

Select Case Index
Case 0
               
    If ValidaDatos = False Then Exit Sub
           
    NombreCeco = fpayuda(0).Caption
    '--> Validar ceco
    Sql = ""
    Sql = " '" & LimpiaDato(Trim(fpText.text)) & "' "
    
    If Option1(0).Value = True Then
        
        Set RS = vg_db.Execute("sgpadm_Sel_ValidarCecoOrgCompras " & Sql & "")
        If Not RS.EOF Then
           MsgBox "Existe ceco organización ..." & " - " & RS!ID_Orgcompra & " - " & RS!NOMBRE_ORG, vbExclamation + vbOKOnly, Me.Caption
           RS.Close
           Set RS = Nothing
           Exit Sub
        End If
        RS.Close
        Set RS = Nothing
    
    Else
    
        '--> Validar si Ceco corresponde Propuesta
        Set RS = vg_db.Execute("sgpadm_Sel_ValidarCecoPropuesta " & Sql & "")
        If RS.EOF Then
           MsgBox "Ceco no corresponde propuesta ...", vbExclamation + vbOKOnly, Me.Caption
           RS.Close
           Set RS = Nothing
           Exit Sub
        End If
        RS.Close
        Set RS = Nothing
    
        '--> Validar que Ceco exista
        Sql = ""
        Sql = " '" & LimpiaDato(Trim(fpText.text)) & "', "
        Sql = Sql + " '" & LimpiaDato(Trim(fpText1.text)) & "' "
        Set RS = vg_db.Execute("sgpadm_Sel_ValidarCecoPropuestaOrgCompras " & Sql & "")
        If RS.EOF Then
           MsgBox "Contrato no existe información ...", vbExclamation + vbOKOnly, Me.Caption
           RS.Close
           Set RS = Nothing
           Exit Sub
        End If
        RS.Close
        Set RS = Nothing
    
    End If
    
    If MsgBox("Esta seguro crear Ceco...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '--> crear ceco org. compras & crear lista
    Sql = ""
    Sql = " '" & LimpiaDato(Trim(fpText.text)) & "', "
    Sql = Sql & " '" & LimpiaDato(Trim(fpText1.text)) & "', "
    Sql = Sql & " '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "' "
    
    If Option1(0).Value = True Then
    
        Set RS = vg_db.Execute("sgpadm_Ins_OrganizacionCompras_V02 " & Sql & "")
    
    Else
    
        Set RS = vg_db.Execute("sgpadm_Del_OrganizacionCompras " & Sql & "")
    
    End If
    
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
       
       Else
          
          MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing

Case 1

    Unload Me
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Function ValidaDatos() As Boolean

On Error GoTo Man_Error

Dim RS      As New ADODB.Recordset
Dim Dias    As Long
Dim i       As Long
Dim Fecha   As String
Dim mes     As String
Dim Año     As String

Let ValidaDatos = True
 
'-------> Validar Ceco
Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
If RS.EOF Then
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   MsgBox "No existe Contrato...", vbExclamation + vbOKOnly, MsgTitulo
   Let ValidaDatos = False
   Exit Function
End If
RS.Close
Set RS = Nothing

'-------> Validar Org. Compras
If Trim(LimpiaDato(fpText1.text)) = "" Then
    Call MsgBox("Debe Ingresar Org. Compras", vbInformation, Me.Caption)
    Call fpText1.SetFocus
    Let ValidaDatos = False
    Exit Function
End If

Exit Function
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  

End Function

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()
    
On Error GoTo Man_Error
    
Me.Caption = "Asignar Lista Convenios Ceco Propuesta"
    
Call fg_carga("")
Me.HelpContextID = vg_OpcM
Call fg_centra(Me)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgpadm_s_cliente_V02 45, '" & LimpiaDato(fpText.text) & "', ''")
If RS.EOF Then
   
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   fpayuda(1).Caption = ""
   fpText1.text = ""
   Exit Sub

End If
fpayuda(0).Caption = Trim(RS!Cli_nombre)
fpText.text = RS!Cli_codigo
RS.Close
Set RS = Nothing
 
fpText1.text = ""
fpayuda(1).Caption = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgpadm_Sel_BuscarOrgCompras_V02 '" & LimpiaDato(fpText1.text) & "'")
If RS.EOF Then
   RS.Close
   Set RS = Nothing
   fpayuda(1).Caption = ""
   Exit Sub
End If
fpayuda(1).Caption = Trim(RS!ID_Orgcompra)
fpText1.text = RS!ID_Orgcompra
RS.Close
Set RS = Nothing
 
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Index

Case 0
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("b_clientes", "cli_", "Clientes", "Cliente_SitioRemoto")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    fpText1.text = ""
    Let fpayuda(1).Caption = ""
    fpText1.SetFocus

Case 1
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("Org. Compras", "", "Org. Compras", "Celo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText1.text = vg_codigo
    fpayuda(1).Caption = vg_nombre

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub
