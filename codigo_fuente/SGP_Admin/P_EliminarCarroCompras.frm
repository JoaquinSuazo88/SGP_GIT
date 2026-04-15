VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form P_EliminarCarroCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar Carros de Compras"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12660
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   12660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   6015
      Left            =   120
      TabIndex        =   13
      Top             =   1560
      Width           =   12375
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5535
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   11895
         _Version        =   393216
         _ExtentX        =   20981
         _ExtentY        =   9763
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
         MaxCols         =   7
         SpreadDesigner  =   "P_EliminarCarroCompras.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   12375
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1980
         TabIndex        =   1
         Top             =   735
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   2
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
         Text            =   "01/09/2013"
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   9165
         TabIndex        =   2
         Top             =   735
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   2
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
         Text            =   "28/09/2013"
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1995
         TabIndex        =   0
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   10800
         TabIndex        =   3
         Top             =   720
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
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
         Index           =   0
         Left            =   720
         TabIndex        =   11
         Top             =   420
         Width           =   1155
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
         Left            =   3735
         TabIndex        =   10
         Top             =   315
         Width           =   6735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3285
         Picture         =   "P_EliminarCarroCompras.frx":1A07
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha desde"
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
         Left            =   705
         TabIndex        =   9
         Top             =   825
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha hasta"
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
         Left            =   7860
         TabIndex        =   8
         Top             =   825
         Width           =   1065
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
         Left            =   3780
         TabIndex        =   12
         Top             =   360
         Width           =   6735
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
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
      Left            =   8400
      TabIndex        =   5
      Top             =   7800
      Width           =   1935
   End
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
      Left            =   10560
      TabIndex        =   6
      Top             =   7800
      Width           =   1935
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "P_EliminarCarroCompras.frx":1D11
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "P_EliminarCarroCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Public lc_Aux As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim Sql         As String
Dim i           As Long
Dim seleccion   As String
Dim Ceco        As String
Dim Id_Pedido   As Long
Dim EstBorrado  As Boolean

Select Case Index

Case 0

  If ValidaDatos("2") = False Then Exit Sub


  '-------> Rescata Ceco Seleccionado
  seleccion = 0
  EstBorrado = True
  fg_carga ""
  
  For i = 1 To vaSpread1.MaxRows
  
      vaSpread1.Row = i
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
      If seleccion = 1 And vaSpread1.RowHidden = False Then
          
          Ceco = ""
          vaSpread1.Col = 2
          Ceco = vaSpread1.text
          
          Id_Pedido = 0
          vaSpread1.Col = 4
          Id_Pedido = vaSpread1.text
          
          Sql = ""
          Sql = Sql + " '" & Ceco & "', "
          Sql = Sql + " " & Id_Pedido & ", "
          Sql = Sql & " '" & UCase(LimpiaDato(Trim(vg_NUsr))) & "' "
          Set RS = vg_db.Execute("sgpadm_Del_CarroCompras " & Sql & "")
          
          If Not RS.EOF Then
  
            If RS(0) > 0 Then
          
                vaSpread1.Col = 7
                vaSpread1.text = RS(0) & " " & RS(1)
       
                vaSpread1.Row = i: vaSpread1.Col = -1
                vaSpread1.BackColor = &H8080FF
            
                EstBorrado = False
                
            Else
          
                vaSpread1.Col = 7
                vaSpread1.text = "Finalizo sin problema" 'RS(0) & " " & RS(1)
                
                vaSpread1.Row = i: vaSpread1.Col = -1
                vaSpread1.BackColor = &HC0FFFF
                
            End If
  
          End If
       
      
      End If
  
      DoEvents
       
  Next i

  fg_descarga
    
  MsgBox IIf(EstBorrado = True, "Proceso finalizo sin problema...", "Proceso finalizo con problema..."), IIf(EstBorrado = True, vbOKOnly, vbCritical) + vbOKOnly, MsgTitulo
Case 1

  Unload Me
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()
    
On Error GoTo Man_Error
    
Dim RS As New ADODB.Recordset
Dim Sql As String

Me.Caption = "Eliminar Carros de Compras"
    
Call fg_carga("")
Me.HelpContextID = vg_OpcM
Call fg_centra(Me)

vaSpread1.MaxRows = 0

Sql = " sgpadm_Sel_leemaximofecha "
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
   FpFecDesde.text = Format(RS(0), "dd/mm/yyyy")
   FpFecHasta.text = Format(RS(1), "dd/mm/yyyy")
End If
  
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Set RS = vg_db.Execute("sgpadm_Sel_BuscarOrgCompras_V02 '" & LimpiaDato(fpText.text) & "'")
If RS.EOF Then
   RS.Close
   Set RS = Nothing
   fpayuda(0).Caption = ""
   Exit Sub
End If
fpayuda(0).Caption = Trim(RS!ID_ORGCOMPRA)
fpText.text = RS!ID_ORGCOMPRA
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

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    Call B_TabEst.LlenaDatos("Org. Compras", "", "Org. Compras", "Celo")
    Call B_TabEst.Show(1)
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText.text = vg_codigo
    fpayuda(0).Caption = vg_nombre
    FpFecDesde.SetFocus
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim Sql       As String
Dim i         As Long
Dim xmlceco   As String
Dim seleccion As String
Dim codCeco   As String

Select Case Button.Index
Case 1

  If ValidaDatos("1") = False Then Exit Sub
      
  vaSpread1.MaxRows = 0
  vaSpread1.Row = -1: vaSpread1.Col = -1
  vaSpread1.BackColor = &HC0FFFF
   
  Sql = ""
  Sql = Sql & " '" & Trim(LimpiaDato(fpText.text)) & "' , "
  Sql = Sql & " '" & Format(FpFecDesde, ("YYYYmmdd")) & "', "
  Sql = Sql & " '" & Format(FpFecHasta, ("YYYYmmdd")) & "' "
  Set RS = vg_db.Execute("sgpadm_Sel_CarrosCompras  " & Sql & "")
  If Not RS.EOF Then
     
     Do While Not RS.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.text = "0"
      
      vaSpread1.Col = 2
      vaSpread1.text = RS!cencos
      
      vaSpread1.Col = 3
      vaSpread1.text = Trim(RS!cli_nombre)
      
      vaSpread1.Col = 4
      vaSpread1.text = RS!idpedido
      
      vaSpread1.Col = 5
      vaSpread1.text = Trim(RS!tipopedido)
      
      vaSpread1.Col = 6
      vaSpread1.text = Trim(RS!estado)
      
      vaSpread1.Col = 7
      vaSpread1.text = ""
      
      RS.MoveNext
     Loop
     Command1(0).Enabled = True
  
  Else
     
     vaSpread1.MaxRows = 0
     Command1(0).Enabled = False
     MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
  
  End If
  RS.Close
  Set RS = Nothing

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidaDatos(ByVal op As String) As Boolean

On Error GoTo Man_Error

Let ValidaDatos = True
Dim i As Long
Dim estseleccion As Boolean
  
'-------> Validar org. compras
If Trim(fpText.text) = "" Then
   MsgBox "Debe ingresar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo
   Call fpText.SetFocus
   Let ValidaDatos = False
   Exit Function
End If
  
'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   Call FpFecHasta.SetFocus
   Let ValidaDatos = False
   Exit Function
End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   Let FpFecDesde.text = Format(Now, "dd/mm/yyyy")
   Call FpFecDesde.SetFocus
   Let ValidaDatos = False
   Exit Function
End If

If op = 2 Then

   estseleccion = False
   For i = 1 To vaSpread1.MaxRows
   
        vaSpread1.Row = i
        vaSpread1.Col = 1
   
        If vaSpread1.text = "1" Then
   
            estseleccion = True
            Exit For
   
        End If
   
   Next i
   
   If Not estseleccion Then
   
      MsgBox "Debe seleccionar un carro de la lista...", vbExclamation + vbOKOnly, MsgTitulo
      Let ValidaDatos = False
      Exit Function
   
   End If
   
End If

Exit Function
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Function

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    Dim i As Long
    vaSpread1.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    Next
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
