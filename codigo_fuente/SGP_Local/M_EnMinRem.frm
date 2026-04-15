VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_EnMinRem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio Bloque Minuta"
   ClientHeight    =   6240
   ClientLeft      =   2100
   ClientTop       =   2325
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5745
      Left            =   30
      TabIndex        =   2
      Top             =   360
      Width           =   15195
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   5505
         Visible         =   0   'False
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1230
         TabIndex        =   0
         Top             =   270
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
         _ExtentY        =   556
         Enabled         =   0   'False
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
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4215
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   14925
         _Version        =   393216
         _ExtentX        =   26326
         _ExtentY        =   7435
         _StockProps     =   64
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   9
         SpreadDesigner  =   "M_EnMinRem.frx":0000
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   13425
         Top             =   5400
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00D9D9FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   11280
         Top             =   5400
         Visible         =   0   'False
         Width           =   300
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
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2745
         Picture         =   "M_EnMinRem.frx":1B82
         Top             =   180
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3255
         TabIndex        =   5
         Top             =   270
         Width           =   3975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha(mm/aa)"
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
         Left            =   120
         TabIndex        =   4
         Top             =   5235
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3300
         TabIndex        =   6
         Top             =   315
         Width           =   3975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_EnMinRem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Text
Dim Msgtitulo As String, modo As String
Dim Fecha As String

Private Sub Form_Activate()
    Call fg_descarga
End Sub

Private Sub Form_Load()
    Me.HelpContextID = vg_OpcM
    Me.Height = 6720
    Me.Width = 15645
    Call fg_centra(Me)
    Msgtitulo = "Envio Minuta Bloque"
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Enviar minuta ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    fpText.Enabled = ModCasino
    Image1(0).Enabled = ModCasino
    fpText.text = MuestraCasino(1)
    fpayuda(0).Caption = MuestraCasino(2)
    MoverDatoGrilla
End Sub

Private Sub fpText_Change()
    Dim RS As New ADODB.Recordset
    RS.Open "SELECT cli_nombre " & _
            "FROM b_clientes WITH ( NOLOCK ) " & _
            "WHERE cli_codigo = '" & Trim(LimpiaDato(fpText.text)) & "' " & _
            "AND cli_tipo = 0 and cli_tipominuta = 1", vg_db, adOpenStatic
    If RS.EOF Then
        RS.Close
        Set RS = Nothing
        fpayuda(0).Caption = ""
        MoverDatoGrilla
        Exit Sub
    End If
    fpayuda(0).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
    MoverDatoGrilla
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
    Select Case Index
        Case 0
            vg_left = fpayuda(0).Left + 6300
            vg_nombre = "": vg_codigo = ""
            B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
            B_TabEst.Show 1
            Me.Refresh
            If vg_codigo = "" Then Exit Sub
            fpText.text = vg_codigo
            fpayuda(0).Caption = vg_nombre
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index
        Case 2 '-------> Envio minuta
            Call EnviaMinuta
        Case 4 '-------> Salir
           Me.Hide
           Unload Me
    End Select

End Sub

Private Sub EnviaMinuta()
    Dim RS As New ADODB.Recordset
    Dim Item As String
    Dim i As Long, spid  As Long
    Dim isel As Boolean
    Dim id As Long
    Dim IdMinutaBloque As Long
    
On Error GoTo Man_Error

    If Len(fpText.text) = 0 Or Trim(fpayuda(0).Caption) = "" Then
        Call MsgBox("Debe Ingresar Centro de costo", vbInformation, Me.Caption)
        Exit Sub
    End If
    '-------> Validar que exista un item seleccionado
    isel = False
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then isel = True: Exit For
    Next i
    If Not isel Then fg_descarga: MsgBox "Debe seleccionar a lo menos un items de la lista", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    
    '-------> validar si existe minuta bloque
    isel = True
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
            vaSpread1.Col = 2
            IdMinutaBloque = CDbl(vaSpread1.text)
        
            vaSpread1.Col = 9
            vaSpread1.text = ""
    
            Set RS = vg_db.Execute("sgp_Sel_ValidarMinutaBloque 1, " & IdMinutaBloque & ", '" & Trim(LimpiaDato(fpText.text)) & "'")
            If RS.EOF Then
               vaSpread1.text = "No existe Minuta, para ser enviada"
               isel = False
               vaSpread1.BackColor = Shape1(1).FillColor
            Else
               vaSpread1.BackColor = Shape1(0).FillColor
            End If
            RS.Close: Set RS = Nothing
        End If
    Next i
    If Not isel Then fg_descarga: MsgBox "No existe Minuta, para ser enviada", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    
    '-------> validar si minuta esta bloqueda
    isel = True
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
            vaSpread1.Col = 2
            IdMinutaBloque = CDbl(vaSpread1.text)
            
            vaSpread1.Col = 9
            vaSpread1.text = ""
    
            Set RS = vg_db.Execute("sgp_Sel_ValidarMinutaBloque 2, " & IdMinutaBloque & ", '" & Trim(LimpiaDato(fpText.text)) & "'")
            If RS.EOF Then
               vaSpread1.text = "Minuta esta registrada como enviada"
               isel = False
               vaSpread1.BackColor = Shape1(1).FillColor
            Else
               vaSpread1.BackColor = Shape1(0).FillColor
            End If
            RS.Close: Set RS = Nothing
        End If
    Next i
    If Not isel Then fg_descarga: MsgBox "Minuta esta registrada como enviada", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    isel = True
    '-------> Generar envio minuta bloque
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
            vaSpread1.Col = 2
            IdMinutaBloque = CDbl(vaSpread1.text)
            
            vaSpread1.Col = 9
            vaSpread1.text = ""
    
            Set RS = vg_db.Execute("sgp_Upd_EnvioMinutaBloque '" & Trim(LimpiaDato(fpText.text)) & "', " & IdMinutaBloque & "")
            Call fg_descarga
            If Not RS.EOF Then
               If UCase(RS(0)) = "OK" Then
                  vaSpread1.text = "Registro enviado"
                  vaSpread1.BackColor = Shape1(0).FillColor
               Else
                  isel = False
                  vaSpread1.text = "Registro finalizo con error" & RS(0)
                  vaSpread1.BackColor = Shape1(1).FillColor
               End If
            End If
            RS.Close: Set RS = Nothing
        End If
    Next i
    Call fg_descarga
    If isel Then
       Call MsgBox("Proceso de Envio Minuta Finalizado", vbInformation + vbOKOnly, Msgtitulo)
       MoverDatoGrilla
    Else
       Call MsgBox("Proceso Finalizado con problema", vbInformation + vbOKOnly, Msgtitulo)
    End If
    Bar1(0).Visible = False: Bar1(0).Value = 0
    Label1(1).Visible = False
    Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, Msgtitulo
End Sub

Sub MoverDatoGrilla()
Dim RS As New ADODB.Recordset
Dim i As Long
Dim arr
Dim str_ As String
Dim fd As String
Dim fh As String
fg_carga ""
Set RS = vg_db.Execute("sgp_Sel_ListarMinutaBloqueEnviar '" & Trim(LimpiaDato(fpText.text)) & "'")
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = Shape1(0).FillColor
vaSpread1.MaxRows = 0
If Not RS.EOF Then
    arr = RS.GetRows
    RS.Close: Set RS = Nothing
    For i = 0 To UBound(arr, 2)
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 2
        vaSpread1.text = CStr(arr(0, i))
        
        vaSpread1.Col = 3
        vaSpread1.text = arr(1, i)
        
        vaSpread1.Col = 4
        vaSpread1.text = Trim(arr(2, i))
        
        vaSpread1.Col = 5
        vaSpread1.text = arr(3, i)
        
        vaSpread1.Col = 6
        vaSpread1.text = Trim(arr(4, i))
        
        vaSpread1.Col = 7
        vaSpread1.text = Trim(arr(5, i))

        vaSpread1.Col = 8
        vaSpread1.text = Trim(arr(6, i))

        vaSpread1.Col = 9
        vaSpread1.text = ""
    Next i
End If
If RS.State = 1 Then
   RS.Close: Set RS = Nothing
End If
fg_descarga
End Sub
