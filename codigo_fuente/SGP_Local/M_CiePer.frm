VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_CiePer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario de Cierre de Mes"
   ClientHeight    =   7320
   ClientLeft      =   5280
   ClientTop       =   3600
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6765
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6850
      Begin VB.Frame Frame7 
         Height          =   2895
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Visible         =   0   'False
         Width           =   6375
         Begin VB.CommandButton Cmd1 
            Caption         =   "&Aceptar"
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
            Left            =   3240
            TabIndex        =   13
            Top             =   2280
            Width           =   1305
         End
         Begin VB.CommandButton Cmd2 
            Caption         =   "&Cancelar"
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
            Left            =   4680
            TabIndex        =   15
            Top             =   2280
            Width           =   1425
         End
         Begin EditLib.fpText Nombre 
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   11
            Top             =   1440
            Width           =   1695
            _Version        =   196608
            _ExtentX        =   2990
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
            MaxLength       =   255
            MultiLine       =   0   'False
            PasswordChar    =   "*"
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   1
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText Nombre 
            Height          =   315
            Index           =   0
            Left            =   2790
            TabIndex        =   9
            Top             =   1080
            Width           =   1695
            _Version        =   196608
            _ExtentX        =   2990
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
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
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
            MaxLength       =   20
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   1
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Lb1 
            AutoSize        =   -1  'True
            Caption         =   "Perido Reabrir : "
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
            Index           =   5
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1410
         End
         Begin VB.Label Lb1 
            AutoSize        =   -1  'True
            Caption         =   "Password"
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
            Left            =   1440
            TabIndex        =   14
            Top             =   1500
            Width           =   930
         End
         Begin VB.Label Lb1 
            AutoSize        =   -1  'True
            Caption         =   "Login"
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
            Left            =   1440
            TabIndex        =   12
            Top             =   1125
            Width           =   585
         End
         Begin VB.Label Lb1 
            AutoSize        =   -1  'True
            Caption         =   "Para reabrir el periodo, tiene comunicarse su monitor"
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
            Index           =   8
            Left            =   240
            TabIndex        =   10
            Top             =   360
            Width           =   4500
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   915
         TabIndex        =   3
         Top             =   210
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
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5325
         Left            =   630
         TabIndex        =   1
         Top             =   840
         Width           =   5610
         _Version        =   393216
         _ExtentX        =   9895
         _ExtentY        =   9393
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   4
         MaxRows         =   25
         SpreadDesigner  =   "M_CiePer.frx":0000
      End
      Begin VB.Label Label1 
         Caption         =   "Espere un momento. Realizando cierre mensual..."
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
         Left            =   1080
         TabIndex        =   7
         Top             =   6360
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2145
         Picture         =   "M_CiePer.frx":0445
         Top             =   120
         Width           =   480
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
         TabIndex        =   5
         Top             =   280
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2655
         TabIndex        =   4
         Top             =   210
         Width           =   3975
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2700
         TabIndex        =   6
         Top             =   255
         Width           =   3975
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_CiePer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS        As New ADODB.Recordset
Dim modo      As String
Dim Fecha     As String
Dim MsgTitulo As String
Dim IrowRea   As Long

Private Sub Cmd1_Click()

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim FechaAbrir  As Long
Dim FechaHabi   As Long
Dim Ceco        As String
Dim fecper      As String

    '-------> Validar usuario
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_param WHERE par_valor = '" & LimpiaDato(Trim(Nombre(0).text)) & "' AND par_codigo = 'usulimbas' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If RS.EOF Then
       
       MsgBox "Usuario no existe..."
       RS.Close
       Set RS = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    RS.Close
    Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_param WHERE par_codigo = 'parconreme' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    
    If Not RS.EOF And UCase(Nombre(1).text) <> UCase(fg_Desencripta(TipoDato(RS!par_valor, ""))) Then
       
       MsgBox "La clave no corresponde al login..."
       RS.Close
       Set RS = Nothing
       Nombre(0).text = ""
       Nombre(0).SetFocus
       Exit Sub
    
    End If
    
    RS.Close
    Set RS = Nothing
               
    With vaSpread1
        
        '-------> Actualizando cerrando periodo y abriendo proximo periodo
    
        .Row = IrowRea
        .Col = 3
        fecper = "XXX" & Format(.text, "yyyymm")
    
        FechaAbrir = Val(Format(.text, "yyyymmdd"))
            
        .Row = IrowRea
        .Lock = False
            
        .Row = IrowRea + 1
        .Col = 3
    
        FechaHabi = Val(Format(.text, "yyyymmdd"))
            
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgp_Upd_Reabrir_Inhabilitar_SolNotaCredito '" & MuestraCasino(1) & "', " & FechaAbrir & ", " & FechaHabi & ", '" & fecper & "', " & vg_codbod & "")
        
        If Not RS.EOF Then
                
           If RS(0) > 0 And Trim(RS(1)) <> "" Then
                   
              RS.Close
              Set RS = Nothing
                      
              MsgBox "Existe error en la actualización cambio de periodo. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo
              
              Frame7.Visible = False
              Nombre(0).text = ""
              Nombre(1).text = ""
              Lb1(5).Caption = ""
              
              Toolbar1.Enabled = True
              Frame1.Enabled = True
              vaSpread1.Enabled = True
              
              Exit Sub
                
           End If
                
        End If
        RS.Close
        Set RS = Nothing
                
        '-------> Traer periodo
        Partida.StatusBar1.Panels(7).text = "Periodo : "
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
        If Not RS.EOF Then
    
            Partida.StatusBar1.Panels(7).text = "Periodo : " & Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)
    
        End If
        RS.Close
        Set RS = Nothing
            
        .Row = IrowRea
        .Col = 4
        .text = "Abierto"
    
    
        .Row = IrowRea + 1
        .Lock = True
        .Col = 4
        .text = "Inhabilitado"
        '-------> Fin actualizando cerrando periodo y abriendo proximo periodo
    
    End With
 
    Frame7.Visible = False
    Nombre(0).text = ""
    Nombre(1).text = ""
    Lb1(5).Caption = ""
    
    Toolbar1.Enabled = True
    Frame1.Enabled = True
    vaSpread1.Enabled = True
    
Exit Sub
Man_Error:
fg_descarga

Frame7.Visible = False
Nombre(0).text = ""
Nombre(1).text = ""
Lb1(5).Caption = ""

Toolbar1.Enabled = True
Frame1.Enabled = True
vaSpread1.Enabled = True

MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
    
End Sub

Private Sub Cmd2_Click()

On Error GoTo Man_Error

Frame7.Visible = False
Lb1(5).Caption = ""

Frame1.Enabled = True
Toolbar1.Enabled = True
vaSpread1.Enabled = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 7755
Me.Width = 7080
fg_centra Me
modo = "M"
MsgTitulo = "Calendario de Cierre de Mes"
Gl_Mo_Botones Me, 10
Gl_Ac_Botones Me, 10, 5, modo
fpText.Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText.text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
LlenarDatos

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change()

On Error GoTo Man_Error

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT * FROM b_clientes WHERE cli_codigo='" & fpText.text & "' AND cli_tipo=0", vg_db, adOpenStatic
If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "":: Exit Sub
fpayuda(0).Caption = Trim(RS!cli_nombre)
RS.Close
Set RS = Nothing

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
        vg_nombre = ""
        vg_codigo = ""
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

Dim RS           As New ADODB.Recordset
Dim fecper       As Long
Dim FecPerMostar As String
Dim iRow         As Long

With vaSpread1
    
    Select Case Button.Index
        
        Case 1 '-------> Cerrar Periodo
            
            If CierreAjuste Then Exit Sub
            
            If .MaxRows < 0 Then Exit Sub
            
            Dim fecha1 As Long, fecha2 As Long
            
            fecper = 0
            fecha1 = 0
            fecha2 = 0
            
            iRow = .ActiveRow
            .Row = iRow '.ActiveRow
            .Col = 4
            If Trim(.text) <> "Abierto" Then Exit Sub
            
            .Col = 1
            fecper = Format(.text, "yyyymm")
            
            .Col = 2
            fecha1 = Format(.text, "yyyymmdd")
            
            .Col = 3
            fecha2 = Format(.text, "yyyymmdd")
            
            'Validar inventario calendarizado 20201001
            If CierrePeriodo(fecha2, vg_codbod, 38) Then
        
               MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
           
            End If
        
            If CierrePeriodo(fecha2, vg_codbod, 4) Then
            
               MsgBox "No ha realizado la toma de inventario", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
            If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 9) Then
            
               MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
            If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 46) Then
            
               MsgBox "Existen documentos pendientes, en la salida ventas especiles. Debe cerrar las salidas ventas especiales", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
            If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 2) Then
            
               MsgBox "Existe información posterior, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
'Mod Ini 20240801            'Validar el proceso de generación caratula se realizo 20231123
'Mod Ini 20240801                        If RS.State = 1 Then RS.Close
'Mod Ini 20240801                        RS.CursorLocation = adUseClient
'Mod Ini 20240801                        vg_db.CursorLocation = adUseClient

'Mod Ini 20240801                        RS.Open "SELECT DISTINCT cencos FROM log_procesos WHERE cencos = '" & MuestraCasino(1) & "' AND tipo_proceso = '3' AND num_documento = '" & Format(.text, "yyyymmdd") & "' AND estado = '1' AND (anulado) IS NULL", vg_db, adOpenStatic
'Mod Ini 20240801                        If Not RS.EOF Then
            
'Mod Ini 20240801                           RS.Close
'Mod Ini 20240801                           Set RS = Nothing
'Mod Ini 20240801                           MsgBox "No ha realizado el envio de la caratula de inventario en la toma inventario. proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
'Mod Ini 20240801                           Exit Sub
            
'Mod Ini 20240801                        End If
'Mod Ini 20240801                        RS.Close
'Mod Ini 20240801                        Set RS = Nothing

            If MsgBox("Esta Seguro Cerrar Mes", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
            
            Label1.Visible = True
            Frame1.Enabled = False
            If Not GeneraMDBCierreMes(Me) Then
             
               Label1.Visible = False
               Frame1.Enabled = True
               MsgBox "Se produjo un error en el proceso de cierre diario, reintente nuevamente cerrar su periodo..", vbCritical + vbOKOnly, MsgTitulo
               Exit Sub
               
            End If
            
            '-------> Actualizando cerrando periodo y abriendo proximo periodo
            vg_db.BeginTrans
            
            vg_db.Execute "UPDATE b_cierreperiodo SET cie_estado=0 WHERE cie_cencos='" & Trim(fpText.text) & "' AND cie_fecter=" & Val(Format(.text, "yyyymmdd")) & ""
            
            .Row = iRow
            .Lock = True
            
            .Col = 4
            .text = "Cerrado"
            .Row = iRow + 1 '.ActiveRow + 1
            .Col = 3
            
            vg_db.Execute "UPDATE b_cierreperiodo SET cie_estado=1 WHERE cie_cencos='" & Trim(fpText.text) & "' AND cie_fecter=" & Val(Format(.text, "yyyymmdd")) & ""
            
            CalcularProvisiones Trim(fpText.text), fecper, fecha1, fecha2, False
        '    '-------> Actualizando cerrando periodo y abriendo proximo periodo, siempre y cuando la fecha del cierre diario se distinta al cierre del periodo mensual
        '    vg_db.Execute "UPDATE a_param SET par_valor='" & fg_Encripta(LimpiaDato(CDate(vg_ciedia) + 1)) & "' WHERE par_codigo='ciediario' AND par_cencos='" & MuestraCasino(1) & "' AND fg_Desencripta(TipoDato(par_valor, ""))<cdate('" & fg_Ctod1(fecha2) & "')"
            vg_db.CommitTrans
            
            '-------> Traer periodo
            Partida.StatusBar1.Panels(7).text = "Periodo : "
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
            If Not RS.EOF Then
   
                Partida.StatusBar1.Panels(7).text = "Periodo : " & Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)

            End If
            RS.Close
            Set RS = Nothing
            
            .Row = iRow + 1
            .Lock = False
            .Col = 4
            .text = "Abierto"
            
            Label1.Visible = False
            Frame1.Enabled = True
            MsgBox "Proceso cierre de periodo, finalizo correctamente", vbInformation + vbOKOnly, MsgTitulo
            '-------> Fin actualizando cerrando periodo y abriendo proximo periodo
        
        Case 2 '-------> Abrir Periodo
            
            If .MaxRows < 0 Then Exit Sub
            
            iRow = .ActiveRow
            IrowRea = iRow
            .Row = iRow '.ActiveRow
            .Col = 1
            fecper = Format(.text, "yyyymm")
            FecPerMostar = Format(.text, "mm/yyyy")
            
            .Col = 4
            If Trim(.text) <> "Cerrado" Then Exit Sub
            
            .Col = 3
            
            If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 2) Then
            
               MsgBox "Existe información posterior, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
            If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 9) Then
            
               MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
            If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 46) Then
            
               MsgBox "Existen documentos pendientes, en la salida ventas especiales. Debe cerrar las salidas ventas especiales", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            
            If MsgBox("Esta Seguro Abrir Mes", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
            
'Mod Ini 20240801                                    Toolbar1.Enabled = False
'Mod Ini 20240801                                    vaSpread1.Enabled = False
'Mod Ini 20240801                                    Lb1(5).Caption = "Perido Reabrir : " & FecPerMostar
'Mod Ini 20240801                                    Frame7.Visible = True
            
            '-------> Actualizando cerrando periodo y abriendo proximo periodo
            vg_db.BeginTrans

            .Row = iRow
            .Col = 3

            vg_db.Execute "UPDATE b_cierreperiodo SET cie_estado=1, cie_proantali =0, cie_gdpenmesali =0, cie_gdpenmesantali =0, cie_sncpenmesali =0, cie_sncpenmesantali =0, cie_proantgrl =0, cie_gdpenmesgrl =0, cie_gdpenmesantgrl =0, cie_sncpenmesgrl =0, cie_sncpenmesantgrl =0, cie_proantdes =0, cie_gdpenmesdes =0, cie_gdpenmesantdes =0, cie_sncpenmesdes =0, cie_sncpenmesantdes =0 WHERE cie_cencos='" & Trim(fpText.text) & "' AND cie_fecter=" & Val(Format(.text, "yyyymmdd")) & ""

            .Row = iRow
            .Lock = False

            .Col = 4
            .text = "Abierto"
            .Row = iRow + 1 '.ActiveRow + 1

            .Col = 3

            vg_db.Execute "UPDATE b_cierreperiodo SET cie_estado=2 WHERE cie_cencos='" & Trim(fpText.text) & "' AND cie_fecter=" & Val(Format(.text, "yyyymmdd")) & ""

            '-------> Actualizar solicitud nota de credito
            vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = '' WHERE toc_tipinf = 'C' AND toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND toc_codbod = " & vg_codbod & " " & _
                          "AND toc_docsnc = '" & "XXX" & fecper & "'"

            vg_db.CommitTrans

            '-------> Traer periodo
            Partida.StatusBar1.Panels(7).text = "Periodo : "
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
            If Not RS.EOF Then

                Partida.StatusBar1.Panels(7).text = "Periodo : " & Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)

            End If
            RS.Close
            Set RS = Nothing

            .Row = iRow + 1
            .Lock = True
            .Col = 4
            .text = "Inhabilitado"
            '-------> Fin actualizando cerrando periodo y abriendo proximo periodo
        
        Case 4
            
            LlenarDatos
        
        Case 6
            
            If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
            LlenarDatos
            modo = ""
            Gl_Ac_Botones Me, 10, 5, modo
        
        Case 8
            
            Dim cieper As Long, fecini As Long, fecter As Long
            fg_carga ""
            
            vg_db.BeginTrans
            For i = 1 To .MaxRows
                
                .Row = i
                .Col = 4
                
                If Trim(.text) <> "Cerrado" Then
                   
                   If Trim(.text) = "Abierto" Then .Col = 3: If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 3) Then vg_db.RollbackTrans: fg_descarga: MsgBox "Existe información para esa fecha, no podra modificarse hasta el proximo periodo", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
                   
                   .Col = 1
                   cieper = 0
                   cieper = Format(.text, "yyyymm")
                   
                   .Col = 2
                   fecini = 0
                   fecini = Format(.text, "yyyymmdd")
                   
                   .Col = 3
                   fecter = 0
                   fecter = Format(.text, "yyyymmdd")
                   
                   vg_db.Execute "UPDATE b_cierreperiodo SET cie_fecini=" & fecini & ", cie_fecter=" & fecter & " WHERE cie_cencos='" & fpText.text & "' AND cie_periodo=" & cieper & ""
                
                End If
            
            Next i
            vg_db.CommitTrans
            
            modo = ""
            Gl_Ac_Botones Me, 10, 5, modo
            fg_descarga
        
        Case 11
            
            If .MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
            I_CierrePeriodo Trim(fpText.text)
        
        Case 14
            
            Me.Hide
            Unload Me
    
    End Select

End With

Exit Sub
Man_Error:

Label1.Visible = False
Frame1.Enabled = True

If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
Label1.Visible = False
Frame1.Enabled = True

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Sub LlenarDatos()

On Error GoTo Man_Error

With vaSpread1
    
    .Visible = False: .MaxRows = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM b_cierreperiodo WHERE  cie_cencos='" & Trim(fpText.text) & "' ORDER BY cie_periodo", vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
          
          .Col = 1
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignCenter
          .text = Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)
          
          .Col = 2
          .CellType = CellTypeStaticText 'CellTypeDate
    '      .TypeDateCentury = False
          .TypeHAlign = TypeHAlignCenter
    '      .TypeSpin = False
    '      .TypeDateFormat = TypeDateFormatDDMMYY
    '      .TypeDateMax = Mid(RS!cie_fecter, 5, 2) & Mid(RS!cie_fecter, 7, 2) & Mid(RS!cie_fecter, 1, 4)
    '      .TypeDateMin = Mid(RS!cie_fecini, 5, 2) & Mid(RS!cie_fecini, 7, 2) & Mid(RS!cie_fecini, 1, 4)
          .text = Mid(RS!cie_fecini, 7, 2) & "/" & Mid(RS!cie_fecini, 5, 2) & "/" & Mid(RS!cie_fecini, 1, 4)
    '      .TypeDateCentury = True
          .Lock = True
          
          .Col = 3
          .CellType = CellTypeDate
          .TypeDateCentury = False
          .TypeHAlign = TypeHAlignCenter
          .TypeSpin = False
          .TypeDateFormat = TypeDateFormatDDMMYY
          .TypeDateMin = "01011973":  .TypeDateMax = "31125000"
    '      .TypeDateMax = IIf(RS!cie_estado <> 1, Mid(RS!cie_fecter, 5, 2) & Mid(RS!cie_fecter, 7, 2) & Mid(RS!cie_fecter, 1, 4), Format(dEoM(Mid(RS!cie_fecter, 7, 2) & "/" & Mid(RS!cie_fecter, 5, 2) & "/" & Mid(RS!cie_fecter, 1, 4)), "mmddyyyy"))
          .TypeDateMax = Format(dEoM(Mid(RS!cie_fecter, 7, 2) & "/" & Mid(RS!cie_fecter, 5, 2) & "/" & Mid(RS!cie_fecter, 1, 4)), "mmddyyyy")
          .TypeDateMin = Mid(RS!cie_fecini, 5, 2) & fg_pone_cero(Str(Val(Mid(RS!cie_fecini, 7, 2)) + 1), 2) & Mid(RS!cie_fecini, 1, 4)
          .text = Mid(RS!cie_fecter, 7, 2) & "/" & Mid(RS!cie_fecter, 5, 2) & "/" & Mid(RS!cie_fecter, 3, 2)
          .TypeDateCentury = True
          .Lock = IIf(RS!cie_estado = 0 Or RS!cie_estado = 2, True, False)
          
          .Col = 4
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignCenter
          
          If RS!cie_estado = 0 Then
             
             .text = "Cerrado"
          
          ElseIf RS!cie_estado = 1 Then
             
             .text = "Abierto"
          
          Else
             
             .text = "Inhabilitado"
          
          End If
          RS.MoveNext
       
       Loop
    
    End If
    
    RS.Close: Set RS = Nothing
    .Visible = True
    
End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

With vaSpread1
    .Row = Row: .Col = Col
    If ChangeMade = False Then Fecha = .text
    
    If ChangeMade = True Then
       
       If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 5) Then
       
          MsgBox "Existe un cierre de mes, no podra modificarse hasta el proximo periodo", vbExclamation + vbOKOnly, MsgTitulo
          .text = Fecha
          Exit Sub
       
       End If
       
       If CierrePeriodo(Format(.text, "yyyymmdd"), vg_codbod, 3) Then
       
          MsgBox "Existe información para esa fecha, no podra modificarse hasta el proximo periodo", vbExclamation + vbOKOnly, MsgTitulo
          .text = Fecha
          Exit Sub
       
       End If
       
       Dim i As Long, fecini As Long, fecfin As Long, diatop As Long, lEOM As Boolean
       fecini = 0: fecfin = 0: diatop = 0
       For i = Row To .MaxRows
           
           .Row = i
           
           If i = Row Then
              
              .Col = 3
    '          .TypeDateMax = Mid(.Text, 4, 2) & Mid(.Text, 1, 2) & Mid(.Text, 7, 4)
              fecini = Mid(.text, 7, 4) & Mid(.text, 4, 2) & fg_pone_cero(Str(Val(Mid(.text, 1, 2))), 2)
              fecfin = Mid(.text, 7, 4) & Mid(.text, 4, 2) & Mid(.text, 1, 2)
              lEOM = IIf(fg_Ctod(fecfin) = dEoM(fg_Ctod(fecfin)), True, False)
              diatop = Val(Mid(fecfin, 7, 2))
           
           Else
    
    '          If (fecfin + 1) > Format(dEoM(fg_Ctod(fecfin)), "yyyymmdd") Or lEOM Then
    '             fecini = Format(dBoM(bEOM(fg_Ctod(fecfin))), "yyyymmdd")
    '          Else
    '             fecini = Format(dEoM(fg_Ctod(fecfin)), "yyyymm") & fg_pone_cero(Str(Val(diatop + 1)), 2) 'fg_pone_cero(Str(Val(Mid(fecini, 7, 2))), 2)
    '          End If
                
                If lEOM Then
                    
                    fecfin = Format(dEoM(fg_Ctod(fecini)), "yyyymmdd")
                
                Else
                    
                    If fecini > Format(dEoM(fg_Ctod(fecini)), "yyyymm") & fg_pone_cero(Str(Val(diatop)), 2) Then
                        
                        fecfin = Format(dEoM(fg_Ctod(fecini)) + 1, "yyyymm") & fg_pone_cero(Str(Val(diatop)), 2)
                        If fecfin > Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then fecfin = Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd")
                    
                    Else
                        
                        fecfin = Format(dEoM(fg_Ctod(fecini)), "yyyymm") & fg_pone_cero(Str(Val(diatop)), 2)
                    
                    End If
                
                End If
    
    '          If fecfin + 1 > Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
    '             fecfin = Format(dEoM("01/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd")
    '          End If
              
              .Col = 3
              .TypeDateCentury = False
              .TypeHAlign = TypeHAlignCenter
              .TypeSpin = False
              .TypeDateFormat = TypeDateFormatDDMMYY
              a = "01011973": b = "31125000"
              .TypeDateMin = a: .TypeDateMax = b
              .TypeDateMin = Mid(fecini, 5, 2) & fg_pone_cero(Str(Val(Mid(fecini, 7, 2) + 1)), 2) & Mid(fecini, 1, 4)
              .TypeDateMax = Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "mmddyyyy")
              .text = Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 3, 2)
              .TypeDateCentury = True
    '          .Text = IIf(Mid(fecfin, 7, 2) > 27, Mid(fecini, 7, 2), fg_pone_cero(Str(Val(Mid(fecini, 7, 2) + 1)), 2)) & "/" & Mid(fecini, 5, 2) & "/"
              .Col = 2
              If (fecfin + 1) > Format(dEoM(Mid(fecfin, 7, 2) & "/" & Mid(fecfin, 5, 2) & "/" & Mid(fecfin, 1, 4)), "yyyymmdd") Then
                 
                 .text = IIf(Mid(fecfin, 7, 2) > 27, Mid(fecini, 7, 2), fg_pone_cero(Str(Val(diatop)), 2)) & "/" & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
              
              Else
                 
                 .text = IIf(Mid(fecfin, 7, 2) > 27, Mid(fecini, 7, 2), fg_pone_cero(Str(Val(diatop + 1)), 2)) & "/" & Mid(fecini, 5, 2) & "/" & Mid(fecini, 1, 4)
              
              End If
           
           End If
           fecini = Format(CDate(fg_Ctod(fecfin)) + 1, "yyyymmdd")
       
       Next i
       If modo = "" Then modo = "M"
       Gl_Ac_Botones Me, 10, 0, modo
    
    End If

End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub
