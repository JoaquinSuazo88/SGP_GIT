VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form C_ConsultarActualizarConveniosPel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar & Actualizar Convenios con problema de envio"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10995
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
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
      Height          =   540
      Left            =   9525
      TabIndex        =   11
      Top             =   8160
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   8160
      TabIndex        =   10
      Top             =   8160
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   10695
      Begin VB.Frame Frame4 
         Caption         =   "Estado Convenios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   8040
         TabIndex        =   20
         Top             =   4680
         Width           =   2535
         Begin VB.Label Label1 
            Caption         =   "Envio Parcial"
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
            Index           =   2
            Left            =   795
            TabIndex        =   23
            Top             =   780
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   0
            Left            =   240
            Picture         =   "C_ConsultarActualizarConveniosPel.frx":0000
            Stretch         =   -1  'True
            Top             =   720
            Width           =   360
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   1
            Left            =   240
            Picture         =   "C_ConsultarActualizarConveniosPel.frx":628A
            Stretch         =   -1  'True
            Top             =   1200
            Width           =   360
         End
         Begin VB.Label Label1 
            Caption         =   "Envio Total"
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
            Index           =   3
            Left            =   795
            TabIndex        =   22
            Top             =   1245
            Width           =   1095
         End
         Begin VB.Image Image1 
            Height          =   360
            Index           =   2
            Left            =   240
            Picture         =   "C_ConsultarActualizarConveniosPel.frx":C514
            Stretch         =   -1  'True
            Top             =   300
            Width           =   360
         End
         Begin VB.Label Label1 
            Caption         =   "Envio Error"
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
            Index           =   4
            Left            =   840
            TabIndex        =   21
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   4440
         TabIndex        =   8
         Top             =   5880
         Width           =   2100
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   1995
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   2640
         TabIndex        =   6
         Top             =   5880
         Width           =   1740
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   1635
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5415
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   6735
         _Version        =   393216
         _ExtentX        =   11880
         _ExtentY        =   9551
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
         SpreadDesigner  =   "C_ConsultarActualizarConveniosPel.frx":1279E
      End
      Begin MSComDlg.CommonDialog CD 
         Left            =   0
         Top             =   5520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   135
         Left            =   360
         TabIndex        =   24
         Top             =   6120
         Visible         =   0   'False
         Width           =   1455
         _Version        =   393216
         _ExtentX        =   2566
         _ExtentY        =   238
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
         MaxRows         =   1
         SpreadDesigner  =   "C_ConsultarActualizarConveniosPel.frx":141B2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   3495
         TabIndex        =   3
         Top             =   600
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   9960
         TabIndex        =   5
         Top             =   600
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
      Begin EditLib.fpDateTime FpFecHasta 
         Height          =   315
         Left            =   7335
         TabIndex        =   18
         Top             =   600
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         Text            =   "13/07/2004"
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Fin Validez"
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
         Left            =   5520
         TabIndex        =   19
         Top             =   675
         Width           =   1530
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicio Validez"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   675
         Width           =   1740
      End
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
            Picture         =   "C_ConsultarActualizarConveniosPel.frx":168DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ": "
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
      Left            =   6960
      TabIndex        =   17
      Top             =   8160
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ": "
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
      Left            =   3600
      TabIndex        =   16
      Top             =   8520
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   ": "
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
      Index           =   4
      Left            =   3600
      TabIndex        =   15
      Top             =   8160
      Width           =   135
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conv. con error (X)"
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
      Left            =   4920
      TabIndex        =   14
      Top             =   8160
      Width           =   1635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conv. en proceso de integración (P)"
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
      Left            =   120
      TabIndex        =   13
      Top             =   8520
      Width           =   3075
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Conv. pendiente de integrar (E)"
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
      TabIndex        =   12
      Top             =   8160
      Width           =   2685
   End
End
Attribute VB_Name = "C_ConsultarActualizarConveniosPel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim MsgTitulo As String
Public lc_Aux As String
Dim Est As Boolean

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim i          As Long
Dim ISeleccion As Boolean
Dim MyBuffer   As String
Dim RS         As New ADODB.Recordset
Dim IdLote     As Double
Dim estado     As String

ISeleccion = False

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 6
    estado = vaSpread1.text
    
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False And estado = "X" Then
    
       ISeleccion = True
       Exit For
       
    End If

Next i

If Not ISeleccion Then

   MsgBox "Debe haber por lo menos un ítem seleccionado de la lista, para actualizar...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

If MsgBox("Esta seguro realizar reenvio del lote a la PEL...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdatePel>"

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 6
    estado = vaSpread1.text
    
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False And estado = "X" Then
    
       vaSpread1.Col = 2
       IdLote = vaSpread1.text
       
       MyBuffer = MyBuffer & " <Pel"
       MyBuffer = MyBuffer & " IdLote = " & Chr(34) & IdLote & Chr(34)
       MyBuffer = MyBuffer & "/>"

    End If
    
Next i

MyBuffer = MyBuffer & "</UpdatePel>"
      
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Upd_XmlConveniosPel_V02 '" & MyBuffer & "'")

If Not RS.EOF Then

   If RS(0) > 0 Or RS(0) < 0 Then
        
     fg_descarga
      
     Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Me.HelpContextID, "", "", "")

     MsgBox RS(1) & VgLinea, vbCritical, MsgTitulo
          
     RS.Close
     Set RS = Nothing
                 
     Exit Sub
              
   Else
        
      If RS(2) > 0 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")

         MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
      
      Else
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_No_Encontraron_Datos_Actualizar"), Me.HelpContextID, "", "", "")
         
         MsgBox "Proceso finalizado, no se encontraron datos que actualizar...", vbInformation + vbOKOnly, Me.Caption
         
      End If
              
   End If

End If

RS.Close
Set RS = Nothing

fg_descarga

Toolbar2_ButtonClick Toolbar2.Buttons(1)

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

Me.Hide
Unload Me

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me

MsgTitulo = "Consultar & Actualizar Convenios con problema de envio"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0
Est = True
Me.HelpContextID = vg_OpcM

Command1.Enabled = False

If Mid(ValidarUsuarioAcceso(Me.HelpContextID, vg_NUsr), 3, 1) = "1" Then

   Command1.Enabled = True

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TextDet1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet1(Index).text, ",")

If Index = 2 Then
   
   TextDet1(3).text = ""

ElseIf Index = 3 Then
   
   TextDet1(2).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 7
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3
    
    vaSpread1.Visible = False
    
    If Trim(TextDet1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 Or Index = 3, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 7
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 7
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 7
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 7
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 7
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
    
    If Trim(TextDet1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 7
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim TotalConvPendienteIntegracion As Double
Dim TotalConvProcesoIntegracion As Double
Dim RotalConvError As Double

Select Case Button.Index

Case 1
  
  fg_carga ""
  
  vaSpread1.MaxRows = 0
    
  
  TotalConvPendienteIntegracion = 0
  TotalConvProcesoIntegracion = 0
  RotalConvError = 0
  Est = True

  Label2(4).Caption = " : " & Format(TotalConvPendienteIntegracion, fg_Pict(6, 0))
  Label2(5).Caption = " : " & Format(TotalConvProcesoIntegracion, fg_Pict(6, 0))
  Label2(6).Caption = " : " & Format(RotalConvError, fg_Pict(6, 0))

  If Not ValidarDatos Then Exit Sub
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
    
  Set RS = vg_db.Execute("sgpadm_Sel_ConsultarConveniosPel '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  If Not RS.EOF Then
     
     Do While Not RS.EOF
      
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
      
        vaSpread1.Col = 1
        vaSpread1.text = "0"
      
        vaSpread1.Col = 2
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(1)
      
        vaSpread1.Col = 3
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = Trim(RS(0))
      
        
        vaSpread2.Row = 1
        vaSpread2.Col = IIf(IsNull(RS(2)) Or RS(2) = "C", 2, IIf(RS(2) = "X", 3, 1))
        
        vaSpread1.Col = 4
        vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
        
        vaSpread1.Col = 5
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(3)
      
        vaSpread1.Col = 6
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = RS(2)
      
        vaSpread1.Col = 7
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.text = 0
      
        Select Case RS(2)
        
            Case "E"
          
                TotalConvPendienteIntegracion = TotalConvPendienteIntegracion + RS(3)
        
            Case "P"
           
                TotalConvProcesoIntegracion = TotalConvProcesoIntegracion + RS(3)
        
            Case "X"
                
                RotalConvError = RotalConvError + RS(3)
        
        End Select
        
        RS.MoveNext
        
     Loop
     
  Else
     
     vaSpread1.MaxRows = 0
     MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
  
  End If
  RS.Close
  Set RS = Nothing
  
  fg_descarga

  Label2(4).Caption = " : " & Format(TotalConvPendienteIntegracion, fg_Pict(6, 0))
  Label2(5).Caption = " : " & Format(TotalConvProcesoIntegracion, fg_Pict(6, 0))
  Label2(6).Caption = " : " & Format(RotalConvError, fg_Pict(6, 0))
  
  Est = False

End Select


Exit Sub
Man_Error:
fg_descarga
Est = True
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarDatos() As Boolean

Dim seleccion As Integer
Dim i As Long

ValidarDatos = True

'-------> Validar fechas
If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
   
   MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If
    
If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
   
   MsgBox "La fecha hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

If DateDiff("d", Format(FpFecDesde.text, "dd/mm/yyyy"), Format(FpFecHasta.text, "dd/mm/yyyy")) + 1 > 7 Then
      
   MsgBox "Ha sobrepasado los 7 días de la semana", vbExclamation + vbOKOnly, MsgTitulo
   ValidarDatos = False
   Exit Function

End If

End Function

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Dim estado As String

Est = True

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
       
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        vaSpread1.Col = 6
        estado = vaSpread1.text
        
        vaSpread1.Col = 1
        
        If vaSpread1.RowHidden = False And estado = "X" Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        vaSpread1.Col = 6
        estado = vaSpread1.text
        
        vaSpread1.Col = 1
        
        If vaSpread1.RowHidden = False And estado = "X" Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Est = False

Exit Sub
Man_Error:
    Est = False
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

Dim EstSel As Boolean

If Est Or ButtonDown = 0 Or vaSpread1.MaxRows < 1 Then Exit Sub

Dim i As Long
Dim estado As String

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    
    vaSpread1.Col = 6
    estado = vaSpread1.text
    
    vaSpread1.Col = 1
    
'    If i <> Row Then
       
       If vaSpread1.text = "1" And estado <> "X" Then
          
          Est = True
          vaSpread1.text = "0"
          Est = False
        
        End If
    
'    End If

Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim estado As String
Dim IdLote As Double
Dim i      As Long

Select Case Col

    Case Is <> 1

        For i = 1 To vaSpread1.maxcols
        
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" Then
            
               MsgBox "Para seleccionar el detalle convenio, no debe haber seleccionado ningun convenio...", vbExclamation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
        
        Next i

        vaSpread1.Row = Row
        
        vaSpread1.Col = 6
        estado = vaSpread1.text
        
        If estado <> "X" Then
        
           MsgBox "Debe seleccionar un convenio con error...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If

        vaSpread1.Col = 2
        IdLote = vaSpread1.text
        
        '--> Validar Acceso
        If Mid(ValidarUsuarioAcceso(1199200, vg_NUsr), 1, 1) <> "1" Then
       
           MsgBox "No tiene acceso detalle Convenio...", vbInformation, MsgTitulo
           Exit Sub
       
        End If
        
        vg_codigo = ""
        Call C_DetalleConveniosConErrorPel.LlenarDatos(IdLote)
        Call C_DetalleConveniosConErrorPel.Show(1)

        If vg_codigo <> "" Then
        
           Toolbar2_ButtonClick Toolbar2.Buttons(1)
        
        End If
        
        vg_codigo = ""
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
