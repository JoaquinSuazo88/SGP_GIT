VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_ExpDetMinBloque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Detalle Minuta Bloque"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16170
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   16170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD 
      Left            =   1320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
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
      Height          =   2145
      Left            =   4065
      TabIndex        =   15
      Top             =   105
      Width           =   8010
      Begin VB.Frame Frame3 
         Height          =   615
         Left            =   1200
         TabIndex        =   22
         Top             =   1440
         Width           =   4335
         Begin VB.OptionButton Option2 
            Caption         =   "Sin Total Día"
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
            Left            =   2400
            TabIndex        =   24
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Con Total Día"
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
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Resumido"
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
         Left            =   3720
         TabIndex        =   21
         Top             =   1200
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado"
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
         Index           =   0
         Left            =   2040
         TabIndex        =   20
         Top             =   1200
         Value           =   -1  'True
         Width           =   1335
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1965
         TabIndex        =   1
         Top             =   750
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
         Left            =   5730
         TabIndex        =   2
         Top             =   750
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
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   7350
         TabIndex        =   16
         Top             =   750
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Left            =   1965
         TabIndex        =   0
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
         Left            =   630
         TabIndex        =   19
         Top             =   795
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
         Left            =   4425
         TabIndex        =   18
         Top             =   795
         Width           =   1065
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
         Left            =   690
         TabIndex        =   17
         Top             =   390
         Width           =   1155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Detalle Minuta Bloque"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4950
      Left            =   105
      TabIndex        =   9
      Top             =   2325
      Width           =   15975
      Begin VB.CommandButton Command1 
         Caption         =   "Ejecutar"
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
         Left            =   12915
         TabIndex        =   7
         Top             =   4200
         Width           =   1275
      End
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
         Left            =   14280
         TabIndex        =   8
         Top             =   4200
         Width           =   1275
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   1785
         TabIndex        =   12
         Top             =   3990
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Left            =   6300
         TabIndex        =   11
         Top             =   3990
         Width           =   3030
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   2925
         End
      End
      Begin VB.Frame Frame7 
         Height          =   435
         Left            =   9345
         TabIndex        =   10
         Top             =   3990
         Width           =   3030
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   2925
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3615
         Left            =   315
         TabIndex        =   3
         Top             =   315
         Width           =   15510
         _Version        =   393216
         _ExtentX        =   27358
         _ExtentY        =   6376
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
         MaxCols         =   11
         SpreadDesigner  =   "I_ExpDetMinBloque.frx":0000
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   3780
         TabIndex        =   13
         Top             =   4620
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lbl_proceso 
         Alignment       =   2  'Center
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   9660
         TabIndex        =   14
         Top             =   4515
         Visible         =   0   'False
         Width           =   435
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
            Picture         =   "I_ExpDetMinBloque.frx":1B6F
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "I_ExpDetMinBloque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim seleccion       As String
Dim Fecha           As Date
Dim i               As Long
Dim j               As Long
Dim Ceco            As String
Dim Regimen         As Long
Dim Servicio        As Long
Dim FecIni          As String
Dim FecFin          As String
Dim Bloque          As String
Dim Conta           As Long
Dim Sql             As String
Dim EstCopiado      As Boolean
Dim LargoDia        As Long
Dim FechaDesFin     As Date
Dim Id_Bloque       As Long
Dim XmlMinBlo       As String
Dim NomArchivoExcel As String
Dim Extension       As String

Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
    
  If vaSpread2.MaxRows > 1 Then
     vaSpread2.Row = 1
     vaSpread2.Col = 8
     Fecha = vaSpread2.text
  End If
  
  If vaSpread2.MaxRows < 1 Then
     MsgBox "Debe seleccionar datos del encabezado...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  End If
  
  '-------> Validar que exista un dato seleccionado
  seleccion = 0
  For i = 1 To vaSpread2.MaxRows
       
       vaSpread2.Row = i
       vaSpread2.Col = 1 'Seleccion
       seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
       If seleccion = 1 And vaSpread2.RowHidden = False Then
          Exit For
       End If
  
  Next i
  
  If seleccion = 0 Then
     
     MsgBox " Se debe seleccionar un Bloque por lo menos", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  '-------> Rescata Ceco Seleccionado
  seleccion = 0
  XmlMinBlo = ""
  XmlMinBlo = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
  XmlMinBlo = xmlceco & "<MinBlo>"

  j = 0
  For i = 1 To vaSpread2.MaxRows
  
      vaSpread2.Row = i
      vaSpread2.Col = 1 'Seleccion
      seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
      If seleccion = 1 And vaSpread2.RowHidden = False Then
          
          Ceco = ""
          vaSpread2.Col = 2
          Ceco = vaSpread2.text
          
          Id_Bloque = 0
          vaSpread2.Col = 4
          Id_Bloque = vaSpread2.text
          
         
          XmlMinBlo = XmlMinBlo & "<C"

          XmlMinBlo = XmlMinBlo & " c = " & Chr(34) & Ceco & Chr(34)
          XmlMinBlo = XmlMinBlo & " b = " & Chr(34) & Id_Bloque & Chr(34)
          XmlMinBlo = XmlMinBlo & "/>"
          
          vaSpread2.Row = i: vaSpread2.Col = -1
          vaSpread2.BackColor = &HC0FFFF
       
      End If
  
      DoEvents
       
  Next i

  XmlMinBlo = XmlMinBlo & "</MinBlo>"
  
  '-------> Validar cantidad registro se sobre pase hoja excel
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  
  Sql = ""
  Sql = Sql + XmlMinBlo
  Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarNRegMinutaBloque '" & Sql & "'")
  If Not RS.EOF Then
  
     If RS!nReg > 1020000 Then
        MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Ceco", vbCritical
        Exit Sub
     End If
  
  End If
  
  '-------> Guardar nombre archivo excel
  NomArchivoExcel = ""
  CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
  CD.Filter = "Todos los archivos *.xls,*.xlsx"
  On Error Resume Next
  CD.ShowSave
           
 '-------> JPAZ Permite controlar Boton Cancelar
 If Err.Number = 32755 Then
    MsgBox "Proceso cancelado"
    Exit Sub
 End If
            
 If CD.FileName = "" Then
    MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
    Exit Sub
 Else
    Extension = ""
    Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
    If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
       MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
       Exit Sub
    End If
    NomArchivoExcel = CD.FileName
 End If
          
'  ProgressBar1.Scrolling = ccScrollingSmooth
'  ProgressBar1.Max = 100
'  ProgressBar1.Visible = True
'  ProgressBar1.Value = 0
'  lbl_proceso.Caption = "0 %"
'  lbl_proceso.Visible = True
  
  Toolbar2.Enabled = False
  FpFecDesde.Enabled = False
  FpFecHasta.Enabled = False

  fg_carga ""
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  
  Sql = ""
  Sql = Sql + XmlMinBlo
  Set RS = vg_db.Execute("sgpadm_Sel_XmlExportarDetMinutaBloque_V04 '" & Sql & "', 1, '" & Trim(LimpiaDato(fpText.text)) & "', '" & IIf(Option1(0).Value = True, "D", "R") & "', '" & IIf(Option2(0).Value = True, "1", "0") & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  
  '-------> Create an instance of Excel and add a workbook
  Set xlApp = CreateObject("Excel.Application")
  Set xlWb = xlApp.Workbooks.Add
  Set xlWs = xlWb.Worksheets("Hoja1")
  
  '-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
  xlApp.UserControl = True
    
  '-------> Check version of Excel
  Call encabezado(RS, xlWs)
          
  xlWs.Cells(2, 1).CopyFromRecordset RS
  '-------> Auto-fit the column widths and row heights
  xlApp.Selection.CurrentRegion.Columns.AutoFit
  xlApp.Selection.CurrentRegion.Rows.AutoFit
    
  xlApp.Columns("A:A").Select
  xlApp.Selection.Delete Shift:=xlToLeft
  
  xlWb.Close True, NomArchivoExcel

  Dim XL As New excel.Application 'Crea el objeto excel
  XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
  XL.Visible = True
  XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
  '-------> Close ADO objects
  RS.Close
  Set RS = Nothing
    
  ' -- Cerrar Excel
  xlApp.Quit
  '-------> Release Excel references
  Set xlWs = Nothing
  Set xlWb = Nothing
  Set xlApp = Nothing
  
  fg_descarga
  MsgBox "Proceso Finalizado", vbInformation, Me.Caption
                
  ProgressBar1.Visible = False
  lbl_proceso.Visible = False
  
  Toolbar2.Enabled = True
  FpFecDesde.Enabled = True
  FpFecHasta.Enabled = True
                
Exit Sub
Man_Error:
    fg_descarga
    
    ProgressBar1.Visible = False
    lbl_proceso.Visible = False
    
    Toolbar2.Enabled = True
    FpFecDesde.Enabled = True
    FpFecHasta.Enabled = True
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub encabezado(ByRef RS As ADODB.Recordset, ByRef xlWs As Object)
On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = RS.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = RS.Fields(icol - 1).Name
Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo
End Sub

Private Sub Command2_Click()
On Error GoTo Man_Error

'-------> Salir de la opción
Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Copiar Minuta Bloque Ceco"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread2.MaxRows = 0

TextDet1(2).text = ""
TextDet1(5).text = ""
TextDet1(6).text = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecDesde_Change()
On Error GoTo Man_Error

vaSpread2.MaxRows = 0
If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecHasta_Change()
On Error GoTo Man_Error

vaSpread2.MaxRows = 0
If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
    Frame3.Visible = True
Case 1
    Frame3.Visible = False
End Select
End Sub

Private Sub TextDet1_Change(Index As Integer)

On Error GoTo Man_Error

If Index = 2 Then
   TextDet1(5).text = ""
   TextDet1(6).text = ""
ElseIf Index = 5 Then
   TextDet1(2).text = ""
   TextDet1(6).text = ""
ElseIf Index = 6 Then
   TextDet1(2).text = ""
   TextDet1(5).text = ""
End If
Select Case Index
Case 2, 5, 6
    vaSpread2.Visible = False
    If Trim(TextDet1(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index
           indactivo = UCase(Trim(vaSpread2.Value)) Like "*" & UCase(TextDet1(Index).text) & "*"
           vaSpread2.Col = 1
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index + 1, 1
    End If
'    vaSpread1_Click Index, 0
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(TextDet1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(TextDet1(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(TextDet1(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
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

  '-------> Validar org. compras
  If Trim(fpText.text) = "" Then
     MsgBox "Debe ingresar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  End If
  
  '-------> Validar fechas
  If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
     MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  End If
    
  If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
     MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  End If

  vaSpread2.MaxRows = 0
  vaSpread2.Row = -1: vaSpread2.Col = -1
  vaSpread2.BackColor = &HC0FFFF
   
  TextDet1(2).text = ""
  TextDet1(5).text = ""
  TextDet1(6).text = ""
   
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
   
  Sql = ""
  Sql = Sql & Trim(LimpiaDato(fpText.text))
  
  Set RS = vg_db.Execute("sgpadm_Sel_OrgComprasCecoMinutaBloque '" & Sql & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  If Not RS.EOF Then
  Do While Not RS.EOF
      
      vaSpread2.MaxRows = vaSpread2.MaxRows + 1
      vaSpread2.Row = vaSpread2.MaxRows
      
      vaSpread2.Col = 1
      vaSpread2.text = "0"
      
      vaSpread2.Col = 2
      vaSpread2.text = RS!Ceco
      
      vaSpread2.Col = 3
      vaSpread2.text = Trim(RS!Cli_nombre)
      
      vaSpread2.Col = 4
      vaSpread2.text = RS!Id_Bloque
      
      vaSpread2.Col = 5
      vaSpread2.text = RS!Regimen & " - " & Trim(RS!reg_nombre)
      
      vaSpread2.Col = 6
      vaSpread2.text = RS!Servicio & " - " & Trim(RS!ser_nombre)
      
      vaSpread2.Col = 7
      vaSpread2.text = RS!fechadesde
         
      vaSpread2.Col = 8
      vaSpread2.text = RS!fechahasta
         
      vaSpread2.Col = 9
      vaSpread2.text = RS!Regimen
         
      vaSpread2.Col = 10
      vaSpread2.text = RS!Servicio
         
      RS.MoveNext
  Loop
  Else
     vaSpread2.MaxRows = 0
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

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim i As Long

If Col = 1 And Row = 0 Then
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        If i <> Row Then vaSpread2.text = IIf(vaSpread2.text = "0", "1", "0")
    Next i
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

