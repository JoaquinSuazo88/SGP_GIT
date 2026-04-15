VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form P_GenInfOpt 
   BackColor       =   &H80000009&
   Caption         =   "Generar Datos Optimum"
   ClientHeight    =   8385
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10710
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8385
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   375
      Left            =   7080
      TabIndex        =   17
      Top             =   8040
      Visible         =   0   'False
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"P_GenInfOpt.frx":0000
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
      Height          =   375
      Index           =   1
      Left            =   9240
      TabIndex        =   2
      Top             =   7920
      Width           =   1335
   End
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
      Height          =   375
      Index           =   0
      Left            =   7680
      TabIndex        =   1
      Top             =   7920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   375
         Left            =   6000
         TabIndex        =   20
         Top             =   2640
         Width           =   735
         _Version        =   196608
         _ExtentX        =   1296
         _ExtentY        =   661
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "0"
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483648"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   9000
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox ChTGramaje 
         Caption         =   "Incorpora Tabla Gramaje"
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
         Left            =   7560
         TabIndex        =   10
         Top             =   2760
         Width           =   2535
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3495
         Left            =   360
         TabIndex        =   9
         Top             =   3120
         Width           =   9735
         _Version        =   393216
         _ExtentX        =   17171
         _ExtentY        =   6165
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
         MaxCols         =   6
         SpreadDesigner  =   "P_GenInfOpt.frx":0082
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2175
         Left            =   360
         TabIndex        =   4
         Top             =   600
         Width           =   2775
         _Version        =   393216
         _ExtentX        =   4895
         _ExtentY        =   3836
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
         MaxCols         =   2
         SpreadDesigner  =   "P_GenInfOpt.frx":1A7B
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1395
         TabIndex        =   5
         Top             =   6960
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
         Left            =   8460
         TabIndex        =   6
         Top             =   6960
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Extracion"
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
         Left            =   4680
         TabIndex        =   19
         Top             =   2760
         Width           =   1065
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   0
         Top             =   0
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000013&
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
         Index           =   6
         Left            =   8160
         TabIndex        =   16
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   8160
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicio Proceso : "
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
         Left            =   6120
         TabIndex        =   14
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Fin    Proceso : "
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
         Left            =   6120
         TabIndex        =   13
         Top             =   840
         Width           =   2055
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
         Index           =   2
         Left            =   7155
         TabIndex        =   8
         Top             =   7050
         Width           =   1065
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
         Left            =   120
         TabIndex        =   7
         Top             =   7050
         Width           =   1110
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   210
      Left            =   120
      TabIndex        =   11
      Top             =   8040
      Visible         =   0   'False
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   370
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
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
      Left            =   240
      TabIndex        =   12
      Top             =   7680
      Visible         =   0   'False
      Width           =   6615
   End
End
Attribute VB_Name = "P_GenInfOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error
 
Dim i As Long
Dim seleccion As Integer

Select Case Index

Case 0
    '-------> Validar NºExtracion
    If Trim(fpLongInteger1.Value) = "" Or fpLongInteger1.Value = 0 Then
        MsgBox "Debe seleccionar Nº Extracion...", vbExclamation + vbOKOnly, MsgTitulo
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

    '-------> Validar que exista un dato seleccionado
    seleccion = 0
    For i = 1 To vaSpread1.MaxRows
       
        vaSpread1.Row = i
        vaSpread1.Col = 1 'Seleccion
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
        If seleccion = 1 Then
            Exit For
        End If
  
    Next i
  
    If seleccion = 0 Then
        
        MsgBox " Se debe seleccionar un Org. Compras por lo menos", vbExclamation + vbOKOnly, MsgTitulo
        Exit Sub
  
    End If

    '-------> Validar que exista un dato seleccionado
    seleccion = 0
    For i = 1 To vaSpread2.MaxRows
       
        vaSpread2.Row = i
        vaSpread2.Col = 1 'Seleccion
        seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
        If seleccion = 1 Then
            Exit For
        End If
  
    Next i
  
    If seleccion = 0 Then
        
        MsgBox " Se debe seleccionar un ceco por lo menos", vbExclamation + vbOKOnly, MsgTitulo
        Exit Sub
  
    End If

    '-------> Validar que no existan problema de asociación ceco con optimum
    seleccion = 0
    For i = 1 To vaSpread2.MaxRows
    
        vaSpread2.Row = i
        vaSpread2.Col = 1 'Seleccion
        seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
        If seleccion = 1 Then
           
           vaSpread2.Col = 5
           
           If vaSpread2.BackColor = Shape1(1).FillColor Then
              
              MsgBox " De los Ceco seleccionado tiene problema de asociación con OPTIMUM, revise la columna observación. ", vbExclamation + vbOKOnly, MsgTitulo
              
              vaSpread2.SetActiveCell 2, vaSpread2.Row: vaSpread2.SetFocus
              
              Exit Sub
           
           End If
        
        End If
      
    Next i
    
    GenerarArchivoOptimum
    
Case 1 '-------> Salir de la opción
    Unload Me

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
MsgTitulo = "Generar Datos Optimum"
fg_centra Me

GenerarOrgCompras

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Sub GenerarOrgCompras()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &HC0FFFF

vaSpread1.MaxRows = 0
vaSpread2.MaxRows = 0

Sql = ""
Sql = "sgpadm_Sel_OrgCompras_V02 "
Sql = Sql & " '' "
Set RS = vg_db.Execute(Sql)

Do While Not RS.EOF

    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows

    vaSpread1.Col = 1
    vaSpread1.text = RS(2)
    
    vaSpread1.Col = 2
    vaSpread1.text = RS(0)

    RS.MoveNext

Loop
RS.Close
Set RS = Nothing

vaSpread1.Col = -1
vaSpread1.Row = -1
vaSpread1.Lock = True

CargarCecoxOrgCompras

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Private Sub FpFecDesde_Change()

If IsDate(FpFecDesde.text) = False Then Exit Sub

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

If IsDate(FpFecHasta.text) = False Then Exit Sub

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
'On Error GoTo Man_Error
'
'Dim i As Long
'
'Select Case BlockCol
'
'Case 1
'
'    vaSpread1.Col = 1
'    For i = BlockRow To BlockRow2
'        vaSpread1.Row = i
'        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
'    Next
'
'End Select
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

'On Error GoTo Man_Error
'
'CargarCecoxOrgCompras
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Err.Description, vbCritical, Msgtitulo
'
End Sub

Sub CargarCecoxOrgCompras()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String

seleccion = 0
  
fg_carga ""

'xmlorgcom = ""
'xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
'xmlorgcom = xmlorgcom & "<Org>"
'
'For i = 1 To vaSpread1.MaxRows
'
'    vaSpread1.Row = i
'    vaSpread1.Col = 1 'Seleccion
'    seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'    If seleccion = 1 And vaSpread1.RowHidden = False Then
'
'       xmlorgcom = xmlorgcom & "<O"
'       vaSpread1.Row = i
'       vaSpread1.Col = 2 'Org. Compras
'       orgcom = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'       xmlorgcom = xmlorgcom & " Org = " & Chr(34) & orgcom & Chr(34)
'       xmlorgcom = xmlorgcom & "/>"
'
'    End If
'
'Next i
'xmlorgcom = xmlorgcom & "</Org>"
   
vaSpread2.MaxRows = 0
vaSpread2.Row = -1: vaSpread2.Col = -1
vaSpread2.BackColor = &HC0FFFF
   
Sql = ""
Sql = " sgpadm_Sel_XmlOrgComprasCeco "
'Sql = Sql & " '" & xmlorgcom & "' "
Set RS = vg_db.Execute(Sql)

Do While Not RS.EOF

    vaSpread2.MaxRows = vaSpread2.MaxRows + 1
    vaSpread2.Row = vaSpread2.MaxRows

    vaSpread2.Col = 1
    vaSpread2.text = RS(4)
    
    vaSpread2.Col = 2
    vaSpread2.text = RS(0)

    vaSpread2.Col = 3
    vaSpread2.text = RS(1)

    vaSpread2.Col = 4
    vaSpread2.text = RS(2)
    
    vaSpread2.Col = 5
    If RS(3) = "1" Then
       vaSpread2.BackColor = Shape1(1).FillColor
    End If
    vaSpread2.text = IIf(RS(3) = "0", "", "Ceco no esta asociado OPTIMUM")

    RS.MoveNext

Loop
RS.Close
Set RS = Nothing

vaSpread2.Col = -1
vaSpread2.Row = -1
vaSpread2.Lock = True

fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Sub GenerarArchivoOptimum()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

Label1(5).Caption = ""
Label1(6).Caption = ""

'--> Tabla Optimum Ingredientes
'Generar_LSRInventTable_1_1_TablaIngredientes
Me.Refresh

'--> Recetas
'Generar_LSRInventTable_1_1_Recetas
Me.Refresh
'Generar_LinkRecipeItem_Offer_Recetas
Me.Refresh
'Generar_BOMVersion_1_1_Recetas
Me.Refresh
'Generar_BOM_1_1_Recetas
Me.Refresh
'Generar_LSRBOM_1_1_Recetas
Me.Refresh
'Generar_DatoSalidaBOMVersion_1_1_Recetas
Me.Refresh
'Generar_BOMTable_1_1_Recetas
Me.Refresh
'Generar_Recipe_Item
Me.Refresh
Generar_Recipe_method
Me.Refresh

''--> Ingredientes
'Generar_LSRInventTable_1_1_Ingredientes
'Me.Refresh
'Generar_BOMTable_1_1_Ingredientes
'Me.Refresh
'Generar_BOM_1_1_Ingredientes
'Me.Refresh
'Generar_LSRBOM_1_1_Ingredientes
'Me.Refresh
'Generar_SalidaBOMTable_1_1_Ingredientes
'Me.Refresh
'Generar_BOMVersion_1_1_Ingredientes
'Me.Refresh
'Generar_Ingredient_Item
'Me.Refresh
'
''--> BomReceta & Ingredientes
'Generar_LSRBOMversion
'Me.Refresh
'
''--> Aportes nutricionales
'Generar_CategoryNutrition
'Me.Refresh
'Generar_GroupNutrition
'Me.Refresh
'Generar_Nutrition_IngredientsHeader
'Me.Refresh
'Generar_Nutrition_IngredientsLines
'Me.Refresh
'Generar_Nutrients_Table
'Me.Refresh
'
''--> Planificación
'Generar_SubMenu_Planificacion
'Me.Refresh
'Generar_Menu_Header_Planificacion
'Me.Refresh
'Generar_Day_Plan_Header_Planificacion
'Me.Refresh
'Generar_Day_Plan_Lines_Planificacion
'Me.Refresh
'Generar_Menu_Dish_Planificacion
'Me.Refresh
'Generar_Dish_Table_Planificacion
'Me.Refresh
'--> Poner fecha inicio proceso
Set RS1 = vg_db.Execute("select getdate() as fecini")
Label1(6).Caption = RS1(0)
RS1.Close
Set RS1 = Nothing

Call MsgBox("Archivo generado en carpeta" & Chr(13) & dir_trabajo & "InformesOptimum\", vbInformation)

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
  
End Sub

Sub Generar_LSRInventTable_1_1_TablaIngredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.Value

Label2.Visible = True
Label2.Caption = "Generar LSRInventTable_1-1 Tabla Ingredientes"

Sql = ""
'20151216 Sql = " sgpadm_Sel_XmlIngLSRInventTableOptimum_1_1_V01 "
'20160106 Sql = " sgpadm_Sel_XmlIngLSRInventTableOptimum_1_1_V03 "
'20160111 Sql = " sgpadm_Sel_XmlIngLSRInventTableOptimum_1_1_V04 "
'20160121 Sql = " sgpadm_Sel_XmlIngLSRInventTableOptimum_1_1_V05 "
'20160502 este script va la ola cl17 y cl37 Sql = " sgpadm_Sel_XmlIngLSRInventTableOptimum_1_1_V06 "
Sql = " sgpadm_Sel_XmlIngLSRInventTableOptimum_1_1_V07 "
'20160502 Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "

RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)
'RS.Close
Set RS = Nothing


'--> Poner fecha inicio proceso
Set RS1 = vg_db.Execute("select getdate() as fecini")
Label1(5).Caption = RS1(0)
RS1.Close
Set RS1 = Nothing

fg_descarga

Exit Sub
Man_Error:
    
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_LSRInventTable_1_1_Ingredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Ingrediente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Ingrediente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Ingrediente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Ingrediente.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar LSRInventTable_1-1 Ingredientes"

Sql = ""
'20151201 Sql = " sgpadm_Sel_XmlIngLSRInventTable_1_1_V01 "
'20151211 Sql = " sgpadm_Sel_XmlIngLSRInventTable_1_1_V02 "
'20151215 Sql = " sgpadm_Sel_XmlIngLSRInventTable_1_1_V03 "
'20160510 Sql = " sgpadm_Sel_XmlIngLSRInventTable_1_1_V04 "
Sql = " sgpadm_Sel_XmlIngLSRInventTable_1_1_V05 "
Sql = Sql & " '" & Extracion & "' "
'Sql = Sql & " '" & xmlorgcom & "', "
'Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
'Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
'Sql = Sql & " '" & OpTablaGramaje & "' "

RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

'--> Poner fecha inicio proceso
Set RS1 = vg_db.Execute("select getdate() as fecini")
Label1(5).Caption = RS1(0)
RS1.Close
Set RS1 = Nothing

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1
  
Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Exclude from portion weight]
   Glosa = Glosa & RS(1) & ";" ' [Main ingredients]
   Glosa = Glosa & RS(2) & ";" ' [Style]
   Glosa = Glosa & RS(3) & ";" ' [Category]
   Glosa = Glosa & RS(4) & ";" ' [CostRange]
   Glosa = Glosa & RS(5) & ";" ' [Diet]
   Glosa = Glosa & RS(6) & ";" ' [Item number]
   Glosa = Glosa & RS(7) & ";" ' [Item name]
   Glosa = Glosa & RS(8) & ";" ' [Initialize location from BOM]
   Glosa = Glosa & RS(9) & ";" ' [Prevent BOM explosion for requisition]
   Glosa = Glosa & RS(10) & ";" ' [Recipe weight unit]
   Glosa = Glosa & RS(11) & ";" ' [Remove from freezer]
   Glosa = Glosa & RS(12) & ";" ' [Recipe target cost]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_BOMTable_1_1_Ingredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim Extracion As String
  
fg_carga ""

Label2.Visible = True
Label2.Caption = "Generar BOMTable_1-1 Ingredientes"

Extracion = fpLongInteger1.text

Sql = ""
'20151124 Sql = " sgpadm_Sel_IngBOMTable_1_1 "
'20151201 Sql = " sgpadm_Sel_IngBOMTable_1_1_V01 "
'20160121 Sql = " sgpadm_Sel_IngBOMTable_1_1_V02 "
'20160510 Sql = " sgpadm_Sel_IngBOMTable_1_1_V03 "
Sql = " sgpadm_Sel_IngBOMTable_1_1_V04 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)
RS.Close
Set RS = Nothing

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_SalidaBOMTable_1_1_Ingredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim Extracion As String
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\BOMTable_1-1_Ingrediente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\BOMTable_1-1_Ingrediente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\BOMTable_1-1_Ingrediente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\BOMTable_1-1_Ingrediente.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar BOMTable_1-1 Ingredientes"

Extracion = fpLongInteger1.text
Sql = ""
'20160516 Sql = " sgpadm_Sel_SalidaIngBOMTable_1_1_V01 "
Sql = " sgpadm_Sel_SalidaIngBOMTable_1_1_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)

   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [BOM]
   Glosa = Glosa & RS(1) & ";" ' [Name]
   Glosa = Glosa & RS(2) & ";" ' [Item group]
   Glosa = Glosa & RS(3) & ";" ' [Check]
   Glosa = Glosa & RS(4) & ";" ' [Approved by]
   Glosa = Glosa & RS(5) & ";" ' [Approved]
   Glosa = Glosa & RS(6) & ";" ' [Approve BOM]
   Glosa = Glosa & RS(7) & ";" ' [SITE]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1
   
Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_BOMVersion_1_1_Ingredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String

fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\BOMVersion_1-1_Ingrediente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\BOMVersion_1-1_Ingrediente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\BOMVersion_1-1_Ingrediente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\BOMVersion_1-1_Ingrediente.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar BOMVersion_1-1 Ingredientes"

Extracion = fpLongInteger1.text
Sql = ""
'20151124 Sql = " sgpadm_Sel_IngBOMVersion_1_1 "
'20160511 Sql = " sgpadm_Sel_IngBOMVersion_1_1_V01 "
Sql = " sgpadm_Sel_IngBOMVersion_1_1_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [To date]
   Glosa = Glosa & RS(1) & ";" ' [From date]
   Glosa = Glosa & RS(2) & ";" ' [Item number]
   Glosa = Glosa & RS(3) & ";" ' [BOM]
   Glosa = Glosa & RS(4) & ";" ' [Name]
   Glosa = Glosa & RS(5) & ";" ' [Active]
   Glosa = Glosa & RS(6) & ";" ' [Approved]
   Glosa = Glosa & RS(7) & ";" ' [Approved by]
   Glosa = Glosa & RS(8) & ";" ' [Construction]
   Glosa = Glosa & RS(9) & ";" ' [From qty.]
   Glosa = Glosa & RS(10) & ";" ' [Dimension No.]
   Glosa = Glosa & RS(11) & ";" ' [Approve BOM version]
   Glosa = Glosa & RS(12) & ";" ' [Activate BOM version]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Ingredient_Item()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String

fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Ingredient_Item.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Ingredient_Item.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Ingredient_Item.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Ingredient_Item.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Ingredient_Item"

Extracion = fpLongInteger1.text
Sql = ""
'20160105 Sql = " sgpadm_Sel_Ingredient_Item_V01 "
'20160511 Sql = " sgpadm_Sel_Ingredient_Item_V02 "
Sql = " sgpadm_Sel_Ingredient_Item_V03 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" '[Item number]
   Glosa = Glosa & RS(1) & ";" '[Item name*]
   Glosa = Glosa & RS(2) & ";" '[Long description (local language)]
   Glosa = Glosa & RS(3) & ";" '[Description additional language]
   Glosa = Glosa & RS(4) & ";" '[Item type]
   Glosa = Glosa & RS(5) & ";" '[Item classification**]
   Glosa = Glosa & RS(6) & ";" '[Counting group]
   Glosa = Glosa & RS(7) & ";" '[Unit***]
   Glosa = Glosa & RS(8) & ";" '[Content]
   Glosa = Glosa & RS(9) & ";" '[Content unit****]
   Glosa = Glosa & RS(10) & ";" '[Net weight]
   Glosa = Glosa & RS(11) & ";" '[Purchase unit****]
   Glosa = Glosa & RS(12) & ";" '[Purchase unit factor]
   Glosa = Glosa & RS(13) & ";" '[Primary Vendor]
   Glosa = Glosa & RS(14) & ";" '[External item number]
   Glosa = Glosa & RS(15) & ";" '[Primary vendor item name]
   Glosa = Glosa & RS(16) & ";" '[Purchase price]
   Glosa = Glosa & RS(17) & ";" '[Bar code]
   Glosa = Glosa & RS(18) & ";" '[Country/region]
   Glosa = Glosa & RS(19) & ";" '[Commodity]
   Glosa = Glosa & RS(20) & ";" '[Item sales tax group]
   Glosa = Glosa & RS(21) & ";" '[Upstream]
   Glosa = Glosa & RS(22) & ";" '[Brand signature]
   Glosa = Glosa & RS(23) & ";" '[Brand rebate]
   Glosa = Glosa & RS(24) & ";" '[Brand rebate percentage]
   Glosa = Glosa & RS(25) & ";" '[Manufacturer rebate]
   Glosa = Glosa & RS(26) & ";" '[Manufacturer rebate percentage]
   Glosa = Glosa & RS(27) & ";" '[Calories]
   Glosa = Glosa & RS(28) & ";" '[Dangerous]
   Glosa = Glosa & RS(29) & ";" '[Characteristic 1]
   Glosa = Glosa & RS(30) & ";" '[Characteristic 2]
   Glosa = Glosa & RS(31) & ";" '[Characteristic 3]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_BOM_1_1_Ingredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim xmlorgcom As String
Dim Extracion As String
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\BOM_1-1_Ingrediente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\BOM_1-1_Ingrediente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\BOM_1-1_Ingrediente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\BOM_1-1_Ingrediente.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar BOM_1-1 Ingredientes"

Extracion = fpLongInteger1.text

Sql = ""
'20151124 Sql = " sgpadm_Sel_IngBOM_1_1 "
'20151215 Sql = " sgpadm_Sel_IngBOM_1_1_V01 "
'20151222 Sql = " sgpadm_Sel_IngBOM_1_1_V02 "
'20160104 Sql = " sgpadm_Sel_IngBOM_1_1_V03 "
'20160304 Sql = " sgpadm_Sel_IngBOM_1_1_V04 "
'20160510 Sql = " sgpadm_Sel_IngBOM_1_1_V05 "
Sql = " sgpadm_Sel_IngBOM_1_1_V06 "
'20160510 Sql = Sql & " '" & xmlorgcom & "' "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Line No]
   Glosa = Glosa & RS(1) & ";" ' [Line type]
   Glosa = Glosa & RS(2) & ";" ' [Consumption is]
   Glosa = Glosa & RS(3) & ";" ' [Item number]
   Glosa = Glosa & RS(4) & ";" ' [Quantity]
   Glosa = Glosa & RS(5) & ";" ' [Calculation]
   Glosa = Glosa & RS(6) & ";" ' [Height]
   Glosa = Glosa & RS(7) & ";" ' [Width]
   Glosa = Glosa & RS(8) & ";" ' [Depth]
   Glosa = Glosa & RS(9) & ";" ' [Density]
   Glosa = Glosa & RS(10) & ";" ' [Constant]
   Glosa = Glosa & RS(11) & ";" ' [Rounding-up]
   Glosa = Glosa & RS(12) & ";" ' [Multiples]
   Glosa = Glosa & RS(13) & ";" ' [Position]
   Glosa = Glosa & RS(14) & ";" ' [From date]
   Glosa = Glosa & RS(15) & ";" ' [To date]
   Glosa = Glosa & RS(16) & ";" ' [Vendor account]
   Glosa = Glosa & RS(17) & ";" ' [Unit]
   Glosa = Glosa & RS(18) & ";" ' [BOM]
   Glosa = Glosa & RS(19) & ";" ' [Formula]
   Glosa = Glosa & RS(20) & ";" ' [Per series]
   Glosa = Glosa & RS(21) & ";" ' [Sub-BOM]
   Glosa = Glosa & RS(22) & ";" ' [Dimension No.]
   Glosa = Glosa & RS(23) & ";" ' [Variable scrap]
   Glosa = Glosa & RS(24) & ";" ' [Constant scrap]
   Glosa = Glosa & RS(25) & ";" ' [Flushing principle]
   Glosa = Glosa & RS(26) & ";" ' [Set subproduction to Consumed]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_LSRBOM_1_1_Ingredientes()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim xmlorgcom As String
Dim Extracion As String

fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\LSRBOM_1-1_Ingrediente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\LSRBOM_1-1_Ingrediente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\LSRBOM_1-1_Ingrediente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\LSRBOM_1-1_Ingrediente.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar LsrBOM_1-1 Ingredientes"

Extracion = fpLongInteger1.text
Sql = ""
'20151124 Sql = " sgpadm_Sel_IngLSRBOM_1_1 "
'20151215 Sql = " sgpadm_Sel_IngLSRBOM_1_1_V01 "
'20151216 Sql = " sgpadm_Sel_IngLSRBOM_1_1_V02 "
'20151222 Sql = " sgpadm_Sel_IngLSRBOM_1_1_V03 "
'20151230 Sql = " sgpadm_Sel_IngLSRBOM_1_1_V04 "
'20160304 Sql = " sgpadm_Sel_IngLSRBOM_1_1_V05 "
'20160511 Sql = " sgpadm_Sel_IngLSRBOM_1_1_V06 "
Sql = " sgpadm_Sel_IngLSRBOM_1_1_V07 "
'20160511 Sql = Sql & " '" & xmlorgcom & "' "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [BOM]
   Glosa = Glosa & RS(1) & ";" ' [BOM type]
   Glosa = Glosa & RS(2) & ";" ' [Wastage %]
   Glosa = Glosa & RS(3) & ";" ' [Exclude from portion weight]
   Glosa = Glosa & RS(4) & ";" ' [Item number]
   Glosa = Glosa & RS(5) & ";" ' [Line NO]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_LSRBOMversion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim xmlorgcom As String
Dim Extracion As String

fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\LSRBOMversion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\LSRBOMversion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\LSRBOMversion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\LSRBOMversion.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar LSRBOMversion"

Extracion = fpLongInteger1.text
Sql = ""
'20160511 Sql = " sgpadm_Sel_LSRBOMversion_V01 "
Sql = " sgpadm_Sel_LSRBOMversion_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" 'Item number
   Glosa = Glosa & RS(1) & ";" 'BOM
   Glosa = Glosa & RS(2) & ";" 'Production time (min)
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_CategoryNutrition()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String

fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Category Nutrition_Nutriente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Category Nutrition_Nutriente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Category Nutrition_Nutriente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Category Nutrition_Nutriente.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Category Nutrition"

Extracion = fpLongInteger1.text
Sql = ""
' 20160512 Sql = " sgpadm_Sel_NutCategoryNutrition_V01 "
Sql = " sgpadm_Sel_NutCategoryNutrition_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)

   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Category Nutrition]
   Glosa = Glosa & RS(1) & ";" ' [Description]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_GroupNutrition()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Group Nutrition_Nutriente.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Group Nutrition_Nutriente.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Group Nutrition_Nutriente.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Group Nutrition_Nutriente.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Group Nutrition"

Extracion = fpLongInteger1.text
Sql = ""
'20160512 Sql = " sgpadm_Sel_NutGroupNutrition_V01 "
Sql = " sgpadm_Sel_NutGroupNutrition_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Category Nutrition]
   Glosa = Glosa & RS(1) & ";" ' [Group Nutrition]
   Glosa = Glosa & RS(2) & ";" ' [Description]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Nutrition_IngredientsHeader()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Nutrition - IngredientsHeader.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Nutrition - IngredientsHeader.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Nutrition - IngredientsHeader.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Nutrition - IngredientsHeader.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Nutrition_Ingredients Header"

Sql = ""
'20160512 Sql = " sgpadm_Sel_NutIngredientsHeader_V01 "
Sql = " sgpadm_Sel_NutIngredientsHeader_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Nutrition*]
   Glosa = Glosa & RS(1) & ";" ' [Description]
   Glosa = Glosa & RS(2) & ";" ' [Category Nutrition]
   Glosa = Glosa & RS(3) & ";" ' [Group Nutrition]
   Glosa = Glosa & RS(4) & ";" ' [Serving size]
   Glosa = Glosa & RS(5) & ";" ' [Unit]
   Glosa = Glosa & RS(6) & ";" ' [Serving size in G]
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Nutrition_IngredientsLines()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String

fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Nutrition - IngredientsLines.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Nutrition - IngredientsLines.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Nutrition - IngredientsLines.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Nutrition - IngredientsLines.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Nutrition_Ingredients Lines"

Extracion = fpLongInteger1.text
Sql = ""
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
'20160512 Sql = " sgpadm_Sel_NutIngredientsLines "
Sql = " sgpadm_Sel_NutIngredientsLines_V01 "
Sql = Sql & " '" & Extracion & "' "

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Set RS = vg_db.Execute(Sql)

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)

   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Nutrition]
   Glosa = Glosa & RS(1) & ";" ' [Line No]
   Glosa = Glosa & RS(2) & ";" ' [Value per portion]
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Nutrients_Table()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Nutrients Table.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Nutrients Table.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Nutrients Table.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Nutrients Table.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Nurients Table"

Extracion = fpLongInteger1.text
Sql = ""
'20160512 Sql = " sgpadm_Sel_NutTable "
Sql = " sgpadm_Sel_NutTable_V01 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Line No]
   Glosa = Glosa & RS(1) & ";" ' [Description]
   Glosa = Glosa & RS(2) & ";" ' [Nutritional report]
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_LSRInventTable_1_1_Recetas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Receta.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Receta.csv")
End If
     
Open dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Receta.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\LSRInventTable_1-1_Receta.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar LsrInventTable_1-1 Recetas"

Sql = ""
'20151030 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1 "
'20151117 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v01 "
'20151211 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v02 "
'20151216 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v03 "
'20151221 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v04 "
'20160105 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v05 "
'20160114 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v06 "
'20160120 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v07 "
'20160121 Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v08 "
'20160503 con este procemiento se bajo la primera hola Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v09 "
Sql = " sgpadm_Sel_XmlRecLSRInventTable_1_1_v10 "
'20160503 Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open Sql, vg_db, adOpenStatic

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Exclude from portion weight]
   Glosa = Glosa & RS(1) & ";" ' [Main ingredients]
   Glosa = Glosa & RS(2) & ";" ' [Style]
   Glosa = Glosa & RS(3) & ";" ' [Category]
   Glosa = Glosa & RS(4) & ";" ' [CostRange]
   Glosa = Glosa & RS(5) & ";" ' [Diet]
   Glosa = Glosa & RS(6) & ";" ' [Item number]
   Glosa = Glosa & RS(7) & ";" ' [Item name]
   Glosa = Glosa & RS(8) & ";" ' [Initialize location from BOM]
   Glosa = Glosa & RS(9) & ";" ' [Prevent BOM explosion for requisition]
   Glosa = Glosa & RS(10) & ";" ' [Recipe weight unit]
   Glosa = Glosa & RS(11) & ";" ' [Remove from freezer]
   Glosa = Glosa & RS(12) & ";" ' [Recipe target cost]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_LinkRecipeItem_Offer_Recetas()

On Error GoTo Man_Error

fg_carga ""

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Link RecipeItem-Offer.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Link RecipeItem-Offer.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Link RecipeItem-Offer.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Link RecipeItem-Offer.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Link RecipeItem-Offer Recetas"

Sql = ""
'20151030 Sql = " sgpadm_Sel_RecLinkRecipeItem_Offer "
'20151117 Sql = " sgpadm_Sel_RecLinkRecipeItem_Offer_V01 "
'20160504 Sql = " sgpadm_Sel_RecLinkRecipeItem_Offer_V02 "
Sql = " sgpadm_Sel_RecLinkRecipeItem_Offer_V03 "
'20160504 Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open Sql, vg_db, adOpenStatic

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Item number]
   Glosa = Glosa & RS(1) & ";" ' [Offer Id]
   Glosa = Glosa & RS(2) & ";" ' [Description]
   Glosa = Glosa & RS(3) & ";" ' [Main ingredients]
   Glosa = Glosa & RS(4) & ";" ' [Style]
   Glosa = Glosa & RS(5) & ";" ' [Category]
   Glosa = Glosa & RS(6) & ";" ' [CostRange]
   Glosa = Glosa & RS(7) & ";" ' [Diet]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_BOMVersion_1_1_Recetas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar BOMVersion_1-1 Recetas"

Sql = ""
'20151030 Sql = " sgpadm_Sel_RecBOMVersion_1_1 "
'20151118 Sql = " sgpadm_Sel_RecBOMVersion_1_1_V01 "
'20160505 Sql = " sgpadm_Sel_RecBOMVersion_1_1_V02 "
Sql = " sgpadm_Sel_RecBOMVersion_1_1_V03 "
20160505 'Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open Sql, vg_db, adOpenStatic
RS.Close
Set RS = Nothing

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_DatoSalidaBOMVersion_1_1_Recetas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\BOMVersion_1-1_Receta.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\BOMVersion_1-1_Receta.csv")
End If
     
Open dir_trabajo & "InformesOptimum\BOMVersion_1-1_Receta.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\BOMVersion_1-1_Receta.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar BOMVersion_1-1 Recetas"

Sql = ""
'20160509 Sql = " sgpadm_Sel_SalidaRecBOMVersion_1_1_V01 "
Sql = " sgpadm_Sel_SalidaRecBOMVersion_1_1_V02 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open Sql, vg_db, adOpenStatic

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [To date]
   Glosa = Glosa & RS(1) & ";" ' [From date]
   Glosa = Glosa & RS(2) & ";" ' [Item number]
   Glosa = Glosa & RS(3) & ";" ' [BOM]
   Glosa = Glosa & RS(4) & ";" ' [Name]
   Glosa = Glosa & RS(5) & ";" ' [Active]
   Glosa = Glosa & RS(6) & ";" ' [Approved]
   Glosa = Glosa & RS(7) & ";" ' [Approved by]
   Glosa = Glosa & RS(8) & ";" ' [Construction]
   Glosa = Glosa & RS(9) & ";" ' [From qty.]
   Glosa = Glosa & RS(10) & ";" ' [Dimension No.]
   Glosa = Glosa & RS(11) & ";" ' [Approve BOM version]
   Glosa = Glosa & RS(12) & ";" ' [Activate BOM version]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1
   
Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_BOMTable_1_1_Recetas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\BOMTable_1-1_Receta.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\BOMTable_1-1_Receta.csv")
End If
     
Open dir_trabajo & "InformesOptimum\BOMTable_1-1_Receta.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\BOMTable_1-1_Receta.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar BOMTable_1-1 Recetas"

Extracion = fpLongInteger1.text

Sql = ""

'20151030 Sql = " sgpadm_Sel_RecBOMTable_1_1 "
'20151119 Sql = " sgpadm_Sel_RecBOMTable_1_1_V01  "
'20151211 Sql = " sgpadm_Sel_RecBOMTable_1_1_V02  "
'20160509 Sql = " sgpadm_Sel_RecBOMTable_1_1_V03  "
Sql = " sgpadm_Sel_RecBOMTable_1_1_V04 "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [BOM]
   Glosa = Glosa & RS(1) & ";" ' [Name]
   Glosa = Glosa & RS(2) & ";" ' [Item group]
   Glosa = Glosa & RS(3) & ";" ' [Check]
   Glosa = Glosa & RS(4) & ";" ' [Approved by]
   Glosa = Glosa & RS(5) & ";" ' [Approved]
   Glosa = Glosa & RS(6) & ";" ' [Approve BOM]
   Glosa = Glosa & RS(7) & ";" ' [SITE]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Recipe_Item()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim Glosa As String
Dim i As Long
Dim Extracion As String
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Recipe_Item.csv ")) Then
    Kill (dir_trabajo & "InformesOptimum\Recipe_Item.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Recipe_Item.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Recipe_Item.csv" For Append As #1
    
Label2.Visible = True
Label2.Caption = "Generar Recipe_Item Recetas"
Extracion = fpLongInteger1.text

Sql = ""

'20160105 Sql = " sgpadm_Sel_Recipe_Item_V01 "
'20160510 Sql = " sgpadm_Sel_Recipe_Item_V02 "
Sql = " sgpadm_Sel_Recipe_Item_V03 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
            
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" '[Item number]
   Glosa = Glosa & RS(1) & ";" '[Item name*]
   Glosa = Glosa & RS(2) & ";" '[Long description (local language)]
   Glosa = Glosa & RS(3) & ";" '[Description additional language]
   Glosa = Glosa & RS(4) & ";" '[Item type]
   Glosa = Glosa & RS(5) & ";" '[Item classification**]
   Glosa = Glosa & RS(6) & ";" '[Counting group]
   Glosa = Glosa & RS(7) & ";" '[Unit***]
   Glosa = Glosa & RS(8) & ";" '[Content]
   Glosa = Glosa & RS(9) & ";" '[Content unit****]
   Glosa = Glosa & RS(10) & ";" '[Net weight]
   Glosa = Glosa & RS(11) & ";" '[Purchase unit****]
   Glosa = Glosa & RS(12) & ";" '[Purchase unit factor]
   Glosa = Glosa & RS(13) & ";" '[Primary Vendor]
   Glosa = Glosa & RS(14) & ";" '[External item number]
   Glosa = Glosa & RS(15) & ";" '[Primary vendor item name]
   Glosa = Glosa & RS(16) & ";" '[Purchase price]
   Glosa = Glosa & RS(17) & ";" '[Bar code]
   Glosa = Glosa & RS(18) & ";" '[Country/region]
   Glosa = Glosa & RS(19) & ";" '[Commodity]
   Glosa = Glosa & RS(20) & ";" '[Item sales tax group]
   Glosa = Glosa & RS(21) & ";" '[Upstream]
   Glosa = Glosa & RS(22) & ";" '[Brand signature]
   Glosa = Glosa & RS(23) & ";" '[Brand rebate]
   Glosa = Glosa & RS(24) & ";" '[Brand rebate percentage]
   Glosa = Glosa & RS(25) & ";" '[Manufacturer rebate]
   Glosa = Glosa & RS(26) & ";" '[Manufacturer rebate percentage]
   Glosa = Glosa & RS(27) & ";" '[Calories]
   Glosa = Glosa & RS(28) & ";" '[Dangerous]
   Glosa = Glosa & RS(29) & ";" '[Characteristic 1]
   Glosa = Glosa & RS(30) & ";" '[Characteristic 2]
   Glosa = Glosa & RS(31) & ";" '[Characteristic 3]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo
End Sub

Sub Generar_BOM_1_1_Recetas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\BOM_1-1_Receta.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\BOM_1-1_Receta.csv")
End If
     
Open dir_trabajo & "InformesOptimum\BOM_1-1_Receta.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\BOM_1-1_Receta.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar BOM_1-1 Recetas"

Sql = ""
'20151030 Sql = " sgpadm_Sel_XmlRecBOM_1_1 "
'20151123 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V01 "
'20151207 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V02 "
'20151215 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V03 "
'20151215 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V04 "
'20151222 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V05 "
'20160108 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V06 "
'20160121 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V07 "
'20160209 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V08 "
'20160509 Sql = " sgpadm_Sel_XmlRecBOM_1_1_V09 "
Sql = " sgpadm_Sel_XmlRecBOM_1_1_V10 "
'Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)

   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Line No]
   Glosa = Glosa & RS(1) & ";" ' [Line type]
   Glosa = Glosa & RS(2) & ";" ' [Consumption is]
   Glosa = Glosa & RS(3) & ";" ' [Item number]
   Glosa = Glosa & RS(4) & ";" ' [Quantity]
   Glosa = Glosa & RS(5) & ";" ' [Calculation]
   Glosa = Glosa & RS(6) & ";" ' [Height]
   Glosa = Glosa & RS(7) & ";" ' [Width]
   Glosa = Glosa & RS(8) & ";" ' [Depth]
   Glosa = Glosa & RS(9) & ";" ' [Density]
   Glosa = Glosa & RS(10) & ";" ' [Constant]
   Glosa = Glosa & RS(11) & ";" ' [Rounding-up]
   Glosa = Glosa & RS(12) & ";" ' [Multiples]
   Glosa = Glosa & RS(13) & ";" ' [Position]
   Glosa = Glosa & RS(14) & ";" ' [From date]
   Glosa = Glosa & RS(15) & ";" ' [To date]
   Glosa = Glosa & RS(16) & ";" ' [Vendor account]
   Glosa = Glosa & RS(17) & ";" ' [Unit]
   Glosa = Glosa & RS(18) & ";" ' [BOM]
   Glosa = Glosa & RS(19) & ";" ' [Formula]
   Glosa = Glosa & RS(20) & ";" ' [Per series]
   Glosa = Glosa & RS(21) & ";" ' [Sub-BOM]
   Glosa = Glosa & RS(22) & ";" ' [Dimension No.]
   Glosa = Glosa & RS(23) & ";" ' [Variable scrap]
   Glosa = Glosa & RS(24) & ";" ' [Constant scrap]
   Glosa = Glosa & RS(25) & ";" ' [Flushing principle]
   Glosa = Glosa & RS(26) & ";" ' [Set subproduction to Consumed]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_LSRBOM_1_1_Recetas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\LSRBOM_Receta.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\LSRBOM_1-1_Receta.csv")
End If
     
Open dir_trabajo & "InformesOptimum\LSRBOM_1-1_Receta.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\LSRBOM_1-1_Receta.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar LSRBOM_1-1 Recetas"

Sql = ""
'20151030 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1 "
'20151123 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V01 "
'20151207 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V02 "
'20151215 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V03 "
'20151222 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V04 "
'20151230 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V05 "
'20160121 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V06 "
'20160509 Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V07 "
Sql = " sgpadm_Sel_XmlRecLSRBOM_1_1_V08 "
'20160509 Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [BOM]
   Glosa = Glosa & RS(1) & ";" ' [BOM type]
'   Glosa = Glosa & RS(2) & ";" ' [Wastage %]
   Glosa = Glosa & 0 & ";" ' [Wastage %]
   Glosa = Glosa & RS(3) & ";" ' [Exclude from portion weight]
   Glosa = Glosa & RS(4) & ";" ' [Item number]
   Glosa = Glosa & RS(5) & ";" ' [Line No]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Recipe_method()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje As String
Dim MetodoPreparación As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Recipe_method.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Recipe_method.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Recipe_method.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Recipe_method.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Recipe method"

Sql = ""
'20160510 Sql = " sgpadm_Sel_Recipe_method_v01 "
'20160629 Sql = " sgpadm_Sel_Recipe_method_v02 "
Sql = " sgpadm_Sel_Recipe_method_v03 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open Sql, vg_db, adOpenStatic

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Item Number]
   Glosa = Glosa & RS(1) & ";" '  [Bom]
   Glosa = Glosa & RS(2) & ";" ' [Warehouse]
   Sql = ""
   RichTextBox1.TextRTF = ""
   RichTextBox1.TextRTF = IIf(IsNull(RS(3)), "", (RS(3)))
   vaSpread1.Row = 1
   vaSpread1.Col = 6
   vaSpread1.text = ""
   vaSpread1.text = Replace(RichTextBox1.text, Chr(13) + Chr(10), " ")
   metodopreparacion = "" 'vbCrLf
   metodopreparacion = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(RichTextBox1.text, Chr(13) + Chr(10), " "), Chr(27), ""), "*", vbCrLf), " f3", "o"), Chr(225) + "7", "-"), "f1", "ñ"), "e1", "a"), "e9", "e"), "b7", " ")
   Text1.text = metodopreparacion
   vg_db.Execute ("update b_receta set rec_metodopreparacion = '" & Text1.text & "' where rec_codigo = " & RS(4) & " ")

   Glosa = Glosa & metodopreparacion & ";"
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_SubMenu_Planificacion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\SubMenu_Planificacion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\SubMenu_Planificacion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\SubMenu_Planificacion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\SubMenu_Planificacion.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Sub Menu Planificación"

Sql = ""
'20151130 Sql = " sgpadm_Sel_PlanSubMenu_V01 "
'20160512 Sql = " sgpadm_Sel_PlanSubMenu_V02 "
Sql = " sgpadm_Sel_PlanSubMenu_V03 "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Submenu *]
   Glosa = Glosa & RS(1) & ";" ' [Submenu order]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1
   
Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Dish_Table_Planificacion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Dish_Table_Planificacion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Dish_Table_Planificacion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Dish_Table_Planificacion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Dish_Table_Planificacion.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Dish Table Planificación"

Sql = ""
'20151130 Sql = " sgpadm_Sel_XmlPlanDish_Table_V01 "
'20151223 Sql = " sgpadm_Sel_XmlPlanDish_Table_V02 "
'20160104 Sql = " sgpadm_Sel_XmlPlanDish_Table_V03 "
'20160516 Sql = " sgpadm_Sel_XmlPlanDish_Table_V04 "
Sql = " sgpadm_Sel_XmlPlanDish_Table_V05 "
'20160516 Sql = Sql & " '" & xmlorgcom & "', "
'20160516 Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
'20160516 Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
'20160516 Sql = Sql & " '" & OpTablaGramaje & "' "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Dish ID*]
   Glosa = Glosa & RS(1) & ";" ' [Description]
   Glosa = Glosa & RS(2) & ";" ' [Submenu**]
   Glosa = Glosa & RS(3) & ";" ' [Category filter]
   Glosa = Glosa & RS(4) & ";" ' [Main ingredient filter]
   Glosa = Glosa & RS(5) & ";" ' [Style filter]
   Glosa = Glosa & RS(6) & ";" ' [Cost range]
   Glosa = Glosa & RS(7) & ";" ' [Diet]
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Menu_Header_Planificacion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Menu_Header_Planificacion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Menu_Header_Planificacion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Menu_Header_Planificacion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Menu_Header_Planificacion.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Menu Header Planificación"

Sql = ""
'20151202 Sql = " sgpadm_Sel_XmlPlanMenu_Header_V01 "
'20151211 Sql = " sgpadm_Sel_XmlPlanMenu_Header_V02 "
'20151215 Sql = " sgpadm_Sel_XmlPlanMenu_Header_V03 "
'20160512 Sql = " sgpadm_Sel_XmlPlanMenu_Header_V04 "
Sql = " sgpadm_Sel_XmlPlanMenu_Header_V05 "
'Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
'Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Menu]
   Glosa = Glosa & RS(1) & ";" ' [Description]
   Glosa = Glosa & RS(2) & ";" ' [Warehouse]
   Glosa = Glosa & RS(3) & ";" ' [Journal name]
   Glosa = Glosa & RS(4) & ";" ' [Consumption type]
   Glosa = Glosa & RS(5) & ";" ' [Consumption group]
   Glosa = Glosa & RS(6) & ";" ' [Menu target cost]
   Glosa = Glosa & RS(7) & ";" ' [All days]
   Glosa = Glosa & RS(8) & ";" ' [Monday]
   Glosa = Glosa & RS(9) & ";" ' [Tuesday]
   Glosa = Glosa & RS(10) & ";" ' [Wednesday]
   Glosa = Glosa & RS(11) & ";" ' [Thursday]
   Glosa = Glosa & RS(12) & ";" ' [Friday]
   Glosa = Glosa & RS(13) & ";" ' [Saturday]
   Glosa = Glosa & RS(14) & ";" ' [Sunday]
'   Glosa = Glosa & RS(15) & ";" ' [Consumption type]
'   Glosa = Glosa & RS(16) & ";" ' [Consumption group]
'   Glosa = Glosa & RS(17) & ";" ' [Time from]
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Menu_Dish_Planificacion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Menu_Dish_Planificacion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Menu_Dish_Planificacion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Menu_Dish_Planificacion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Menu_Dish_Planificacion.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Menu Dish Planificación"

Sql = ""
'20151124 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V01 "
'20151203 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V02 "
'20151211 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V03 "
'20151215 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V04 "
'20151223 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V05 "
'20160108 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V06 "
'20160209 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V07 "
'20160516 Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V08 "
Sql = " sgpadm_Sel_XmlPlanMenu_Dish_V09 "
'20160516 Sql = Sql & " '" & xmlorgcom & "', "
'Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
'Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
'Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [MenuId]
   Glosa = Glosa & RS(1) & ";" ' [DishId]
   Glosa = Glosa & RS(2) & ";" ' [DishDescription]
   Glosa = Glosa & RS(3) & ";" ' [RecommPriceInclVAT]
   Glosa = Glosa & RS(4) & ";" ' [SubmenuId]
   Glosa = Glosa & RS(5) & ";" ' [SequenceNo]
   Glosa = Glosa & RS(6) & ";" ' [sdxPortionFactorDish]
   Glosa = Glosa & RS(7) & ";" ' [sdxDishTargetCost]
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Day_Plan_Header_Planificacion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Day_Plan_Header_Planificacion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Day_Plan_Header_Planificacion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Day_Plan_Header_Planificacion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Day_Plan_Header_Planificacion.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Day Plan Header Planificación"

Sql = ""
'20151202 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Header_V01 "
'20151211 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Header_V02 "
'20150513 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Header_V03 "
Sql = " sgpadm_Sel_XmlPlanDay_Plan_Header_V04 "
'Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & OpTablaGramaje & "', "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Day plan date]
   Glosa = Glosa & RS(1) & ";" ' [Menu]
   Glosa = Glosa & RS(2) & ";" ' [Warehouse]
   Glosa = Glosa & RS(3) & ";" ' [Total amount]
   Glosa = Glosa & RS(4) & ";" ' [Plan lines exist]
   Glosa = Glosa & RS(5) & ";" ' [Filter Date Id]
   Glosa = Glosa & RS(6) & ";" ' [ReplicationCounter]
   Glosa = Glosa & RS(7) & ";" ' [Button grid id]
   Glosa = Glosa & " " & ";"   ' Blanco
   Glosa = Glosa & RS(3) & ";" '[Number of Guests]
   
   Print #1, Glosa
    
   RS.MoveNext
   i = i + 1

Loop
Close #1
RS.Close
Set RS = Nothing


ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Sub Generar_Day_Plan_Lines_Planificacion()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String
Dim i As Long
Dim seleccion As Integer
Dim xmlorgcom  As String
Dim orgcom As String
Dim Ceco As String
Dim Glosa As String
Dim OpTablaGramaje  As String
Dim Extracion As String

seleccion = 0
  
fg_carga ""

xmlorgcom = ""
xmlorgcom = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
xmlorgcom = xmlorgcom & "<OrgC>"
  
For i = 1 To vaSpread2.MaxRows
       
    vaSpread2.Row = i
    vaSpread2.Col = 1 'Seleccion
    seleccion = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    If seleccion = 1 And vaSpread2.RowHidden = False Then
       
       xmlorgcom = xmlorgcom & "<OC"
       vaSpread2.Row = i
       
       vaSpread2.Col = 2 'Ceco
       Ceco = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       vaSpread2.Col = 4 'Org. Compras
       orgcom = IIf(vaSpread2.text = "", 0, vaSpread2.text)
       
       xmlorgcom = xmlorgcom & " C = " & Chr(34) & Ceco & Chr(34)
       xmlorgcom = xmlorgcom & " O = " & Chr(34) & orgcom & Chr(34)
       
       xmlorgcom = xmlorgcom & "/>"
    
    End If
   
Next i
xmlorgcom = xmlorgcom & "</OrgC>"

Set AnJes = CreateObject("scripting.filesystemobject")
If Not AnJes.FolderExists(dir_trabajo & "InformesOptimum") Then
   Call AnJes.CreateFolder(dir_trabajo & "InformesOptimum")
End If

If (AnJes.FileExists(dir_trabajo & "InformesOptimum\Day_Plan_Lines_Planificacion.csv")) Then
    Kill (dir_trabajo & "InformesOptimum\Day_Plan_Lines_Planificacion.csv")
End If
     
Open dir_trabajo & "InformesOptimum\Day_Plan_Lines_Planificacion.csv" For Append As #1
Close #1
Open dir_trabajo & "InformesOptimum\Day_Plan_Lines_Planificacion.csv" For Append As #1
    
OpTablaGramaje = IIf(ChTGramaje.Value = 0, "0", "1")
Extracion = fpLongInteger1.text

Label2.Visible = True
Label2.Caption = "Generar Day Plan Lines Planificación"

Sql = ""
'20151030 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines "
'20151124 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V01 "
'20151211 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V02 "
'20151215 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V03 "
'20151223 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V04 "
'20160120 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V05 "
'20160516 Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V06 "
Sql = " sgpadm_Sel_XmlPlanDay_Plan_Lines_V07 "
'20160516 Sql = Sql & " '" & xmlorgcom & "', "
Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ", "
Sql = Sql & " '" & Extracion & "' "
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute(Sql)

ProgressBar1.Scrolling = ccScrollingSmooth
ProgressBar1.Max = 100
ProgressBar1.Visible = True
ProgressBar1.Value = 0
i = 1

Do While Not RS.EOF

   ProgressBar1.Value = Val((i / RS.RecordCount) * 100)
   
   Glosa = ""
   Glosa = Glosa & RS(0) & ";" ' [Day plan date]
   Glosa = Glosa & RS(1) & ";" ' [Menu]
   Glosa = Glosa & RS(2) & ";" ' [LineNo]
   Glosa = Glosa & RS(3) & ";" ' [Dish]
   Glosa = Glosa & RS(4) & ";" ' [Item number]
   Glosa = Glosa & RS(5) & ";" ' [Planned qty]
   Glosa = Glosa & RS(6) & ";" ' [Portion Factor*]
   Glosa = Glosa & RS(7) & ";" ' [Unit price]
   Glosa = Glosa & RS(8) & ";" ' [Actual unit cost]
   Glosa = Glosa & RS(9) & ";" ' [Warehouse]
   Glosa = Glosa & RS(10) & ";" ' [Submenu]
   Glosa = Glosa & RS(11) & ";" ' [ReplicationCounter]
   Glosa = Glosa & RS(12) & ";" ' [Actual prepared qty]
   Glosa = Glosa & RS(13) & ";" ' [Requisition qty]
   Glosa = Glosa & RS(14) & ";" ' [Sequence]
   Glosa = Glosa & RS(15) & ";" ' [Purchase requisition ID]
   Glosa = Glosa & RS(16) & ";" ' [Ordered qty]
   Glosa = Glosa & RS(17) & ";" ' [Unit]
   Glosa = Glosa & RS(18) & ";" ' [Button grid id]
   Glosa = Glosa & RS(19) & ";" ' [Button id]
   Glosa = Glosa & RS(20) & ";" ' [Production]
   
   Print #1, Glosa
   i = i + 1
    
   RS.MoveNext

Loop
Close #1
RS.Close
Set RS = Nothing

ProgressBar1.Visible = False
'Label2.Visible = False

fg_descarga

Exit Sub
Man_Error:
    Close #1
    fg_descarga
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

'On Error GoTo Man_Error
'
'Dim i As Long
'
'Select Case BlockCol
'
'Case 1
'
'    vaSpread2.Col = 1
'    For i = BlockRow To BlockRow2
'        vaSpread2.Row = i
'        vaSpread2.Value = IIf(vaSpread2.Value = "1", "0", "1")
'    Next
'
'End Select
'
'Exit Sub
'Man_Error:
'    fg_descarga
'    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo

End Sub
