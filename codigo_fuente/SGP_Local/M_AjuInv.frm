VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_AjuInv 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajuste de Inventario"
   ClientHeight    =   6390
   ClientLeft      =   1440
   ClientTop       =   2415
   ClientWidth     =   11670
   Icon            =   "M_AjuInv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6390
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   2415
      Left            =   3960
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton Command1 
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
         Index           =   1
         Left            =   2760
         TabIndex        =   6
         Top             =   1920
         Width           =   1425
      End
      Begin VB.CommandButton Command1 
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
         Index           =   0
         Left            =   1320
         TabIndex        =   5
         Top             =   1920
         Width           =   1425
      End
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   1
         Left            =   1830
         TabIndex        =   4
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
         Left            =   1830
         TabIndex        =   3
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Como realizo un ajuste de precio, tiene  comunicarse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   4
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   4515
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   480
         TabIndex        =   19
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   480
         TabIndex        =   18
         Top             =   1500
         Width           =   930
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   45
      TabIndex        =   10
      Top             =   345
      Width           =   11520
      Begin EditLib.fpText Combo1 
         Height          =   345
         Index           =   0
         Left            =   4320
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   195
         Width           =   2970
         _Version        =   196608
         _ExtentX        =   5239
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         ThreeDOutsideHighlightColor=   -2147483628
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
         CaretInsert     =   0
         CaretOverWrite  =   3
         UserEntry       =   0
         HideSelection   =   -1  'True
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   2
         MarginTop       =   2
         MarginRight     =   2
         MarginBottom    =   2
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   2
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   255
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
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   1785
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   195
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
         _ExtentY        =   609
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   0   'False
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
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         ControlType     =   2
         Text            =   ""
         DateCalcMethod  =   0
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
         Caption         =   "Bodega"
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
         Left            =   3570
         TabIndex        =   14
         Top             =   240
         Width           =   660
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
         Left            =   1125
         TabIndex        =   13
         Top             =   240
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5310
      Left            =   45
      TabIndex        =   12
      Top             =   945
      Width           =   11535
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   120
         TabIndex        =   16
         Top             =   4680
         Width           =   1785
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   1680
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   1935
         TabIndex        =   15
         Top             =   4680
         Width           =   4125
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   4020
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4365
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   11295
         _Version        =   393216
         _ExtentX        =   19923
         _ExtentY        =   7699
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         FormulaSync     =   0   'False
         MaxCols         =   9
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_AjuInv.frx":0442
         TextTipDelay    =   200
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   420
         Left            =   7320
         TabIndex        =   9
         Top             =   4800
         Width           =   3840
         _ExtentX        =   6773
         _ExtentY        =   741
         ButtonWidth     =   3307
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
         BorderStyle     =   1
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
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_AjuInv.frx":0C37
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_AjuInv.frx":0F51
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_AjuInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim MsgTitulo As String, est As Boolean, estpre As Boolean, estexi As Boolean, estgra As Boolean
Dim aAp As String
Dim precer As Boolean
    
Private Sub Command1_Click(Index As Integer)

Dim RS As New ADODB.Recordset

Select Case Index

Case 0
    
    '-------> Validar usuario
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_param WHERE par_valor = '" & LimpiaDato(Trim(Nombre(0).text)) & "' AND par_codigo = 'usulimbas' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If RS.EOF Then MsgBox "Login no existe...": RS.Close: Set RS = Nothing: Nombre(0).text = "": Nombre(0).SetFocus: Exit Sub
    RS.Close: Set RS = Nothing
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT * FROM a_param WHERE par_codigo = 'parconajpi' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If Not RS.EOF And UCase(Nombre(1).text) <> UCase(fg_Desencripta(TipoDato(RS!par_valor, ""))) Then MsgBox "La clave no corresponde al login...": RS.Close: Set RS = Nothing: Nombre(0).text = "": Nombre(0).SetFocus: Exit Sub
    RS.Close: Set RS = Nothing
    
    Frame5.Visible = False
    Nombre(0).text = ""
    Nombre(1).text = ""
    '-------> Grabar datos cuando existe modificación en precio
    GrabarDatos
    Toolbar1.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True

Case 1
    
    Frame5.Visible = False
    Nombre(0).text = ""
    Nombre(1).text = ""
    Toolbar1.Enabled = True
    Frame1.Enabled = True
    Frame2.Enabled = True

End Select

End Sub

Private Sub Form_Activate()

fg_descarga
If vg_pais = "CO" And vg_invrot = "1" Then
   
   If Not estexi Then
      
      If estgra Then Toolbar1_ButtonClick Toolbar1.Buttons(1)
      If precer Then Toolbar1_ButtonClick Toolbar1.Buttons(6)
   
   End If

End If

End Sub

Private Sub Form_Load()

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Me.Height = 6870
Me.Width = 11760
Dim X As Boolean, estfpro As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
est = False
estpre = False
estexi = True
estgra = True

fg_centra Me
Me.HelpContextID = vg_OpcM
MsgTitulo = "Ajuste de Inventario"

precer = True
Gl_Mo_Botones Me, 5
Gl_Ac_Botones Me, 5, 1, ""
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

Dim Sql As String, v_fecinv  As Variant, v_codbod As Long, i As Long, difer As Double, lisnom As String
Dim liscod As String, aju_tipo As String, codaux As Long, z As Long

Date1(0).text = M_TomInv.Date1(0).text
Combo1(0).text = Left(M_TomInv.Combo1(0).List(M_TomInv.Combo1(0).ListIndex), 50)
v_codbod = fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)
v_fecinv = Format(Date1(0).text, "yyyymmdd")
aAp = ""
'--------- Muestra inventario guardado -----------
opfampro = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), " ORDER BY dev.dev_numlin", " ORDER BY pro.pro_ctacon, pro.pro_codtip, pro.pro_nombre")
estfpro = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
   
RS.Open "SELECT dev.dev_codmer, pro.pro_nombre, pro.pro_codtip, pro.pro_ctacon, dev.dev_precos, uni.uni_nombre, dev.dev_canmer, aju.aju_tipo, aju.aju_codigo " & _
        "FROM  b_totventas tov, b_detventas dev, b_productos pro, a_unidad uni, a_tipoajuste aju " & _
        "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
        "AND   pro.pro_codigo = dev.dev_codmer AND uni.uni_codigo = pro.pro_coduni AND tov.tov_codser = aju.aju_codigo " & _
        "AND   tov.tov_fecemi = '" & Format(Date1(0).text, "yyyymmdd") & "' AND tov_codbod = " & v_codbod & " AND tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' " & opfampro & "", vg_db, adOpenStatic

vaSpread1.MaxRows = 0
i = 1
codtip = 0

If Not RS.EOF Then
   
   estexi = True
    
    Do While Not RS.EOF
       
       vaSpread1.MaxRows = i
       vaSpread1.Row = vaSpread1.MaxRows
       
       If RS!pro_codtip <> codtip And estfpro Then
          
          vaSpread1.Col = 1: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.text = fg_BuscaenArbol(RS!pro_codtip, "a_tipopro", "tip_codigo")
          vaSpread1.Col = 3: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 4: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 5: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 6: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 7: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 8: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 9: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          codtip = RS!pro_codtip
          i = i + 1
          vaSpread1.MaxRows = i
          vaSpread1.Row = i
        
        End If
        
        vaSpread1.Font.Bold = False
        vaSpread1.Col = 1: vaSpread1.text = RS!dev_codmer
        vaSpread1.Col = 2: vaSpread1.text = RS!pro_nombre
        vaSpread1.Col = 3: vaSpread1.text = RS!uni_nombre
        vaSpread1.Col = 4: vaSpread1.text = IIf(RS!aju_tipo = "D", Format(RS!dev_canmer * -1, fg_Pict(9, vg_DCa)), Format(RS!dev_canmer, fg_Pict(9, vg_DCa))): vaSpread1.ForeColor = IIf(RS!aju_tipo = "D", RGB(255, 0, 0), RGB(0, 0, 0))
        vaSpread1.Col = 5: vaSpread1.text = Format(RS!dev_precos, fg_Pict(9, 2)) 'vg_DCa))
        vaSpread1.Col = 8: vaSpread1.text = Format(RS!dev_precos, fg_Pict(9, 2)) 'vg_DCa))
        vaSpread1.Col = 9: vaSpread1.text = Format(0, fg_Pict(9, vg_DCa))
        lisnom = "": liscod = ""
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        RS2.Open "SELECT aju_codigo, aju_nombre FROM a_tipoajuste WHERE aju_tipaju = 1 AND aju_tipo = '" & RS!aju_tipo & "'", vg_db, adOpenStatic
        Do While Not RS2.EOF
            
            vaSpread1.Col = 6: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS2!aju_nombre)
            vaSpread1.Col = 7: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS2!aju_codigo
            RS2.MoveNext
        
        Loop
        RS2.Close
        Set RS2 = Nothing
        
        vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = lisnom
        vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = liscod
        
        For z = 0 To vaSpread1.TypeComboBoxCount
            
            vaSpread1.TypeComboBoxCurSel = z
            If Val(vaSpread1.text) = RS!aju_codigo Then codaux = z: Exit For
            codaux = -1
        
        Next z
        
        vaSpread1.Col = 6
        vaSpread1.TypeComboBoxCurSel = codaux
        RS.MoveNext
        i = i + 1
    
    Loop
    Toolbar2.Enabled = False
    Gl_Ac_Botones Me, 5, 2, ""

Else
    
    Dim CodAju As Long
    
    estexi = False
    '-------> Crear tabla temporal
    aAp = Trim(vg_NUsr) & "_tmp_AjusteInventarioN1"
    fg_CheckTmp aAp
 
    vg_db.Execute "SELECT DISTINCT a.tin_fectom, a.tin_codbod, a.tin_codpro, c.ppd_propon, c.ppd_propon AS aux_propon, a.tin_stosis, a.tin_stofis " & _
                  "INTO " & aAp & " " & _
                  "FROM  b_tomainv a INNER JOIN b_productos b " & _
                  "ON    a.tin_codpro = b.pro_codigo " & _
                  "LEFT  JOIN b_productospmpdia c " & _
                  "ON    b.pro_codigo = c.ppd_codpro  " & _
                  "AND   c.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                  "AND   c.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " " & _
                  "WHERE a.tin_fectom = " & v_fecinv & " " & _
                  "AND   a.tin_codbod = " & vg_codbod & " " & _
                  "AND   Round(a.tin_stosis, " & vg_DCa & ") <> Round(a.tin_stofis, " & vg_DCa & ")"
                      
    '--------- Muestra las diferencias del ultimo inventario -----------
    opfampro = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), " ORDER BY b.pro_nombre", " ORDER BY b.pro_ctacon, b.pro_codtip, b.pro_nombre")
    estfpro = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
           
    RS1.Open "SELECT DISTINCT a.tin_codpro, b.pro_nombre, b.pro_codtip, b.pro_ctacon, d.ppd_propon, c.uni_nombre, " & _
             "a.tin_stosis , a.tin_stofis " & _
             "FROM b_tomainv a " & _
             "INNER JOIN b_productos  b " & _
             "ON a.tin_codpro = b.pro_codigo " & _
             "INNER JOIN a_unidad c " & _
             "ON b.pro_coduni = c.uni_codigo " & _
             "LEFT JOIN b_productospmpdia d " & _
             "ON a.tin_codpro = d.ppd_codpro " & _
             "AND   d.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " " & _
             "AND   d.ppd_cencos = '" & MuestraCasino(1) & "' " & _
             "WHERE a.tin_fectom = " & v_fecinv & " " & _
             "AND   a.tin_codbod= " & v_codbod & " " & _
             "AND   Round(a.tin_stosis, " & vg_DCa & ") <> Round(a.tin_stofis, " & vg_DCa & ") " & _
             "" & opfampro & "", vg_db, adOpenStatic
        
    Do While Not RS1.EOF
        
        vaSpread1.MaxRows = i
        vaSpread1.Row = vaSpread1.MaxRows
        
        If RS1!pro_codtip <> codtip And estfpro Then
           
           vaSpread1.Col = 1: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.text = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
           vaSpread1.Col = 3: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 4: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 5: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 6: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 7: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 8: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           vaSpread1.Col = 9: vaSpread1.Font.Bold = True: vaSpread1.text = ""
           codtip = RS1!pro_codtip
           i = i + 1: vaSpread1.MaxRows = i: vaSpread1.Row = i
        
        End If
        
        vaSpread1.Font.Bold = False
        difer = Format(RS1!tin_stofis - RS1!tin_stosis, fg_Pict(9, vg_DCa))
        vaSpread1.Col = 1: vaSpread1.text = RS1!tin_codpro
        vaSpread1.Col = 2: vaSpread1.text = RS1!pro_nombre
        vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nombre
        vaSpread1.Col = 4: vaSpread1.text = Format(difer, fg_Pict(9, vg_DCa)): vaSpread1.ForeColor = IIf(difer < 0, RGB(255, 0, 0), RGB(0, 0, 0))
        vaSpread1.Col = 5: vaSpread1.Lock = False: vaSpread1.text = Format(IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon), fg_Pict(9, 2)) 'vg_DCa))
        vaSpread1.Col = 8: vaSpread1.text = Format(IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon), fg_Pict(9, 2)) 'vg_DCa))
        vaSpread1.Col = 9: vaSpread1.text = Format(RS1!tin_stosis, fg_Pict(9, vg_DCa))
        lisnom = "": liscod = ""
        aju_tipo = IIf(difer < 0, "D", "A")
        CodAju = IIf(difer < 0, 4, 3)
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS2.Open "SELECT aju_codigo, aju_nombre FROM a_tipoajuste WHERE aju_tipaju = 1 and aju_tipo = '" & aju_tipo & "' and aju_codigo = " & CodAju & "", vg_db, adOpenStatic
        
        Do While Not RS2.EOF
            
            vaSpread1.Col = 6: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS2!aju_nombre)
            vaSpread1.Col = 7: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS2!aju_codigo
            If RS1!ppd_propon = 0 And RS1!tin_stosis = 0 And liscod = "3" Then
            
               Exit Do
               
            End If
            
            RS2.MoveNext
        
        Loop
        RS2.Close
        Set RS2 = Nothing
        
        vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = lisnom
        vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = liscod
        
        If RS1!tin_stosis = 0 Or RS1!ppd_propon = 0 Then
            
            For z = 0 To vaSpread1.TypeComboBoxCount
                
                vaSpread1.TypeComboBoxCurSel = z
'                If RS1!ppd_propon = 0 Then codaux = z: vaSpread1.Col = 5: vaSpread1.Lock = False: Exit For
                If RS1!ppd_propon = 0 Then codaux = z: Exit For
                codaux = -1
            
            Next z
            vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = codaux
            vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = codaux
        
        End If
        
        RS1.MoveNext
        i = i + 1
    
    Loop
    RS1.Close
    Set RS1 = Nothing
    '----------------------------------------------------------------------
    
    Gl_Ac_Botones Me, 5, 1, ""

End If
vaSpread1.Row = -1
vaSpread1.Col = 6
vaSpread1.Lock = IIf(RS.RecordCount > 0, True, False)

'Gl_Ac_Botones Me, 5, IIf(RS.RecordCount > 0 Or CDate(Date1(0).text) <> (CDate(vg_ciedia) - 1), 2, 1), ""

Toolbar2.Enabled = IIf(Toolbar1.Buttons(1).Visible = True, True, False)
RS.Close
Set RS = Nothing

End Sub

Private Sub Form_Resize()

If Me.WindowState = 2 Then
    
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)

ElseIf Me.WindowState = 0 Then
    
    Frame2.Left = 45
    Frame1.Left = 45

End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""

End Sub

Private Sub Text1_Change(Index As Integer)

Select Case Index

Case 1, 2
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
           vaSpread1.Col = 1
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
              'Activar familia productos
              
              If M_TomInv.Check2.Value = 1 Then
                 
                 For j = i To 1 Step -1
                     
                     vaSpread1.Row = j: vaSpread1.Col = 1
                     If Trim(vaSpread1.text) = "" Then vaSpread1.RowHidden = False: Exit For
                 
                 Next j
              
              End If
           
           Else
              
              If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index, 1
    
    End If
'    vaSpread1_Click index, 0
    vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    
    vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

End Sub

Public Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim CodBOD As Long, nombod As String, fecemi As Date, pmp As Double, auxpmp As Double, stosis As Double
Dim RS As New ADODB.Recordset

CodBOD = Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, ""))
nombod = Trim(Combo1(0).text)
fecemi = Format(Date1(0).text, "dd/mm/yyyy")
TraerFechaCierre

Select Case Button.Index

Case 1 '-------> Graba
    
    '-------> Asignar concepto automatico
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), CodBOD, 9) Then
       
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), vg_codbod, 46) Then
    
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If

    '-------> Validar si existe ajuste inventario
    If CierrePeriodo(Format(Date1(0).text, "yyyymmdd"), CodBOD, 32) Then
    
       MsgBox "Existe ajuste Inventario, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    estpre = False
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.RowHidden = False
        vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = 0
        vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = 0
        
        '-------> validar si hay cambio de precio
        If vaSpread1.Font.Bold <> True Then
           
           vaSpread1.Col = 5: pmp = Round(IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text), 2)
           vaSpread1.Col = 8: auxpmp = Round(vaSpread1.text, 2)
           vaSpread1.Col = 9: stosis = Round(vaSpread1.text, vg_DCa)
           If pmp <> auxpmp And stosis > 0 Then estpre = True
        
        End If
    
    Next i
    
    precer = True
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 6
        If vaSpread1.TypeComboBoxCurSel = -1 And vaSpread1.Font.Bold <> True Then
           
           MsgBox "Falta seleccionar concepto...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        vaSpread1.Col = 5
        If Val(vaSpread1.Value) <= 0 And vaSpread1.Font.Bold <> True Then
        
           precer = False
           MsgBox "Falta ingresar precio...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
       
       End If
       
    Next i
    
    If estpre Then
       
       Toolbar1.Enabled = False
       Frame1.Enabled = False
       Frame2.Enabled = False
       Label1(4).Caption = "Se ha Realizo un ajuste de precio, este " & VgLinea & " ajuste debe ser autorizado por el monitor sgp"
       Frame5.Visible = True
    
    Else
       
       GrabarDatos
    
    End If
    
    '-------> INI: Mover estado a la tabla parametro toma inventario
    vg_db.Execute "update a_param set par_valor = '0' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
    '-------> FIN: Mover estado a la tabla parametro toma inventario

    '-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    Set RS = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(fecemi, "yyyymmdd") & ", '1'")
    If Not RS.EOF Then
    
       If RS(0) > 0 And Trim(RS(1)) <> "" Then
       
          RS.Close
          Set RS = Nothing
          
          MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
    
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    '-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
    
    
Case 3 '-------> Imprimir
    
    I_Ajuste CodBOD & "|" & nombod, CVDate(fecemi)

Case 6 '-------> Salir
    
    Me.Hide
    Unload Me

End Select

End Sub

Sub GrabarDatos()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset

fg_carga ""

Dim rutcli As String, tipdoc As String, NumDoc As Long, CodBOD   As Long, fecemi As Date, codser As Long, i As Long, canact As Double, z As Long, nombod As String, fecnum As Long, aumdes As Long
Dim numlin As Long, codmer As String, canmer As Double, canaux As Double, propon As Double, predoc As Double, auxpre As Double, descri As String, diablq As Date, Folio As Long, total As Double, ptotal As Double
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0

'-------> bloqueo barra comando principal
Toolbar1.Enabled = False
CodBOD = Val(fg_codigocbo(M_TomInv.Combo1, 0, 10, ""))
nombod = Trim(Combo1(0).text)
fecemi = Format(Date1(0).text, "dd/mm/yyyy")
fecnum = Format(Date1(0).text, "yyyymmdd")
rutcli = MuestraCasino(1)
tipdoc = "AI"

If RS4.State = 1 Then RS4.Close
RS4.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'-------> Ini : Consultar si exite ajuste de inventario

Set RS4 = vg_db.Execute("select top 1 tov_rutcli " & _
                     "From b_totventas " & _
                     "where tov_rutcli = '" & rutcli & "' " & _
                     "and   tov_codbod = " & CodBOD & " " & _
                     "and   tov_tipdoc = '" & tipdoc & "' " & _
                     "and   tov_fecemi = '" & Format(fecemi, "yyyymmdd") & "' " & _
                     "and   tov_estdoc not in ('A', 'P')")

If Not RS4.EOF Then

   vg_db.Execute "UPDATE b_totventas set tov_estdoc = 'A' where tov_rutcli = '" & rutcli & "' " & _
                                                         "and   tov_codbod = " & CodBOD & " " & _
                                                         "and   tov_tipdoc = '" & tipdoc & "' " & _
                                                         "and   tov_fecemi = '" & Format(fecemi, "yyyymmdd") & "' " & _
                                                         "and   tov_estdoc not in ('A', 'P')"

End If

RS4.Close
Set RS4 = Nothing

'-------> Fin : Consultar si existe ajuste de inventario

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT aju_codigo, aju_nombre, aju_tipo FROM a_tipoajuste WHERE aju_tipaju = 1 and aju_codigo in (3,4)", vg_db, adOpenStatic

If Not RS1.EOF Then
    
'    vg_db.BeginTrans
    
    Do While Not RS1.EOF
        
        codser = 0
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 7
            codser = RS1!aju_codigo
            
            If codser = Val(vaSpread1.text) And vaSpread1.Font.Bold <> True Then
                
                rutcli = MuestraCasino(1)
                aumdes = IIf(RS1!aju_tipo = "D", 0, 1)
                tipdoc = "AI"
paso:
                NumDoc = TraerCorrelativo(vg_codbod, "AI")
                vg_db.Execute "UPDATE b_parametros SET par_correlativo = " & NumDoc & " WHERE par_codbod = " & vg_codbod & " AND par_tipdoc = 'AI'"
                '-------> Encabezado
                                  
                vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                              "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & CodBOD & ", '" & Format(fecemi, "yyyymmdd") & "', " & aumdes & ", " & codser & ", 0, '', '', 0)"
                
                '-------> Detalle
                
                total = 0
                
                For z = 1 To vaSpread1.MaxRows
                    
                    vaSpread1.Row = z
                    vaSpread1.Col = 7
                    
                    If codser = Val(vaSpread1.text) Then
                        
                        numlin = z
                        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
                        vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.text))
                        vaSpread1.Col = 4: canmer = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
                        vaSpread1.Col = 5: predoc = Format(vaSpread1.text, fg_Pict(9, 2)) 'vg_DCa))
                        vaSpread1.Col = 8: auxpre = Format(vaSpread1.text, fg_Pict(9, 2)) 'vg_DCa))
                        canaux = IIf(canmer < 0, canmer * -1, canmer)
                        ptotal = canaux * predoc
                        total = total + ptotal
                        
                        vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_coding, dev_codmer, dev_canmin, dev_canmer, dev_porcen, dev_precos, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_codsec, dev_acepre) " & _
                                      "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", null, '" & codmer & "', " & canaux & ", " & canaux & ", 0, " & predoc & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', null, '1')"
                        'vaSpread1.Col = 1
                        'vg_db.Execute "UPDATE b_tomainv SET tin_stosis = tin_stosis+" & canmer & " WHERE tin_fectom = " & fecnum & " " & _
                                      "AND tin_codbod = " & CodBod & " AND tin_codpro = '" & Trim(vaSpread1.Text) & "'"
                        '------- Reemplaza precio promedio si es inventario inicial -----------
                        vaSpread1.Col = 5
                        propon = Round(vaSpread1.text, vg_DCa)
'                            vaSpread1.Col = 7
'                            If vaSpread1.text = "3" Then
                        
                        If predoc <> auxpre Then
                           
                           Dim pmp As Double, auxCanmer As Double, auxPropon As Double
                            
                           If RS2.State = 1 Then RS2.Close
                           RS2.CursorLocation = adUseClient
                           vg_db.CursorLocation = adUseClient
   
                           RS2.Open "SELECT pro_facing FROM b_productos WHERE pro_codigo = '" & codmer & "'", vg_db, adOpenStatic
                           'PMP Ingrediente
                           
                           If Not RS2.EOF Then
                              
                              If RS3.State = 1 Then RS3.Close
                              RS3.CursorLocation = adUseClient
                              vg_db.CursorLocation = adUseClient

                              RS3.Open "SELECT DISTINCT ppd_cencos FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND ppd_codpro = '" & codmer & "'", vg_db, adOpenStatic
                              
                              If RS3.EOF Then
                                 
                                 vg_db.Execute "INSERT INTO b_productospmpdia VALUES ('" & MuestraCasino(1) & "', '" & codmer & "', " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", " & propon & ", 0, " & propon & ", '" & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & "')"
                              
                              Else
                                 
                                 '------- Actuliza codigo compra y pedido de ultimo producto para ingrediente
                                 vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & propon & " WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND ppd_codpro = '" & codmer & "'"
                              
                              End If
                              RS3.Close
                              Set RS3 = Nothing
                              
                              If vg_tipbase = "1" Then
                                 
                                 vg_db.Execute "UPDATE b_contlistpreing a, b_productosing b SET a.cpi_codped = '" & codmer & "', a.cpi_codcom = '" & codmer & "' " & _
                                               "WHERE b.pri_coding = a.cpi_coding AND b.pri_codpro = '" & codmer & "' AND a.cpi_cencos = '" & MuestraCasino(1) & "'"
                              
                              Else
                                 
                                 vg_db.Execute "UPDATE b_contlistpreing SET b_contlistpreing.cpi_codped = '" & codmer & "', b_contlistpreing.cpi_codcom = '" & codmer & "' " & _
                                               "FROM b_contlistpreing a, b_productosing b WHERE b.pri_coding = a.cpi_coding AND b.pri_codpro = '" & codmer & "' AND a.cpi_cencos = '" & MuestraCasino(1) & "'"
                              
                              End If
                              
                              If RS3.State = 1 Then RS3.Close
                              RS3.CursorLocation = adUseClient
                              vg_db.CursorLocation = adUseClient

                              RS3.Open "SELECT DISTINCT pri_coding FROM b_productosing WHERE pri_codpro = '" & codmer & "'", vg_db, adOpenStatic
                              
                              If Not RS3.EOF Then
                                 
                                 If RS4.State = 1 Then RS4.Close
                                 RS4.CursorLocation = adUseClient
                                 vg_db.CursorLocation = adUseClient

                                 RS4.Open "SELECT Round(AVG(a.ppd_propon/c.pro_facing), 2) AS cosing " & _
                                          "FROM  b_productospmpdia a, b_productosing b, b_productos c " & _
                                          "WHERE b.pri_codpro = c.pro_codigo " & _
                                          "AND   c.pro_codigo = a.ppd_codpro " & _
                                          "AND   a.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                                          "AND   a.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " " & _
                                          "AND   a.ppd_propon > 0 AND b.pri_coding = '" & RS3!pri_coding & "'", vg_db, adOpenStatic
                                 
                                 If Not RS4.EOF Then
                                    
                                    vg_db.Execute "UPDATE b_contlistpreing SET cpi_feccos = " & Format(Date, "yyyymmdd") & ", cpi_precos = " & IIf(IsNull(RS4!cosing), 0, RS4!cosing) & " " & _
                                                  "WHERE cpi_cencos = '" & MuestraCasino(1) & "' AND cpi_coding = '" & RS3!pri_coding & "'"
                                 
                                 End If
                                 RS4.Close
                                 Set RS4 = Nothing
                              
                              End If
                              RS3.Close
                              Set RS3 = Nothing
                              
                              vaSpread1.Col = 1
                              vg_db.Execute "UPDATE b_tomainv SET tin_propon=" & propon & " WHERE tin_fectom = " & fecnum & " " & _
                                            "AND tin_codbod = " & CodBOD & " AND tin_codpro = '" & Trim(vaSpread1.text) & "'"
                           
                           End If
                           RS2.Close
                           Set RS2 = Nothing
                           
                           '-------> Actualizar tipo movimiento 3 corresponde cambio de ajuste y precio
                           vg_db.Execute "UPDATE b_detventas SET dev_acepre = " & IIf(canmer = 0, "2", "3") & " " & _
                                         "WHERE dev_rutcli = '" & rutcli & "' AND dev_tipdoc = '" & tipdoc & "' AND dev_numdoc = " & NumDoc & " AND dev_numlin = " & numlin & ""
                        
                        End If
                        vaSpread1.Col = 1
                        
                        '-------> Control de Stock ---------
                        ValidaBod CodBOD, Trim(LimpiaDato(codmer))
                        vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer+" & canmer & " " & _
                                      "WHERE bod_codpro = '" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod = " & vg_codbod & ""
                    
                    End If
                
                Next z
                
                '-------> Total
                vg_db.Execute "UPDATE b_totventas SET tov_totdoc = " & total & " WHERE tov_rutcli = '" & rutcli & "' " & _
                              "AND tov_tipdoc = 'AI' AND tov_numdoc = " & NumDoc & " AND tov_codbod = " & vg_codbod & ""
                
                Exit For
            
            End If
        
        Next i
        
        RS1.MoveNext
    
    Loop
    
    '-------> Actualizar b_productospmpdia stock
    If vg_tipbase = "1" Then
       
       vg_db.Execute "UPDATE b_productospmpdia INNER JOIN b_bodegas ON b_productospmpdia.ppd_codpro = b_bodegas.bod_codpro SET b_productospmpdia.ppd_saldo = b_bodegas.bod_canmer " & _
                     "WHERE b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & "  AND b_bodegas.bod_codbod = " & vg_codbod & ""
    
    Else

       vg_db.Execute "UPDATE b_productospmpdia SET b_productospmpdia.ppd_saldo = b.bod_canmer FROM b_productospmpdia a, b_bodegas b " & _
                     "WHERE a.ppd_codpro = b.bod_codpro AND b.bod_codbod = " & vg_codbod & " AND a.ppd_cencos = '" & MuestraCasino(1) & "' AND a.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & ""
    
    End If
'    vg_db.CommitTrans
    estgra = False

Else
    
    RS1.Close
    Set RS1 = Nothing
    '-------> desbloqueo barra comando principal
    Toolbar1.Enabled = True
    fg_descarga
    Exit Sub

End If
RS1.Close
Set RS1 = Nothing

'-------> desbloqueo barra comando principal
Toolbar1.Enabled = True
fg_descarga
vaSpread1.Row = -1
vaSpread1.Col = 5: vaSpread1.Lock = True
vaSpread1.Col = 6: vaSpread1.Lock = True
Gl_Ac_Botones Me, 5, 2, ""
I_Ajuste CodBOD & "|" & nombod, CVDate(fecemi)
Toolbar2.Enabled = False

Exit Sub
Man_Error:
vg_swpegreceta = 0
If Err = -2147467259 Then GoTo paso
If Err = 3034 Then
   'vg_db.RollbackTrans
   Exit Sub
End If

'vg_db.RollbackTrans
'-------> desbloqueo barra comando principal
Toolbar1.Enabled = True
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i As Long, z As Long, v_fecinv As Variant
Dim codtip  As Long
Dim lisnom  As String, liscod As String
Dim difer As Double, pmp As Double
Dim estfpro  As Boolean

Select Case Button.Index

Case 1 '-------> Agregar Producto
    
    vg_nombre = "": vg_codigo = ""
    v_fecinv = Format(Date1(0).text, "yyyymmdd")
    vg_left = Toolbar2.Width
    B_TabEst.LlenaDatos Trim(CStr(v_fecinv)), Trim(Str(vg_codbod)), "Productos", "ProInv1"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    Me.Refresh
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    Next i
    
    '-------> Traer productos pmp
    pmp = 0
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS1.Open "SELECT ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
             "FROM b_productospmpdia " & _
             "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
             "AND   ppd_codpro = '" & vg_codigo & "' " & _
             "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
             "GROUP BY ppd_propon " & _
             "HAVING (ppd_propon) > 0", vg_db, adOpenStatic
    If Not RS1.EOF Then pmp = IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon)
    RS1.Close
    Set RS1 = Nothing
    '-------> Insertar producto en tabla temporal
'    vg_db.Execute "INSERT INTO " & aAp & " VALUES (" & v_fecinv & ", " & vg_codbod & ", '" & vg_codigo & "', " & pmp & ", 0 ,0)"
    vg_db.Execute "INSERT INTO " & aAp & " SELECT DISTINCT " & v_fecinv & " AS tin_fectom, " & vg_codbod & " AS tin_codbod, '" & vg_codigo & "' AS tin_codpro, " & pmp & " AS ppd_propon, " & pmp & " AS aux_propon, tin_stosis, tin_stofis FROM b_tomainv WHERE tin_fectom = " & v_fecinv & " AND tin_codbod = " & vg_codbod & " AND tin_codpro = '" & vg_codigo & "'"
    '--------- Muestra las diferencias del ultimo inventario -----------
    opfampro = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), " ORDER BY b.pro_nombre", " ORDER BY b.pro_ctacon, b.pro_codtip, b.pro_nombre")
    estfpro = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT a.tin_codpro, b.pro_nombre, b.pro_codtip, b.pro_ctacon, a.ppd_propon, a.aux_propon, c.uni_nombre, a.tin_stosis, a.tin_stofis " & _
             "FROM " & aAp & " a, b_productos b, a_unidad c " & _
             "WHERE a.tin_codpro = b.pro_codigo " & _
             "AND   b.pro_coduni = c.uni_codigo " & _
             "AND   a.tin_fectom = " & v_fecinv & " " & _
             "AND   a.tin_codbod = " & vg_codbod & " " & _
             " " & opfampro & "", vg_db, adOpenStatic
    i = 1
    vaSpread1.Visible = False
    vaSpread1.MaxRows = 0
    
    Do While Not RS1.EOF
       
       vaSpread1.MaxRows = i
       vaSpread1.Row = vaSpread1.MaxRows
       
       If RS1!pro_codtip <> codtip And estfpro Then
          
          vaSpread1.Col = 1: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 2: vaSpread1.Font.Bold = True: vaSpread1.text = fg_BuscaenArbol(RS1!pro_codtip, "a_tipopro", "tip_codigo")
          vaSpread1.Col = 3: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 4: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 5: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 6: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 7: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 8: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          vaSpread1.Col = 9: vaSpread1.Font.Bold = True: vaSpread1.text = ""
          codtip = RS1!pro_codtip
          i = i + 1: vaSpread1.MaxRows = i: vaSpread1.Row = i
       
       End If
       vaSpread1.Font.Bold = False
       difer = Format(RS1!tin_stofis - RS1!tin_stosis, fg_Pict(9, vg_DCa))
       vaSpread1.Col = 1: vaSpread1.text = RS1!tin_codpro
       vaSpread1.Col = 2: vaSpread1.text = RS1!pro_nombre
       vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nombre
       vaSpread1.Col = 4: vaSpread1.text = Format(difer, fg_Pict(9, vg_DCa)): vaSpread1.ForeColor = IIf(difer < 0, RGB(255, 0, 0), RGB(0, 0, 0))
       vaSpread1.Col = 5: vaSpread1.Lock = False: vaSpread1.text = Format(IIf(IsNull(RS1!ppd_propon), 0, RS1!ppd_propon), fg_Pict(9, 2)) ' vg_DCa))
       vaSpread1.Col = 8: vaSpread1.text = Format(IIf(IsNull(RS1!aux_propon), 0, RS1!aux_propon), fg_Pict(9, 2)) 'vg_DCa))
       vaSpread1.Col = 9: vaSpread1.text = Format(RS1!tin_stosis, fg_Pict(9, vg_DCa))
       lisnom = "": liscod = ""
       aju_tipo = IIf(difer < 0, "D", "A")
       
       If RS2.State = 1 Then RS2.Close
       RS2.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

       RS2.Open "SELECT aju_codigo, aju_nombre FROM a_tipoajuste WHERE aju_tipaju = 1 and aju_tipo = '" & aju_tipo & "'", vg_db, adOpenStatic
       
       Do While Not RS2.EOF
          
          vaSpread1.Col = 6: lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(RS2!aju_nombre)
          vaSpread1.Col = 7: liscod = liscod & IIf(liscod <> "", Chr$(9), "") & RS2!aju_codigo
          If RS1!ppd_propon = 0 And RS1!tin_stosis = 0 And liscod = "3" Then Exit Do
          RS2.MoveNext
       
       Loop
       RS2.Close
       Set RS2 = Nothing
       
       vaSpread1.Col = 6: vaSpread1.TypeComboBoxList = lisnom
       vaSpread1.Col = 7: vaSpread1.TypeComboBoxList = liscod
       
       If RS1!tin_stosis = 0 Or RS1!ppd_propon = 0 Then
          
          For z = 0 To vaSpread1.TypeComboBoxCount
              
              vaSpread1.TypeComboBoxCurSel = z
'             If RS1!ppd_propon = 0 Then codaux = z: vaSpread1.Col = 5: vaSpread1.Lock = False: Exit For
              If RS1!ppd_propon = 0 Then codaux = z: Exit For
              codaux = -1
          
          Next z
          
          vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = codaux
          vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = codaux
       
       End If
       RS1.MoveNext: i = i + 1
    
    Loop
    RS1.Close
    Set RS1 = Nothing
    vaSpread1.Visible = True

Case 2
    
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If Trim(vaSpread1.text) = "" Then
    
       MsgBox "No puede eliminar familia producto...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    vaSpread1.Col = 4
    If Val(vaSpread1.text) <> 0 Then
    
       MsgBox "No puede eliminar producto con diferencia...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    
    End If
    
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '-------> Eliminar dato tabla temporal
    vaSpread1.Col = 1
    If Trim(vaSpread1.text) <> "" Then vg_db.Execute "DELETE " & aAp & " FROM " & aAp & " WHERE tin_codpro = '" & Trim(vaSpread1.text) & "'"
    i = vaSpread1.Row
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    
    If (vaSpread1.ActiveRow - 1) >= 0 Then
        vaSpread1.Row = i: vaSpread1.Col = 1
        If Trim(vaSpread1.text) <> "" Then Exit Sub
        vaSpread1.Row = IIf(vaSpread1.ActiveRow - 1 = 0, 1, (i - 1))
        vaSpread1.Col = 1
        
        If Trim(vaSpread1.text) = "" Then
        
           vaSpread1.DeleteRows (vaSpread1.Row), 1
           vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        
        End If
        
    End If
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus

End Select

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)

If est Then Exit Sub

Select Case Col

Case 6
    
    Dim indice As Long
    vaSpread1.Row = Row
    vaSpread1.Col = 6: indice = vaSpread1.TypeComboBoxCurSel
    vaSpread1.Col = 7: vaSpread1.TypeComboBoxCurSel = indice
    vaSpread1.Col = 1
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS1.Open "SELECT b.ppd_propon FROM b_productos a, b_productospmpdia b " & _
             "WHERE a.pro_codigo = b.ppd_codpro " & _
             "AND   a.pro_codigo = '" & Trim(vaSpread1.text) & "' " & _
             "AND   b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
             "AND   b.ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & "", vg_db, adOpenStatic
    
    If Not RS1.EOF Then
       
       vaSpread1.Col = 7
       
       If RS1!ppd_propon = 0 Then
            
            vaSpread1.Col = 5: vaSpread1.Lock = False
            vaSpread1.SetActiveCell 5, Row - 1
       
       Else
          
          vaSpread1.Col = 5: vaSpread1.Lock = True
          vaSpread1.text = RS1!ppd_propon
       
       End If
    
    End If
    RS1.Close
    Set RS1 = Nothing

End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

If est Or vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = Row

Select Case Col

Case 5
    
    '-------> Actualizar Tabla temporal
    Dim codigo As String, pmp As Double, auxpmp As Double, stosis As Double
    vaSpread1.Col = 1: codigo = Trim(vaSpread1.text)
    vaSpread1.Col = 5: pmp = Round(vaSpread1.text, 2)
    vaSpread1.Col = 8: auxpmp = Round(vaSpread1.text, 2)
    vaSpread1.Col = 9: stosis = Round(vaSpread1.text, vg_DCa)
'    vaSpread1.Col = 5
    vg_db.Execute "UPDATE " & aAp & " SET ppd_propon = " & pmp & ", aux_propon = " & auxpmp & " WHERE tin_fectom = " & Format(Date1(0).text, "yyyymmdd") & " AND tin_codbod = " & vg_codbod & " AND tin_codpro = '" & codigo & "'"
'    If pmp <> auxpmp And pmp > 0 Then estpre = True
'    If pmp <> auxpmp And pmp > 0 Then estpre = True

End Select

End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

If Row = 0 Then Exit Sub

Select Case Col

Case 6
    
    vaSpread1.Row = Row: vaSpread1.Col = Col
    If vaSpread1.ColWidth(Col) > (vaSpread1.MaxTextCellWidth - 2) Then Exit Sub
    TipWidth = vaSpread1.MaxTextColWidth(Col)
    ShowTip = True
    MultiLine = 2
    TipText = vaSpread1.text

End Select

End Sub
