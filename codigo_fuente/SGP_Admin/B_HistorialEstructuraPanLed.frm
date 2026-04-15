VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_HistorialEstructuraPanLed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Historial Estructura Servicio"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11175
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   11175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10215
         _Version        =   393216
         _ExtentX        =   18018
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
         MaxCols         =   6
         SpreadDesigner  =   "B_HistorialEstructuraPanLed.frx":0000
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
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_HistorialEstructuraPanLed.frx":18D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_HistorialEstructuraPanLed.frx":1BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_HistorialEstructuraPanLed.frx":1F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "B_HistorialEstructuraPanLed.frx":2062
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   4320
      Left            =   10635
      TabIndex        =   2
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   7620
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Confirmar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "B_HistorialEstructuraPanLed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
 
fg_centra Me
'Me.Left = vg_left

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub LlenarHistorial(Ceco As String)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("sgpadm_Sel_HistorialEstructuraPanLed '" & LimpiaDato(Ceco) & "'")
If Not RS.EOF Then

   Do While Not RS.EOF
   
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      
      vaSpread1.Col = 1
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS!Ceco
      
      
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS!CecoDesc
      
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS!Regimen
      
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS!RegimenDesc
      
      vaSpread1.Col = 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS!Servicio
      
      vaSpread1.Col = 6
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS!ServicioDesc
      
      RS.MoveNext
   
   Loop
   
End If

RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error
 
Select Case Button.Index

Case 1
    
    MoverDatos
    
Case 2

    vg_codcasino = ""
    vg_codregimen = 0
    vg_codservicio = 0
    
End Select

Cerrar

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub Cerrar()

On Error GoTo Man_Error

Me.Hide
Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Sub MoverDatos()

On Error GoTo Man_Error

Screen.MousePointer = 11
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow

vaSpread1.Col = 1
vg_codcasino = vaSpread1.text
    
vaSpread1.Col = 3
vg_codregimen = vaSpread1.text
    
vaSpread1.Col = 5
vg_codservicio = vaSpread1.text

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

MoverDatos
Cerrar

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
