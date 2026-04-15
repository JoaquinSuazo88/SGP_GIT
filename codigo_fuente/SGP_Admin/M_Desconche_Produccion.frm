VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Desconche_Produccion 
   Caption         =   "Mantenedor Costo Mermas"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8685
   ScaleWidth      =   13305
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   12855
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   720
         TabIndex        =   11
         Top             =   6120
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   930
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   5655
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   12375
         _Version        =   393216
         _ExtentX        =   21828
         _ExtentY        =   9975
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
         MaxCols         =   10
         SpreadDesigner  =   "M_Desconche_Produccion.frx":0000
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1676
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
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   0
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
         ThreeDOnFocusInvert=   0   'False
         ThreeDFrameColor=   -2147483637
         Appearance      =   1
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   -1  'True
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime FpFecDesde 
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   600
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
         Left            =   6000
         TabIndex        =   10
         Top             =   600
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
      Begin VB.Label Label1 
         Caption         =   "Fecha Hasta"
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
         Left            =   4680
         TabIndex        =   9
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Desde"
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
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   1215
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
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   6735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2520
         Picture         =   "M_Desconche_Produccion.frx":1B6F
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Segmento"
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
         TabIndex        =   1
         Top             =   360
         Width           =   855
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
         Left            =   3165
         TabIndex        =   6
         Top             =   285
         Width           =   6735
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   13305
      _ExtentX        =   23469
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Desconche_Produccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private NomFor As String
Private BtnX   As Variant
Dim OpGr       As Boolean

Public lc_Aux  As String
Dim modo       As String
Dim Est        As Boolean
Dim MsgTitulo  As String

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
    Call fg_carga("")
    MsgTitulo = "Mantención Costos de Mermas"
    Me.HelpContextID = vg_OpcM
    Call fg_centra(Me)
    Let Me.Height = 9270
    Let Me.Width = 13545
    Est = True
    modo = ""
    
    Let FpFecDesde.text = Format(Date, "dd/mm/yyyy")
    Let FpFecHasta.text = Format(Date, "dd/mm/yyyy")
    Let vaSpread1.MaxRows = 0
       
    Gl_Mo_Botones Me, 1
    Gl_Ac_Botones Me, 1, 1, modo
    
    Call fg_descarga
    
    OpGr = False

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

MoverDatosGrilla

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
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

On Error GoTo Man_Error

If IsDate(FpFecHasta.text) = False Then Exit Sub

MoverDatosGrilla

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
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

Private Sub fpLongInteger1_Change(Index As Integer)
            
On Error GoTo Man_Error
            
Dim RS As New ADODB.Recordset

vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_s_segmento 1, " & Val(fpLongInteger1(0).Value) & ",''")
fpayuda(1).Caption = ""

If Not RS.EOF Then

   fpayuda(1).Caption = Trim(RS!seg_nombre)

End If

RS.Close
Set RS = Nothing

MoverDatosGrilla

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cliente"

End Sub

Sub MoverDatosGrilla()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_Sel_ExisteCostoMermas " & Val(fpLongInteger1(0).Value) & ", '" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "'")

If RS.EOF Then

    RS.Close
    Set RS = Nothing

    Exit Sub
    
End If
RS.Close
Set RS = Nothing

OpGr = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
            
Set RS = vg_db.Execute("sgpadm_Sel_CostoMermas " & Val(fpLongInteger1(0).Value) & ", '" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "'")

If Not RS.EOF Then

    Do While Not RS.EOF = True
       
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       vaSpread1.Row = vaSpread1.MaxRows
             
       vaSpread1.Col = 1
       vaSpread1.text = (RS!CodServicio)
       
       vaSpread1.Col = 2
       vaSpread1.text = Trim(RS!nomservicio)
       
       vaSpread1.Col = 3
       vaSpread1.text = RS!Costo_Desconche
       
       vaSpread1.Col = 4
       vaSpread1.text = RS!Costo_Produccion
       
       vaSpread1.Col = 5
       vaSpread1.text = RS!Costo_Pan
       
       vaSpread1.Col = 6
       vaSpread1.text = RS!Fecha_Creacion
       
       vaSpread1.Col = 7
       vaSpread1.text = RS!Fecha_Modificacion
       
       vaSpread1.Col = 8
       vaSpread1.text = RS!Activo
       
       vaSpread1.Col = 9
       vaSpread1.text = 0
       
       vaSpread1.Col = 10
       vaSpread1.text = ""
       
       RS.MoveNext
    
    Loop

End If

RS.Close
Set RS = Nothing

OpGr = False

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = ""
    vg_codigo = ""
    
    B_TabEst.LlenaDatos "a_segmento", "seg_", "Segmento", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(1).Caption = Trim(vg_nombre)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Or vaSpread1.MaxRows = 0 Then

   Text1(1).text = ""
   
   Exit Sub

End If
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 9
    vaSpread1.text = 0
    
Next

Select Case Index

Case 1
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 1, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 9
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 9
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 9
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 9
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 9
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
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 9
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim i          As Long
Dim Servicio   As Long
Dim Desconche  As Double
Dim Produccion As Double
Dim Pan        As Double
Dim Activo     As String
Dim Actualiza  As String

Dim MyBuffer   As String

Select Case Button.Index

    Case 1

        If Val(fpLongInteger1(0).Value) = 0 Or Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
        
           MsgBox "Debe seleccionar dato del encabezado como segmento, fecha desde y hasta ...", vbExclamation + vbOKOnly, MsgTitulo
           
           Exit Sub
        
        End If

        vaSpread1.MaxRows = 0
        OpGr = True
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
            
        Set RS = vg_db.Execute("sgpadm_Sel_CostoMermas " & Val(fpLongInteger1(0).Value) & ", '" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "'")

        If Not RS.EOF Then

            Do While Not RS.EOF = True
       
                vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                vaSpread1.Row = vaSpread1.MaxRows
             
                vaSpread1.Col = 1
                vaSpread1.text = (RS!CodServicio)
       
                vaSpread1.Col = 2
                vaSpread1.text = Trim(RS!nomservicio)
       
                vaSpread1.Col = 3
                vaSpread1.text = Format(RS!Costo_Desconche, fg_Pict(6, 2))
       
                vaSpread1.Col = 4
                vaSpread1.text = Format(RS!Costo_Produccion, fg_Pict(6, 2))
       
                vaSpread1.Col = 5
                vaSpread1.text = Format(RS!Costo_Pan, fg_Pict(6, 2))
       
                vaSpread1.Col = 6
                vaSpread1.text = RS!Fecha_Creacion
       
                vaSpread1.Col = 7
                vaSpread1.text = RS!Fecha_Modificacion
       
                vaSpread1.Col = 8
                vaSpread1.text = RS!Activo
       
                vaSpread1.Col = 9
                vaSpread1.text = 0
       
                vaSpread1.Col = 10
                vaSpread1.text = ""
       
                RS.MoveNext
    
            Loop

        End If

        RS.Close
        Set RS = Nothing
    
        OpGr = False
        
        modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
        Frame1.Enabled = False
    
    Case 3
        
        If Val(fpLongInteger1(0).Value) = 0 Or Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
        
           MsgBox "Debe seleccionar dato del encabezado...", vbExclamation + vbOKOnly, MsgTitulo
           
           Exit Sub
        
        End If
                
        If vaSpread1.MaxRows = 0 Then
        
           MsgBox "No existe información detalle de la grilla...", vbExclamation + vbOKOnly, MsgTitulo
           
           Exit Sub
        
        End If
        
        modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
        Frame1.Enabled = False
    
    Case 5
        
        If Val(fpLongInteger1(0).Value) = 0 Or Trim(FpFecDesde.text) = "" Or Trim(FpFecHasta.text) = "" Then
           
           Exit Sub
           
        End If
        
        If Not Est < 1 Then
           
           MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        If vaSpread1.MaxRows = 0 Then
        
           MsgBox "No existe información detalle de la grilla...", vbExclamation + vbOKOnly, MsgTitulo
           
           Exit Sub
        
        End If
        
        If MsgBox("Desactiva registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
        
           Exit Sub
        
        End If
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Set RS = vg_db.Execute("sgpadm_Del_CostoMermas " & Val(fpLongInteger1(0).Value) & ", '" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "', '" & vg_NUsr & "'")

        If Not RS.EOF Then
      
           If Trim(RS(0)) <> "" Then
                    
              'registrar Log sistema desactivar el item
              Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), Me.HelpContextID, "", "", "")
      
              OpGr = True
              
              For i = 1 To vaSpread1.MaxRows
              
                  vaSpread1.Row = i
                  vaSpread1.Col = 8
                  vaSpread1.text = "0"
                  
              Next i
                            
              MoverDatosGrilla
        
              OpGr = False
              
              
              MsgBox "Registro desactivado exitosamente", vbInformation + vbOKOnly, MsgTitulo
           
           Else
         
              'registrar Log sistema error desactivar el item
               Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
          
               MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
                        
           End If
      
        End If
        RS.Close
        Set RS = Nothing
        
        modo = ""
        
        Gl_Ac_Botones Me, 1, 1, modo
    
    Case 7
        
        MoverDatosGrilla
    
    Case 10
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        MoverDatosGrilla
        
        modo = ""
        
        Gl_Ac_Botones Me, 1, 1, modo
               
        Frame1.Enabled = True
    
    Case 12
        
        fg_carga ""
        
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<GrabaCostoMerma>"

        For i = 1 To vaSpread1.MaxRows
            
            DoEvents
            
            vaSpread1.Row = i
            
            vaSpread1.Col = 1
            Servicio = vaSpread1.text
            
            vaSpread1.Col = 3
            Desconche = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
            
            vaSpread1.Col = 4
            Produccion = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
            
            vaSpread1.Col = 5
            Pan = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
            
            vaSpread1.Col = 8
            Activo = vaSpread1.text
            
            vaSpread1.Col = 10
            Actualiza = vaSpread1.text
            
            MyBuffer = MyBuffer & " <Detalle"
            MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
            MyBuffer = MyBuffer & " Desco = " & Chr(34) & Desconche & Chr(34)
            MyBuffer = MyBuffer & " Pro = " & Chr(34) & Produccion & Chr(34)
            MyBuffer = MyBuffer & " Pan = " & Chr(34) & Pan & Chr(34)
            MyBuffer = MyBuffer & " Act = " & Chr(34) & Activo & Chr(34)
            MyBuffer = MyBuffer & " Actua = " & Chr(34) & Actualiza & Chr(34)
            MyBuffer = MyBuffer & "/>"

        Next i
        
        MyBuffer = MyBuffer & "</GrabaCostoMerma>"

        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS = vg_db.Execute("sgpadm_Ins_XmlCostoMermas '" & MyBuffer & "', " & Val(fpLongInteger1(0).Value) & ", '" & Format(FpFecDesde.text, "yyyymmdd") & "', '" & Format(FpFecHasta.text, "yyyymmdd") & "', '" & vg_NUsr & "'")
        If Not RS.EOF Then
                           
           If RS(0) > 0 Then
                  
              MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
               
           Else
           
              MsgBox "El dato se grabo exitosamente ", vbInformation + vbOKOnly, MsgTitulo
           
           End If
            
        End If
        RS.Close
        Set RS = Nothing
        
        MoverDatosGrilla
        
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
        
        Frame1.Enabled = True
        
        fg_descarga
    
    Case 15
               
        I_ParamCostoMermas
    
    Case 18
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
If Err = -2147467259 Or 2147217900 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

If (Col <> 8) Or Row = 0 Or OpGr Then Exit Sub

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo


vaSpread1.Row = Row
vaSpread1.Col = 10
vaSpread1.text = "1"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

vaSpread1.Row = Row
vaSpread1.Col = 10
vaSpread1.text = "1"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
