VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_RCDiar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre Díario"
   ClientHeight    =   5085
   ClientLeft      =   6090
   ClientTop       =   1620
   ClientWidth     =   6180
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6180
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   4545
      Left            =   90
      TabIndex        =   0
      Top             =   495
      Width           =   6015
      Begin VB.CommandButton Command1 
         Caption         =   "Desactivar proceso de inventario"
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
         Left            =   1440
         TabIndex        =   10
         Top             =   4200
         Width           =   3135
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2325
         Left            =   210
         TabIndex        =   1
         Top             =   1140
         Width           =   5595
         _Version        =   393216
         _ExtentX        =   9869
         _ExtentY        =   4101
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   14
         MaxRows         =   6
         ScrollBars      =   0
         SpreadDesigner  =   "M_RCDiar.frx":0000
         UserResize      =   0
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1635
         TabIndex        =   2
         Top             =   420
         Width           =   945
         _Version        =   196608
         _ExtentX        =   1658
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
         ButtonStyle     =   1
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
         Text            =   "11/2024"
         DateCalcMethod  =   4
         DateTimeFormat  =   5
         UserDefinedFormat=   "mm/yyyy"
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
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   3795
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   3255
         Top             =   585
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Día Cerrado y enviado"
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
         Left            =   3600
         TabIndex        =   9
         Top             =   570
         Width           =   1935
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
         Left            =   210
         TabIndex        =   8
         Top             =   3525
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Día Cerrado y no enviado"
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
         Left            =   3600
         TabIndex        =   6
         Top             =   840
         Width           =   2205
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   3255
         Top             =   870
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Día Habilitado"
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
         Left            =   3615
         TabIndex        =   5
         Top             =   240
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   3255
         Top             =   270
         Width           =   300
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
         Index           =   0
         Left            =   210
         TabIndex        =   3
         Top             =   495
         Width           =   1230
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_RCDiar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim rst As New ADODB.Recordset
Dim modo As String, Fecha As String, MsgTitulo As String
Dim est As Boolean

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset

'If est Then Exit Sub

If MsgBox("Esta Seguro desactivar el proceso inventario... ?? ", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

'-------> INI: Mover estado a la tabla parametro toma inventario
vg_db.Execute "update a_param set par_valor = '0' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
'-------> FIN: Mover estado a la tabla parametro toma inventario

''-------> INI : Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado
'If RS1.State = 1 Then RS1.Close
'RS1.CursorLocation = adUseClient
'vg_db.CursorLocation = adUseClient
'
'Set RS1 = vg_db.Execute("sgp_Upd_ValidarInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", '1'")
'If Not RS1.EOF Then
'
'   If RS1(0) > 0 And Trim(RS1(1)) <> "" Then
'
'      RS1.Close
'      Set RS1 = Nothing
'
'      MsgBox "Existe error grabar inventario calendarizado..", vbExclamation + vbOKOnly, MsgTitulo
'      Exit Sub
'
'    End If
'
'End If
'RS1.Close
'Set RS1 = Nothing
''-------> FIN : Validar inventario calendarizado Validar inventario calendarizado toma inv & ajuste actualizar cambio de estado

Command1.Enabled = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 5520
Me.Width = 6270
fg_centra Me
modo = "M"
MsgTitulo = "Cierre Diario"
est = True
Gl_Mo_Botones Me, 10
Toolbar1.Buttons(1).ToolTipText = "Cerrar Día"
Toolbar1.Buttons(2).ToolTipText = "Reabrir Día"
Gl_Ac_Botones Me, 10, 6, modo
fpDateTime1.DateTimeFormat = UserDefined
fpDateTime1.UserDefinedFormat = "mm/yyyy"
fpDateTime1.text = Format(Date, "mm/yyyy")
ArmarCalendario
ActivarBotonCierreInventario

est = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Sub ActivarBotonCierreInventario()

On Error GoTo Man_Error

'Activar boton para cerrar inventario
Command1.Visible = True
Command1.Enabled = False

If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 42) And _
   ValidaPCServidor = True Then
        
   Command1.Visible = True
   Command1.Enabled = True
          
End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Sub ArmarCalendario()

On Error GoTo Man_Error

Dim sql1 As String
'-------> Armar calendario
Dim i As Long, j As Long, nrosem As Long, diafin As Long, indexit As Boolean
With vaSpread1
    .TextTip = 2 'SS_TEXTTIP_FLOATINGFOCUSONLY
    ' Control displays text tips after 250 milliseconds
    .TextTipDelay = 0
    .Row = -1: .Col = -1:
    .BackColor = &H8000000F
    diafin = fg_mes(Format(fpDateTime1.text, "mm") & Format(fpDateTime1.text, "yyyy"))
    nrosem = 1
    .Visible = False
    For i = 1 To 6
        For j = 1 To 14
            .Row = i
            .Col = j
            .text = ""
        Next j
    Next i
    For i = 1 To diafin
        Select Case fg_Dia(Format(fpDateTime1.text, "yyyymm") & fg_pone_cero(i, 2))
        Case 1
            .Row = nrosem
            .Col = 7 'domingo
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
            nrosem = nrosem + 1
        Case 2
            .Row = nrosem
            .Col = 1 'lunes
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 3
            .Row = nrosem
            .Col = 2 'martes
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 4
            .Row = nrosem
            .Col = 3 'miercoles
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 5
            .Row = nrosem
            .Col = 4 'jueves
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 6
            .Row = nrosem
            .Col = 5 'viernes
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        Case 7
            .Row = nrosem
            .Col = 6 'sabado
            .BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignCenter
            .text = CStr(i)
        End Select
        If i = 1 Then
           Toolbar1.Buttons(1).Visible = IIf(.BackColor = Shape1(1).FillColor, False, True)
           Toolbar1.Buttons(2).Visible = IIf(.BackColor = Shape1(1).FillColor, True, False)
        End If
    Next i
    .RetainSelBlock = False
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    sql1 = IIf(vg_tipbase = "1", " val(format(fecha, 'yyyymm')) ", " substring(CONVERT(varchar(10), fecha,112),1,6) ")
    RS1.Open "SELECT DISTINCT fecha, estenv, fecsub FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND " & sql1 & " = '" & Format(fpDateTime1.text, "yyyymm") & "' ORDER BY fecha", vg_db, adOpenStatic
    indexit = False
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          indexit = False
          For j = 1 To 6
              For i = 1 To 7
                  .Row = j
                  .Col = i
                  If Val(.text) = Val(Mid(RS1!Fecha, 1, 2)) And RS1!Fecha <= (CDate(vg_ciedia) - 1) Then
                     .BackColor = IIf(RS1!estenv = "0", Shape1(1).FillColor, Shape1(2).FillColor)
                     .Col = i + 7
                     .text = IIf(IsNull(RS1!fecsub) Or Trim(RS1!fecsub) = "", "", Trim(RS1!fecsub))
                     indexit = True: Exit For
                  End If
              Next i
              If indexit Then Exit For
          Next j
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
    .Visible = True
End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpDateTime1_Change()

On Error GoTo Man_Error

ArmarCalendario

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS           As New ADODB.Recordset
Dim sql1         As String
Dim sql2         As String
Dim vec_tipact() As Variant
Dim i            As Long
Dim vRet         As Variant
Dim EstCon       As Boolean

If Button.Index = 1 Or Button.Index = 2 Then

    '-------> Validar PC Servidor esta en blanco nombre de maquina
    If ValidaPCServidorVacio = False Then
    
        If MsgBox("Desea que este PC realice Cierre Diario ?? ", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
        Dim sEquipo As String * 255
        GetComputerName sEquipo, 255
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'SvrAppCont'")
        If RS.EOF Then
   
           vg_db.Execute ("sgp_Ins_Param 'SvrAppCont', 'Identifica PC Servidor.', 'C', '" & Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1) & "', '" & MuestraCasino(1) & "'")

        Else
        
           vg_db.Execute ("sgp_Upd_Param 1, '" & MuestraCasino(1) & "',  'SvrAppCont', '', '',  '" + Left$(sEquipo, InStr(sEquipo, vbNullChar) - 1) + "'")
        
        End If
        RS.Close
        Set RS = Nothing
        
    End If
    
    '------->Validar PC Servidor
    If ValidaPCServidor = False Then
        
        MsgBox "Debe realizar el cierre diario en el computador configurado como Servidor...", vbExclamation + vbOKOnly, MsgTitulo
        Exit Sub
        
    End If

End If
                   
Select Case Button.Index

Case 1 '-------> Cerrar día
    
    Me.Label1(1).Caption = ""
    Me.Bar1(0).Visible = False
    '-------> Validar que no existan salidas pedientes
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    If vaSpread1.BackColor = -2147483633 Then Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(vg_ciedia) <> CDate(fg_pone_cero(vaSpread1.text, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy"))) Then MsgBox "Día no corresponde al cierre diario", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 0) Then
       
       MsgBox "No ha cerrado periodo anterior...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 14) Then
    
       MsgBox "Existen documentos pendientes, en la salida producción. Debe cerrar las salidas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 46) Then
    
       MsgBox "Existen documentos pendientes, en la ventas servicios especiales. Debe cerrar las ventas servicios especiales", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 15) Then
    
       MsgBox "Existen ventas cafeteria, sin cerrar", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 27) Then
       
       If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 8) Then
       
          MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
       End If
       
    End If
       
    'Validar si existe un inventario en proceso 20201001
    If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 38) Then
        
       MsgBox "Existe un inventario en proceso...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
       
    '-------> Validar Raciones Servicio Principales
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 43) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) Then
    
       Exit Sub
    
    End If
    
    '-------> Validar Raciones no Vendidas o mermas por preparación
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 47) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) Then
    
       Exit Sub
    
    End If
    
    '-------> Validar Raciones no Vendidas Desconche
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 50) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) Then
    
       Exit Sub
    
    End If
    
    '-------> Validar Salida producciónes Servicio Principales, si hay raciones en servicio
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 51) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) Then
    
       Exit Sub
    
    End If
    
    '-------> Validar Raciones Vendida Principales
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 44) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) Then
    
       Exit Sub
    
    End If
         
         
    '-------> Validar actividades diarias
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT COUNT(*) as NREG FROM b_casinotipoactividades WHERE cta_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    
    If Not RS.EOF Then
       
       ReDim vec_tipact(RS!nreg + 2, 1)
    
    End If
    RS.Close: Set RS = Nothing
    
    For i = 1 To UBound(vec_tipact)
        
        vec_tipact(i, 1) = 0
    
    Next i
    
    i = 1
         
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
   
    RS.Open "SELECT * FROM b_casinotipoactividades WHERE cta_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          vec_tipact(i, 1) = RS!cta_tipact
          RS.MoveNext: i = i + 1
       
       Loop
    
    End If
    RS.Close: Set RS = Nothing
    
    If Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 29) Then
       
       For i = 1 To UBound(vec_tipact)
           
           If vec_tipact(i, 1) > 0 Then
           
           If vec_tipact(i, 1) = 1 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 17) Then MsgBox "No se han ingresado documentos de proveedores para el periodo " + vg_ciedia + VgLinea + VgLinea + Space(10) + "Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 2 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 18) Then Exit Sub ': MsgBox "salidas a producción", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 3 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 19) Then MsgBox "No existen devoluciones a bodega para el periodo " + vg_ciedia + VgLinea + VgLinea + Space(10) + "Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 4 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 20) Then MsgBox "No existen mermas para el periodo " + vg_ciedia + VgLinea + VgLinea + Space(10) + "Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 5 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 21) Then Exit Sub
           
           ElseIf vec_tipact(i, 1) = 6 Then

'              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 22) Then Exit Sub ': MsgBox "control de raciones", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 36) Then Exit Sub
           
           ElseIf vec_tipact(i, 1) = 7 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 23) Then MsgBox "No existen registros de venta cafeteria para el periodo " + vg_ciedia + VgLinea + VgLinea + Space(29) + "Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 8 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 24) Then Exit Sub 'MsgBox "venta servicios de contrato", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 9 Then
              
              If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 25) Then MsgBox "No existe venta directa para el periodo " + vg_ciedia + VgLinea + VgLinea + Space(15) + "Proceso Cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           ElseIf vec_tipact(i, 1) = 10 Then
              
              If CierrePeriodo(vg_ciedia, vg_codbod, 26) Then
                 
                 vg_invrot = "1"
                 M_TomInv.Show 1
                 
                 If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 8) Or CierrePeriodo(vg_ciedia, vg_codbod, 26) Then
                   
                   Exit Sub
                 
                 End If
              
              End If
           
           End If
           
           End If
      
      Next i
    
    End If
    
    '-------> Validar ajuste precios venta
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 37) Then
    
'       Exit Sub
    
    End If
    
     
    If MsgBox("Para ejecutar este proceso, todos los otros equipos con SGP no deben estar utilizando ninguna funcionalidad del sistema.                       ¿Está seguro de cerrar el día?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    '-------> enviar mensaje inventario calendarizado
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
        
    EstCon = False
    
    Set RS = vg_db.Execute("sgp_Sel_EnviarMensajeInventarioCalendarizado '" & MuestraCasino(1) & "', " & Format(vg_ciedia, "yyyymmdd") & "")
    If Not RS.EOF Then
       
       If RS!Glosa = "antes" Then
       
          EstCon = True
          
          If MsgBox("Desea tomar inventario antes la fecha propuesta - " & RS!Fecha_Inventario & " ...", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
          
             '-------> INI: Mover estado = 1 a la tabla parametro toma inventario
             vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
             '-------> FIN: Mover estado = 1 a la tabla parametro toma inventario

          '   Exit Sub
             
          End If
       
       ElseIf RS!Glosa = "hoy" Then
    
          EstCon = True
          
          If MsgBox("Desea tomar inventario con la fecha propuesta - " & RS!Fecha_Inventario & " ...", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
          
             '-------> INI: Mover estado = 1 a la tabla parametro toma inventario
             vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
             '-------> FIN: Mover estado = 1 a la tabla parametro toma inventario

           '  Exit Sub
             
          End If
       
       ElseIf RS!Glosa = "despues" Then
       
          EstCon = True
          
          If MsgBox("Desea tomar inventario despues de la fecha propuesta - " & RS!Fecha_Inventario & " ...", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
          
             '-------> INI: Mover estado = 1 a la tabla parametro toma inventario
             vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
             '-------> FIN: Mover estado = 1 a la tabla parametro toma inventario

            ' Exit Sub
             
          End If
          
       End If
    
    End If
    RS.Close
    Set RS = Nothing
        
    '-------> Validar inventario calendarizado
    If CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 39) Then
       
       EstCon = True
       
       '-------> INI: Mover estado = 1 a la tabla parametro toma inventario
       vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
       '-------> FIN: Mover estado = 1 a la tabla parametro toma inventario
             
'       MsgBox "Debe realizar inventario calendarizado. Proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
       MsgBox "Debe realizar inventario calendarizado, el cierre se bloqueara despues de este cierre y no se permite avanzar hasta que tome su inventario...", vbExclamation + vbOKOnly, MsgTitulo
       
'       Exit Sub
    
    End If

    '-------> preguntar si desea tomar inventario
    If Not EstCon And Not CierrePeriodo(Format(vg_ciedia, "yyyymmdd"), vg_codbod, 41) Then
    
       If MsgBox("Desea tomar inventario despues de este cierre...", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
          
          '-------> INI: Mover estado = 1 a la tabla parametro toma inventario
          vg_db.Execute "update a_param set par_valor = '1' where par_codigo = 'partominv' and par_cencos = '" & MuestraCasino(1) & "'"
          '-------> FIN: Mover estado = 1 a la tabla parametro toma inventario
           
       End If
       
    End If
    
    '-------> Rutina recalculo PMP día
    Toolbar1.Enabled = False
    Frame2.Enabled = False
    
    If vg_tipbase = "1" Then
       
       CalcularPMPDiaAccess Me, True, True
    
    Else
       
       '-------> Reprocesar dia por integración PEL
       Set RS = vg_db.Execute("SELECT  min(CONVERT(VARCHAR(8), lfs.Fecha, 112)) AS fecha " & _
                              "FROM    dbo.Log_FacturaSAP AS lfs " & _
                              "INNER JOIN dbo.b_clientes AS bc ON lfs.Ceco = bc.cli_codigo " & _
                              "WHERE   bc.cli_activo = '1' " & _
                              "AND CONVERT(VARCHAR(8), lfs.Fecha, 112) > ( SELECT  MAX(bt.tin_fectom) " & _
                                                    "FROM    dbo.b_tomainv AS bt " & _
                                                    "Where bt.tin_codbod = bc.cli_codbod ) " & _
                             "AND bc.cli_tipo = 0 " & _
                            "AND lfs.Estado = 3 " & _
                            "AND bc.cli_codigo = '" & MuestraCasino(1) & "'")
       
       If Not RS.EOF Then
          
          If IsNull(RS!Fecha) Then
             
             RS.Close: Set RS = Nothing
             
             If Not CalcularPMPDiaSql(Me, True, True) Then
             
                Toolbar1.Buttons(1).Visible = False
                Toolbar1.Buttons(2).Visible = True
                Toolbar1.Enabled = True: Frame2.Enabled = True
                
                If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
                
                Exit Sub
             
             End If
          
          Else
             
             If RS!Fecha < Format(CDate(vg_ciedia), "yyyymmdd") Then
                
                RS.Close: Set RS = Nothing
                CalcularPMPDiaSqlPEL Me, True, True
             
             Else
                
                RS.Close: Set RS = Nothing
                
                If Not CalcularPMPDiaSql(Me, True, True) Then
                
                   If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
                   
                   Toolbar1.Buttons(1).Visible = False
                   Toolbar1.Buttons(2).Visible = True
                   Toolbar1.Enabled = True: Frame2.Enabled = True
                   
                   Exit Sub
                   
                End If
             
             End If
          
          End If
       
       Else
          
          RS.Close: Set RS = Nothing
          
          If Not CalcularPMPDiaSql(Me, True, True) Then
          
             If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
             
             Toolbar1.Buttons(1).Visible = False
             Toolbar1.Buttons(2).Visible = True
             Toolbar1.Enabled = True: Frame2.Enabled = True
                   
             Exit Sub
             
          End If
       
       End If

    End If
    '-------> Mover zero al stock negativo
    vg_db.Execute "UPDATE b_bodegas set bod_canmer = 0 WHERE bod_codbod = " & vg_codbod & " AND round(bod_canmer, " & vg_DCa & ") < 0"
    
    If ValidarInventarioRotativo(MuestraCasino(1)) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 28) And Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 29) Then
       
       '-------> Actualizar b_productospmpdia stock
       If vg_tipbase = "1" Then
          
          vg_db.Execute "UPDATE b_productospmpdia INNER JOIN b_tomainv ON b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro " & _
                        "SET   b_productospmpdia.ppd_saldo  = b_tomainv.tin_stofis " & _
                        "WHERE b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                        "AND   b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                        "AND   b_tomainv.tin_codbod         = " & vg_codbod & " " & _
                        "AND   b_tomainv.tin_fectom         = " & Format(CDate(vg_ciedia), "yyyymmdd") & ""
       
       Else
          
          vg_db.Execute "Update b_productospmpdia " & _
                        "Set    b_productospmpdia.ppd_saldo = b_tomainv.tin_stofis " & _
                        "From   b_productospmpdia, b_tomainv " & _
                        "Where  b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro " & _
                        "AND    b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                        "AND    b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' " & _
                        "AND    b_tomainv.tin_codbod         = " & vg_codbod & " " & _
                        "AND    b_tomainv.tin_fectom         = " & Format(CDate(vg_ciedia), "yyyymmdd") & ""
       
       End If
    
    End If
    
    Call Proceso
    ActivarBotonCierreInventario


Case 2 '-------> Reabrir día
    
    Dim mensaje As String
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    If vaSpread1.BackColor = -2147483633 Then Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor And CDate(vg_ciedia) - 1 <> CDate(Format(fg_pone_cero(vaSpread1.text, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "dd/mm/yyyy")) Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(vg_ciedia) <> CDate(fg_pone_cero(vaSpread1.text, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy"))) And CDate(vg_ciedia) - 1 <> CDate(Format(fg_pone_cero(vaSpread1.text, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "dd/mm/yyyy")) Then MsgBox "Día no corresponde al cierre diario", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor _
       And CDate(vg_ciedia) - 1 = CDate(Format(fg_pone_cero(vaSpread1.text, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "dd/mm/yyyy")) _
       And Format(vg_ciedia, "mmyyyy") <> CDate(Format(fg_pone_cero(vaSpread1.text, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "mmyyyy")) Then
       
       If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 2) Then MsgBox "Existe información posterior, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    ElseIf ValidarInventarioRotativo(MuestraCasino(1)) Then
       
       If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 2) Then MsgBox "Existe información posterior, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 29) And Not ValidarInventarioRotativo(MuestraCasino(1)) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 4) Then
       
       MsgBox "Existe una toma inventario, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    ElseIf ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29) And Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 4) Then
       
       MsgBox "Existe una toma inventario, proceso cancelado", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    'Validar si existe un inventario en proceso 20201001
    If CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 38) Then
        
       MsgBox "Existe un inventario en proceso...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    mensaje = ""
    If Not CierrePeriodo(Format(CDate(vg_ciedia) - 1, "yyyymmdd"), vg_codbod, 29) And ValidarInventarioRotativo(MuestraCasino(1)) Then mensaje = "Si reabre el día, se borrar la toma inventario"
    
    If MsgBox(mensaje & VgLinea & VgLinea & "Esta Seguro Reabrir Día", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    fg_carga ""
    Toolbar1.Enabled = False: Frame2.Enabled = False
    Me.Caption = MsgTitulo
    
    vg_db.Execute "UPDATE a_param SET par_valor='" & fg_Encripta(LimpiaDato(CDate(vg_ciedia) - 1)) & "' WHERE par_codigo='ciediario' AND par_cencos='" & MuestraCasino(1) & "'"
    '-------> Grabar log_cierrediario
    sql1 = IIf(vg_tipbase = "1", " '" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "h:m:s") & "' ", " '" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "' ")
    sql2 = ""
    sql2 = IIf(vg_tipbase = "1", " CDATE('" & (CDate(vg_ciedia) - 1) & "') ", " '" & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & "' ")
    vg_db.Execute "INSERT INTO log_cierrediario VALUES (" & sql1 & ", " & sql2 & ", '" & Trim(vg_NUsr) & "', '2.- Reabrir Día', '" & MuestraCasino(1) & "')"
    If vg_tipbase = "1" Then
        vg_db.Execute "DELETE log_enviocierrediario FROM log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql2 & ""
    Else
       vg_db.Execute "DELETE log_enviocierrediario WHERE cencos = '" & MuestraCasino(1) & "' AND fecha = " & sql2 & ""
    End If
    
    '-------> Traer fecha cierre día
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    vg_ciedia = ""
    RS.Open "SELECT DISTINCT par_nombre, par_valor FROM a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If Not RS.EOF Then
       vg_ciedia = fg_Desencripta(TipoDato(RS!par_valor, ""))
       Partida.StatusBar1.Panels(8).text = Trim(RS!par_nombre) & " : " & CDate(vg_ciedia) - 1
    End If
    RS.Close: Set RS = Nothing
    
    If Not CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 29) And ValidarInventarioRotativo(MuestraCasino(1)) Then
       Dim estinv As Boolean
       '-------> borrar toma inventario
       estinv = CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 30) '------> rutina validar si so existe inventario

       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       'Detalle - Devuelve stock
       If vg_tipbase = "1" Then
          RS1.Open "SELECT dev.dev_codmer, dev.dev_canmer, aju.aju_tipo FROM b_totventas tov, b_detventas dev, a_tipoajuste aju " & _
                   "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
                   "AND   tov.tov_codser = aju.aju_codigo AND tov.tov_fecemi = Cdate('" & vg_ciedia & "') AND tov_codbod = " & vg_codbod & " " & _
                   "AND   tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' ORDER BY dev.dev_numlin", vg_db, adOpenStatic
       Else
          RS1.Open "SELECT dev.dev_codmer, dev.dev_canmer, aju.aju_tipo FROM b_totventas tov, b_detventas dev, a_tipoajuste aju " & _
                   "WHERE tov.tov_rutcli = dev.dev_rutcli AND tov.tov_tipdoc = dev.dev_tipdoc AND tov.tov_numdoc = dev.dev_numdoc " & _
                   "AND   tov.tov_codser = aju.aju_codigo AND tov.tov_fecemi = '" & Format(vg_ciedia, "yyyymmdd") & "' AND tov_codbod = " & vg_codbod & " " & _
                   "AND   tov.tov_tipdoc = 'AI' AND tov.tov_estdoc <> 'A' AND tov.tov_estdoc <> 'P' ORDER BY dev.dev_numlin", vg_db, adOpenStatic
       End If
       If Not RS1.EOF Then
           Do While Not RS1.EOF
               vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer" & IIf(RS1!aju_tipo = "A", "-", "+") & RS1!dev_canmer & " " & _
                             "WHERE bod_codpro = '" & RS1!dev_codmer & "' AND bod_codbod = " & vg_codbod & ""
               RS1.MoveNext
           Loop
       End If
       RS1.Close: Set RS1 = Nothing
       '-------> Borrar toma inventario
       vg_db.Execute "DELETE b_tomainv FROM b_tomainv WHERE tin_fectom = " & Format(vg_ciedia, "yyyymmdd") & " AND tin_codbod = " & vg_codbod & ""
       '-------> Borrar ajuste inventario
       If vg_tipbase = "1" Then
          
          vg_db.Execute "DELETE dev.* FROM b_totventas tov inner join b_detventas dev " & _
                        "ON (tov.tov_numdoc = dev.dev_numdoc) AND (tov.tov_tipdoc = dev.dev_tipdoc) " & _
                        "AND (tov.tov_rutcli = dev.dev_rutcli) " & _
                        "WHERE tov.tov_fecemi = Cdate('" & vg_ciedia & "') AND tov.tov_codbod = " & vg_codbod & " " & _
                        "AND tov.tov_tipdoc = 'AI'"
          vg_db.Execute "DELETE FROM b_totventas WHERE tov_fecemi = Cdate('" & vg_ciedia & "') AND tov_codbod = " & vg_codbod & " " & _
                        "AND tov_tipdoc = 'AI'"
       
       Else
          
          vg_db.Execute "DELETE b_detventas FROM b_totventas tov inner join b_detventas dev " & _
                        "ON (tov.tov_numdoc = dev.dev_numdoc) AND (tov.tov_tipdoc = dev.dev_tipdoc) " & _
                        "AND (tov.tov_rutcli = dev.dev_rutcli) " & _
                        "WHERE tov.tov_fecemi = '" & Format(vg_ciedia, "yyyymmdd") & "' AND tov.tov_codbod = " & vg_codbod & " " & _
                        "AND tov.tov_tipdoc = 'AI'"
          vg_db.Execute "DELETE FROM b_totventas WHERE tov_fecemi = '" & Format(vg_ciedia, "yyyymmdd") & "' AND tov_codbod = " & vg_codbod & " " & _
                        "AND tov_tipdoc = 'AI'"
       
       End If
'       vg_db.Execute "UPDATE b_productospmpdia SET ppd_saldo = 0 WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & ""
       'vg_db.CommitTrans
'       If estinv Then '-------> si variable es true recalcula día
'          Label1(1).Visible = True
'          Label1(1).Caption = "Un momento, Recalculando día..."
'          If vg_tipbase = "1" Then
'             CalcularPMPDiaAccess Me, False, true
'          Else
'             CalcularPMPDiaSql Me, False, true
'          End If
'          Label1(1).Caption = ""
'          Label1(1).Visible = False
'       End If
    End If
    
    Me.Caption = MsgTitulo
'28/06/2019
    vg_db.Execute "sgp_Upd_ReabrirCierreDiario '" & MuestraCasino(1) & "', " & Format(CDate(vg_ciedia) - IIf(ValidarInventarioRotativo(MuestraCasino(1)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31), 0, 1), "yyyymmdd") & ""
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.BackColor = Shape1(0).FillColor
    
    If Trim(vg_ciedia) = "" Then fg_descarga: MsgBox "No esta activo la fecha cierre día, Comunicase con departamento de informatica" & VgLinea & Space(40) & "Proceso cancelado ...", vbCritical + vbOKOnly, "Menú Principal": End
    
'    vg_ciedia = CDate(vg_ciedia) - 1
'    Partida.StatusBar1.Panels(8).text = "Cierre Diario : " & CDate(vg_ciedia) - 1
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Enabled = True: Frame2.Enabled = True
    If ConsultaProcess("sgpsdx.exe") Then KillProcess ("sgpsdx.exe")
    If Dir(dir_trabajo & "EnvioWebReporting" & "\" & MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".mdb") <> "" Then Kill dir_trabajo & "EnvioWebReporting" & "\" & MuestraCasino(1) & Format(vg_ciedia, "yyyymmdd") & ".mdb" ' borrar base datos si existe
    fg_descarga

Case 3 '-------> Actualizar lista
    
    ArmarCalendario
    ActivarBotonCierreInventario

Case 14 '-------> Salir
   
   Me.Hide
   Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    
On Error GoTo Man_Error

    If vaSpread1.MaxRows < 1 Then Exit Sub
    vaSpread1.Row = Row
    vaSpread1.Col = Col
    If vaSpread1.BackColor = -2147483633 Then Exit Sub
    Toolbar1.Buttons(1).Visible = IIf(vaSpread1.BackColor = Shape1(1).FillColor Or vaSpread1.BackColor = Shape1(2).FillColor, False, True)
    Toolbar1.Buttons(2).Visible = IIf(vaSpread1.BackColor = Shape1(1).FillColor Or vaSpread1.BackColor = Shape1(2).FillColor, True, False)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    
On Error GoTo Man_Error

    If vaSpread1.MaxRows < 1 Then Exit Sub
        ArmarCalendario
    vaSpread1.RetainSelBlock = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

Dim Nombre As String, Dia As String, Glosa As String
With vaSpread1
    If .MaxRows < 1 Then Exit Sub
    .Row = Row
    TipWidth = 4000
    ShowTip = True
    MultiLine = 2
    .Col = Col: Dia = Trim(.text)
    Glosa = IIf(.BackColor = Shape1(0).FillColor, "", IIf(.BackColor = Shape1(1).FillColor, "No Enviado : ", "Enviado : "))
    .Col = Col + 7: Nombre = Trim(.text)
    TipText = "Día   : " & Dia & vbCrLf & Glosa & Trim(Nombre)
End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Sub

Function Proceso()
    
On Error GoTo Man_Error

    Dim sql1   As String
    Dim sql2   As String
    Dim sql3   As String
    Dim RS_Pla As New ADODB.Recordset
    Dim RS     As New ADODB.Recordset
    Dim vRet   As Variant
    
With vaSpread1
    
    .Row = .ActiveRow
    .Col = .ActiveCol
    .BackColor = Shape1(1).FillColor
    Call Cierre_Dia
    M_RCDiar.Refresh
'    .SetActiveCell IIf(.ActiveCol > 7, 1, .ActiveCol + 1), IIf(.ActiveRow > 6, 1, .ActiveRow)
    sql1 = IIf(vg_tipbase = "2", "'" & Format(vg_ciedia, "yyyymmdd") & "'", " CDATE('" & vg_ciedia & "') ")

    If rst.State = 1 Then rst.Close
    rst.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    rst.Open "SELECT DISTINCT CFI_Fecha FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & MuestraCasino(1) & "' " & _
             "AND CFI_Fecha >= " & sql1 & " ORDER BY CFI_Fecha", vg_db, adOpenStatic
    
    Do While Not rst.EOF
        
        If rst.Fields("CFI_Fecha") = (CDate(vg_ciedia)) And CierrePeriodo(Format(CDate(vg_ciedia), "yyyymmdd"), vg_codbod, 31) Then
           
           '-------> Validar si existe minuta en ese día feriado.
           sql2 = IIf(vg_tipbase = "1", " AND cdate(a.CFI_Fecha) = '" & vg_ciedia & "' ", " AND Convert(VarChar(10), a.CFI_Fecha, 103) = '" & vg_ciedia & "' ")
           sql3 = Format(vg_ciedia, "yyyymmdd")
           
           If RS_Pla.State = 1 Then RS_Pla.Close
           RS_Pla.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
    
           RS_Pla.Open "SELECT DISTINCT a.CFI_Fecha FROM b_Fecha_Inhabiles a WHERE a.CFI_CeCo = '" & MuestraCasino(1) & "' " & sql2 & " AND a.CFI_CeCo IN (SELECT DISTINCT a.min_cencos FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos =  '" & MuestraCasino(1) & "' AND a.min_fecmin = " & sql3 & " AND b.mid_tipmin = '2')", vg_db, adOpenStatic
           If Not RS_Pla.EOF Then RS_Pla.Close: Set RS_Pla = Nothing: Exit Do
           RS_Pla.Close: Set RS_Pla = Nothing
           If vg_tipbase = "2" Then
              
              '-------> Reprocesar dia por integración PEL
              If RS.State = 1 Then RS.Close
              RS.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient

              Set RS = vg_db.Execute("SELECT  min(CONVERT(VARCHAR(8), lfs.Fecha, 112)) AS fecha " & _
                                    "FROM    dbo.Log_FacturaSAP AS lfs " & _
                                    "INNER JOIN dbo.b_clientes AS bc ON lfs.Ceco = bc.cli_codigo " & _
                                    "WHERE   bc.cli_activo = '1' " & _
                                    "AND CONVERT(VARCHAR(8), lfs.Fecha, 112) > ( SELECT  MAX(bt.tin_fectom) " & _
                                                    "FROM    dbo.b_tomainv AS bt " & _
                                                    "Where bt.tin_codbod = bc.cli_codbod ) " & _
                                    "AND bc.cli_tipo = 0 " & _
                                    "AND lfs.Estado = 3 " & _
                                    "AND bc.cli_codigo = '" & MuestraCasino(1) & "'")
              
              If Not RS.EOF Then
                 
                 If IsNull(RS!Fecha) Then
                    
                    RS.Close: Set RS = Nothing
                                    
                    If Not CalcularPMPDiaSql(Me, True, True) Then
                    
                        Toolbar1.Buttons(1).Visible = False
                        Toolbar1.Buttons(2).Visible = True
                        Toolbar1.Enabled = True: Frame2.Enabled = True
                
                        If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
                
                        Exit Function
                    
                    End If
                 
                 Else
                    If RS!Fecha < Format(CDate(vg_ciedia), "yyyymmdd") Then
                       
                       RS.Close: Set RS = Nothing
                       CalcularPMPDiaSqlPEL Me, True, True
                    
                    Else
                       
                       RS.Close: Set RS = Nothing
                       
                       If Not CalcularPMPDiaSql(Me, True, True) Then
                    
                          Toolbar1.Buttons(1).Visible = False
                          Toolbar1.Buttons(2).Visible = True
                          Toolbar1.Enabled = True: Frame2.Enabled = True
                
                          If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
                          
                          Exit Function
                    
                       End If
                    End If
                 
                 End If
              
              Else
                 
                 RS.Close: Set RS = Nothing
                 
                 If CalcularPMPDiaSql(Me, True, True) Then
              
                    Toolbar1.Buttons(1).Visible = False
                    Toolbar1.Buttons(2).Visible = True
                    Toolbar1.Enabled = True: Frame2.Enabled = True
                
                    If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
                    
                    Exit Function
              
                 End If
                 
              End If
           ElseIf vg_tipbase = "1" Then
              
              CalcularPMPDiaAccess Me, True, True
            
            End If
            .Row = .ActiveRow
            .Col = .ActiveCol + 1
            .BackColor = Shape1(1).FillColor
            Call Cierre_Dia
'            .SetActiveCell IIf(.ActiveCol > 7, 1, .ActiveCol + 1), IIf(.ActiveRow > 6, 1, .ActiveRow)
            vg_ciedia = (CDate(vg_ciedia) - 1) + 1
            M_RCDiar.Refresh
        Else
           Exit Do
        End If
        rst.MoveNext
     Loop
     
     rst.Close: Set rst = Nothing
     
     If Not ConsultaProcess("sgpsdx.exe") Then On Error Resume Next: vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")
     
     MsgBox "Proceso de Cierre Día Finalizado", vbInformation + vbOKOnly, MsgTitulo

End With

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Function

Function Cierre_Dia()
    
On Error GoTo Man_Error

    '-------> Traer fecha cierre día
    vg_ciedia = ""
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open "SELECT DISTINCT par_nombre, par_valor FROM a_param WHERE par_codigo = 'ciediario' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    
    If Not RS.EOF Then
       vg_ciedia = fg_Desencripta(TipoDato(RS!par_valor, ""))
       Partida.StatusBar1.Panels(8).text = Trim(RS!par_nombre) & " : " & CDate(vg_ciedia) - 1
    End If
    RS.Close: Set RS = Nothing
    
    If Trim(vg_ciedia) = "" Then MsgBox "No esta activo la fecha cierre día, Comunicase con departamento de informatica" & VgLinea & Space(40) & "Proceso cancelado ...", vbCritical + vbOKOnly, "Menú Principal": End
    
'    vg_ciedia = CDate(vg_ciedia) + 1
'    Partida.StatusBar1.Panels(8).text = "Cierre Diario : " & vg_ciedia
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Enabled = True: Frame2.Enabled = True

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"

End Function
