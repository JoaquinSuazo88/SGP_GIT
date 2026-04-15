VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_PtoLecVal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Punto Lectura de Vales"
   ClientHeight    =   8115
   ClientLeft      =   3195
   ClientTop       =   1710
   ClientWidth     =   13020
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   13020
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   12735
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   1
         Top             =   600
         Width           =   4815
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1440
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   2280
         MaxLength       =   100
         TabIndex        =   0
         Top             =   240
         Width           =   6975
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   2175
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   11655
         _Version        =   393216
         _ExtentX        =   20558
         _ExtentY        =   3836
         _StockProps     =   64
         ButtonDrawMode  =   2
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
         SpreadDesigner  =   "M_PtoLecVal.frx":0000
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   3135
         Left            =   600
         TabIndex        =   12
         Top             =   4080
         Width           =   11655
         _Version        =   393216
         _ExtentX        =   20558
         _ExtentY        =   5530
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
         MaxCols         =   5
         SpreadDesigner  =   "M_PtoLecVal.frx":188E
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ubicación"
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
         Left            =   600
         TabIndex        =   11
         Top             =   670
         Width           =   870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicio"
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
         Left            =   600
         TabIndex        =   10
         Top             =   3720
         Width           =   705
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   300
         TabIndex        =   9
         Top             =   1485
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Servicios"
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
         Index           =   1
         Left            =   600
         TabIndex        =   6
         Top             =   4080
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Punto de Atención"
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
         Left            =   600
         TabIndex        =   5
         Top             =   1035
         Width           =   1590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre PC"
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
         Left            =   600
         TabIndex        =   4
         Top             =   315
         Width           =   960
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   13020
      _ExtentX        =   22966
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_PtoLecVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim Msgtitulo  As String
Dim est As Boolean
Dim GraUbicacion As Boolean
Dim GraDetalle As Boolean

Private Sub Combo1_Click(Index As Integer)
'If est Then Exit Sub
'MoverDatosGrillas
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Msgtitulo = "Punto Lectura de Vales"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 4, modo
est = True
GraUbicacion = False
GraDetalle = False
Dim opusu As Boolean
Dim dwLen As Long
Dim strString As String
Dim RS As New ADODB.Recordset
Dim PtoAte As Long
PtoAte = 0
dwLen = MAX_COMPUTERNAME_LENGTH + 1
strString = String(dwLen, "X")
GetComputerName strString, dwLen
strString = Left(strString, dwLen)
Text1.text = Trim(strString)

''-------> Cargar Combo Punto de Atención
'CargarDatoCombo Combo1, 0, "a_pto_lectura_vales_pto_atencion", strString, "PtoLecVal", "N"
'If Combo1(0).ListCount > -1 Then
'   Combo1(0).ListIndex = 0
'End If

MoverDatosGrillasPtoVta
vaSpread2.MaxRows = 0
'MoverDatosGrillasServicio 0, Ptoate
est = False
End Sub

Sub MoverDatosGrillasPtoVta()
Dim RS As New ADODB.Recordset
Dim opusu As Boolean
Dim PtoAte As Long
est = True
RS.Open RutinaLectura.PtoLecturaVales(1, 0, Text1.text), vg_db, adOpenStatic
If Not RS.EOF Then Text2.text = IIf(IsNull(RS!lec_ubicacion), "", Trim(RS!lec_ubicacion))
RS.Close: Set RS = Nothing
Text2.Enabled = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, falso)
Label1(3).Caption = "Servicio "
'------->Cargar Punto Atención
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
With vaSpread1
'    opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", Falso, True)
    .MaxRows = 0
    RS.Open RutinaLectura.PuntoAtencion(5, 0, Text1.text), vg_db, adOpenStatic
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .Lock = opusu
       .Value = "0" 'IIf(IsNull(RS!activo), "0", "1")
       
       .Col = 2
       .CellType = CellTypeStaticText
       .Value = RS!ate_codatencion
       If i = 1 And Fila = 0 Then
          PtoAte = RS!ate_codatencion
       End If
       .Col = 3
       .CellType = CellTypeStaticText
       .Value = IIf(IsNull(RS!ate_descripcion), "", Trim(RS!ate_descripcion))
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    If .MaxRows > 0 Then
       .Row = Fila: .Col = 2: .SetActiveCell 1, .Row
    End If
    Gl_Ac_Botones Me, 1, IIf(.MaxRows > 0, 4, 3), modo
End With
vaSpread1.Visible = True
est = False
End Sub

Sub MoverDatosGrillasServicio(Fila As Long, PtoAte As Long)
Dim RS As New ADODB.Recordset
Dim opusu As Boolean
Dim CodLecVales As Long
Label1(3).Caption = "Servicio "
opusu = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", falso, True)
CodLecVales = 0
'Leer punto vales
RS.Open RutinaLectura.PtoLecturaVales(1, 0, Text1.text), vg_db, adOpenStatic
If Not RS.EOF Then CodLecVales = RS!lec_codlecvales
RS.Close: Set RS = Nothing

'Cargar Servicios
vaSpread2.Visible = False
vaSpread2.MaxRows = 0
With vaSpread2
    .MaxRows = 0
    'RS.Open RutinaLectura.Minutas(13, fg_codigocbo(Combo1, 0, 10, 0), 0, 0, Text1.text), vg_db, adOpenStatic
    RS.Open RutinaLectura.Minutas(13, CodLecVales, PtoAte, 0, Text1.text), vg_db, adOpenStatic
    Do While Not RS.EOF
       .MaxRows = .MaxRows + 1
       .Row = .MaxRows
       .Col = 1
       .Lock = opusu
       .Value = IIf(IsNull(RS!ate_codatencion) Or RS!ate_codatencion < 1, "0", "1")
       
       .Col = 2
       .CellType = CellTypeStaticText
       .Value = RS!reg_codigo
       
       .Col = 3
       .CellType = CellTypeStaticText
       .Value = IIf(IsNull(RS!reg_nombre), "", Trim(RS!reg_nombre))
       
       .Col = 4
       .CellType = CellTypeStaticText
       .Value = RS!ser_codigo
       
       .Col = 5
       .CellType = CellTypeStaticText
       .Value = IIf(IsNull(RS!ser_nombre), "", Trim(RS!ser_nombre))
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
'    Gl_Ac_Botones Me, 1, IIf(.MaxRows > 0, 8, 3), modo
End With
vaSpread2.Visible = True
End Sub

Private Sub Text2_Change()
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    GraUbicacion = True
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim codigo As Long
Dim i As Long
Dim Validar As Boolean
Dim PtoAte As Long
Dim CodLecVales As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1 'Incluir
Case 3 'Alterar
    '-------> Validar Servicio
    Validar = False
    '-------> Validar Pto. Atención
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then Validar = True
    Next i
    If Not Validar Then
        MsgBox "Debe seleccionar al menos un punto atención.", vbExclamation + vbOKOnly, Msgtitulo: vaSpread2.SetFocus: Exit Sub
    End If
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    GraDetalle = True
    GraUbicacion = True
Case 5 'Borrar
    '-------> Validar Pto. Atención
    Validar = False
    PtoAte = 0
    '-------> Validar Pto. Atención
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then vaSpread1.Col = 2: PtoAte = IIf(vaSpread1.text = "", 0, vaSpread1.text): Validar = True
    Next i
    If Not Validar Then
        MsgBox "Debe seleccionar al menos un punto atención.", vbExclamation + vbOKOnly, Msgtitulo: vaSpread2.SetFocus: Exit Sub
    End If
    '-------> Validar Servicio
    Validar = False
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        If vaSpread2.text = "1" Then Validar = True
    Next i
    If Not Validar Then
        MsgBox "Debe haber servicios seleccionado.", vbExclamation + vbOKOnly, Msgtitulo: vaSpread2.SetFocus: Exit Sub
    End If
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    'Leer punto vales
    RS.Open RutinaLectura.PtoLecturaVales(1, 0, Text1.text), vg_db, adOpenStatic
    If Not RS.EOF Then codigo = RS!lec_codlecvales
    RS.Close: Set RS = Nothing
    'Borrar datos vales atención
    vg_db.Execute "DELETE a_pto_lectura_vales_pto_atencion from a_pto_lectura_vales_pto_atencion where lec_codlecvales = " & codigo & " and ate_codatencion = " & PtoAte & ""
    'Borrar datos vales servicios
    vg_db.Execute "DELETE a_pto_lectura_vales_servicio from a_pto_lectura_vales_servicio where lec_codlecvales = " & codigo & " and ate_codatencion = " & PtoAte & ""
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    MoverDatosGrillasPtoVta
    vaSpread2.MaxRows = 0
    MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, Msgtitulo
    GraUbicacion = False
    GraDetalle = False
Case 7 'Actualizar Lista
    If vaSpread1.MaxRows < 1 Then Exit Sub
'    vaSpread1.Row = vaSpread1.ActiveRow
'    vaSpread1.Col = 2
'    MoverDatosGrillasServicio vaSpread1.Row, IIf(vaSpread1.text = "", 0, vaSpread1.text)
    MoverDatosGrillasPtoVta
    vaSpread2.MaxRows = 0
    Label1(3).Caption = "Servicio "
    GraUbicacion = False
    GraDetalle = False
Case 10 'Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    MoverDatosGrillasPtoVta
    vaSpread2.MaxRows = 0
    Label1(3).Caption = "Servicio "
    GraUbicacion = False
    GraDetalle = False
'   If vaSpread1.MaxRows < 1 Then Exit Sub
'    vaSpread1.Row = vaSpread1.ActiveRow
'    vaSpread1.Col = 2
'    MoverDatosGrillasServicio vaSpread1.Row, IIf(vaSpread1.text = "", 0, vaSpread1.text)
Case 12
    GrabaRegistro
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    'Leer punto vales
    RS.Open RutinaLectura.PtoLecturaVales(1, 0, Text1.text), vg_db, adOpenStatic
    If Not RS.EOF Then CodLecVales = RS!lec_codlecvales
    RS.Close: Set RS = Nothing
    I_PtoLecturaVales CodLecVales, Text1.text
    GraUbicacion = False
    GraDetalle = False
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub GrabaRegistro()
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim i As Long
Dim codigo As Long
Dim PtoAte As Long
Dim codreg As Long
Dim codser As Long
Dim nommaq As String
Dim Validar As Boolean
If Trim(LimpiaDato(Text2.text)) = "" Then MsgBox "Favor ingresar ubicación, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, Msgtitulo: Text2.SetFocus: Exit Sub
If GraUbicacion Or GraDetalle Then
   modo = "A"
   codigo = 0
   PtoAte = 0
   RS.Open RutinaLectura.PtoLecturaVales(1, 0, Text1.text), vg_db, adOpenStatic
   If Not RS.EOF Then modo = "M": codigo = RS!lec_codlecvales
   RS.Close: Set RS = Nothing
   '-------> Traer ultimo código
   If modo = "A" Then
      RS.Open RutinaLectura.PtoLecturaVales(6, 0, ""), vg_db, adOpenStatic
      If Not RS.EOF Then RS.MoveFirst: codigo = RS!lec_codlecvales + 1 Else codigo = 1
      RS.Close: Set RS = Nothing
   End If
   '-------> Grabar Datos
   If modo = "A" Then
      vg_db.Execute "INSERT INTO a_pto_lectura_vales (lec_codlecvales, lec_nombrepc, lec_ubicacion, lec_activo) VALUES (" & codigo & ", '" & Trim(Text1.text) & "', '" & Trim(LimpiaDato(Text2.text)) & "', 1)"
   Else
      vg_db.Execute "UPDATE a_pto_lectura_vales SET lec_ubicacion = '" & Trim(LimpiaDato(Text2.text)) & "' WHERE lec_codlecvales = " & codigo & ""
   End If
   If Not GraDetalle Then
      '-------> Activar Pto. Atención
      For i = 1 To vaSpread1.MaxRows
          vaSpread1.Row = i
          vaSpread1.Col = 1
          vaSpread1.Lock = False
      Next i
      modo = "": Gl_Ac_Botones Me, 1, 4, modo
      MsgBox "Registro guardo exitosamente", vbInformation + vbOKOnly, Msgtitulo
      GraUbicacion = False
      GraDetalle = False
      Exit Sub
   End If
End If
Validar = False
'-------> Validar Pto. Atención
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If vaSpread1.text = "1" Then Validar = True
Next i
If Not Validar Then
    MsgBox "Debe seleccionar al menos un punto atencion", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.SetFocus: Exit Sub
End If
'-------> Validar Servicio
Validar = False
'-------> Validar Pto. Atención
For i = 1 To vaSpread2.MaxRows
    vaSpread2.Row = i
    vaSpread2.Col = 1
    If vaSpread2.text = "1" Then Validar = True
Next i
If Not Validar Then
    MsgBox "Debe seleccionar al menos un servicio", vbExclamation + vbOKOnly, Msgtitulo: vaSpread2.SetFocus: Exit Sub
End If
'vg_db.Execute "DELETE a_pto_lectura_vales_pto_atencion from a_pto_lectura_vales_pto_atencion where lec_codlecvales = " & codigo & " and ate_codatencion = " & fg_codigocbo(Combo1, 0, 10, 0) & ""
'vg_db.Execute "INSERT INTO a_pto_lectura_vales_pto_atencion (lec_codlecvales, ate_codatencion) values (" & codigo & ", '" & fg_codigocbo(Combo1, 0, 10, 0) & "')"
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If vaSpread1.text = "1" Then vaSpread1.Col = 2: PtoAte = vaSpread1.text: Exit For
Next i
'Borrar datos vales atención
vg_db.Execute "DELETE a_pto_lectura_vales_pto_atencion from a_pto_lectura_vales_pto_atencion where lec_codlecvales = " & codigo & " and ate_codatencion = " & PtoAte & ""
vg_db.Execute "INSERT INTO a_pto_lectura_vales_pto_atencion (lec_codlecvales, ate_codatencion) values (" & codigo & ", " & PtoAte & ")"

'Borrar datos vales servicios
vg_db.Execute "DELETE a_pto_lectura_vales_servicio from a_pto_lectura_vales_servicio where lec_codlecvales = " & codigo & " and ate_codatencion = " & PtoAte & ""
With vaSpread2
    For i = 1 To .MaxRows
        .Row = i
        .Col = 1
        If .text = "1" Then
           .Col = 2: codreg = .Value
           .Col = 4: codser = .Value
           vg_db.Execute "INSERT INTO a_pto_lectura_vales_servicio (lec_codlecvales, reg_codigo, ser_codigo, ate_codatencion) values (" & codigo & ", " & codreg & ", " & codser & ", " & PtoAte & ")"
        End If
    Next i
End With
'-------> Activar Pto. Atención
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    vaSpread1.Lock = False
Next i
modo = "": Gl_Ac_Botones Me, 1, 4, modo
MsgBox "Registro guardo exitosamente", vbInformation + vbOKOnly, Msgtitulo
GraUbicacion = False
GraDetalle = False
Exit Sub
Man_Error:
MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If ButtonDown = 0 Then Exit Sub
If est Then Exit Sub
Dim i As Long
Dim PtoAte As Long
Dim NomPtoAte As String
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    If i <> Row Then
'       vaSpread1.Lock = True
       If vaSpread1.text = "1" Then
          est = True
          vaSpread1.text = "0"
          est = False
        End If
    End If
Next i
'If modo = "" Then modo = "M"
'Gl_Ac_Botones Me, 1, 0, modo
vaSpread1.Row = Row
vaSpread1.Col = 2
PtoAte = IIf(vaSpread1.text = "", 0, vaSpread1.text)
vaSpread1.Col = 3
NomPtoAte = IIf(vaSpread1.text = "", 0, vaSpread1.text)
est = True
MoverDatosGrillasServicio 0, PtoAte
Label1(3).Caption = "Servicio " & PtoAte & " - " & Trim(NomPtoAte)
est = False
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
'If vaSpread1.MaxRows < 1 Then Exit Sub
'   vaSpread1.Row = vaSpread1.ActiveRow
'   vaSpread1.Col = 2
'   MoverDatosGrillasServicio vaSpread1.Row, IIf(vaSpread1.text = "", 0, vaSpread1.text)
End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
est = True
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i
    vaSpread1.Col = 1
    vaSpread1.Lock = True
Next i
GraDetalle = True
est = False
End Sub
