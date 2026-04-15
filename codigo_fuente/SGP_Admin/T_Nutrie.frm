VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form T_Nutrie 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nutrientes"
   ClientHeight    =   4815
   ClientLeft      =   1860
   ClientTop       =   1995
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   10905
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   3450
      Left            =   0
      TabIndex        =   7
      Top             =   1350
      Width           =   10875
      _Version        =   393216
      _ExtentX        =   19182
      _ExtentY        =   6085
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      EditEnterAction =   4
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
      MaxCols         =   7
      SpreadDesigner  =   "T_Nutrie.frx":0000
      ScrollBarTrack  =   3
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   971
      Left            =   1800
      TabIndex        =   0
      Top             =   360
      Width           =   6015
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "T_Nutrie.frx":1B82
         Left            =   1965
         List            =   "T_Nutrie.frx":1B8C
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1965
         TabIndex        =   1
         Top             =   555
         Width           =   2505
         _Version        =   196608
         _ExtentX        =   4410
         _ExtentY        =   870
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
         NoSpecialKeys   =   0
         AutoAdvance     =   -1  'True
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
         ControlType     =   0
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   4545
         TabIndex        =   4
         Top             =   645
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Texto"
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
         Left            =   480
         TabIndex        =   3
         Top             =   645
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Buscar Columna"
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
         Left            =   480
         TabIndex        =   2
         Top             =   345
         Width           =   1380
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10905
      _ExtentX        =   19235
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "T_Nutrie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo As String
Dim OpGr As Boolean

Private Sub GrabaRegistro(Fila As Long)

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim codigo      As Long
Dim Nombre      As String
Dim NomCor      As String
Dim indpri      As Long
Dim secnro      As Long
Dim IndPriSello As String
Dim SecNroSello As Long

OpGr = True
vaSpread1.Row = Fila

vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

vaSpread1.Col = 2
Nombre = Trim(LimpiaDato(Mid(vaSpread1.Value, 1, 30)))

vaSpread1.Col = 3
NomCor = Trim(LimpiaDato(Mid(vaSpread1.Value, 1, 5)))

vaSpread1.Col = 4
If vaSpread1.text = "1" Then
    
   vaSpread1.Col = 4
   indpri = 1

Else
    
   vaSpread1.Col = 4
   indpri = 0

End If

vaSpread1.Col = 5
secnro = Trim(LimpiaDato(vaSpread1.Value))

vaSpread1.Col = 6
If vaSpread1.text = "1" Then
   
   IndPriSello = "1"
   
Else

   IndPriSello = "0"

End If

vaSpread1.Col = 7
SecNroSello = IIf(Trim(vaSpread1.Value) = "", 0, vaSpread1.Value)

If Trim(Nombre) = "" Then

   MsgBox "Falta nombre...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 2
   vaSpread1.SetActiveCell 2, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If Trim(NomCor) = "" Then

  MsgBox "Falta nombre corto...", vbExclamation + vbOKOnly, MsgTitulo
  vaSpread1.Row = Fila
  vaSpread1.Col = 3
  vaSpread1.SetActiveCell 2, vaSpread1.Row
  vaSpread1.SetFocus
  OpGr = False
  Exit Sub

End If

If secnro < 1 Then

   MsgBox "Falta orden ...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 5
   vaSpread1.SetActiveCell 2, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

If SecNroSello < 1 Then

   MsgBox "Falta orden sello...", vbExclamation + vbOKOnly, MsgTitulo
   vaSpread1.Row = Fila
   vaSpread1.Col = 7
   vaSpread1.SetActiveCell 2, vaSpread1.Row
   vaSpread1.SetFocus
   OpGr = False
   Exit Sub

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 7
    If vaSpread1.text = SecNroSello And i <> Fila Then
    
       MsgBox "Ya existe un orden sello en otra fila...", vbExclamation + vbOKOnly, MsgTitulo
       vaSpread1.Row = Fila
       vaSpread1.Col = 7
       vaSpread1.SetActiveCell 2, vaSpread1.Row
       vaSpread1.SetFocus
       OpGr = False
       Exit Sub
    
    End If
    
Next i

If modo = "A" Then
   
   indice = 0
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgpadm_iu_nutriente 'A', 0, '" & Trim(Nombre) & "', '" & Trim(NomCor) & "', " & indpri & ", " & secnro & ", '" & IndPriSello & "', " & SecNroSello & "")
   If Not RS.EOF Then
      
      codigo = RS!indice
      vaSpread1.Col = 1
      vaSpread1.Value = codigo
   
   End If
   RS.Close
   Set RS = Nothing
   
Else
   
   vg_db.Execute "sgpadm_iu_nutriente 'M', " & codigo & ", '" & Trim(Nombre) & "', '" & Trim(NomCor) & "', " & indpri & ", " & secnro & ", '" & IndPriSello & "', " & SecNroSello & ""

End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True
fpText1.Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 5380
Me.Width = 10995
MsgTitulo = "Nutriente"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1
MoverDatosGrillas
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState = 0 Then
   
   Frame1.Move 1200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440

ElseIf Me.WindowState = 2 Then
   
   Frame1.Move 4200, 360, 6015, 971
   vaSpread1.Move 0, 1440, ScaleWidth, ScaleHeight - 1440

End If
Toolbar1.Refresh

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub fpText1_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
vaSpread1.Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
     
     Set RS = vg_db.Execute("sgpadm_s_nutriente 3, 0,'%" & UCase(LimpiaDato(fpText1.text)) & "%'")

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
     
     Set RS = vg_db.Execute("sgpadm_s_nutriente 4, 0,'%" & UCase(LimpiaDato(fpText1.text)) & "%'")

End If
If RS.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS!nReg
i = 1
If Not RS.EOF Then
    Do While Not RS.EOF
        
        vaSpread1.Row = i
        i = i + 1
        vaSpread1.Col = 1
        vaSpread1.TypeHAlign = 1
        vaSpread1.Value = RS!nut_codigo
        
        vaSpread1.Col = 2
        vaSpread1.Value = Trim(RS!nut_nombre)
        
        vaSpread1.Col = 3
        vaSpread1.Value = Trim(RS!nut_nomuni)
        
        vaSpread1.Col = 4
        If RS!nut_indpri = 1 Then
        
            vaSpread1.CellType = 10
            vaSpread1.TypeCheckText = ""
            vaSpread1.TypeCheckCenter = True
            vaSpread1.text = "1"
        
        Else
        
            vaSpread1.CellType = 10
            vaSpread1.TypeCheckText = ""
            vaSpread1.TypeCheckCenter = True
            vaSpread1.text = "0"

        End If
      
        vaSpread1.Col = 5
        vaSpread1.Value = RS!nut_secnro
      
        vaSpread1.Col = 6
        If RS!nut_IndicePrincipalSello = True Then
        
            vaSpread1.CellType = 10
            vaSpread1.TypeCheckText = ""
            vaSpread1.TypeCheckCenter = True
            vaSpread1.text = "1"
        
        Else
        
            vaSpread1.CellType = 10
            vaSpread1.TypeCheckText = ""
            vaSpread1.TypeCheckCenter = True
            vaSpread1.text = "0"
        
        End If
      
        vaSpread1.Col = 7
        vaSpread1.Value = RS!nut_OrdenPrincipalSello
      
      RS.MoveNext
   
   Loop
   Gl_Ac_Botones Me, 1, 1, modo
   
End If
RS.Close
Set RS = Nothing
vaSpread1.Visible = True
If fpText1.text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim codigo As Long
Dim Nombre As String
Dim NomCor As String

Select Case Button.Index

    Case 1
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 2
        vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
        vaSpread1.SetFocus
    
    Case 3
        
        modo = "M"
        If vaSpread1.MaxRows < 1 Then Exit Sub
        Gl_Ac_Botones Me, 1, 0, modo
    
    Case 5
        
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = Val(vaSpread1.Value)
        
        vg_db.Execute "DELETE FROM a_nutriente WHERE nut_codigo=" & codigo
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
    
    Case 7
        
        fpText1.text = ""
        MoverDatosGrillas
    
    Case 10
        
        If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        If modo = "A" Then
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.DeleteRows vaSpread1.Row, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        
        Else
            
            Cancela
        
        End If
        modo = ""
        Gl_Ac_Botones Me, 1, 1, modo
        Combo1.Enabled = True
        fpText1.Enabled = True
    
    Case 12
        
        GrabaRegistro vaSpread1.ActiveRow
    
    Case 15
        
        If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        I_Nutri
    
    Case 18
        
        Me.Hide
        Unload Me

End Select
Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then
   
   MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error"

Else
    
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End If

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If (Col <> 4 And Col <> 6) Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub MoverDatosGrillas()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

OpGr = True
vaSpread1.Visible = False
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0,''")

Do While Not RS.EOF
    
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    
    vaSpread1.Col = 1
    vaSpread1.Value = RS!nut_codigo
    
    vaSpread1.Col = 2
    vaSpread1.Value = Trim(RS!nut_nombre)
    
    vaSpread1.Col = 3
    vaSpread1.Value = Trim(RS!nut_nomuni)
    
    vaSpread1.Col = 4
    
    If RS!nut_indpri = 1 Then
        
        vaSpread1.CellType = 10
        vaSpread1.TypeCheckText = ""
        vaSpread1.TypeCheckCenter = True
        vaSpread1.text = "1"
    
    Else
    
        vaSpread1.CellType = 10
        vaSpread1.TypeCheckText = ""
        vaSpread1.TypeCheckCenter = True
        vaSpread1.text = "0"
    
    End If
    
    vaSpread1.Col = 5
    vaSpread1.Value = Trim(RS!nut_secnro)
    
    vaSpread1.Col = 6
    If RS!nut_IndicePrincipalSello = True Then
        
        vaSpread1.CellType = 10
        vaSpread1.TypeCheckText = ""
        vaSpread1.TypeCheckCenter = True
        vaSpread1.text = "1"
    
    Else
    
        vaSpread1.CellType = 10
        vaSpread1.TypeCheckText = ""
        vaSpread1.TypeCheckCenter = True
        vaSpread1.text = "0"
    
    
    End If
    
    vaSpread1.Col = 7
    vaSpread1.Value = RS!nut_OrdenPrincipalSello
    
    RS.MoveNext
    
Loop
RS.Close
Set RS = Nothing

Gl_Ac_Botones Me, 1, 1, modo
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
vaSpread1.Visible = True
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    
    GrabaRegistro Row

ElseIf Toolbar1.Buttons(12).Visible = False Then
    
    Cancela

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Cancela()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_s_nutriente 5, " & codigo & ",''")
If Not RS.EOF Then
    
    vaSpread1.Col = 2
    vaSpread1.Value = Trim(RS!nut_nombre)
    
    vaSpread1.Col = 3
    vaSpread1.Value = Trim(RS!nut_nomuni)
    
    vaSpread1.Col = 4
    vaSpread1.CellType = 10
    vaSpread1.TypeCheckText = ""
    vaSpread1.TypeCheckCenter = True
    vaSpread1.text = Trim(Str(RS!nut_indpri))
    
    vaSpread1.Col = 5
    vaSpread1.Value = Trim(RS!nut_secnro)
    
    vaSpread1.Col = 6
    vaSpread1.CellType = 10
    vaSpread1.TypeCheckText = ""
    vaSpread1.TypeCheckCenter = True
    vaSpread1.text = IIf(RS!nut_IndicePrincipalSello = True, "1", "0")
    
    vaSpread1.Col = 7
    vaSpread1.Value = RS!nut_OrdenPrincipalSello
    
End If
RS.Close
Set RS = Nothing
OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub
