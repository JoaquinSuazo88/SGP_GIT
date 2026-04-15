VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_FechaGrupoDespachoCeco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantención días de grupo despacho casino"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9555
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   2520
      TabIndex        =   14
      Top             =   4680
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton Command1 
         Caption         =   "&No"
         Height          =   255
         Index           =   1
         Left            =   3960
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Si"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   16
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   6000
      TabIndex        =   6
      Top             =   600
      Width           =   3375
      Begin VB.CheckBox Check2 
         Caption         =   "Dom"
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
         Height          =   255
         Index           =   6
         Left            =   1560
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sab"
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
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Vie"
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
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Jue"
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
         Height          =   255
         Index           =   3
         Left            =   2280
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Mie"
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
         Height          =   255
         Index           =   2
         Left            =   1560
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Mar"
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
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Lun"
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
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   735
      End
   End
   Begin EditLib.fpText fpText1 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
      _Version        =   196608
      _ExtentX        =   2990
      _ExtentY        =   556
      Enabled         =   0   'False
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
      AutoCase        =   0
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   14595
      _ExtentX        =   25744
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4575
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   9255
      _Version        =   393216
      _ExtentX        =   16325
      _ExtentY        =   8070
      _StockProps     =   64
      AutoClipboard   =   0   'False
      ButtonDrawMode  =   1
      DisplayRowHeaders=   0   'False
      EditEnterAction =   5
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   14
      MaxRows         =   1
      ProcessTab      =   -1  'True
      ScrollBars      =   2
      SpreadDesigner  =   "M_FechaGrupoDespachoCeco.frx":0000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   75
   End
   Begin VB.Label Label1 
      Caption         =   "Ceco : "
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
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Ultima Fecha Generada"
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
      TabIndex        =   3
      Top             =   1320
      Width           =   2175
   End
End
Attribute VB_Name = "M_FechaGrupoDespachoCeco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim modo             As String
Dim MsgTitulo        As String
Dim Est              As Boolean
Dim EstA             As Boolean
Public lc_Aux        As String
Dim fecha_parametro  As String
Dim codigoceco       As String
Dim nombrecodigoCeco As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS                As New ADODB.Recordset
Dim i                 As Long
Dim codgru            As Long
Dim MyBuffer          As String
Dim Sql               As String
Dim EstDia            As Boolean
Dim Lun               As String
Dim Mar               As String
Dim Mie               As String
Dim Jue               As String
Dim Vie               As String
Dim Sab               As String
Dim Dom               As String
Dim DesGrDes          As String
Dim estext            As String
Dim fechamodif        As String
Dim actualiza         As String
Dim buffact           As String
Dim collecact         As String
Dim modulo            As String
Dim buffantes         As String
Dim buffactual        As String
Dim existeActualizado As Integer
Dim fechaavalidar     As String
Dim FechaMenor        As Long

Frame2.Visible = False

Select Case Index

    Case 0
    
        vg_BorradoDatos = True
        Frame2.Visible = False
        vaSpread1.Row = -1
        vaSpread1.Col = -1
        vaSpread1.Lock = False
        vaSpread1.Enabled = True
        Toolbar1.Enabled = True
        
        '--> Validar que exista un dato seleccionado como modificado
        EstDia = False
        For i = 1 To vaSpread1.MaxRows
    
            vaSpread1.Row = i
            vaSpread1.Col = 13
    
            If vaSpread1.text = "1" Then
    
               EstDia = True
    
            End If
    
        Next i
    
        If EstDia = False Then
    
           MsgBox "Debe haber a lo menos un datos modificado. Proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
           Exit Sub
    
        End If
    
        '--> Validar que los días sean igual a los parametros CD
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Sql = "sgpadm_Sel_TraerDiaParametroCD '" & codigoceco & "'"
        Set RS = vg_db.Execute(Sql)
    
        If Not RS.EOF Then
    
           Lun = IIf(RS("lu") = False, 0, 1)
           Mar = IIf(RS("ma") = False, 0, 1)
           Mie = IIf(RS("mi") = False, 0, 1)
           Jue = IIf(RS("ju") = False, 0, 1)
           Vie = IIf(RS("vi") = False, 0, 1)
           Sab = IIf(RS("sa") = False, 0, 1)
           Dom = IIf(RS("do") = False, 0, 1)
    
        End If
    
        RS.Close
        Set RS = Nothing
    
        EstDia = True
        For i = 1 To vaSpread1.MaxRows
    
            vaSpread1.Row = i
            vaSpread1.Col = 13
    
            If vaSpread1.text = "1" Then
    
               vaSpread1.Col = 3
               DesGrDes = vaSpread1.text
    
               vaSpread1.Col = 4
               If vaSpread1.text <> Lun And Lun <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Lunes no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 5
               If vaSpread1.text <> Mar And Mar <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Martes no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 6
               If vaSpread1.text <> Mie And Mie <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Miercoles no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 7
               If vaSpread1.text <> Jue And Jue <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Jueves no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 8
               If vaSpread1.text <> Vie And Vie <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Viernes no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 9
               If vaSpread1.text <> Sab And Sab <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Sabado no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 10
               If vaSpread1.text <> Dom And Dom <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Domingo no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
            End If
    
        Next i
    
        If EstDia = False Then
    
           fg_descarga
           Exit Sub
    
        End If
    
        EstDia = False
    
        '--> Grabar solo parametros grupo despacho
        For i = 1 To vaSpread1.MaxRows
    
            vaSpread1.Row = i
            vaSpread1.Col = 13
    
            If Trim(vaSpread1.text) = "1" Then
    
               Lun = 0
               Mar = 0
               Mie = 0
               Jue = 0
               Vie = 0
               Sab = 0
               Dom = 0
               codgru = 0
    
               vaSpread1.Col = 2
               codgru = vaSpread1.text
    
               vaSpread1.Col = 4
               Lun = vaSpread1.text
    
               vaSpread1.Col = 5
               Mar = vaSpread1.text
    
               vaSpread1.Col = 6
               Mie = vaSpread1.text
    
               vaSpread1.Col = 7
               Jue = vaSpread1.text
    
               vaSpread1.Col = 8
               Vie = vaSpread1.text
    
               vaSpread1.Col = 9
               Sab = vaSpread1.text
    
               vaSpread1.Col = 10
               Dom = vaSpread1.text
    
               '--> Grabar Parametros grupo despacho
               If Lun <> 0 Or Mar <> 0 Or Mie <> 0 Or Jue <> 0 Or Vie <> 0 Or Sab <> 0 Or Dom <> 0 Then
    
                  If RS.State = 1 Then RS.Close
                  RS.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
    
                  Sql = ""
                  Sql = " sgpadm_Ins_ParametroGrupoDespachoCeco "
                  Sql = Sql & " '" & codigoceco & "',"
                  Sql = Sql & codgru & ","
                  Sql = Sql & Lun & ","
                  Sql = Sql & Mar & ","
                  Sql = Sql & Mie & ","
                  Sql = Sql & Jue & ","
                  Sql = Sql & Vie & ","
                  Sql = Sql & Sab & ","
                  Sql = Sql & Dom
    
                  Set RS = vg_db.Execute(Sql)
    
                  If Not RS.EOF Then
    
                     If RS(0) > 0 Then
    
                        RS.Close
                        Set RS = Nothing
                        MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
                        Exit Sub
    
                     Else
    
                        EstDia = True
    
                     End If
    
                  End If
                  RS.Close
                  Set RS = Nothing
    
              Else
    
                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
    
                 Sql = ""
                 Sql = " sgpadm_Del_ParamtroGrupoDespachoCecos "
                 Sql = Sql & " '" & codigoceco & "', "
                 Sql = Sql & " " & codgru & " "
    
                 Set RS = vg_db.Execute(Sql)
    
                 If Not RS.EOF Then
    
                    If RS(0) > 0 Then
    
                       RS.Close
                       Set RS = Nothing
                       MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
                       Exit Sub
    
                    Else
    
                       EstDia = True
    
                    End If
    
                 End If
                 RS.Close
                 Set RS = Nothing
    
              End If
    
            End If
    
        Next i
    
        '--> Grabar solo parametros grupo despacho
        FechaMenor = 0
        xmlfamilia = ""
        xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
    
        For i = 1 To vaSpread1.MaxRows
    
           estext = False
           vaSpread1.Row = i
    
           vaSpread1.Col = 12 'Marca Actualiza
           fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
            vaSpread1.Col = 13 'Marca Actualiza
            actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
            If fechaavalidar <> "0" And actualiza = 1 Then
    
              vaSpread1.Col = 12 'fecha
              fechamodif = Format(vaSpread1.text, "YYYYMMDD")
    
              If Format(vaSpread1.text, "YYYYMMDD") < FechaMenor Or FechaMenor = 0 Then
    
                 FechaMenor = fechamodif
    
              End If
    
              vaSpread1.Col = 2 'ceco
              codgrupo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
              xmlfamilia = xmlfamilia & " <RutaCeco"
              xmlfamilia = xmlfamilia & " Ceco = " & Chr(34) & codigoceco & Chr(34)
              xmlfamilia = xmlfamilia & " GrpDes = " & Chr(34) & codgrupo & Chr(34)
              xmlfamilia = xmlfamilia & "/>"
    
            End If
    
        Next i
    
        xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
    
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Sql = "sgpadm_Ins_XmlGenerarRutaGrupoDespachoCeco "
        Sql = Sql & " '" & xmlfamilia & "',"
        Sql = Sql & fechamodif & ","
        Sql = Sql & fecha_parametro & ","
        Sql = Sql & IIf(vg_BorradoDatos, 1, 0) & ", "
        Sql = Sql & FechaMenor & ", "
        Sql = Sql & vg_NUsr
    
        Set RS = vg_db.Execute(Sql)
        If Not RS.EOF Then
    
           If RS(0) > 0 Then
    
              RS.Close
              Set RS = Nothing
              MsgBox RS(1)
              Exit Sub
    
           End If
    
        End If
    
        RS.Close
        Set RS = Nothing
    
        'registrar Log sistema actualizar lista
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
    
        MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
    
        modo = ""
        Gl_Ac_Botones Me, 1, 13, modo
    
        fg_descarga
    
  
    Case 1
        
        vg_BorradoDatos = False
'        ValidarRutasNormal = False
        Frame2.Visible = False
        vaSpread1.Row = -1
        vaSpread1.Col = -1
        vaSpread1.Lock = False
        vaSpread1.Enabled = True
        Toolbar1.Enabled = True
        
        Exit Sub

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = 1196005
MsgTitulo = "Mantención días de grupo despacho casino"
fg_centra Me
modo = ""

Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 13, modo

Label1(2).Caption = codigoceco & " - " & nombrecodigoCeco

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Sub Moverdetalle(Ceco As String, NombreCeco As String)

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim i      As Long
Dim AssMod As Boolean
   
fg_carga ""

codigoceco = Ceco
nombrecodigoCeco = NombreCeco

Me.HelpContextID = 1196005
'-------> Dar acceso modificar rutas
AssMod = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
    
vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = "sgpadm_Sel_Parametros 'fecrutagde'"
Set RS = vg_db.Execute(Sql)
If Not RS.EOF Then
     
   fecha_parametro = RS(0)
  
End If
RS.Close
Set RS = Nothing
   
If fecha_parametro <> "" Then
     
   fpText1 = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
  
End If

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = "sgpadm_Sel_TraerDiaParametroCD '" & codigoceco & "'"
Set RS = vg_db.Execute(Sql)

If Not RS.EOF Then

   Check2(0).Value = IIf(RS("lu") = False, 0, 1)
   Check2(1).Value = IIf(RS("ma") = False, 0, 1)
   Check2(2).Value = IIf(RS("mi") = False, 0, 1)
   Check2(3).Value = IIf(RS("ju") = False, 0, 1)
   Check2(4).Value = IIf(RS("vi") = False, 0, 1)
   Check2(5).Value = IIf(RS("sa") = False, 0, 1)
   Check2(6).Value = IIf(RS("do") = False, 0, 1)
       
End If
    
RS.Close
Set RS = Nothing
    
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Parametros_GrupoDespacho_casino '" & Ceco & "'")

Do While Not RS.EOF
             
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 2 'codigo familia sgp
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = RS("Idgrupodespacho")
                 
   vaSpread1.Col = 3 ' Nombre familia sgp
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.text = RS("Nombre")
        
   vaSpread1.Col = 4 ' Lunes
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("lu")
        
   vaSpread1.Col = 5 ' Martes
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("ma")
        
   vaSpread1.Col = 6 ' Miercoles
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("mi")
        
   vaSpread1.Col = 7 ' Jueves
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("ju")
        
   vaSpread1.Col = 8 ' Viernes
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("vi")
   
   vaSpread1.Col = 9 ' Sabado
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("sa")
   
   vaSpread1.Col = 10 ' Domingo
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = RS("do")
   
   vaSpread1.Col = 12
   vaSpread1.CellType = CellTypeStaticText
   vaSpread1.Lock = IIf(AssMod = True, False, True)
   vaSpread1.text = ""
   
   vaSpread1.Col = 13
   vaSpread1.text = ""
   
   RS.MoveNext
     
Loop
    
RS.Close
Set RS = Nothing

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS                As New ADODB.Recordset
Dim i                 As Long
Dim codgru            As Long
Dim MyBuffer          As String
Dim Sql               As String
Dim EstDia            As Boolean
Dim Lun               As String
Dim Mar               As String
Dim Mie               As String
Dim Jue               As String
Dim Vie               As String
Dim Sab               As String
Dim Dom               As String
Dim DesGrDes          As String
Dim estext            As String
Dim fechamodif        As String
Dim actualiza         As String
Dim buffact           As String
Dim collecact         As String
Dim modulo            As String
Dim buffantes         As String
Dim buffactual        As String
Dim existeActualizado As Integer
Dim fechaavalidar     As String
Dim FechaMenor        As Long

Select Case Button.Index
     
    Case 3 '-------> Modificar
        
        'registrar Log sistema modificar
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
        modo = "M"
        Gl_Ac_Botones Me, 1, 0, modo
    
    Case 7, 10 '-------> Actualizar lista y cancelar

        'registrar Log sistema actualizar lista
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Actualizar"), Me.HelpContextID, "", "", "")

        Moverdetalle codigoceco, nombrecodigoCeco
        Gl_Ac_Botones Me, 1, 13, modo

    Case 12 '------> Confirmar

        fg_carga ""
    
        'registrar Log sistema modificar
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Modificar"), Me.HelpContextID, "", "", "")
    
        '--> Validar que no exista fecha rutas normal

        EstDia = False
        fechaavalidar = ""
        actualiza = 0
        FechaMenor = 0
        xmlfamilia = ""
        xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
            
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            
            vaSpread1.Col = 12 'Marca Actualiza
            fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                
            vaSpread1.Col = 13 'Marca Actualiza
            actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                
            If fechaavalidar <> "0" And actualiza = 1 Then
                
               vaSpread1.Col = 12 'fecha
               fechamodif = Format(vaSpread1.text, "YYYYMMDD")
                  
               If Format(vaSpread1.text, "YYYYMMDD") <= FechaMenor Or FechaMenor = 0 Then
               
                  FechaMenor = Format(vaSpread1.text, "YYYYMMDD")
                  Fecha = vaSpread1.text
               
               End If
               
               vaSpread1.Col = 2 'ceco
               codgrupo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                  
               xmlfamilia = xmlfamilia & " <RutaCeco"
               xmlfamilia = xmlfamilia & " Ceco = " & Chr(34) & codigoceco & Chr(34)
               xmlfamilia = xmlfamilia & " GrpDes = " & Chr(34) & codgrupo & Chr(34)
               xmlfamilia = xmlfamilia & " Fec = " & Chr(34) & fechamodif & Chr(34)
               xmlfamilia = xmlfamilia & "/>"
                
            
            End If
                
        Next i
            
        xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
                  
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        vg_BorradoDatos = False
        
        Set RS = vg_db.Execute("sgpadm_Sel_XmlValidasiExisteRutaNormal '" & xmlfamilia & "'")
        If Not RS.EOF Then
        
           If RS!Ceco > 0 And RS!estado <> "0" Then
                
              fg_descarga
              RS.Close
              Set RS = Nothing
              
              Frame2.Visible = True
'              vaSpread1.Enabled = False
              vaSpread1.Row = -1
              vaSpread1.Col = -1
              vaSpread1.Lock = True
              Toolbar1.Enabled = False
              Label2.Caption = "Existen rutas normales apartir de la fechas " & Fecha & VgLinea & "solicitadas por pantalla. Desea borrar las rutas normales S/N??? "
              
              vg_BorradoDatos = False
              Exit Sub
            
           Else
           
              vg_BorradoDatos = False
           
           End If
               
        End If
           
        ' Close ADO objects
        RS.Close
        Set RS = Nothing
        
'        Call Command1(0)
        vg_BorradoDatos = False
    
        '--> Validar que exista un dato seleccionado como modificado
'        vaSpread1.Row = -1
'        vaSpread1.Col = -1
'        vaSpread1.Lock = False
'        Toolbar1.Enabled = True
        
        EstDia = False
        For i = 1 To vaSpread1.MaxRows
    
            vaSpread1.Row = i
            vaSpread1.Col = 13
    
            If vaSpread1.text = "1" Then
    
               EstDia = True
    
            End If
    
        Next i
    
        If EstDia = False Then
    
           MsgBox "Debe haber a lo menos un datos modificado. Proceso cancelado", vbCritical + vbOKOnly, MsgTitulo
           Exit Sub
    
        End If
    
        '--> Validar que los días sean igual a los parametros CD
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Sql = "sgpadm_Sel_TraerDiaParametroCD '" & codigoceco & "'"
        Set RS = vg_db.Execute(Sql)
    
        If Not RS.EOF Then
    
           Lun = IIf(RS("lu") = False, 0, 1)
           Mar = IIf(RS("ma") = False, 0, 1)
           Mie = IIf(RS("mi") = False, 0, 1)
           Jue = IIf(RS("ju") = False, 0, 1)
           Vie = IIf(RS("vi") = False, 0, 1)
           Sab = IIf(RS("sa") = False, 0, 1)
           Dom = IIf(RS("do") = False, 0, 1)
    
        End If
    
        RS.Close
        Set RS = Nothing
    
        EstDia = True
        For i = 1 To vaSpread1.MaxRows
    
            vaSpread1.Row = i
            vaSpread1.Col = 13
    
            If vaSpread1.text = "1" Then
    
               vaSpread1.Col = 3
               DesGrDes = vaSpread1.text
    
               vaSpread1.Col = 4
               If vaSpread1.text <> Lun And Lun <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Lunes no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 5
               If vaSpread1.text <> Mar And Mar <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Martes no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 6
               If vaSpread1.text <> Mie And Mie <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Miercoles no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 7
               If vaSpread1.text <> Jue And Jue <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Jueves no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 8
               If vaSpread1.text <> Vie And Vie <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Viernes no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 9
               If vaSpread1.text <> Sab And Sab <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Sabado no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
               vaSpread1.Col = 10
               If vaSpread1.text <> Dom And Dom <> "1" Then
    
                  MsgBox "Grupo : " & DesGrDes & " del día Domingo no corresponde a los parametros CD", vbCritical + vbOKOnly, MsgTitulo
                  EstDia = False
    
               End If
    
            End If
    
        Next i
    
        If EstDia = False Then
    
           fg_descarga
           Exit Sub
    
        End If
    
        EstDia = False
    
        '--> Grabar solo parametros grupo despacho
        For i = 1 To vaSpread1.MaxRows
    
            vaSpread1.Row = i
            vaSpread1.Col = 13
    
            If Trim(vaSpread1.text) = "1" Then
    
               Lun = 0
               Mar = 0
               Mie = 0
               Jue = 0
               Vie = 0
               Sab = 0
               Dom = 0
               codgru = 0
    
               vaSpread1.Col = 2
               codgru = vaSpread1.text
    
               vaSpread1.Col = 4
               Lun = vaSpread1.text
    
               vaSpread1.Col = 5
               Mar = vaSpread1.text
    
               vaSpread1.Col = 6
               Mie = vaSpread1.text
    
               vaSpread1.Col = 7
               Jue = vaSpread1.text
    
               vaSpread1.Col = 8
               Vie = vaSpread1.text
    
               vaSpread1.Col = 9
               Sab = vaSpread1.text
    
               vaSpread1.Col = 10
               Dom = vaSpread1.text
    
               '--> Grabar Parametros grupo despacho
               If Lun <> 0 Or Mar <> 0 Or Mie <> 0 Or Jue <> 0 Or Vie <> 0 Or Sab <> 0 Or Dom <> 0 Then
    
                  If RS.State = 1 Then RS.Close
                  RS.CursorLocation = adUseClient
                  vg_db.CursorLocation = adUseClient
    
                  Sql = ""
                  Sql = " sgpadm_Ins_ParametroGrupoDespachoCeco "
                  Sql = Sql & " '" & codigoceco & "',"
                  Sql = Sql & codgru & ","
                  Sql = Sql & Lun & ","
                  Sql = Sql & Mar & ","
                  Sql = Sql & Mie & ","
                  Sql = Sql & Jue & ","
                  Sql = Sql & Vie & ","
                  Sql = Sql & Sab & ","
                  Sql = Sql & Dom
    
                  Set RS = vg_db.Execute(Sql)
    
                  If Not RS.EOF Then
    
                     If RS(0) > 0 Then
    
                        RS.Close
                        Set RS = Nothing
                        MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
                        Exit Sub
    
                     Else
    
                        EstDia = True
    
                     End If
    
                  End If
                  RS.Close
                  Set RS = Nothing
    
              Else
    
                 If RS.State = 1 Then RS.Close
                 RS.CursorLocation = adUseClient
                 vg_db.CursorLocation = adUseClient
    
                 Sql = ""
                 Sql = " sgpadm_Del_ParamtroGrupoDespachoCecos "
                 Sql = Sql & " '" & codigoceco & "', "
                 Sql = Sql & " " & codgru & " "
    
                 Set RS = vg_db.Execute(Sql)
    
                 If Not RS.EOF Then
    
                    If RS(0) > 0 Then
    
                       RS.Close
                       Set RS = Nothing
                       MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, Me.Caption
                       Exit Sub
    
                    Else
    
                       EstDia = True
    
                    End If
    
                 End If
                 RS.Close
                 Set RS = Nothing
    
              End If
    
            End If
    
        Next i
    
        '--> Grabar solo parametros grupo despacho
        FechaMenor = 0
        xmlfamilia = ""
        xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
    
        For i = 1 To vaSpread1.MaxRows
    
           estext = False
           vaSpread1.Row = i
    
           vaSpread1.Col = 12 'Marca Actualiza
           fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
            vaSpread1.Col = 13 'Marca Actualiza
            actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
            If fechaavalidar <> "0" And actualiza = 1 Then
    
              vaSpread1.Col = 12 'fecha
              fechamodif = Format(vaSpread1.text, "YYYYMMDD")
    
              If Format(vaSpread1.text, "YYYYMMDD") < FechaMenor Or FechaMenor = 0 Then
    
                 FechaMenor = fechamodif
    
              End If
    
              vaSpread1.Col = 2 'ceco
              codgrupo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
    
              xmlfamilia = xmlfamilia & " <RutaCeco"
              xmlfamilia = xmlfamilia & " Ceco = " & Chr(34) & codigoceco & Chr(34)
              xmlfamilia = xmlfamilia & " GrpDes = " & Chr(34) & codgrupo & Chr(34)
              xmlfamilia = xmlfamilia & "/>"
    
            End If
    
        Next i
    
        xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
    
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
    
        Sql = "sgpadm_Ins_XmlGenerarRutaGrupoDespachoCeco "
        Sql = Sql & " '" & xmlfamilia & "',"
        Sql = Sql & fechamodif & ","
        Sql = Sql & fecha_parametro & ","
        Sql = Sql & IIf(vg_BorradoDatos, 1, 0) & ", "
        Sql = Sql & FechaMenor & ", "
        Sql = Sql & vg_NUsr
    
        Set RS = vg_db.Execute(Sql)
        If Not RS.EOF Then
    
           If RS(0) > 0 Then
    
              RS.Close
              Set RS = Nothing
              MsgBox RS(1)
              Exit Sub
    
           End If
    
        End If
    
        RS.Close
        Set RS = Nothing
    
        'registrar Log sistema actualizar lista
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Me.HelpContextID, "", "", "")
    
        vaSpread1.Row = -1
        vaSpread1.Col = -1
        vaSpread1.Lock = True
        Toolbar1.Enabled = False
        
        MsgBox "Proceso finalizado sin problema", vbInformation + vbOKOnly, Me.Caption
    
        vaSpread1.Row = -1
        vaSpread1.Col = -1
        vaSpread1.Lock = False
        Toolbar1.Enabled = True
        
        modo = ""
        Gl_Ac_Botones Me, 1, 13, modo
    
        fg_descarga
    '
    Case 15 '-------> Imprimir

        'registrar Log sistema actualizar lista
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), Me.HelpContextID, "", "", "")

        I_ParametroGrupoDespachoCeco codigoceco, nombrecodigoCeco

    Case 18 '-------> Salir

        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga

vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = False
Toolbar1.Enabled = True

MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim Sql As String

If vaSpread1.MaxRows < 1 Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol
If vaSpread1.Lock = True Then Exit Sub

Select Case Col

    Case 4, 5, 6, 7, 8, 9, 10
               
        If fecha_parametro <> "" Then
           
           Sql = " sgpadm_sel_maximafechadesdedelcecos " & codigoceco
           
           Set RS = vg_db.Execute(Sql)
                  
           If Not RS.EOF Then
              
              If IsNull(RS(0)) Then
                 
                 vaSpread1.Col = 12 ' Fecha Hastas
                 vaSpread1.Lock = False
                 vaSpread1.CellType = CellTypeDate
                 vaSpread1.TypeCurrencyMin = fecha_parametro
                 vaSpread1.text = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
                 vaSpread1.SetFocus
                       
                 vaSpread1.Col = 13
                 vaSpread1.text = "1"
                       
              Else
                    
                  vaSpread1.Col = 12 ' Fecha Hastas
                  vaSpread1.Lock = False
                  vaSpread1.CellType = CellTypeDate
                  vaSpread1.text = Format(RS(0), "DD/MM/YyyY")
                  vaSpread1.SetFocus
                       
                  vaSpread1.Col = 13
                  vaSpread1.text = "1"
                  
              End If
                     
           End If
                 
           RS.Close
                 
        End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim Sql As String

If vaSpread1.MaxRows < 1 Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol
If vaSpread1.Lock = True Then Exit Sub

Select Case Col

    Case 4, 5, 6, 7, 8, 9, 10
               
        If fecha_parametro <> "" Then
           
           Sql = " sgpadm_sel_maximafechadesdedelcecos " & codigoceco
           
           Set RS = vg_db.Execute(Sql)
                  
           If Not RS.EOF Then
              
              If IsNull(RS(0)) Then
                 
                 vaSpread1.Col = 12 ' Fecha Hastas
                 vaSpread1.Lock = False
                 vaSpread1.CellType = CellTypeDate
                 vaSpread1.TypeCurrencyMin = fecha_parametro
                 vaSpread1.text = Mid(fecha_parametro, 7, 2) + "/" + Mid(fecha_parametro, 5, 2) + "/" + Mid(fecha_parametro, 1, 4)
                 vaSpread1.SetFocus
                       
                 vaSpread1.Col = 13
                 vaSpread1.text = "1"
                       
              Else
                    
                  vaSpread1.Col = 12 ' Fecha Hastas
                  vaSpread1.Lock = False
                  vaSpread1.CellType = CellTypeDate
                  vaSpread1.TypeDateFormat = TypeDateFormatDDMMYY 'SS_CELL_DATE_FORMAT_DDMONYY
                  'Set minimum date
                  vaSpread1.TypeDateMin = Str(Format(RS(0), "mmddyyyy"))
                  vaSpread1.text = Format(RS(0), "DD/MM/YyyY")
                  vaSpread1.SetFocus
                       
                  vaSpread1.Col = 13
                  vaSpread1.text = "1"
                  
              End If
                     
           End If
                 
           RS.Close
                 
        End If

        If modo = "" Then modo = "M"
    
        If Toolbar1.Buttons(12).Visible = False Then
       
           Gl_Ac_Botones Me, 1, 0, modo
    
        End If
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function ValidarRutasNormal() As Boolean

'Dim RS            As New ADODB.Recordset
'Dim Sql           As String
'Dim EstDia        As Boolean
'Dim fechaavalidar As String
'Dim actualiza     As Long
'Dim xmlfamilia    As String
'Dim i             As Long
'Dim codgrupo      As Long
'Dim FechaMenor    As Long
'Dim Fecha         As String
'
''Dim xlApp         As Object
''Dim xlWb          As Object
''Dim xlWs          As Object
'
'ValidarRutasNormal = True
'
'EstDia = False
'fechaavalidar = ""
'actualiza = 0
'FechaMenor = 0
'xmlfamilia = ""
'xmlfamilia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
'xmlfamilia = xmlfamilia & "<GrabaRutaCeco>"
'
'For i = 1 To vaSpread1.MaxRows
'
'    vaSpread1.Row = i
'
'    vaSpread1.Col = 12 'Marca Actualiza
'    fechaavalidar = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'    vaSpread1.Col = 13 'Marca Actualiza
'    actualiza = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'    If fechaavalidar <> "0" And actualiza = 1 Then
'
'       vaSpread1.Col = 12 'fecha
'       fechamodif = Format(vaSpread1.text, "YYYYMMDD")
'
'       If Format(vaSpread1.text, "YYYYMMDD") <= FechaMenor Or FechaMenor = 0 Then
'
'          FechaMenor = Format(vaSpread1.text, "YYYYMMDD")
'          Fecha = vaSpread1.text
'
'       End If
'
'       vaSpread1.Col = 2 'ceco
'       codgrupo = IIf(vaSpread1.text = "", 0, vaSpread1.text)
'
'       xmlfamilia = xmlfamilia & " <RutaCeco"
'       xmlfamilia = xmlfamilia & " Ceco = " & Chr(34) & codigoceco & Chr(34)
'       xmlfamilia = xmlfamilia & " GrpDes = " & Chr(34) & codgrupo & Chr(34)
'       xmlfamilia = xmlfamilia & " Fec = " & Chr(34) & fechamodif & Chr(34)
'       xmlfamilia = xmlfamilia & "/>"
'
'
'    End If
'
'Next i
'
'xmlfamilia = xmlfamilia & "</GrabaRutaCeco>"
'
'If RS.State = 1 Then RS.Close
'RS.CursorLocation = adUseClient
'vg_db.CursorLocation = adUseClient
'
''Sql = ""
''Sql = "sgpadm_Sel_XmlValidasiExisteRutaNormal "
''Sql = Sql & " '" & xmlfamilia & "'"
''
'
'Set RS = vg_db.Execute("sgpadm_Sel_XmlValidasiExisteRutaNormal '" & xmlfamilia & "'")
'If Not RS.EOF Then
'
'   If RS!Ceco > 0 And RS!estado <> "0" Then
'
'
'      fg_descarga
'      If MsgBox("Existen rutas normales apatir de la fechas, solicitadas por pantalla. Desea borrar las rutas normales S/N???", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
'
''      fg_descarga
''      Frame2.Visible = True
''      Label2.Caption = "Existen rutas normales apartir de la fechas " & Fecha & VgLinea & "solicitadas por pantalla. Desea borrar las rutas normales S/N??? "
'
''      vg_BorradoDatos = False
''      ValidarRutasNormal = False
''      Exit Function
''
''          Else
''
''             vg_BorradoDatos = True
''
''          End If
'
'   End If
'
'End If
'
'' Close ADO objects
'RS.Close
'Set RS = Nothing
'
'
'If Est Then
'
'   'Label2.Caption = "Existen rutas normales apartir de la fechas " & Fecha & VgLinea & "solicitadas por pantalla. Desea borrar las rutas normales S/N??? "
'   If MsgBox("Existen rutas normales apatir de la fechas " & Fecha & VgLinea & " solicitadas por pantalla. Desea borrar las rutas normales S/N???", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
'
'      vg_BorradoDatos = False
'      ValidarRutasNormal = False
'      Exit Function
'
'   End If
'End If
'
'
'Exit Function
'Man_Error:
'
'ValidarRutasNormal = False
'fg_descarga
'MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
'ins_log_error Date & Time & Err & ":  " & Error$(Err)
'
End Function
