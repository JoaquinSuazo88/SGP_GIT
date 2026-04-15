VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_ImpTxt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Formatos Sap"
   ClientHeight    =   2370
   ClientLeft      =   4680
   ClientTop       =   2400
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   8625
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1635
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   5895
         _Version        =   196608
         _ExtentX        =   10398
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
         ThreeDInsideHighlightColor=   -2147483633
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
         ButtonStyle     =   2
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
         NoSpecialKeys   =   1
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
         ControlType     =   3
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
      Begin MSComDlg.CommonDialog CD 
         Left            =   120
         Top             =   1680
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar prbStatus 
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   1320
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         Caption         =   "lblStatus"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   1080
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Archivo Origen"
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
         Left            =   240
         TabIndex        =   3
         Top             =   435
         Width           =   1275
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   2370
      Left            =   7995
      TabIndex        =   1
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   4180
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "P_ImpTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MsgTitulo As String
Dim nomarc    As String
Dim myPath    As String
Dim Op_Carga  As String

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

fg_centra Me
fg_carga ""

MsgTitulo = IIf(Op_Carga = "1", "Importar Formato SAP", "Importar Formato SAP JUSTICIA")

Me.Caption = MsgTitulo
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = True
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Toolbar1.Buttons(1).Enabled = False
lblStatus.Visible = False: prbStatus.Visible = False
fg_descarga

End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

Dim fromRihgt As String
CD.Filter = "Todos los archivos (*.txt)|*.txt"
CD.DefaultExt = "*.txt"
'CD.InitDir = dir_trabajo ' & "Actualizar"
CD.FilterIndex = 2
CD.Flags = cdlOFNFileMustExist
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.ShowOpen
If CD.FileName = "" Then
   
   Toolbar1.Buttons(1).Enabled = False
   fpText1.text = ""

Else
    
    fromRight = InStrRev(CD.FileName, "\", , vbTextCompare)
    If fromRight > 1 Then
       
       myPath = Left(CD.FileName, fromRight)
    
    End If
   Toolbar1.Buttons(1).Enabled = True
   fpText1.text = CD.FileName 'Dir(CD.FileName)
'   CD

End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim esterr    As Boolean
Dim lngRow    As Long
Dim codsap    As String
Dim auxcodsap As String
Dim codsgp    As String
Dim codsiges  As String
Dim nomtaberr As String
Dim MyBuffer  As String
Dim i         As Long

Select Case Button.Index

Case 1 And Op_Carga = "1"
    
    fg_carga ""
    prbStatus.Max = 1
    If Trim(fpText1.text) = "" Then MsgBox "Debe seleccionar archivo(TXT)", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    
    Open fpText1.text For Input As #1
    Do While Not EOF(1)
        
        Line Input #1, strLineReg: prbStatus.Max = prbStatus.Max + 1
    
    Loop
    Close #1
    
    lblStatus.Visible = True: prbStatus.Visible = True: prbStatus.Min = 0: lngRow = 0
    Open myPath & "LogError" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & "_" & Mid(Dir(CD.FileName), 1, Len(Dir(CD.FileName)) - 4) & ".txt" For Output As #2 'Crear archivos de errores
    Open fpText1.text For Input As #1
    esterr = True
    
    If Not EOF(1) Then
        
        Do While Not EOF(1)
            
            Line Input #1, strLineReg
            lblStatus.Caption = "Procesando registros, " & Trim(Str(lngRow)) & "/" & Trim(Str(prbStatus.Max))
            DoEvents
            codsap = Mid(strLineReg, 1, InStr(strLineReg, ";") - 1)
            codsgp = Mid(strLineReg, InStr(strLineReg, ";") + 1)
            Set RS = vg_db.Execute("sgpadm_s_formatocomprassap 5, '" & codsap & "', '" & codsgp & "'")
            
            If Not RS.EOF Then
               
               If Trim(RS(0)) = "1" And Trim(RS(1)) = "1" Then 'Si ambos campos estan en uno quiere decir que existe en ambas bases.
                  
                  vg_db.Execute ("sgpadm_i_formatocompras_sap_sgp '" & codsap & "', '" & codsgp & "'")
               
               Else
                  
                  Print #2, codsap & "|" & IIf(Trim(RS(0)) = "1", "Código correcto", "Código sap no existe") & "|" & codsgp & "|" & IIf(Trim(RS(1)) = "1", "Código Correcto", "Código sgp No existe")
                  esterr = False
               
               End If
            
            End If
            
            RS.Close: Set RS = Nothing
            strLineReg = ""
            lngRow = lngRow + 1
            prbStatus.Value = lngRow
        
        Loop
    
    End If
    lblStatus.Visible = False: prbStatus.Visible = False
    Close #1
    Close #2
    
    '-------> Borrar tablas de errores si no existen problema
    If esterr Then
       
       If Dir(myPath & "LogError" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & "_" & Mid(Dir(CD.FileName), 1, Len(Dir(CD.FileName)) - 4) & ".txt") <> "" Then Kill myPath & "LogError" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & "_" & Mid(Dir(CD.FileName), 1, Len(Dir(CD.FileName)) - 4) & ".txt"
    
    Else
       
       nomtaberr = myPath & "LogError" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & "_" & Mid(Dir(CD.FileName), 1, Len(Dir(CD.FileName)) - 4) & ".txt"
    
    End If
    fg_descarga
    MsgBox IIf(esterr, "Proceso finalizo sin problema..", "Proceso finalizo con problema " & VgLinea & VgLinea & " El archivo con errores fue generado en la siguiente carpeta " & VgLinea & nomtaberr), IIf(esterr, vbInformation, vbCritical) + vbOKOnly, MsgTitulo: Exit Sub

Case 1 And Op_Carga = "2"
    
    fg_carga ""
    prbStatus.Max = 1
    i = 1
    Dim LineaSplitted() As String
    
    If Trim(fpText1.text) = "" Then MsgBox "Debe seleccionar archivo(TXT)", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    
    If Not Unix2Dos(fpText1.text) Then MsgBox "Problema formato", vbInformation + vbOKOnly, MsgTitulo: Exit Sub

    Open fpText1.text For Input As #1
    Do While Not EOF(1)
        
        Line Input #1, strLineReg
        prbStatus.Max = prbStatus.Max + 1
    
        If i = 1 Then
            
            LineaSplitted = Split(strLineReg, "|")
            If Trim(LineaSplitted(1)) <> "Tipo de Material" And Trim(LineaSplitted(3)) <> "Codigo Material" And Trim(LineaSplitted(13)) <> "Nº Antiguo Material" Then
            
               MsgBox " Proceso Cancelado, formato no corresponde", vbCritical + vbOKOnly, MsgTitulo
               Close #1
               fg_descarga
               Exit Sub
               
            End If
        
        End If
        
        i = i + 1
        
    Loop
    Close #1
    
    If i > 1 Then
       
       For i = LBound(LineaSplitted) To UBound(LineaSplitted)

           LineaSplitted(i) = ""

       Next
    
    ElseIf i = 1 Then
    
       MsgBox " Proceso Cancelado, formato no corresponde o bien esta vacio", vbCritical + vbOKOnly, MsgTitulo
       fg_descarga
       Exit Sub

    End If
    
    lblStatus.Visible = True
    prbStatus.Visible = True
    prbStatus.Min = 0
    lngRow = 0
'    Open myPath & "LogError" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & "_" & Mid(Dir(CD.FileName), 1, Len(Dir(CD.FileName)) - 4) & ".txt" For Output As #2 'Crear archivos de errores
    Open fpText1.text For Input As #1
    esterr = True
    auxcodsap = ""
    codsap = ""
    codsiges = ""
    i = 1
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaCodSiges>"
    
    If Not EOF(1) Then
        
        Do While Not EOF(1)
            
            Line Input #1, strLineReg
                        
            lblStatus.Caption = "Procesando registros, " & Trim(Str(lngRow)) & "/" & Trim(Str(prbStatus.Max))
            DoEvents
            
            For i = LBound(LineaSplitted) To UBound(LineaSplitted)

                LineaSplitted(i) = ""

            Next
            
            If Trim(strLineReg) <> "" Then
               LineaSplitted = Split(strLineReg, "|")
                codsap = LineaSplitted(3)
               codsiges = LineaSplitted(13)
            
            If Trim(codsap) <> "Codigo Material" Then
            
               If codsap <> auxcodsap And Trim(codsiges) <> "" And IsNumeric(Trim(codsiges)) And IsNumeric(Trim(codsap)) Then
               
                  MyBuffer = MyBuffer & " <CodSig"
                  MyBuffer = MyBuffer & " CSap = " & Chr(34) & CDbl(Trim(codsap)) & Chr(34)
                  MyBuffer = MyBuffer & " CSige = " & Chr(34) & CDbl(Trim(codsiges)) & Chr(34)
                  MyBuffer = MyBuffer & "/>"
               
                  auxcodsap = codsap
                  
                  i = i + 1
                  
               End If

            End If
            End If
            strLineReg = ""
            lngRow = lngRow + 1
            prbStatus.Value = lngRow
        
        Loop
    
    End If
    lblStatus.Visible = False: prbStatus.Visible = False
    Close #1
    
    MyBuffer = MyBuffer & "</GrabaCodSiges>"

    Set RS = vg_db.Execute("sgpadm_Ins_XmlFormatoComprasSapSiges '" & MyBuffer & "'")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
      
       Else
       
          MsgBox " Proceso Finalizo sin problema", vbInformation + vbOKOnly, MsgTitulo
       
       End If
       
    End If
    RS.Close: Set RS = Nothing
    fg_descarga

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
fg_descarga
If Err = 5 Then Resume Next: Exit Sub
If Err = 55 Then MsgBox "Archivo esta abierto, proceso cancelado": Exit Sub
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume

End Sub

Sub llena_datos(OpCarga As String)

Op_Carga = OpCarga

End Sub
