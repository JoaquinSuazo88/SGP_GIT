VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form P_GReTec 
   Caption         =   "Generación Maestro Reecta Hacia Tecfood"
   ClientHeight    =   6765
   ClientLeft      =   4005
   ClientTop       =   2445
   ClientWidth     =   6090
   LinkTopic       =   "Form1"
   ScaleHeight     =   6765
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Recetas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Index           =   1
      Left            =   30
      TabIndex        =   0
      Top             =   600
      Width           =   6015
      Begin RichTextLib.RichTextBox RichTextBox1 
         Height          =   255
         Left            =   5520
         TabIndex        =   9
         Top             =   5880
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   450
         _Version        =   393217
         TextRTF         =   $"P_GReTec.frx":0000
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1005
         Index           =   3
         Left            =   570
         TabIndex        =   1
         Top             =   240
         Width           =   4725
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            ItemData        =   "P_GReTec.frx":008B
            Left            =   1680
            List            =   "P_GReTec.frx":0098
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            Top             =   600
            Width           =   2895
            _Version        =   196608
            _ExtentX        =   5106
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
         Begin VB.Label Label1 
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
            Height          =   315
            Index           =   1
            Left            =   150
            TabIndex        =   5
            Top             =   345
            Width           =   1485
         End
         Begin VB.Label Label1 
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
            Height          =   315
            Index           =   2
            Left            =   150
            TabIndex        =   4
            Top             =   660
            Width           =   1470
         End
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   1
         Left            =   660
         TabIndex        =   6
         Top             =   5880
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4455
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   1410
         Width           =   5655
         _Version        =   393216
         _ExtentX        =   9975
         _ExtentY        =   7858
         _StockProps     =   64
         BackColorStyle  =   1
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
         MaxCols         =   6
         MaxRows         =   2
         SpreadDesigner  =   "P_GReTec.frx":00B5
         TextTip         =   2
         TextTipDelay    =   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "P_GReTec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Est As Boolean, estado As Boolean
Dim codtip As Long, ibusca As Long, i As Long, j As Long
Dim aAp As String

Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
Select Case Index
Case 1
    fptnombre(1).Enabled = True
    fptnombre(1).text = ""
   If vaSpread1(1).MaxRows > 0 Then vaSpread1(1).SetFocus
End Select
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 7275
Me.Width = 6225
Me.HelpContextID = vg_OpcM
Msgtitulo = "Generación Maestro Producto Hacia Tecfood"
fg_centra Me
Est = True: ibusca = 0
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): btnX.Visible = True: btnX.ToolTipText = "Procesar": btnX.Enabled = True
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(1).ListIndex = 1
MoverDatoGrilla
Est = False
SendKeys "+{Tab}"
End Sub

Private Sub fpTnombre_Change(Index As Integer)
Select Case Index
Case 1
    If LimpiaDato(Trim(fptnombre(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    findstring = Trim(fptnombre(1).text)
    If fptnombre(1).text = "" Then
       vaSpread1(1).Visible = False
       swactiva = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If swactiva = 0 Then swactiva = 1
       Next i
       vaSpread1(1).Visible = True
    Else
       swactiva = 0
       vaSpread1(1).Visible = False
       irow = 0
       For i = 1 To vaSpread1(1).MaxRows
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 1
           vaSpread1(1).text = "0"
           If Combo1(1).ItemData(Combo1(1).ListIndex) = 0 Or Combo1(1).ItemData(Combo1(1).ListIndex) = 1 Then
              vaSpread1(1).Col = IIf(Combo1(1).ItemData(Combo1(1).ListIndex) = 0, 2, 3)
           Else
              findstring = Trim(Str(codtip))
              vaSpread1(1).Col = 4
           End If
           sourcestring = Trim(vaSpread1(1).text)
           indactivo = UCase(Trim(sourcestring)) Like "*" & UCase(findstring) & "*"
           If indactivo = -1 Then
              If swactiva = 0 Then swactiva = 1
              If vaSpread1(1).RowHidden = True Then
                 vaSpread1(1).RowHidden = False
              ElseIf vaSpread1(1).RowHidden = True Then
                 vaSpread1(1).RowHidden = False
              End If
              irow = irow + 1
           Else
              If vaSpread1(1).RowHidden = False Then vaSpread1(1).RowHidden = True
           End If
       Next i
       vaSpread1(1).Visible = True
       End If
End Select
End Sub

Private Sub fptnombre_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 34 And irow > 0 Then vaSpread1(Index).SetFocus
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    If vaSpread1(1).MaxRows < 1 Then Exit Sub
    Dim i As Long, j As Long
    Dim isel As Boolean
    isel = False
    For i = 1 To vaSpread1(1).MaxRows
        vaSpread1(1).Row = i
        vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then isel = True: Exit For
    Next i
    If isel = False Then MsgBox "Debe Seleccionar A lo menor un producto", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    Frame1(1).Enabled = False
    Bar1(1).Visible = True
    Bar1(1).Value = 0
    Dim codrec As String, nomrec As String, nomfan As String, StrFam As String, StrFamb As String
    Dim codcatdie As Long, codtipplato As Long, coddi1 As Long, coddi2 As Long, codti1 As Long, codti2 As Long, codti3 As Long
    Dim opgraba  As Boolean
    dBo = dir_trabajo + BaseDeDato
    For j = 1 To vaSpread1(1).MaxRows
        Bar1(1).Value = Val((j / vaSpread1(1).MaxRows) * 100)
        vaSpread1(1).Row = j: vaSpread1(1).Col = 1
        If vaSpread1(1).text = "1" And vaSpread1(1).RowHidden = False Then
           DoEvents
           
'tecfood           vg_dbtec.BeginTrans

'           vg_db.BeginTrans
' Actualizar metodo preparación tecfood
           codrec = 0: vaSpread1(1).Col = 2: codrec = vaSpread1(1).text
            RichTextBox1.TextRTF = ""
            RS.Open "select rec_metpre from b_receta Where rec_codigo=" & codrec & " and rec_metpre Is Not Null", vg_db, adOpenStatic
            If Not RS.EOF Then RichTextBox1.TextRTF = RS!rec_metpre
            RS.Close: Set RS = Nothing
'tecfood            If Trim(RichTextBox1.TextRTF) <> "" Then
'tecfood               Dim strSQL As String
'tecfood               Dim rsUpd As New ADODB.Recordset
'tecfood               Dim oStr As New ADODB.Stream
              
'tecfood               RS.Open "select cdprato from prato where usr_cdpratinte='" & codrec & "' and cdnvprato='6' and cdfilauxprat='P'", vg_dbtec, adOpenStatic
'tecfood               If Not RS.EOF Then
'tecfood                  strSQL = "SELECT * "
'tecfood                  strSQL = strSQL & "FROM prepprapad "
'tecfood                  strSQL = strSQL & "WHERE cdprato = '" & RS!cdprato & "' and cdfilauxprat='P'"
'tecfood                  rsUpd.Open strSQL, vg_dbtec, adOpenDynamic, adLockOptimistic
               
'tecfood                  RichTextBox1.SaveFile ("C:\Temp.del")
               
'tecfood                  oStr.Open
'tecfood                  oStr.Type = adTypeBinary
'tecfood                  oStr.LoadFromFile "C:\Temp.del"
'tecfood                  If rsUpd.EOF Then
'tecfood                     rsUpd.AddNew
'tecfood                     rsUpd("cdfilauxprat").Value = "P"
'tecfood                     rsUpd("cdprato").Value = RS!cdprato
'tecfood                     rsUpd("nrprepprat").Value = "01"
'tecfood                     rsUpd("idprepprat").Value = "1"
'tecfood                     rsUpd("txprepprat").Value = oStr.Read
'tecfood                  Else
'tecfood                     rsUpd("txprepprat").Value = oStr.Read
'tecfood                  End If
'tecfood                  rsUpd.Update
'tecfood                  rsUpd.Close
'tecfood                  oStr.Close
'tecfood                  Set oStr = Nothing
'tecfood                  Set rsUpd = Nothing
'tecfood                  Kill "C:\Temp.del"
'tecfood               End If
'tecfood               RS.Close: Set RS = Nothing
'tecfood            End If

' Actualizar Recetas tecfood
'           vaSpread1(1).Col = 2: codrec = vaSpread1(1).Text: vg_codreceta = vaSpread1(1).Text
'           vaSpread1(1).Col = 3: nomrec = Trim(vaSpread1(1).Text)
'           vaSpread1(1).Col = 4: nomfan = Trim(vaSpread1(1).Text)
'           vaSpread1(1).Col = 5: codcatdie = Val(vaSpread1(1).Text)
'           vaSpread1(1).Col = 6: codtipplato = Val(vaSpread1(1).Text)
'           vaSpread1(1).SetActiveCell 2, vaSpread1(1).Row
'           opgraba = False
'            '------- Validar categoria dietetica
'           StrFam = fg_BuscaCodArbol(codcatdie, "a_recetacatdie", "car_codigo")
'           If Len(StrFam) <> 0 Then
'              Do While InStr(StrFam, ";") <> 0
'                 StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
'                 StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
'                 coddi1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
'                 If Val(Mid(StrFamb, 1)) = 0 Then opgraba = True: Exit Do
'                 coddi2 = Val(Mid(StrFamb, 1)): codcatdie = Val(Mid(StrFamb, 1))
'              Loop
'           End If
'           '------- Fin validar categoria dietetica
'           '------- Validar tipo de plato
'           StrFam = fg_BuscaCodArbol(codtipplato, "a_recetatippla", "tip_codigo")
'           If Len(StrFam) <> 0 Then
'              Do While InStr(StrFam, ";") <> 0
'                 StrFamb = Mid(StrFam, 1, InStr(StrFam, ";") - 1)
'                 StrFam = IIf(Len(StrFam) > InStr(StrFam, ";"), Mid(StrFam, InStr(StrFam, ";") + 1), "")
'                 codti1 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
'                 codti2 = Val(Mid(StrFamb, 1, InStr(StrFamb, "&") - 1)): StrFamb = Mid(StrFamb, InStr(StrFamb, "&") + 1)
'                 If codti2 = 0 Then opgraba = True: Exit Do
'                 codti3 = Val(Mid(StrFamb, 1)): codtipplato = IIf(codti3 < 1, codti2, codti3)
'              Loop
'           End If
'            '------- Fin validar tipo de plato
'           If Not opgraba Then
'              If GrabarRecetaTecfood(CStr(codrec), LimpiaDato(Trim(nomrec)), LimpiaDato(Trim(nomfan)), CStr(coddi1), CStr(coddi2), CStr(codti1), CStr(codti2), CStr(codti3), Me, "1") Then vg_dbtec.RollbackTrans: vg_db.RollbackTrans: Exit Sub
'           End If
'           vg_dbtec.CommitTrans
'           vg_db.CommitTrans
        End If
    Next j
    fg_descarga
    Bar1(1).Visible = False
    MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, Msgtitulo
    Frame1(1).Enabled = True
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
fg_descarga
Frame1(1).Enabled = True
Bar1(1).Visible = False
Man_Error:
Select Case Err
Case 35764
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans
    DoEvents
    For i = 1 To 1000000
    Next i
    Resume
Case 76
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans
    Resume Next
Case -2147467259
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub
Case 3034
'tecfood    vg_dbtec.RollbackTrans
    vg_db.RollbackTrans: Exit Sub
End Select
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread1_ButtonClicked(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If Index = 1 Then Exit Sub
vaSpread1(1).Row = Row
Select Case Col
Case 1
    If Row = 0 Or Row = -1 Then x = vaSpread1(1).MaxRows: j = 1 Else x = vaSpread1(1).Row: j = vaSpread1(1).Row
    fg_descarga
End Select
End Sub

Sub MoverDatoGrilla()
On Error GoTo Man_Error
fg_carga "": estado = True: i = 1
'------- Mover recetas
If Est Then
   vaSpread1(1).MaxRows = 0
   RS.Open "select rec_codigo, rec_nombre, rec_nomfan, rec_catdie, rec_tippla from b_receta order by rec_nombre, rec_catdie, rec_tippla", vg_db, adOpenStatic
   If Not RS.EOF Then
      Do While Not RS.EOF
         vaSpread1(1).MaxRows = vaSpread1(1).MaxRows + 1
         vaSpread1(1).Row = vaSpread1(1).MaxRows
              
         vaSpread1(1).Col = 2
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).TypeSpin = False
         vaSpread1(1).TypeIntegerSpinInc = 1
         vaSpread1(1).TypeIntegerSpinWrap = False
         vaSpread1(1).text = RS!rec_codigo

         vaSpread1(1).Col = 3
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = Trim(RS!rec_nombre)
         
         vaSpread1(1).Col = 4
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = Trim(RS!rec_nomfan)
         
         vaSpread1(1).Col = 5
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = RS!rec_catdie
         
         vaSpread1(1).Col = 6
         vaSpread1(1).TypeHAlign = TypeHAlignLeft
         vaSpread1(1).text = RS!rec_tippla
         
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
End If
vaSpread1(1).SetActiveCell 1, i
fg_descarga
Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)
If Col = 1 And Row = 0 Then vaSpread1(Index).Row = -1: vaSpread1(Index).Col = 1: vaSpread1(Index).text = IIf(vaSpread1(Index).Value = "1", "0", "1")
End Sub

Private Sub vaSpread1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Or KeyCode = 13 Then Exit Sub
If TeclasNoPermitidas(KeyCode) = True Then fptnombre(Index).text = IIf(KeyCode = 8, fptnombre(Index).text, fptnombre(Index).text & Chr(KeyCode)): fptnombre(Index).SetFocus: fptnombre(Index).SelStart = Len(fptnombre(Index).text)
End Sub
