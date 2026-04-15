VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_FormatoSalidaDiwo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formato Salida Diwo"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   10605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame6 
         Height          =   435
         Index           =   2
         Left            =   7995
         TabIndex        =   17
         Top             =   5400
         Width           =   1710
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   18
            Top             =   135
            Width           =   1605
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   2
         Left            =   7080
         TabIndex        =   15
         Top             =   5400
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   16
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Index           =   1
         Left            =   5355
         TabIndex        =   13
         Top             =   5400
         Width           =   1710
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   14
            Top             =   135
            Width           =   1605
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   1
         Left            =   4440
         TabIndex        =   11
         Top             =   5400
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   795
         End
      End
      Begin VB.Frame Frame6 
         Height          =   435
         Index           =   0
         Left            =   2595
         TabIndex        =   9
         Top             =   5400
         Width           =   1710
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   1605
         End
      End
      Begin VB.Frame Frame5 
         Height          =   435
         Index           =   0
         Left            =   1680
         TabIndex        =   7
         Top             =   5400
         Width           =   900
         Begin VB.TextBox TextDet1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   795
         End
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
         Height          =   495
         Index           =   1
         Left            =   8520
         TabIndex        =   5
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar Formato"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   7200
         TabIndex        =   4
         Top             =   6000
         Width           =   1215
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3735
         Left            =   240
         TabIndex        =   3
         Top             =   1560
         Width           =   9855
         _Version        =   393216
         _ExtentX        =   17383
         _ExtentY        =   6588
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
         MaxCols         =   8
         SpreadDesigner  =   "I_FormatoSalidaDiwo.frx":0000
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   4020
         TabIndex        =   1
         Top             =   480
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
         BackColor       =   16777215
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
         AutoAdvance     =   0   'False
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
         NullColor       =   -2147483624
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   "10/2019"
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
         AutoMenu        =   0   'False
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   5520
         TabIndex        =   6
         Top             =   480
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Line Line1 
         X1              =   0
         X2              =   10200
         Y1              =   1200
         Y2              =   1200
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
         Left            =   2520
         TabIndex        =   2
         Top             =   525
         Width           =   1230
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "I_FormatoSalidaDiwo.frx":1A27
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "I_FormatoSalidaDiwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim i               As Long
Dim isel            As Boolean
Dim Ceco            As String
Dim Reg             As Long
Dim Ser             As Long
Dim NomArchivoExcel As String
Dim EstTemporal     As Boolean

EstTemporal = False

Select Case Index

    Case 0

        isel = False
        For i = 1 To vaSpread1.MaxRows
        
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" Then
            
               isel = True
               
            End If
        
        Next i

        If Not isel Then

            MsgBox "Debe haber selecionado al menos un datos de la grilla", vbExclamation + vbOKOnly, Me.Caption
            Exit Sub

        End If

        '-------> Guardar nombre archivo excel
        NomArchivoExcel = ""
        CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
        CD.Filter = "Todos los archivos *.csv"
        On Error Resume Next
        CD.ShowSave
                   
        '-------> JPAZ Permite controlar Boton Cancelar
        If Err.Number = 32755 Then
           
           MsgBox "Proceso cancelado"
           Exit Sub
        
        End If
                    
        If CD.FileName = "" Then
           
           MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
           Exit Sub
        
        Else
           
           Extension = ""
           Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
           
           If UCase(Extension) <> "CSV" Then
              
              MsgBox "La extensión del archivo debe ser (*.csv)", vbCritical
              Exit Sub
           
           End If
           NomArchivoExcel = CD.FileName
        
        End If
           
        fg_carga ""
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<ParDiwo>"

        For i = 1 To vaSpread1.MaxRows

            vaSpread1.Row = i
            vaSpread1.Col = 1
    
            If vaSpread1.text = "1" Then
    
                vaSpread1.Col = 2
                Ceco = vaSpread1.text
    
                vaSpread1.Col = 4
                Reg = vaSpread1.text
    
                vaSpread1.Col = 6
                Ser = vaSpread1.text

                MyBuffer = MyBuffer & " <Diwo"
                MyBuffer = MyBuffer & " Cec = " & Chr(34) & Ceco & Chr(34)
                MyBuffer = MyBuffer & " Reg = " & Chr(34) & Reg & Chr(34)
                MyBuffer = MyBuffer & " Ser = " & Chr(34) & Ser & Chr(34)
                MyBuffer = MyBuffer & "/>"

            End If

        Next i

        MyBuffer = MyBuffer & "</ParDiwo>"

        Set RS = vg_db.Execute("sgpadm_Sel_XmlBajadaMinutaDiwo '" & MyBuffer & "', " & Format(fpDateTime1.text, "yyyymm") & "")
        If Not RS.EOF Then
        
           EstTemporal = True
           Open NomArchivoExcel For Output As #1
            
           Do While Not RS.EOF
           
'              Print #1, "CL" & ";" & RS![Ceco] & ";" & RS![CecoDesc] & ";" & RS![Fecha] & ";" & RS![Cód.Regimen] & ";" & RS![Cód.RegimenDesc] & ";" & RS![Cód.Servicio] & ";" & RS![Cód.ServicioDesc] & ";" & RS![Cód.Estructura] & ";" & RS![Cód.EstructuraDesc] & ";" & RS![Cód.Receta] & ";" & RS![Cód.RecetaDesc] & ";" & RS![Alias]
              
              Print #1, RS![Pais] & ";" & RS![Ceco] & ";" & RS![CecoDesc] & ";" & RS![Fecha] & ";" & RS![Cód.Servicio] & ";" & RS![Cód.ServicioDesc] & ";" & RS![Cód.Estructura] & ";" & RS![Cód.EstructuraDesc] & ";" & RS![Cód.Receta] & ";" & RS![Cód.RecetaDesc] & ";" & RS![Alias]
              RS.MoveNext
           
           Loop
            
        End If
        RS.Close
        Set RS = Nothing
        Close #1
        EstTemporal = False
        
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_CSV"), CStr(Me.HelpContextID), "", "", "")
        fg_descarga

        MsgBox "Proceso Termino Correctamente ", vbInformation + vbOKOnly, MsgTitulo
    
    Case 1

        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "", "")
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
    
    If EstTemporal Then
    
       Close #1
        
    End If
    
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()
    
On Error GoTo Man_Error

    Call fg_descarga

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
    Call fg_carga("")
    Me.HelpContextID = vg_OpcM
    MsgTitulo = "Formato Salida Diwo"
    Call fg_centra(Me)
    Let Me.Height = 7380
    Let Me.Width = 10695
    
    fpDateTime1.text = Format(Date, "mm/yyyy")
    vaSpread1.MaxRows = 0

    Call fg_descarga
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub TextDet1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet1(Index).text, ",")

If Index = 2 Then
   
   TextDet1(3).text = ""
   TextDet1(4).text = ""
   TextDet1(5).text = ""
   TextDet1(6).text = ""
   TextDet1(7).text = ""

ElseIf Index = 3 Then
   
   TextDet1(2).text = ""
   TextDet1(4).text = ""
   TextDet1(5).text = ""
   TextDet1(6).text = ""
   TextDet1(7).text = ""

ElseIf Index = 4 Then
   
   TextDet1(2).text = ""
   TextDet1(3).text = ""
   TextDet1(5).text = ""
   TextDet1(6).text = ""
   TextDet1(7).text = ""

ElseIf Index = 5 Then
   
   TextDet1(2).text = ""
   TextDet1(3).text = ""
   TextDet1(4).text = ""
   TextDet1(6).text = ""
   TextDet1(7).text = ""

ElseIf Index = 6 Then
   
   TextDet1(2).text = ""
   TextDet1(3).text = ""
   TextDet1(4).text = ""
   TextDet1(5).text = ""
   TextDet1(7).text = ""

ElseIf Index = 7 Then
   
   TextDet1(2).text = ""
   TextDet1(3).text = ""
   TextDet1(4).text = ""
   TextDet1(5).text = ""
   TextDet1(6).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 8
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4, 5, 6, 7
    
    vaSpread1.Visible = False
    
    If Trim(TextDet1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 1
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 8
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 8
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 8
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 1
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 8
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
    
    If Trim(TextDet1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 8
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Local Error GoTo Error

Dim RS As New ADODB.Recordset
Dim Sql As String
    
    Select Case Button.Index
    
    Case 1 'Mostrar datos en la grilla
        
            vaSpread1.Visible = False
            vaSpread1.MaxRows = 0
            Sql = ""
            Sql = Sql & Format(fpDateTime1.text, "yyyymm")
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_SalidaEstructuraServicioDiwo  " & Sql & "")
            Do While Not RS.EOF
               
               vaSpread1.MaxRows = vaSpread1.MaxRows + 1
               vaSpread1.Row = vaSpread1.MaxRows
               vaSpread1.Col = 1
               vaSpread1.text = "0"
               
               vaSpread1.Col = 2
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Ceco
               
               vaSpread1.Col = 3
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!CecoDesc
               
               vaSpread1.Col = 4
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Regimen
               
               vaSpread1.Col = 5
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!RegimenDesc
               
               vaSpread1.Col = 6
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!Servicio
               
               vaSpread1.Col = 7
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = RS!ServicioDesc
               
               vaSpread1.Col = 8
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = 0
               
               RS.MoveNext
            
            Loop
            RS.Close
            Set RS = Nothing
            vaSpread1.Visible = True
        
'        End If
    
    End Select

Exit Sub

Error:
    fg_descarga
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, MsgTitulo
    Exit Sub

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1

    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
    For i = 1 To vaSpread1.MaxRows 'BlockRow To BlockRow2
        
        vaSpread1.Row = i
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    Frame1.Enabled = True
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub
