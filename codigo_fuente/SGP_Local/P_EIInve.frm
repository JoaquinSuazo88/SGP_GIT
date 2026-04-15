VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_EIInve 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportar Inventario"
   ClientHeight    =   1980
   ClientLeft      =   4200
   ClientTop       =   3525
   ClientWidth     =   8940
   Icon            =   "P_EIInve.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   135
      Left            =   2760
      TabIndex        =   9
      Top             =   1800
      Visible         =   0   'False
      Width           =   375
      _Version        =   393216
      _ExtentX        =   661
      _ExtentY        =   238
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
      SpreadDesigner  =   "P_EIInve.frx":0442
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   8175
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "P_EIInve.frx":0616
         Left            =   1635
         List            =   "P_EIInve.frx":0618
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   5895
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1620
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   315
         Width           =   5940
         _Version        =   196608
         _ExtentX        =   10477
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
         ButtonStyle     =   2
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   2
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   0
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
         ThreeDFrameColor=   -2147483637
         Appearance      =   0
         BorderDropShadow=   1
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483637
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   240
         Left            =   240
         TabIndex        =   3
         Top             =   1365
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Hoja"
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
         Left            =   240
         TabIndex        =   8
         Top             =   810
         Width           =   1110
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1680
         TabIndex        =   7
         Top             =   840
         Width           =   5895
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta Exp."
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
         Left            =   240
         TabIndex        =   5
         Top             =   390
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
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
         Left            =   240
         TabIndex        =   4
         Top             =   1140
         Visible         =   0   'False
         Width           =   600
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   1980
      Left            =   8400
      TabIndex        =   6
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   3493
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "P_EIInve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim RS As New ADODB.Recordset
Dim opcion As String, Msgtitulo As String, fecinv As Long

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1), True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
vaSpread1.MaxRows = 0
End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)

On Error GoTo ManError

Dim List() As String
Dim listcount As Integer
Dim fromRight As Long, i As Long
Dim handle As Integer
Dim myPath As String
Dim f As Boolean

ReDim List(1)

Cd.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
Cd.Filter = "Todos los archivos (*.xls)|*.xls"
Cd.DefaultExt = "*.mdb"
If opcion = "E" Or opcion = "EP" Then
   If opcion = "EP" Then
      Cd.Filename = "Carga Pedido_" & Format(Date, "yyyymmdd")
   End If
   
   Cd.ShowSave
Else
   Cd.ShowOpen
End If

If Cd.Filename = "" Then fpText1.text = "" Else fpText1.text = Cd.Filename 'Dir(CD.Filename)

fromRight = InStrRev(Cd.Filename, "\", , vbTextCompare)
If fromRight > 1 Then
   myPath = Left(Cd.Filename, fromRight)
End If
vaSpread1.MaxRows = 0: vaSpread1.MaxRows = 500
vaSpread1.MaxCols = 0: vaSpread1.MaxCols = 500
f = vaSpread1.GetExcelSheetList(Cd.Filename, List, listcount, (myPath & "log.txt"), handle, True)
If (listcount - 1 > 1) Then
   ReDim List(listcount - 1)
   f = vaSpread1.GetExcelSheetList(Cd.Filename, List, listcount, (myPath & "log.txt"), handle, False)
End If

Combo1.Clear
For i = 0 To listcount - 1
    Combo1.AddItem (List(i))
Next i
If Dir(myPath & "log.txt") <> "" Then Kill myPath & "log.txt"

Exit Sub
ManError:

If Err.Number = -2147417848 Then

   MsgBox "La Planilla excel esta abierta, debe cerrar la pantilla y la esta opción : " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
   
   Exit Sub

End If

MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim oError As Boolean
Dim sql As String
Dim j As Long
Dim i As Long

On Error GoTo ManError

Select Case Button.Index

Case 1
    
    '------- Validar ruta
    If Trim(fpText1.text) = "" Then fg_descarga: MsgBox "Carpeta no existe", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    oError = False
    If opcion = "EP" Then
       If Dir(Cd.Filename) <> "" Then Kill Cd.Filename 'borrar base datos si existe
       If Dir(Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt") <> "" Then Kill Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt"
       Open Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt" For Output As #1
       PB.Min = 0: PB.Value = 0: PB.Visible = True
       i = 1
       
           
           MVI_EstNecCompra.vaSpread1.Row = 0
           MVI_EstNecCompra.vaSpread1.Col = 1
           sql = " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 2
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 3
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 4
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 5
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 6
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 7
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 8
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 9
           sql = sql & " " & MVI_EstNecCompra.vaSpread1.text & "|"
           Print #1, sql
       
       For i = 1 To MVI_EstNecCompra.vaSpread1.MaxRows
           sql = ""
           PB.Value = Val((i / MVI_EstNecCompra.vaSpread1.MaxRows) * 100)
           MVI_EstNecCompra.vaSpread1.Row = i
           MVI_EstNecCompra.vaSpread1.Col = 1
           sql = MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 2
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 3
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 4
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 5
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 6
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 7
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           MVI_EstNecCompra.vaSpread1.Col = 8
           sql = sql & Format(MVI_EstNecCompra.vaSpread1.text, "mm/dd/yyyy") & "|"
           MVI_EstNecCompra.vaSpread1.Col = 9
           sql = sql & MVI_EstNecCompra.vaSpread1.text & "|"
           Print #1, sql
'           Print #1, Trim(RS!pro_codigo) & ";" & Trim(RS!pro_nombre) & ";" & Trim(RS!uni_nomcor) & ";" & Round(RS!tin_stofis, vg_DCa)
       Next i
       Close #1
       Set XL = CreateObject("Excel.application")
       XL.Workbooks.OpenText Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt", , 1, 1, , , , , , , True, "|"
       XL.ActiveWorkbook.SaveAs Filename:=Cd.Filename, _
                                         FileFormat:=xlNormal, password:="", WriteResPassword:="", _
                                         ReadOnlyRecommended:=False, CreateBackup:=False
       XL.Quit
       Set XL = Nothing
       Label1(2).Visible = False: PB.Visible = False
       If Dir(Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt") <> "" Then Kill Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt"
       oError = IIf(MVI_EstNecCompra.vaSpread1.MaxRows = 0, falser, True)
    ElseIf opcion = "E" Then
       '-------> Exportar Inventario
       RS.Open "SELECT b.pro_codigo, b.pro_nombre, c.uni_nomcor, a.tin_stofis " & _
               "FROM b_tomainv a, b_productos b, a_unidad c " & _
               "WHERE a.tin_codpro = b.pro_codigo " & _
               "AND   b.pro_coduni = c.uni_codigo " & _
               "AND   a.tin_fectom = " & fecinv & " " & _
               "AND   a.tin_codbod = " & vg_codbod & "", vg_db, adOpenStatic
       If Not RS.EOF Then
          If Dir(Cd.Filename) <> "" Then Kill Cd.Filename 'borrar base datos si existe
          If Dir(Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt") <> "" Then Kill Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt"
          Open Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt" For Output As #1
          PB.Min = 0: PB.Value = 0: PB.Visible = True
          i = 1
          Print #1, "" & ";" & "CENCOS : " & MuestraCasino(1) & " " & MuestraCasino(2)
          Print #1, "" & ";" & "Fecha Inventario : " & fg_Ctod1(fecinv)
          Print #1,
          Do While Not RS.EOF
             PB.Value = Val((i / RS.RecordCount) * 100)
             Print #1, Trim(RS!pro_codigo) & ";" & Trim(RS!pro_nombre) & ";" & Trim(RS!uni_nomcor) & ";" & Round(RS!tin_stofis, vg_DCa)
             RS.MoveNext: i = i + 1
          Loop
       End If
       RS.Close: Set RS = Nothing: Close #1
       Set XL = CreateObject("Excel.application")
       XL.Workbooks.OpenText Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt", , 1, 1, , , , , , , True, ";"
       XL.ActiveWorkbook.SaveAs Filename:=Cd.Filename, _
                                         FileFormat:=xlNormal, password:="", WriteResPassword:="", _
                                         ReadOnlyRecommended:=False, CreateBackup:=False
       XL.Quit
       Set XL = Nothing
       Label1(2).Visible = False: PB.Visible = False
       If Dir(Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt") <> "" Then Kill Mid((Cd.Filename), 1, Len((Cd.Filename)) - 3) & "txt"
       oError = True
    
    ElseIf opcion = "I" Then
       
       Dim isel     As Integer, filepath As String, codigo As String, stofis As Double
       Dim dbexcel  As Database
       Dim PathXls  As String
       Dim cn       As ADODB.Connection
       Dim cSpi     As Long
       Dim File_Ext As String
       i = 1: PB.Min = 0: PB.Value = 0: PB.Visible = True
       SheetName = Trim(Combo1.text) & "$"
       filepath = Trim(fpText1.text)
'       Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
'       Set RsExcel = dbexcel.OpenRecordset(sheetname)

       Set RsExcel = New ADODB.Recordset
       Set cn = New ADODB.Connection

       PathXls = Trim(fpText1.text)
       File_Ext = UCase(Right(Cd.Filename, Len(Cd.Filename) - (InStrRev(Cd.Filename, "."))))

       With cn
     
         Select Case File_Ext
        
            ' Excel 97/2003
            Case "XLS"
          
              .Provider = "Microsoft.Jet.OLEDB.4.0"
              .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
              .CursorLocation = 3
          
            ' Excel 2010
            Case "XLSX"

              .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
              .ConnectionString = "Extended Properties=Excel 8.0;"
              .CursorLocation = 3
     
         End Select
     
AbrirExcel:

         .Open

       End With

       RsExcel.Open ("SELECT * FROM [" & SheetName & "]"), cn

       If RsExcel.EOF Then Exit Sub

       RsExcel.MoveFirst
       Do While RsExcel.EOF <> True
          DoEvents
         PB.Value = Val((i / RsExcel.RecordCount) * 100)
          If RsExcel.Fields(0).Value = "*" Then Exit Do
          codigo = "": codigo = IIf(Not IsNull(RsExcel.Fields(0).Value), RsExcel.Fields(0).Value, "")
          If Trim(codigo) <> "" Then
             stofis = 0
             If IsNumeric(RsExcel.Fields(3)) Then stofis = RsExcel.Fields(3)
             
             If stofis > 0 Then
                
                vg_db.Execute "UPDATE  b_tomainv SET tin_stofis=" & stofis & " " & _
                              "WHERE tin_fectom = " & fecinv & " " & _
                              "AND   tin_codbod = " & vg_codbod & " " & _
                              "AND   tin_codpro = '" & codigo & "'"
             
             End If
          
          End If
'          Label2(2).Caption = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
          RsExcel.MoveNext: i = i + 1
       
       Loop
       RsExcel.Close: Set RsExcel = Nothing
       Label1(2).Visible = False: PB.Visible = False
       vg_codigo = "X"
       oError = True
    
    End If
    
    If Not oError Then
       
       MsgBox IIf(opcion = "E" Or opcion = "EP", "Proceso de Exportar Falló", "Proceso de Importar Falló"), vbInformation + vbOKOnly, Msgtitulo
    
    Else
       
       MsgBox IIf(opcion = "E" Or opcion = "EP", "Proceso de Exportar Finalizado", "Proceso de Importar Finalizado"), vbInformation + vbOKOnly, Msgtitulo
    
    End If

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
ManError:
If Err.Number = 70 Or Err.Number = 52 Or Err.Number = 1004 Then
   
   If opcion <> "EP" Then RS.Close: Set RS = Nothing
   Close #1
   MsgBox "La Planilla excel esta abierta, debe cerrar " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
          Label1(2).Visible = False: PB.Visible = False

   Exit Sub

ElseIf Err.Number = -2147467259 Then
   
   cn.Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
   cn.ConnectionString = "Extended Properties=Excel 8.0;"
   cn.CursorLocation = 3
    
   GoTo AbrirExcel

ElseIf Err.Number = -2147217900 Then

   Set RsExcel = Nothing
   MsgBox "La Planilla excel esta abierta, debe cerrar la pantilla y la esta opción " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
   
   Exit Sub

End If
MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo

End Sub

Sub Inicio(mtit As String, op As String, Fecha As Long)
Msgtitulo = mtit
Me.Caption = mtit
opcion = op
fecinv = Fecha
If op = "I" Then
   Label1(1).Caption = "Ruta Imp."
Else
   label2(0).Visible = False
   Combo1.Visible = False
   fpayuda(2).Visible = False
End If
End Sub
