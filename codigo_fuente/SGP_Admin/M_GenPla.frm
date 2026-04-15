VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{1DF3AFED-99E0-4474-9900-954B8FD24E86}#1.0#0"; "ChilkatMail2.dll"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_GenPla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Archivos Planos Planificación"
   ClientHeight    =   8295
   ClientLeft      =   3375
   ClientTop       =   2025
   ClientWidth     =   7470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   6615
      Left            =   30
      TabIndex        =   8
      Top             =   1560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Planificación"
      TabPicture(0)   =   "M_GenPla.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Listas de Precios SAC"
      TabPicture(1)   =   "M_GenPla.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   5655
         Left            =   -74880
         TabIndex        =   13
         Top             =   480
         Width           =   7095
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4590
            Left            =   120
            TabIndex        =   14
            Top             =   360
            Width           =   6855
            _Version        =   393216
            _ExtentX        =   12091
            _ExtentY        =   8096
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
            SpreadDesigner  =   "M_GenPla.frx":0038
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Enviados"
            Height          =   195
            Index           =   2
            Left            =   4800
            TabIndex        =   16
            Top             =   5400
            Width           =   660
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0FFC0&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   2
            Left            =   4440
            Top             =   5430
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "No Enviados"
            Height          =   195
            Index           =   0
            Left            =   6015
            TabIndex        =   15
            Top             =   5400
            Width           =   915
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   5655
            Top             =   5430
            Width           =   300
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5985
         Left            =   120
         TabIndex        =   9
         Top             =   540
         Width           =   7125
         Begin MSComctlLib.TreeView TvwDir 
            Height          =   5085
            Left            =   90
            TabIndex        =   10
            Top             =   210
            Width           =   6885
            _ExtentX        =   12144
            _ExtentY        =   8969
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   165
            Index           =   0
            Left            =   90
            TabIndex        =   11
            Top             =   5730
            Visible         =   0   'False
            Width           =   4470
            _ExtentX        =   7885
            _ExtentY        =   291
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Enviando Datos 5 Etapas"
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
            Left            =   120
            TabIndex        =   17
            Top             =   5400
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Enviado"
            Height          =   195
            Index           =   1
            Left            =   6360
            TabIndex        =   12
            Top             =   5385
            Width           =   585
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   6000
            Top             =   5415
            Width           =   300
         End
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5910
      Top             =   7860
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
            Picture         =   "M_GenPla.frx":1A24
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1005
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   7155
      Begin ACTIVEZIPLib.ActiveZip AZ1 
         Left            =   6120
         Top             =   960
         _Version        =   65536
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         ItemData        =   "M_GenPla.frx":1DBE
         Left            =   1710
         List            =   "M_GenPla.frx":1DC0
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   160
         Width           =   4575
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   2730
         TabIndex        =   4
         Top             =   555
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Left            =   1710
         TabIndex        =   2
         Top             =   555
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
         Text            =   "11/2021"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Envió"
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
         TabIndex        =   6
         Top             =   255
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Período"
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
         Left            =   3240
         TabIndex        =   5
         Top             =   615
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Período"
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
         TabIndex        =   1
         Top             =   615
         Width           =   690
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATMAILLib2Ctl.ChilkatMailMan2 oMail 
      Left            =   7200
      OleObjectBlob   =   "M_GenPla.frx":1DC2
      Top             =   7920
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   1230
      OleObjectBlob   =   "M_GenPla.frx":1DE6
      Top             =   0
   End
End
Attribute VB_Name = "M_GenPla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS           As New ADODB.Recordset
Dim RS1          As New ADODB.Recordset
Dim RS2          As New ADODB.Recordset
Dim EstError     As Boolean
Dim MsgTitulo    As String
Dim rootNode     As Node
Dim ArNivTree(6) As Variant
Dim fso
Private BtnX     As Variant

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 8775
Me.Width = 7560
Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Generar Archivos Plano Planificación"
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Enviar", , tbrDefault, "A_Enviar"): BtnX.Visible = True: BtnX.ToolTipText = "Enviar": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
fpDateTime1.text = Format(Date, "mm/yyyy")
Combo1(0).Clear
Combo1(0).AddItem "Planificación y Parametrización" & Space(150) & "(0)"
Combo1(0).AddItem "Solo Parametrización" & Space(150) & "(1)"
Combo1(0).ListIndex = -1
SSTab1.Tab = 0
SSTab1.TabEnabled(1) = False
vaSpread1.MaxRows = 0
End Sub

Private Sub fpDateTime1_Change()
'If IsDate(fpDateTime1(Index).Text) = False Then Exit Sub
TvwDir.Nodes.Clear
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case Button.Index
Case 1
    Dim nEst As Boolean, mdirserver As String, sourcefile As String, aAp As String, dBo As String, spid As Long
    Dim codReg As String, codser As String, codcas As String, tprod As String, treceta As String, logenv As String
    Dim i As Long, subseg As Long, cdcen As String, dtref As String, nrsem As Long, envlip As Boolean
    Dim auxceco As String
    Dim MyBufferProd As String
    Dim MyBufferSubSeg As String
    Dim MyBufferReceta As String
    Dim MyBufferReg As String
    Dim MyBufferSer As String
    Dim MyBufferCeco As String
    Dim CodOpt As String
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    envlip = False
    SSTab1.Tab = 0
    '-------> Borrar tabla de paso estructura servicio
    vg_db.Execute "DELETE paso_enviolistaprecio WHERE elp_spid = @@spid and elp_usu = '" & vg_NUsr & "'"
    '-------> Buscar spid
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then spid = RS!spid
    RS.Close: Set RS = Nothing
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           envlip = True
           vaSpread1.Col = 2: cdcen = Trim(vaSpread1.text)
           vaSpread1.Col = 4: dtref = Trim(vaSpread1.text)
           vaSpread1.Col = 6: nrsem = Val(vaSpread1.text)
           vg_db.Execute ("INSERT INTO paso_enviolistaprecio (elp_spid, elp_usu, elp_cdcen, elp_dtref, elp_nrsem) VALUES(" & spid & ", '" & vg_NUsr & "', '" & cdcen & "', '', '" & dtref & "', " & nrsem & ")")
        End If
    Next i
    
    nEst = False
    For i = 1 To TvwDir.Nodes.count
        If TvwDir.Nodes.item(i).Checked = True And InStr(TvwDir.Nodes.item(i).Key, "CASINO") <> 0 Then nEst = True: Exit For
    Next
    If Not nEst And Not envlip Then MsgBox "No ha seleccionado datos...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If Combo1(0).ListIndex = -1 Or Combo1(0).text = "" Then MsgBox "Seleccione tipo envió...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    fg_carga ""
    Toolbar1.Enabled = False
    Frame1(0).Enabled = True
    '-------> Base de datos origen
    '-------> Acceso base de access dBo = dir_trabajo + BaseDeDato
    dBo = "'' [ODBC;Driver={SQL Server};Server=" + vg_SqlNSvr + ";Database=" + vg_SqlBase + ";UID=" + vg_SqlNUsr + ";PWD=" + vg_SqlPass + "]"
    '-------> Crear directorio si no existe
    mdirserver = Dir(dir_trabajo & "\" & "Actualizar", vbDirectory)
    If mdirserver = "" Then MkDir dir_trabajo & "\" & "Actualizar"
    mdirserver = dir_trabajo & "Actualizar" & "\"
    '-------> Fin crear directorio si no existe
    '-------> Generar base padre
    sourcefile = "minutageneral" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
    If Dir(mdirpc & sourcefile) <> "" Then Kill mdirpc & sourcefile ' borrar base datos si existe
    
    '-------> Filtrar los sub-segmento que seleciono el usuario
    '-------> Borrar tabla de paso regimen que va filtrar sub-segmento
    vg_db.Execute "DELETE paso_regimen WHERE reg_spid = @@spid and reg_usr = '" & vg_NUsr & "'"
    Dim codsse As String
    Dim auxreg As Long
    codsse = "": subseg = 0: auxreg = 0: auxceco = ""
    Let MyBufferSubSeg = ""
    Let MyBufferSubSeg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferSubSeg = MyBufferSubSeg & "<SubSegmento>"
    
    Let MyBufferReg = ""
    Let MyBufferReg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferReg = MyBufferReg & "<Regimen>"
    
    Let MyBufferCeco = ""
    Let MyBufferCeco = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferCeco = MyBufferCeco & "<Ceco>"
    
    For i = 1 To TvwDir.Nodes.count
        DoEvents
        If TvwDir.Nodes.item(i).Checked = True And InStr(TvwDir.Nodes.item(i).Key, "CASINO") <> 0 And subseg <> LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5))) Then
           codsse = codsse & "" & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5))) & ","
           subseg = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5)))
           MyBufferSubSeg = MyBufferSubSeg & " <SubSegmentos"
           MyBufferSubSeg = MyBufferSubSeg & " CodSubSegmento = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5))) & Chr(34)
           Let MyBufferSubSeg = MyBufferSubSeg & "/>"

'           codreg = codreg & LCase(Trim(Mid(TvwDir.Nodes.Item(i).Key, 17, 5))) & ","
           vg_db.Execute "INSERT INTO paso_regimen (reg_spid, reg_usr, reg_codigo) VALUES (" & spid & ", '" & vg_NUsr & "', " & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5))) & ")"
        End If
        If TvwDir.Nodes.item(i).Checked = True And InStr(TvwDir.Nodes.item(i).Key, "CASINO") <> 0 And auxreg <> LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) Then
           auxreg = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5)))
           codReg = codReg & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) & ","
           MyBufferReg = MyBufferReg & " <Regimenes"
           MyBufferReg = MyBufferReg & " CodRegimen = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) & Chr(34)
           Let MyBufferReg = MyBufferReg & "/>"
        End If
        If TvwDir.Nodes.item(i).Checked = True And InStr(TvwDir.Nodes.item(i).Key, "CASINO") <> 0 And auxceco <> LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 7, 10))) Then
           auxceco = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 7, 10)))
           MyBufferCeco = MyBufferCeco & " <Cecos"
           MyBufferCeco = MyBufferCeco & " CodCeco = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 7, 10))) & Chr(34)
           Let MyBufferCeco = MyBufferCeco & "/>"
        End If
    Next i
    Let MyBufferSubSeg = MyBufferSubSeg & "</SubSegmento>"
    Let MyBufferReg = MyBufferReg & "</Regimen>"
    Let MyBufferCeco = MyBufferCeco & "</Ceco>"
    '-------> Validar si existen minutas reales concepto propuesta
    If Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Then
       ws_respuesta = ""
       I_MinutasRealesConRecetasPropuesta Format(fpDateTime1.Value, "yyyymm") & "01", Format(fpDateTime1.Value, "yyyymm") & Mid(dEoM("01/" & Format(fpDateTime1.Value, "mm/yyyy")), 1, 2), vg_NUsr, spid
       If Trim(ws_respuesta) <> "" Then
          MsgBox "Proceso cancelado, ya que existen el concepto propuesta en algunas de las minutas...", vbInformation + vbOKOnly, MsgTitulo
          Toolbar1.Enabled = True
          fg_descarga
          Exit Sub
       End If
    End If
   Set RS = vg_db.Execute("SELECT DISTINCT reg_codigo FROM a_regimen With(NoLock) WHERE reg_codigo IN (" & Mid(codReg, 1, Len(codReg) - 1) & ")")
    If Not RS.EOF Then
       codReg = ""
       Do While Not RS.EOF
          codReg = codReg & RS!reg_codigo & ","
          RS.MoveNext
       Loop
    End If
    RS.Close: Set RS = Nothing
    
    If Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Then
       Set RS = vg_db.Execute("sgpadm_Sel_EnvioMdbProductos ")
       Let MyBufferProd = ""
       Let MyBufferProd = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
       Let MyBufferProd = MyBufferProd & "<Producto>"
       Do While Not RS.EOF
          MyBufferProd = MyBufferProd & " <Productos"
          MyBufferProd = MyBufferProd & " CodProducto = " & Chr(34) & RS(0) & Chr(34)
          Let MyBufferProd = MyBufferProd & "/>"
          RS.MoveNext
       Loop
       RS.Close: Set RS = Nothing
       Let MyBufferProd = MyBufferProd & "</Producto>"
    
       If nEst Then
          Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbMinutaReceta  '" & MyBufferSubSeg & "', " & Val(Format(fpDateTime1.text, "yyyymm")) & "")
          Let MyBufferReceta = ""
          Let MyBufferReceta = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
          Let MyBufferReceta = MyBufferReceta & "<Receta>"
          Do While Not RS.EOF
             MyBufferReceta = MyBufferReceta & " <Recetas"
             MyBufferReceta = MyBufferReceta & " CodReceta = " & Chr(34) & RS(0) & Chr(34)
             Let MyBufferReceta = MyBufferReceta & "/>"
             RS.MoveNext
          Loop
          RS.Close: Set RS = Nothing
          Let MyBufferReceta = MyBufferReceta & "</Receta>"
       
       End If
    End If
    DoEvents
    GenerarBaseEnviado mdirpc & sourcefile, tprod, treceta, dBo, IIf(Not nEst, 0, IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, 2, 3)), Val(Format(fpDateTime1.text, "yyyymm")), codsse, codReg, "'0',", MyBufferProd, MyBufferReceta, MyBufferSubSeg, MyBufferReg, MyBufferCeco
    codcas = "": codReg = "": codser = "": subseg = 0
    Bar1(0).Visible = True: Bar1(0).Value = 0
    Label1(3).Visible = True: Label1(3).Caption = "Enviando Información Contratos 5 Etapas"
    If nEst Then
        '-------> Crear archivo log de envio productos, recetas y planificación
        logenv = "mailSent" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".log"
        If Dir(dir_trabajo & logenv) <> "" Then Kill dir_trabajo & logenv ' borrar base datos si existe
        Open dir_trabajo & logenv For Output As #1 'Crear archivos de errores
        Close #1
        '-------> Xml Regimen
        Let MyBufferReg = ""
        Let MyBufferReg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBufferReg = MyBufferReg & "<Regimen>"
        '-------> Xml Servicio
        Let MyBufferSer = ""
        Let MyBufferSer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBufferSer = MyBufferSer & "<Servicio>"
        
        EstError = True
        For i = 1 To TvwDir.Nodes.count
            DoEvents
            Bar1(0).Value = Val((i / TvwDir.Nodes.count) * 100)
            If TvwDir.Nodes.item(i).Checked = True And InStr(TvwDir.Nodes.item(i).Key, "CASINO") <> 0 Then
               If codcas <> LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 7, 10))) Then
                  If Trim(codcas) <> "" Then
                     Let MyBufferReg = MyBufferReg & "</Regimen>"
                     Let MyBufferSer = MyBufferSer & "</Servicio>"
                     GenerarArchivos subseg, codcas, codReg, codser, sourcefile, mdirpc, dBo, logenv, MyBufferReg, MyBufferSer
                     If Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Or Val(fg_codigocbo(Combo1, 0, 1, "")) = 1 Then
                        '-------> Grabar tabla b_minutacasino de enviado la información
                        vg_db.Execute "DELETE b_minutacasino FROM b_minutacasino WHERE mic_cencos = '" & codcas & "' AND mic_codreg IN (" & codReg & ") AND mic_codser IN (" & codser & ") AND mic_fecmin = " & Val(Format(fpDateTime1.text, "yyyymm")) & ""
                        vg_db.Execute "INSERT INTO b_minutacasino (mic_cencos, mic_codreg, mic_codser, mic_fecmin, mic_fecenv) SELECT '" & codcas & "', reg.reg_codigo, ser.ser_codigo, " & Val(Format(fpDateTime1.text, "yyyymm")) & ", " & Format(Date, "yyyymmdd") & " FROM a_regimen reg, a_servicio ser WHERE reg.reg_codigo IN (" & codReg & ") AND ser.ser_codigo IN (" & codser & ")"
                     End If
                     '-------> Xml Regimen
                     Let MyBufferReg = ""
                     Let MyBufferReg = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
                     Let MyBufferReg = MyBufferReg & "<Regimen>"
                     '-------> Xml Servicio
                     Let MyBufferSer = ""
                     Let MyBufferSer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
                     Let MyBufferSer = MyBufferSer & "<Servicio>"
                     subseg = 0: codcas = "": codReg = "": codser = ""
                  End If
                  subseg = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5)))
                  codcas = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 7, 10)))
                  codReg = codReg & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) & ","
                  codser = codser & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 28, 5))) & ","
                  '-------> Xml Regimen
                  MyBufferReg = MyBufferReg & " <Regimenes"
                  MyBufferReg = MyBufferReg & " CodRegimen = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) & Chr(34)
                  Let MyBufferReg = MyBufferReg & "/>"
                  '-------> Xml Servicio
                  MyBufferSer = MyBufferSer & " <Servicios"
                  MyBufferSer = MyBufferSer & " CodServicio = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 28, 5))) & Chr(34)
                  Let MyBufferSer = MyBufferSer & "/>"
               
               Else
                  subseg = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 2, 5)))
                  codcas = LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 7, 10)))
                  codReg = codReg & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) & ","
                  codser = codser & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 28, 5))) & ","
                  
                  '-------> Xml Regimen
                  MyBufferReg = MyBufferReg & " <Regimenes"
                  MyBufferReg = MyBufferReg & " CodRegimen = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 17, 5))) & Chr(34)
                  Let MyBufferReg = MyBufferReg & "/>"
               
                  '-------> Xml Servicio
                  MyBufferSer = MyBufferSer & " <Servicios"
                  MyBufferSer = MyBufferSer & " CodServicio = " & Chr(34) & LCase(Trim(Mid(TvwDir.Nodes.item(i).Key, 28, 5))) & Chr(34)
                  Let MyBufferSer = MyBufferSer & "/>"
               
               End If
            End If
        Next
        If Trim(codcas) <> "" Then
           Let MyBufferReg = MyBufferReg & "</Regimen>"
           Let MyBufferSer = MyBufferSer & "</Servicio>"
           GenerarArchivos subseg, codcas, codReg, codser, sourcefile, mdirpc, dBo, logenv, MyBufferReg, MyBufferSer
            If Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Or Val(fg_codigocbo(Combo1, 0, 1, "")) = 1 Then
               '-------> Grabar tabla b_minutacasino de enviado la información
               vg_db.Execute "DELETE b_minutacasino FROM b_minutacasino WHERE mic_cencos = '" & codcas & "' AND mic_codreg IN (" & codReg & ") AND mic_codser IN (" & codser & ") AND mic_fecmin = " & Val(Format(fpDateTime1.text, "yyyymm")) & ""
               vg_db.Execute "INSERT INTO b_minutacasino (mic_cencos, mic_codreg, mic_codser, mic_fecmin, mic_fecenv) SELECT '" & codcas & "', reg.reg_codigo, ser.ser_codigo, " & Val(Format(fpDateTime1.text, "yyyymm")) & ", " & Format(Date, "yyyymmdd") & " FROM a_regimen reg, a_servicio ser WHERE reg.reg_codigo IN (" & codReg & ") AND ser.ser_codigo IN (" & codser & ")"
            End If
           
           '-------> Grabar lista de precio siempre y cuando este envio planificación de minutas.
           If Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Then
           End If
           subseg = 0: codcas = "": codReg = "": codser = ""
        End If
        Bar1(0).Visible = False
        Toolbar1.Enabled = True
        Frame1(0).Enabled = True
        fg_descarga
        Bar1(0).Visible = False
    Else
'       GenerarListaPrecioContratono5Etapas sourcefile, mdirpc, dBo, nEst, spid, ""
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           If vaSpread1.text = "1" Then
              vaSpread1.Col = 2: cdcen = Trim(vaSpread1.text)
              vaSpread1.Col = 4: dtref = Trim(vaSpread1.text)
              vaSpread1.Col = 6: nrsem = Val(vaSpread1.text)
              vg_db.Execute "DELETE b_enviolistapreciocecom FROM b_enviolistapreciocecom WHERE lpc_cecom = '" & cdcen & "' AND lpc_periodo = '" & dtref & "' AND lpc_nrosem = " & nrsem & ""
              vg_db.Execute "INSERT INTO b_enviolistapreciocecom (lpc_cecom, lpc_periodo, lpc_nrosem, lpc_fecenv) VALUES ('" & cdcen & "', '" & dtref & "', " & nrsem & ", '" & Format(Date, "yyyymmdd") & "')"
           End If
       Next i
    End If
    '-------> Borrar base patron
    If Dir(mdirpc & sourcefile) <> "" Then Kill mdirpc & sourcefile
    
    '-------> Copiar archivos access \\SQLDES\CXCASINO, luego borrar archivos del PC
    fso.CopyFile mdirpc & "sgp*.zip", mdirserver, True
    If Dir(mdirpc & "sgp*.zip") <> "" Then Kill mdirpc & "sgp*.zip"
    '-------> Fin copiar archivos access \\SQLDES\CXCASINO, luego borrar archivos del PC
    
    Label1(1).Visible = False
    If Trim(sourcefile) <> "" Then MsgBox "Generación Finalizado Sin Problema", vbInformation + vbOKOnly, MsgTitulo
Case 3
    Me.Hide
    Unload Me
End Select

Exit Sub
fg_descarga
Toolbar1.Enabled = True
Frame1(0).Enabled = True
Bar1(0).Visible = False
RS.Close: Set RS = Nothing
Man_Error:
Select Case Err
Case 91
    Resume Next
    Exit Sub
Case 35764
    DoEvents
    For i = 1 To 1000000
    Next i
    Resume
Case 76
    Resume Next
    Exit Sub
Case -2147467259
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    Exit Sub
Case 3034
'    vg_db.RollbackTrans: Exit Sub
End Select
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Function GenerarArchivos(subseg As Long, codcas As String, codReg As String, codser As String, sourcefile As String, mdir As String, dBo As String, logenv As String, MyBufferReg As String, MyBufferSer As String)
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim destinofile As String, destinofilezip As String, cDBI As String, csubjects As String, cBody As String, codzon As String, codtis As Long, CodSeg As Long, socsap As String, nomcli As String, sobrec As String
Dim CHost As String, Cdire As String, Cuser As String, Cpass As String, Cpuer As Long, cecsac As String, minsre As String
Dim codmun As Long, ccisac As Long, codrge As Long
Dim a As Variant
Dim lDir As Variant
Dim CodOpt As String
DoEvents
codtis = 0: CodSeg = 0: socsap = "": sobrec = "": codmun = 0: codrge = 0: minsre = "0"
Set RS = vg_db.Execute("SELECT cli_nombre, cli_codzon, cli_codtis, cli_codseg, cli_socsap, cli_sobrec, cli_codmun, cli_ccisac, cli_cecsac, cli_codreg, cli_minsre, cli_TipoMinuta FROM b_clientes With(NoLock) WHERE cli_codigo = '" & codcas & "'")
If Not RS.EOF Then
   codzon = RS!cli_codzon
   codtis = RS!cli_codtis
   CodSeg = RS!cli_codseg
   codmun = IIf(IsNull(RS!cli_codmun), 0, RS!cli_codmun)
   socsap = IIf(IsNull(RS!cli_socsap), "", RS!cli_socsap)
   nomcli = IIf(IsNull(RS!cli_nombre), "", RS!cli_nombre)
   sobrec = IIf(IsNull(RS!cli_sobrec), "", RS!cli_sobrec)
   ccisac = IIf(IsNull(RS!cli_ccisac), 0, RS!cli_ccisac)
   cecsac = IIf(IsNull(RS!cli_cecsac), "", RS!cli_cecsac)
   codrge = IIf(IsNull(RS!cli_codreg), 0, RS!cli_codreg)
'   minsre = IIf(IsNull(RS!cli_minsre), "0", RS!cli_minsre)
   minsre = IIf(IsNull(RS!cli_tipominuta) Or RS!cli_tipominuta = 0, "0", RS!cli_tipominuta)
End If
RS.Close: Set RS = Nothing

'-------> Mover codigo optimun
CodOpt = ""
Set RS = vg_db.Execute("SELECT isnull(Cecos_AX,'') as Cecos_AX FROM Cecos_Sap_AX WHERE Cecos_Sap = '" & codcas & "' and Sociedad_Sap = '" & socsap & "'")
If Not RS.EOF Then
   CodOpt = RS!Cecos_AX
End If
RS.Close: Set RS = Nothing
           
Label1(1).Visible = True
Label1(1).Caption = Trim(codcas) & " " & Trim(nomcli)
codReg = Mid(codReg, 1, Len(codReg) - 1)
codser = Mid(codser, 1, Len(codser) - 1)
destinofile = "sgp" & codcas & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".kkk"
destinofilezip = "sgp" & codcas & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
'------- verificar si existe archivo mdb destino si existe borrar y copiar
If Dir(mdir & destinofile) <> "" Then Kill mdir & destinofile
FileCopy mdir & sourcefile, mdir & destinofile
cDBI = mdir & destinofile
Set dbi = New ADODB.Connection
dbi.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cDBI & "' ;Persist Security Info=False"
dbi.ConnectionTimeout = 3600
dbi.CommandTimeout = 3600
dbi.Open
If Val(fg_codigocbo(Combo1, 0, 1, "")) = 0 Then
   '------- generar tabla gramaje
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbTablaGramaje '" & MyBufferReg & "', " & subseg & ", " & codzon & "")
   Do While Not RS.EOF
      DoEvents
      dbi.Execute "INSERT INTO b_tablagramaje (tgr_codreg, tgr_codrec, tgr_coding, tgr_codzon, tgr_codins, tgr_cantgr) " & _
                "VALUES (" & RS!tgr_codreg & ", " & RS!tgr_codrec & ", '" & RS!tgr_coding & "', " & RS!tgr_codzon & ", '" & RS!tgr_codins & "', " & RS!tgr_cantgr & ")"
      
'   dbi.Execute "INSERT INTO b_tablagramaje SELECT DISTINCT a.tgr_codreg, a.tgr_codrec, a.tgr_coding, a.tgr_codzon, a.tgr_codins, a.tgr_cantgr FROM b_tablagramajeaux a, a_subsegmentoaux b WHERE a.tgr_subseg = b.sub_codigo AND a.tgr_codzon = " & codzon & " AND b.sub_codigo = " & subseg & " AND a.tgr_codreg in (" & codReg & ")"
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing

   dbi.Execute "INSERT INTO gra_receta (rec_codigo) SELECT DISTINCT tgr_codrec FROM b_tablagramaje"
   dbi.Execute "DELETE b_receta.* FROM b_receta INNER JOIN gra_receta ON  b_receta.rec_codigo = gra_receta.rec_codigo"
   dbi.Execute "DELETE b_recetadet.* FROM b_recetadet INNER JOIN gra_receta ON b_recetadet.red_codigo = gra_receta.rec_codigo"
   '------- insertar receta desde tabla gramaje
   dbi.Execute "INSERT INTO b_receta (rec_codigo, rec_catdie, rec_tippla, rec_nombre, rec_nomfan, rec_metpre, rec_conche, rec_sugere, rec_basrac, rec_tiprec, rec_fecvig, rec_gruvul) SELECT DISTINCT a.rec_codigo, a.rec_catdie, a.rec_tippla, a.rec_nombre, a.rec_nomfan, '', a.rec_conche, a.rec_sugere, a.rec_basrac, a.rec_tiprec, a.rec_fecvig, a.rec_gruvul FROM b_recetaaux a, b_tablagramaje b WHERE a.rec_codigo = b.tgr_codrec"
   dbi.Execute "UPDATE b_receta INNER JOIN b_recetaaux ON b_receta.rec_codigo = b_recetaaux.rec_codigo SET b_receta.rec_metpre=b_recetaaux.rec_metpre"
   dbi.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) SELECT DISTINCT a.red_codigo, a.red_nroite, a.red_codpro, a.red_canpro, a.red_cospro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, 0 FROM b_recetadetaux a, b_tablagramaje b WHERE a.red_codigo = b.tgr_codrec"
   '------- insertar receta desde tabla gramaje con origen regimen
   dbi.Execute "UPDATE b_receta SET rec_tiprec = 1 WHERE rec_codigo in (SELECT DISTINCT tgr_codrec FROM b_tablagramaje)"
   dbi.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec) SELECT DISTINCT a.red_codigo, a.red_nroite, a.red_codpro, a.red_canpro, a.red_cospro, a.red_pctapr, a.red_pctcoc, a.red_pctnut, b.tgr_codreg FROM b_recetadetaux a, b_tablagramaje b WHERE a.red_codigo = b.tgr_codrec"
   dbi.Execute "UPDATE b_recetadet INNER JOIN b_tablagramaje ON (b_recetadet.red_tiprec = b_tablagramaje.tgr_codreg) AND (b_recetadet.red_codpro = b_tablagramaje.tgr_coding) AND (b_recetadet.red_codigo = b_tablagramaje.tgr_codrec) SET b_recetadet.red_codpro = [b_tablagramaje].[tgr_codins], b_recetadet.red_canpro = [b_tablagramaje].[tgr_cantgr]"
   dbi.Execute "UPDATE b_recetadet INNER JOIN b_ingrediente ON b_recetadet.red_codpro = b_ingrediente.ing_codigo SET b_recetadet.red_pctapr = [b_ingrediente].[ing_pctapr], b_recetadet.red_pctcoc = [b_ingrediente].[ing_pctcoc], b_recetadet.red_pctnut = [b_ingrediente].[ing_pctnut] WHERE b_recetadet.red_tiprec > 0"
   '------- Borrar tabla regimen que no tengan relación con el contrato
   dbi.Execute "DELETE a_regimen FROM a_regimen WHERE reg_codigo NOT IN (" & codReg & ")"
   
   '-------> Generar encabezado planificación minutas
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbMinuta '" & MyBufferReg & "', '" & MyBufferSer & "', " & subseg & ", " & Val(Format(fpDateTime1.text, "yyyymm")) & "")
   Do While Not RS.EOF
      DoEvents
      dbi.Execute "INSERT INTO b_minuta (min_codigo, min_cencos, min_codreg, min_codser, min_fecmin, min_indblo, min_racteo, min_racrea) " & _
                "VALUES (" & RS!min_codigo & ", '" & RS!min_subseg & "', " & RS!min_codreg & ", " & RS!min_codser & ", " & RS!min_fecmin & ", " & RS!MIN_INDBLO & ", " & RS!min_racteo & ", " & RS!min_racrea & ")"
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing

   '-------> Generar detalle planificación minutas
   Set RS = vg_db.Execute("sgpadm_Sel_XmlEnvioMdbMinutaDetalle '" & MyBufferReg & "', '" & MyBufferSer & "', " & subseg & ",  " & Val(Format(fpDateTime1.text, "yyyymm")) & "")
   Do While Not RS.EOF
      DoEvents
      dbi.Execute "INSERT INTO b_minutadet (mid_codigo, mid_tipmin, mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec, mid_fecval, mid_tiprec, mid_nummer, mid_rec5eta, mid_cosdes) " & _
                "VALUES (" & RS!mid_codigo & ", '" & RS!mid_tipmin & "', " & RS!mid_numlin & ", " & RS!mid_estser & ", " & RS!mid_codrec & ", " & RS!mid_numrac & ", '" & RS!mid_descri & "', " & RS!mid_cosrec & ", " & RS!mid_fecval & ", " & RS!mid_tiprec & ", " & RS!mid_nummer & ", '" & RS!mid_rec5eta & "', " & RS!mid_cosdes & ")"
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
   
'   '------- Borrar planificación de minutas que no tengan relación con el contrato
'   dbi.Execute "DELETE a.* FROM b_minuta b INNER JOIN b_minutadet a ON b.min_codigo = a.mid_codigo WHERE b.min_cencos NOT IN ('" & subseg & "')"
'   dbi.Execute "DELETE a.* FROM b_minuta b INNER JOIN b_minutadet a ON b.min_codigo = a.mid_codigo WHERE b.min_codreg NOT IN (" & codReg & ")"
'   dbi.Execute "DELETE a.* FROM b_minuta b INNER JOIN b_minutadet a ON b.min_codigo = a.mid_codigo WHERE b.min_codser NOT IN (" & codser & ")"

'   dbi.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos NOT IN ('" & subseg & "')"
'   dbi.Execute "DELETE b_minuta FROM b_minuta WHERE min_codreg NOT IN (" & codReg & ")"
'   dbi.Execute "DELETE b_minuta FROM b_minuta WHERE min_codser NOT IN (" & codser & ")"
End If
'------- Borrar costo patron que no tengan relación con el contrato
dbi.Execute "DELETE b_costopatron FROM b_costopatron WHERE cpa_cencos NOT IN ('" & codcas & "')"
dbi.Execute "DELETE b_costopatron FROM b_costopatron WHERE cpa_codreg NOT IN (" & codReg & ")"
dbi.Execute "DELETE b_costopatron FROM b_costopatron WHERE cpa_codser NOT IN (" & codser & ")"
'dbi.Execute "UPDATE b_costopatron SET cpa_cencos='" & codcas & "'"

'-------> Borrar costo gramo familia producto que no tengan relación con el contrato
dbi.Execute "DELETE b_gramofamproducto FROM b_gramofamproducto WHERE gfp_cencos NOT IN ('" & subseg & "')"
dbi.Execute "UPDATE b_gramofamproducto SET gfp_cencos = '" & codcas & "'"

'-------> Borrar tabla servicio que no tengan relación con el contrato
dbi.Execute "DELETE a_estservicio FROM a_estservicio WHERE ess_codser NOT IN (" & codser & ")"
dbi.Execute "DELETE a_servicio FROM a_servicio WHERE ser_codigo NOT IN (" & codser & ")"
dbi.Execute "INSERT INTO a_servicio SELECT DISTINCT a.ser_codigo, a.ser_nombre, a.ser_orden FROM a_servicio a, a_estservicio b  IN " & dBo & " WHERE a.ser_codigo IN (b.ess_codser) AND a.ser_codigo NOT IN (SELECT DISTINCT ser_codigo FROM a_servicio) AND b.ess_codigo IN (SELECT DISTINCT mid_estser FROM b_minutadet)"
dbi.Execute "INSERT INTO a_estservicio SELECT DISTINCT ess_codser, ess_codigo, ess_nombre, ess_orden FROM a_estservicio IN " & dBo & " WHERE ess_codser IN (SELECT ser_codigo FROM a_servicio) AND ess_codigo NOT IN (SELECT ess_codigo FROM a_estservicio)"
'-------> Borrar tabla tipo servicio y segmento que no tenga relación con el contrato
dbi.Execute "DELETE a_tiposervicio FROM a_tiposervicio WHERE tis_codigo NOT IN (" & codtis & ")"
dbi.Execute "DELETE a_segmento FROM a_segmento WHERE seg_codigo NOT IN (" & CodSeg & ")"
'-------> Borrar tabla casino envia sap
dbi.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos NOT IN ('" & codcas & "')"
'-------> Borrar tabla parametro codigo barra
dbi.Execute "DELETE a_par_codigo_barra FROM a_par_codigo_barra WHERE cli_codigo NOT IN ('" & codcas & "')"

'-------> Actualizar campo cencos de tabla b_minuta
dbi.Execute "UPDATE b_minuta SET min_cencos = '" & codcas & "'"

'-------> Actualizar campo min_racteo  min_racrea si Tipo de Minuta es MIO

dbi.Execute "UPDATE b_minuta SET min_racteo = 0,  min_racrea = 0 WHERE '" & minsre & "' = '2'"

'-------> Actualizar campo mid_numrac si  si Tipo de Minuta es MIO

dbi.Execute "UPDATE b_minutadet SET mid_numrac = 0 WHERE '" & minsre & "' = '2'"

dbi.Execute "INSERT INTO a_param VALUES ('5etapas', 'Casinos 5 Etapas', 'C', 'S')"
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'addreceta', 'Adicional Receta', 'N', pnr_nreceta FROM b_paramnreceta IN " & dBo & " WHERE pnr_codseg IN (" & subseg & ")"
'-------> Generar parametros ejecutivos contables
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'datcont', mid(cli_nomcontable,1,40), 'C', cli_emailcontable FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> generar email envio pedido
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'emailenped', 'Email Envio Pedido', 'C', cli_emailenviopedido FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"

'-------> Insert concepto grupo vulnerable a tabla a_param.
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'opgruvul', 'Opción Grupo Vulnerable', 'C', iif(cli_gruvul = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> Insert concepto modulo paciente.
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modpac', 'Modulo Paciente', 'C', iif(cli_modpac = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> Insert concepto parametro proveedor
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modprove', 'Parametro Modificar Proveedor', 'N', '0' FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> Insert concepto generación pedido Web o SGP
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT DISTINCT 'gpedsgpweb', 'Parametro Generación Pedido x SGP o Web', 'C', cli_opgped FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> Insert concepto Hipersensibilidad Alimentaria tabla a_param.
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'hipali', 'Opción Hipersensibilidad Alimentaria', 'C', iif(cli_hipali = 'S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> Insert concepto Tipo Operación tabla a_param.
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'tipope', 'Tipo Operación 0=Gravada:1=No Gravada', 'C', cli_tipope FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
'-------> Insert concepto Minuta Sitio Remoto tabla a_param.
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'minsre', 'Minuta Sitio Remoto 0=No:1=SI', 'C', cli_minsre FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"


'-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA TEORICA 2013-01-11
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'blockmiteo', 'Bloqueo Minuta Teorica 0=No:1=SI', 'C', cli_blockminteo FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"

'-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA REAL 2013-01-11
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'blockmirea', 'Bloqueo Minuta Real 0=No:1=SI', 'C', cli_blockminreal FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"

'-------> INSERT - MVA - PARAMETRO DE BLOQUEO DE MINUTA (BLOQUEO MINUTA) 2013-01-11
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'blockmicon', 'Bloqueo Minuta 0=No:1=SI', 'C', cli_blockmincontrato FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"

'-------> INSERT - MVA - PARAMETRO DE TRABAJA FIN SEMANA (BLOQUEO MINUTA) 2013-03-08
dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'trabfinsem', 'Trabaja Fin Semana 0=No:1=SI', 'C', cli_blockmintrabajafinsemana FROM b_clientes IN " & dBo & " WHERE cli_codigo = '" & codcas & "' AND (cli_tipo = 0 OR cli_tipo = 2)"


'-------> Mover datos a la tabla centro de costo
dbi.Execute "INSERT INTO a_cencos (cen_codigo, cen_socsap, cen_sobrec, cen_codmun, cen_ccisac, cen_cecsac, cen_codreg, cen_codopt) VALUES ('" & codcas & "', '" & socsap & "', '" & sobrec & "', " & codmun & ", " & ccisac & ", '" & cecsac & "', " & codrge & ", '" & CodOpt & "')"

'-------> Mover datos parametros despachos
'dbi.Execute "INSERT INTO b_paramdesp SELECT DISTINCT pad_cencos, pad_codtip, pad_tipo, pad_diaseg, pad_diario FROM b_parametrodespachos IN " & dBo & " WHERE pad_cencos = '" & codcas & "'"
dbi.Execute "INSERT INTO b_paramdesp SELECT DISTINCT pad_cencos, pad_codtip AS pad_codtip, pad_tipo, pad_diaseg, pad_diario FROM b_parametrodespachos IN " & dBo & " WHERE pad_cencos = '" & codcas & "'"
'-------> Mover datos días inhabiles
dbi.Execute "INSERT INTO b_Fecha_Inhabiles SELECT DISTINCT CFI_CeCo, CFI_Fecha, CFI_Glosa FROM Cas_b_Fecha_Inhabiles IN " & dBo & " WHERE CFI_CeCo = '" & codcas & "'"
'-------> Mover datos casino tipo actividades
dbi.Execute "INSERT INTO b_casinotipoactividades SELECT DISTINCT cta_cencos, cta_tipact FROM b_casinotipoactividades IN " & dBo & " WHERE cta_cencos = '" & codcas & "'"
'-------> Mover datos casino parametro stock
dbi.Execute "INSERT INTO b_casinoparametrostock SELECT DISTINCT cps_cencos, cps_invsto, cps_reqmen, cps_porinv, cps_liscri, cps_diario, cps_ajuimp FROM b_casinoparametrostock IN " & dBo & " WHERE cps_cencos = '" & codcas & "'"
'-------> Mover datos clase documento sap
dbi.Execute "INSERT INTO a_clasedocsap SELECT DISTINCT cds_coddoc, cds_codreg, cds_cdosap FROM a_clasedocsap IN " & dBo & " WHERE cds_codreg = " & codrge & ""

'------->
'-------> Borrar tabla auxiliares
'------->
dbi.Execute "DROP table b_recetaaux"
dbi.Execute "DROP table b_recetadetaux"
dbi.Execute "DROP table b_tablagramajeaux"
dbi.Execute "DROP table b_tablagramajeauxceco"
dbi.Execute "DROP table a_subsegmentoaux"
dbi.Execute "DROP table tmp_receta"
dbi.Execute "DROP table gra_receta"
dbi.Close: Set dbi = Nothing
If Dir(mdir & Mid(destinofile, 1, (Len(destinofile) - 3)) & "ldb") = "" And Trim(Environ("OS")) <> "" Then
   If Dir(mdir & "xxxpla.mdb") <> "" Then Kill mdir & "xxxpla.mdb"
   DBEngine.CompactDatabase mdir & destinofile, mdir & "xxxpla.mdb", dbLangGeneral
   Kill mdir & destinofile
   fso.MoveFile mdir & "xxxpla.mdb", mdir & destinofile
End If
'-------> verificar si existe archivo zip destino si existe borrar
If Dir(mdir & destinofilezip) <> "" Then Kill mdir & destinofilezip
AZ1.CreateZip mdir & destinofilezip, "": AZ1.AddFile mdir & destinofile, "", True, "": AZ1.Close
'-------> verificar si existe archivo mdb destino si existe borrar
If Dir(mdir & destinofile) <> "" Then Kill mdir & destinofile
'-------> leer contrato
DoEvents
Set RS = vg_db.Execute("SELECT * FROM b_clientes With(NoLock) WHERE cli_codigo = '" & codcas & "'")
If Not RS.EOF Then
   If RS!cli_openvio = 1 Then
      csubjects = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "Se Informa que el maestro de planificación del mes " & Format(fpDateTime1.text, "mm/yyyy") & " esta disponible. Para que usted pueda actualizar ", "Se Informa que las parametrizaciones 5 etapas esta disponible. Para que usted pueda actualizar ")
      cBody = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "Se Informa que el maestro de planificación esta disponible. Para que usted pueda actualizar", "Se Informa que las parametrizaciones 5 etapas esta disponible. Para que usted pueda actualizar")
      '-------> Traer datos FTP
      Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param With(NoLock) WHERE upper(par_codigo) LIKE '%" & LimpiaDato(UCase("ftp")) & "%'")
      If RS1.EOF Then fg_descarga: RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: Frame1(0).Enabled = True: Frame1(1).Enabled = True: Bar1(0).Visible = False: Bar1(1).Visible = False: MsgBox "No existe Parametrización FTP, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Function
      Do While Not RS1.EOF
         If RS1!par_codigo = "ftpser" Then CHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
         If RS1!par_codigo = "ftpdir" Then Cdire = fg_Desencripta(TipoDato(RS1!par_valor, ""))
         If RS1!par_codigo = "ftpusu" Then Cuser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
         If RS1!par_codigo = "ftppas" Then Cpass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
         If RS1!par_codigo = "ftppue" Then Cpuer = fg_Desencripta(TipoDato(RS1!par_valor, ""))
         RS1.MoveNext
      Loop
      RS1.Close: Set RS1 = Nothing
      a = oFTP.Version
      oFTP.UseIEProxy = False
      oFTP.Port = Cpuer '21
      oFTP.HostName = CHost '"sgp.sodexhochile.cl" '"64.76.138.76" '"64.76.45.71"  'fg_Desencripta(TipoDato(cHost, ""))
      oFTP.UserName = Cuser '"userftp" '"sodexho"   'fg_Desencripta(TipoDato(cUser, ""))
      oFTP.password = Cpass '"*sdxo7528*" '"*sdxo123*" '"shx873" 'fg_Desencripta(TipoDato(cPass, ""))
      oFTP.Connect
      If oFTP.IsConnected Then
         lDir = oFTP.GetCurrentDirListing("*.*")
         oFTP.SaveLastError ("aaa.xml")
'         a = oFTP.ChangeRemoteDir("/casinos/bd")
         a = oFTP.ChangeRemoteDir(Cdire)
         oFTP.SaveLastError ("aaa.xml")
         lDir = oFTP.GetCurrentDirListing("*.*")
         oFTP.SaveLastError ("aaa.xml")
         a = oFTP.PutFile(mdir & destinofilezip, destinofilezip)
         oFTP.SaveLastError ("aaa.xml")
         oFTP.Disconnect
         If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
            fg_descarga
            MsgBox "Casino : (" & Trim(RS!cli_codigo) & ") " & Trim(RS!cli_nombre) & " no se puede notificar por correo, ya que no tiene asignado el mail", vbInformation + vbOKOnly, MsgTitulo
            fg_carga ""
         Else
'            SendMail1 oMail, csubjects, cBody, mdir & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 0, logenv
            SendMailOutlook oMail, csubjects, csubjects, mdirpc & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1, logenv
         End If
      End If
   ElseIf RS!cli_openvio = 2 Then
      If IsNull(RS!cli_email) Or Trim(RS!cli_email) = "" Then
         fg_descarga
         MsgBox "Casino : (" & Trim(RS!cli_codigo) & ") " & Trim(RS!cli_nombre) & " no será enviado por correo, ya que no tiene asignado el mail, solamente se genero como archivo", vbInformation + vbOKOnly, MsgTitulo
         fg_carga ""
      Else
         csubjects = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "Adjunto archivo planificación " & Format(fpDateTime1.text, "mm/yyyy"), "Se Informa que las parametrizaciones 5 etapas esta disponible. Para que usted pueda actualizar ")
         cBody = IIf(Val(fg_codigocbo(Combo1, 0, 1, "")) = 0, "Adjunto archivo planificación " & Format(fpDateTime1.text, "mm/yyyy") & ". Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar", "Adjunto archivo parametrización 5 etapa esta disponible. Este archivo Ud. tiene que guardar en la siguiente carpeta C:\Archivos de programa\sgp\actualizar")
'         SendMail1 oMail, csubjects, cBody, mdir & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1, logenv
         SendMailOutlook oMail, csubjects, cBody, mdirpc & destinofilezip, Trim(RS!cli_nombre), Trim(RS!cli_email), 1, logenv
      End If
   End If

   If Trim(vg_GlosaEnvioCorreo) <> "" And RS!cli_openvio <> 3 Then
      EstError = False
   End If
   
End If
RS.Close: Set RS = Nothing
End Function

Private Sub GenerarListaPrecioContratono5Etapas(sourcefile As String, mdirpc As String, dBo As String, nEst As Boolean, spid As Long, codcas)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim icopy As Boolean, i As Long, auxcas As String, auxdtref As String, auxnrsem As Long
Dim cdcen As String, cdfil As String, dtref As String, nrsem As Long
Dim cencos As String, nomcencos As String, codpro As String, sourcefilezip As String, destinofile As String, destinofilezip As String, mdirserver As String, lognarchsou As String, socsap As String
Dim fso, codtis As Long, CodSeg As Long
Dim dbosac  As Variant, cDBI As Variant
Set fso = CreateObject("Scripting.FileSystemObject")

'-------> Acceso base SAC
dbosac = "'' [ODBC;Driver={Microsoft ODBC for Oracle};SERVER=" + vgsac_NSvr + ";uid=" + vgsac_NUsr + ";pwd=" + vgsac_Pass + "]"

'Set RS = vg_db.Execute("sgpadm_s_cliente_V02 " & IIf(nEst, 16, 17) & ", '',''")
Set RS = vg_db.Execute("sgpadm_s_cliente_V02listaprecio " & IIf(nEst, 1, 2) & ", '" & codcas & "', '" & Format(dBoM("01/" & Format(fpDateTime1.text, "mm/yyyy")), "mm/dd/yyyy") & "', '" & Format(dEoM("25/" & Format(fpDateTime1.text, "mm/yyyy")), "mm/dd/yyyy") & "', '" & vg_NUsr & "', " & spid & "")
auxdtref = "": auxnrsem = 0: auxcas = ""
Label1(3).Visible = True: Label1(3).Caption = "Enviando Información Contratos no 5 Etapas"
Bar1(0).Visible = True
Bar1(0).Value = 0: icopy = False: i = 1
Do While Not RS.EOF
   DoEvents
     
   Label1(1).Visible = True
   Label1(1).Caption = Trim(RS!cli_codigo) & " " & Trim(RS!cli_nombre)
     
   '-------> Crear archivos *.MDB y *.ZIP
'   destinofile = "sgp" & LCase(RS!cli_codigo) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".mdb"
'   destinofile = "sgp" & LCase(RS!cli_codigo) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".kkk"
'   destinofilezip = "sgp" & LCase(RS!cli_codigo) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
   destinofile = "sgp" & (RS!cli_codigo) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".kkk"
   destinofilezip = "sgp" & (RS!cli_codigo) & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
   '-------> verificar si existe archivo mdb destino si existe borrar y copiar
   DoEvents
   If Dir(mdirpc & destinofile) <> "" Then Kill mdirpc & destinofile
   FileCopy mdirpc & sourcefile, mdirpc & destinofile
   '---------------------------
   '------- Abrir base contrato
   '---------------------------
   cDBI = mdirpc & destinofile
   Set dbi = New ADODB.Connection
   dbi.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cDBI & "' ;Persist Security Info=False"
   dbi.ConnectionTimeout = 3600
   dbi.CommandTimeout = 3600
   dbi.Open
   
   '-------> Mover datos a variable auxiliar
   auxcas = Trim(RS!CADFIL_CDFIL)
   auxdtref = RS!CICCPA_DTREF
   auxnrsem = RS!CICCPA_NRSEM
   '-------> Corte Control
   Do While Not RS.EOF And Trim(RS!CADFIL_CDFIL) = Trim(auxcas)
      DoEvents
      Bar1(0).Value = Val((i / RS!nReg) * 100)
      '-------> Insert tabla lista precio sac
      dbi.Execute "INSERT INTO b_sac_listaprecio (lps_cencos, lps_periodo, lps_codsac, lps_precio) " & _
                  "SELECT DISTINCT CADFIL_CDFIL, '" & Format(fpDateTime1.text, "yyyymm") & "', CPOPRO_CDPRO, FORPRO_VLPCO " & _
                  "FROM vw_sgp_listaprecio IN " & dbosac & " " & _
                  "WHERE TABCEN_CDCEN = '" & RS!TABCEN_CDCEN & "' " & _
                  "AND   CADFIL_CDFIL = '" & RS!CADFIL_CDFIL & "' " & _
                  "AND   CICCPA_DTREF = '" & RS!CICCPA_DTREF & "' " & _
                  "AND   CICCPA_NRSEM = " & RS!CICCPA_NRSEM & ""
      RS.MoveNext: i = i + 1
      If RS.EOF Then Exit Do
   Loop
            
   '-------> Insert tabla formato compras sac
   dbi.Execute "INSERT INTO b_formatocompras (foc_codsac, foc_codcat, foc_nomsac, foc_unisac, foc_vigini, foc_flexec, foc_vigfin, foc_faccon) " & _
               "SELECT DISTINCT foc_codsac, foc_codcat, foc_nomsac, foc_unisac, foc_vigini, foc_flexec, foc_vigfin, foc_faccon FROM b_formatocompras IN " & dBo & " WHERE foc_codsac IN (SELECT DISTINCT lps_codsac FROM b_sac_listaprecio)" ' WHERE lps_cencos = '" & auxcas & "')" ' AND lps_periodo = '" & auxDTREF & "')"
         
   dbi.Execute "INSERT INTO b_formatocomprassgp (fcs_codsac, fcs_codsgp, fcs_sgppre) " & _
               "SELECT DISTINCT a.fcs_codsac, a.fcs_codsgp, a.fcs_sgppre FROM b_formatocomprassgp a, b_formatocompras b IN " & dBo & " WHERE a.fcs_codsac = b.foc_codsac AND b.foc_codsac IN (SELECT DISTINCT lps_codsac FROM b_sac_listaprecio)" ' WHERE lps_cencos = '" & auxcas & "')" ' AND lps_periodo = '" & RS1!auxDTREF & "')"
          
   '-------> Grabar tabla b_minutacasino de enviado la información
   vg_db.Execute "DELETE b_listapreciocasino FROM b_listapreciocasino WHERE lpc_cencos='" & auxcas & "' AND lpc_periodo = '" & auxdtref & "' AND lpc_nrosem = " & auxnrsem & ""
   vg_db.Execute "INSERT INTO b_listapreciocasino (lpc_cencos, lpc_periodo, lpc_nrosem, lpc_fecenv) VALUES ('" & auxcas & "', '" & auxdtref & "', " & auxnrsem & ", '" & Format(Date, "yyyymmdd") & "')"
   DoEvents
      
      
   '-------> Generar parametros ejecutivos contables
   dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'datcont', mid(cli_nomcontable,1,40), 'C', cli_emailcontable FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & auxcas & "' AND (cli_tipo=0 OR cli_tipo=2)"
   dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT '5etapas', 'Casino 5 Etapas', 'C', iif(cli_subseg=0,'N','S') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & auxcas & "' AND (cli_tipo=0 OR cli_tipo=2)"
   dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT par_codigo, par_nombre, par_tipo, par_valor FROM a_param IN " & dBo & " WHERE par_codigo='porprepro'"
   '-------> Insert concepto grupo vulnerable a tabla a_param.
   dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'opgruvul', 'Opción Grupo Vulnerable', 'C', iif(cli_gruvul='S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & auxcas & "' AND (cli_tipo=0 OR cli_tipo=2)"
   '-------> Insert concepto modulo paciente.
   dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modpac', 'Modulo Paciente', 'C', iif(cli_modpac='S','S','N') FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & auxcas & "' AND (cli_tipo=0 OR cli_tipo=2)"
   '-------> Insert concepto parametro proveedor
   dbi.Execute "INSERT INTO a_param (par_codigo, par_nombre, par_tipo, par_valor) SELECT 'modprove', 'Parametro Modificar Proveedor', 'N', '0' FROM b_clientes IN " & dBo & " WHERE cli_codigo='" & auxcas & "' AND (cli_tipo=0 OR cli_tipo=2)"
   codtis = 0: CodSeg = 0: socsap = ""
   RS2.Open "SELECT * FROM b_clientes WHERE cli_codigo='" & auxcas & "'", vg_db, adOpenForwardOnly
   If Not RS2.EOF Then
      codtis = IIf(IsNull(RS2!cli_codtis), 0, RS2!cli_codtis)
      CodSeg = IIf(IsNull(RS2!cli_codseg), 0, RS2!cli_codseg)
      socsap = IIf(IsNull(RS2!cli_socsap), "", RS2!cli_socsap)
   End If
   RS2.Close: Set RS2 = Nothing

   '-------> Borrar tabla casino envia sap
   dbi.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos NOT IN ('" & auxcas & "')"
   '-------> Mover datos a la tabla centro de costo
   dbi.Execute "INSERT INTO a_cencos (cen_codigo, cen_socsap) VALUES ('" & auxcas & "', '" & socsap & "')"
   '-------> borrar información
   If Not nEst Then
      dbi.Execute "DELETE * FROM a_estservicio"
      dbi.Execute "DELETE * FROM a_regimen"
      dbi.Execute "DELETE * FROM a_servicio"
      dbi.Execute "DELETE * FROM b_costopatron"
      dbi.Execute "DELETE * FROM b_gramofamproducto"
      dbi.Execute "DELETE * FROM b_minuta"
      dbi.Execute "DELETE * FROM b_minutadet"
      dbi.Execute "DELETE * FROM b_receta"
      dbi.Execute "DELETE * FROM b_recetadet"
      dbi.Execute "DELETE * FROM b_tablagramaje"
   End If
   '------->
   '-------> Borrar tabla auxiliares
   '------->
   If nEst Then
      dbi.Execute "DROP table b_recetaaux"
      dbi.Execute "DROP table b_recetadetaux"
      dbi.Execute "DROP table b_tablagramajeaux"
      dbi.Execute "DROP table a_subsegmentoaux"
      dbi.Execute "DROP table tmp_receta"
      dbi.Execute "DROP table gra_receta"
   End If
   '-------> Cerrar base access
   dbi.Close: Set dbi = Nothing
   DoEvents
           
   '-------> verificar si existe archivo zip destino si existe borrar
   If Dir(mdirpc & destinofilezip) <> "" Then Kill mdirpc & destinofilezip
   AZ1.CreateZip mdirpc & destinofilezip, "": AZ1.AddFile mdirpc & destinofile, "", True, "": AZ1.Close
   '-------> verificar si existe archivo mdb destino si existe borrar
   If Dir(mdirpc & destinofile) <> "" Then Kill mdirpc & destinofile
   '-------> leer casino
   DoEvents
Loop
RS.Close: Set RS = Nothing
Frame1(0).Enabled = True
Toolbar1.Enabled = True
Bar1(0).Value = 0: Bar1(0).Visible = False

Man_Error:
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
Dim nivel As Long, i As Long, nReg As Long, padre As Variant
Select Case Button.Index
Case 1
    SSTab1.Tab = 0
    TvwDir.Nodes.Clear
    Set RS = vg_db.Execute("SELECT COUNT(min_subseg) AS nreg FROM b_minuta  With(NoLock) WHERE substring(convert(char(8),min_fecmin),1,6) = " & Val(Format(fpDateTime1.text, "yyyymm")) & "")
    If RS.EOF Or RS!nReg = 0 Then RS.Close: Set RS = Nothing: MsgBox "No existe planificación, para este periodo ", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    RS.Close: Set RS = Nothing
    fg_carga ""
    Dim xcencos As String, psubseg As String, pcencos As String, pcodreg As String
    Dim xsubseg As Long, xcodreg As Long, xcodser As Long
    ArNivTree(0) = "Casino"   'Código Servicio
    ArNivTree(1) = 6  'Largo de Subsegmento
    ArNivTree(2) = 16  'Largo de Contrato + Subsegmento
    ArNivTree(3) = 21 'Largo de Regimen + Contrato + Subsegmento
    nivel = 65: i = 1
    xsubseg = 0: xcodreg = 0: xcodser = 0: xcencos = "": psubseg = "": pcencos = "": pcodreg = "": nReg = 0
    Set RS = vg_db.Execute("sgpadm_s_traerenvioplanif " & Val(Format(fpDateTime1.text, "yyyymm")) & "")
    Bar1(0).Visible = True: Bar1(0).Value = 0
    If Not RS.EOF Then
       nReg = RS!nReg
        Do While Not RS.EOF
           Bar1(0).Value = Val((i / nReg) * 100)
            If RS!sub_codigo <> xsubseg Then
                padre = Chr(nivel)
                Set rootNode = TvwDir.Nodes.Add(, , "N" & fg_pone_espacio(RS!sub_codigo, 5), RS!sub_codigo & " - " & Trim(RS!sub_nombre))
                psubseg = "": psubseg = "N" & fg_pone_espacio(RS!sub_codigo, 5): xsubseg = RS!sub_codigo
                xcencos = "": xcodreg = 0: xcodser = 0
            End If
            If Trim(RS!cli_codigo) <> Trim(xcencos) Then
                padre = Chr(nivel)
                Set rootNode = TvwDir.Nodes.Add(psubseg, tvwChild, psubseg & fg_pone_espacio(RS!cli_codigo, 10), Trim(RS!cli_codigo) & " - " & Trim(RS!cli_nombre))
                pcencos = "": pcencos = psubseg & fg_pone_espacio(RS!cli_codigo, 10): xcencos = Trim(RS!cli_codigo)
                If RS!mic_cencos = 1 Then TvwDir.Nodes.item(rootNode.Index).ForeColor = Shape1(0).FillColor: TvwDir.Nodes.item(rootNode.Index).Bold = True
                xcodreg = 0: xcodser = 0
            End If
            If RS!reg_codigo <> xcodreg Then
                Set rootNode = TvwDir.Nodes.Add(pcencos, tvwChild, pcencos & fg_pone_espacio(RS!reg_codigo, 5), RS!reg_codigo & " - " & Trim(RS!reg_nombre))
                pcodreg = "": pcodreg = pcencos & fg_pone_espacio(RS!reg_codigo, 5): xcodreg = RS!reg_codigo
                If RS!mic_cencos = 1 Then TvwDir.Nodes.item(rootNode.Index).ForeColor = Shape1(0).FillColor: TvwDir.Nodes.item(rootNode.Index).Bold = True
                xcodreg = RS!reg_codigo
                xcodser = 0
            End If
            If RS!ser_codigo <> xcodser Then
               If Trim(RS!mic_fecenv) <> "" Then
                  Set rootNode = TvwDir.Nodes.Add(pcodreg, tvwChild, pcodreg & "CASINO" & RS!ser_codigo, RS!ser_codigo & " - " & Trim(RS!ser_nombre) & " - " & Mid(RS!mic_fecenv, 7, 2) & "/" & Mid(RS!mic_fecenv, 5, 2) & "/" & Mid(RS!mic_fecenv, 1, 4))
               Else
                  Set rootNode = TvwDir.Nodes.Add(pcodreg, tvwChild, pcodreg & "CASINO" & RS!ser_codigo, RS!ser_codigo & " - " & Trim(RS!ser_nombre))
               End If
               If RS!mic_cencos = 1 Then TvwDir.Nodes.item(rootNode.Index).ForeColor = Shape1(0).FillColor: TvwDir.Nodes.item(rootNode.Index).Bold = True
               xcodser = RS!ser_codigo
            End If
            RS.MoveNext: i = i + 1
        Loop
    Else
       Bar1(0).Visible = False: fg_descarga
       MsgBox "No existe información parametrización", vbExclamation + vbOKOnly, MsgTitulo
    End If
    RS.Close: Set RS = Nothing
    Bar1(0).Visible = False: fg_descarga
    '-------> Abrir base sac
    fg_descarga
End Select

Exit Sub
Man_Error:
If Err = 3034 Then RS.Close: Set RS = Nothing: Exit Sub
RS.Close: Set RS = Nothing
Resume Next
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub TvwDir_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim cKey As String, lKey As Integer, i As Long, lCheck As Boolean, lGraba As Boolean
TvwDir.Nodes.item(Node.Key).Selected = True
lCheck = TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Checked
cKey = Trim(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Key)
lKey = Len(cKey)

Dim MarcarDesc As Boolean, INiv As Integer, RecNivel As String
If lCheck Then
    MarcarDesc = True: INiv = 1
    RecNivel = Mid(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Key, 1, ArNivTree(INiv))
End If
'------->
For i = 1 To TvwDir.Nodes.count
    If cKey = Mid(TvwDir.Nodes.item(i).Key, 1, lKey) Then
        TvwDir.Nodes.item(i).Checked = lCheck
        lGraba = True
    End If
    '-------> Comando marcas descendentes
    If MarcarDesc And Trim(TvwDir.Nodes.item(i).Key) = RecNivel Then
        INiv = INiv + 1
        RecNivel = Mid(TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Key, 1, ArNivTree(INiv))
        TvwDir.Nodes.item(i).Checked = True
    End If
    '------->
Next i
fg_descarga
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows < 1 Then Exit Sub
If Col = 1 And Row = 0 Then vaSpread1.Row = -1: vaSpread1.Col = 1: vaSpread1.text = IIf(vaSpread1.Value = "1", "0", "1")
End Sub
