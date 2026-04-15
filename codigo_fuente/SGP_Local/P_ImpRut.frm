VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_ImpRut 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importar Datos de Rutas"
   ClientHeight    =   7935
   ClientLeft      =   3840
   ClientTop       =   1755
   ClientWidth     =   8475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   5055
      Left            =   120
      TabIndex        =   15
      Top             =   2760
      Width           =   7575
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4455
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   7335
         _Version        =   393216
         _ExtentX        =   12938
         _ExtentY        =   7858
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
         MaxRows         =   0
         SpreadDesigner  =   "P_ImpRut.frx":0000
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
         Index           =   1
         Left            =   120
         TabIndex        =   17
         Top             =   4800
         Visible         =   0   'False
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   7575
      Begin VB.Frame Frame2 
         Caption         =   "Opción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   7
         Top             =   1560
         Width           =   5895
         Begin VB.OptionButton Option1 
            Caption         =   "Casino"
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
            Left            =   4200
            TabIndex        =   10
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Calendario"
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
            Left            =   2400
            TabIndex        =   9
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Productos"
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
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "P_ImpRut.frx":2F69
         Left            =   1515
         List            =   "P_ImpRut.frx":2F6B
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1035
         Width           =   5895
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Left            =   1515
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   600
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1520
         TabIndex        =   11
         Top             =   195
         Visible         =   0   'False
         Width           =   915
         _Version        =   196608
         _ExtentX        =   1614
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ruta"
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
         Left            =   120
         TabIndex        =   13
         Top             =   285
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2925
         TabIndex        =   12
         Top             =   195
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2460
         Picture         =   "P_ImpRut.frx":2F6D
         Top             =   120
         Visible         =   0   'False
         Width           =   480
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
         Left            =   120
         TabIndex        =   6
         Top             =   1125
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Formato Excel"
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
         TabIndex        =   5
         Top             =   675
         Width           =   1215
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000011&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1560
         TabIndex        =   4
         Top             =   1145
         Width           =   5895
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2955
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   4455
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7935
      Left            =   7845
      TabIndex        =   0
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   13996
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog Cd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "P_ImpRut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Dim RS As New ADODB.Recordset
Dim Msgtitulo  As String
Dim opcion As String
Dim handle As Integer

Private Sub Combo1_Click()
   Dim Est As Boolean
   Dim codpro As String, codcco As String, codrut As Long, codlis As Long, codpne As Long, fecha As String, precio As Double, i As Long, fectop As String, dia As Long, hora As String, pn As String, pa As String, a As String
   Dim cantidad As Double
   Dim filepath As String, FechaWeb As String
   Dim dbexcel As Database, cSpi As Long
   fg_carga ""
   vaSpread2.MaxRows = 500
   x = vaSpread2.ImportExcelSheet(handle, Combo1.ListIndex): fg_descarga
   vaSpread2.RowHeight(0) = 28
   vaSpread2.ColWidth(1) = 10
   vaSpread2.Row = -1
   vaSpread2.Col = -1
   vaSpread2.Lock = True
   If opcion = "acprod" Then
      vaSpread2.MaxCols = 4
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Código Producto"
      vaSpread2.Col = 2
      vaSpread2.text = "Central de Compras"
      vaSpread2.Col = 3
      vaSpread2.text = "Cantidad"
      vaSpread2.Col = 4
      vaSpread2.text = "Fecha Inicio"
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      Est = True
      i = 1
      rsexcel.MoveFirst
      vaSpread2.TextTip = TextTipFloating
      Do While rsexcel.EOF <> True
         If rsexcel.Fields(0).Value = "*" Then Exit Do
         vaSpread2.Row = i
         '-------> validar código producto sac
         vaSpread2.Col = 1
         codpro = ""
         codpro = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, "")
         vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_productos 1, '" & codpro & "'")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Producto No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar código central de compras
         codcco = ""
         vaSpread2.Col = 2
         codcco = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), LimpiaDato(Trim(rsexcel.Fields(1).Value)), "")
         vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_centralcompras 3, 0, '" & Trim(TipoDato(codcco, "")) & "'")
         If RS.EOF Then
            vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Central de Compras No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> Validar si producto esta asociado a central de compras
         Set RS = vg_dbpedweb.Execute("pedweb_s_productoscentral 1, '" & codpro & "', '" & Trim(TipoDato(codcco, "")) & "'")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Producto No corresponde central de compras"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar cantidad
         vaSpread2.Col = 3
         cantidad = 0
         cantidad = Val(vaSpread2.text) 'IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, 0)
         vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Not IsNumeric(cantidad) Or cantidad = 0 Then
            vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato cantidad no corresponde"
            Est = False
         End If
         '-------> validar fecha
         vaSpread2.Col = 4
         fecha = ""
         fecha = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(3).Value), rsexcel.Fields(3).Value, "")
         vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Not IsDate(fecha) Then
            vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato de fecha no corresponde"
            Est = False
         End If
         i = i + 1
         rsexcel.MoveNext
      Loop
      rsexcel.Close: Set rsexcel = Nothing
      If Not Est Then Toolbar1.Buttons(1).Enabled = False Else Toolbar1.Buttons(1).Enabled = True
      fg_descarga
   ElseIf opcion = "ruta" And Option1(0).Value = True Then '-------> Productos
      fg_carga ""
      vaSpread2.MaxCols = 3
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Ruta"
      vaSpread2.Col = 2
      vaSpread2.text = "Codigo Producto"
      vaSpread2.Col = 3
      vaSpread2.text = "Central de Compras"
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      rsexcel.MoveFirst
      Est = True
      i = 1
      vaSpread2.TextTip = TextTipFloating
      Do While rsexcel.EOF <> True
         DoEvents
         If rsexcel.Fields(0).Value = "*" Then Exit Do
         vaSpread2.Row = i
         '-------> validar código ruta
         vaSpread2.Col = 1
         codrut = 0
         codrut = Val(vaSpread2.text) 'IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, 0)
         vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 5, " & codrut & ", ''")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Ruta No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar código producto
         vaSpread2.Col = 2
         codpro = ""
         codpro = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_productos 1, '" & codpro & "'")
         If RS.EOF Then
            vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Producto No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar código central de compras
         vaSpread2.Col = 3
         codcco = ""
         codcco = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
         vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_centralcompras 2, 0, '" & Trim(TipoDato(codcco, "")) & "'")
         If RS.EOF Then
            vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Central de Compras No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing: i = i + 1
         rsexcel.MoveNext
      Loop
      rsexcel.Close: Set rsexcel = Nothing
      If Not Est Then Toolbar1.Buttons(1).Enabled = False Else Toolbar1.Buttons(1).Enabled = True
      fg_descarga
   ElseIf opcion = "ruta" And Option1(1).Value = True Then '-------> Calendario
      fg_carga ""
      vaSpread2.MaxCols = 5
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Código Ruta"
      vaSpread2.Col = 2
      vaSpread2.text = "Fecha Despacho"
      vaSpread2.Col = 3
      vaSpread2.text = "Fecha Tope"
      vaSpread2.Col = 4
      vaSpread2.text = "Día"
      vaSpread2.Col = 5
      vaSpread2.text = "Hora"
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      rsexcel.MoveFirst
      Est = True
      i = 1
      vaSpread2.TextTip = TextTipFloating
      Do While rsexcel.EOF <> True
         If rsexcel.Fields(0).Value = "*" Then Exit Do
         vaSpread2.Row = i
         '-------> validar código ruta
         vaSpread2.Col = 1
         codrut = 0
         codrut = Val(vaSpread2.text) 'IIf(Not IsNumeric(rsexcel.Fields(0).Value) And Not IsNull(rsexcel.Fields(0).Value), Val(rsexcel.Fields(0).Value), 0)
         vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 5, " & codrut & ", ''")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Ruta No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar fecha
         vaSpread2.Col = 2
         fecha = ""
         fecha = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(fecha) = "" Or Not IsDate(fecha) Then
            vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato fecha no corresponde"
            Est = False
         End If
         '-------> validar fecha top
         vaSpread2.Col = 3
         fectop = ""
         fectop = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
         vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(fectop) = "" Or Not IsDate(fectop) Then
            vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato fecha no corresponde"
            Est = False
         End If
         '-------> validar dia
         vaSpread2.Col = 4
         dia = 0
         dia = Val(vaSpread2.text) 'IIf(Not IsNumeric(rsexcel.Fields(3).Value), 0, rsexcel.Fields(3).Value)
         vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Not IsNumeric(dia) Or dia = 0 Then
            vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato día no corresponde"
            Est = False
         End If
         '-------> validar hora
         vaSpread2.Col = 5
         hora = ""
         hora = vaSpread2.text 'IIf(Not IsNumeric(rsexcel.Fields(4).Value), rsexcel.Fields(4).Value, 0)
         vaSpread2.Col = 5: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If hora = "" Or IsNull(hora) Then
            vaSpread2.Col = 5: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato hora no corresponde"
            Est = False
         End If
         rsexcel.MoveNext: i = i + 1
      Loop
      rsexcel.Close: Set rsexcel = Nothing
      If Not Est Then Toolbar1.Buttons(1).Enabled = False Else Toolbar1.Buttons(1).Enabled = True
      fg_descarga
   ElseIf opcion = "ruta" And Option1(2).Value = True Then '-------> Casino
      fg_carga ""
      vaSpread2.MaxCols = 4
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Código Ruta"
      vaSpread2.Col = 2
      vaSpread2.text = "Fecha"
      vaSpread2.Col = 3
      vaSpread2.text = "Código Sac"
      vaSpread2.Col = 4
      vaSpread2.text = "Fecha Web"
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      rsexcel.MoveFirst
      Est = True
      i = 1
      vaSpread2.TextTip = TextTipFloating
      Do While rsexcel.EOF <> True
         If rsexcel.Fields(0).Value = "*" Then Exit Do
         vaSpread2.Row = i
         '-------> validar código ruta
         vaSpread2.Col = 1
         codrut = 0
         codrut = Val(vaSpread2.text) 'IIf(Not IsNumeric(rsexcel.Fields(0).Value) And Not IsNull(rsexcel.Fields(0).Value), Val(rsexcel.Fields(0).Value), 0)
         vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_ruta 5, " & codrut & ", ''")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Ruta No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar fecha
         vaSpread2.Col = 2
         fecha = ""
         fecha = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(fecha) = "" Or Not IsDate(fecha) Then
            vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato fecha no corresponde"
            Est = False
         End If
         '-------> validar código casino
         vaSpread2.Col = 3
         codcco = ""
         codcco = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
         Set RS = vg_dbpedweb.Execute("pedweb_s_clientes 2, '" & codcco & "', ''")
         vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If RS.EOF Then
            vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Centro de costo No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar fecha
         vaSpread2.Col = 4
         FechaWeb = ""
         FechaWeb = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(3).Value), rsexcel.Fields(3).Value, "")
         vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(FechaWeb) = "" Or Not IsDate(FechaWeb) Then
            vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato fecha web no corresponde"
            Est = False
         End If
         rsexcel.MoveNext: i = i + 1
      Loop
      rsexcel.Close: Set rsexcel = Nothing
      If Not Est Then Toolbar1.Buttons(1).Enabled = False Else Toolbar1.Buttons(1).Enabled = True
      fg_descarga
   ElseIf opcion = "lispre" Then '-------> Lista Precio
      fg_carga ""
      vaSpread2.MaxCols = 4
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Cod. Lista"
      vaSpread2.Col = 2
      vaSpread2.text = "Cod. Productos"
      vaSpread2.Col = 3
      vaSpread2.text = "Fecha"
      vaSpread2.Col = 4
      vaSpread2.text = "Precio"
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      vaSpread2.TextTip = TextTipFloating
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      rsexcel.MoveFirst
      Est = True
      i = 1
      vaSpread2.TextTip = TextTipFloating
      Do While rsexcel.EOF <> True
         If rsexcel.Fields(0).Value = "*" Then Exit Do
         vaSpread2.Row = i
         '-------> Validar código lista precio
         vaSpread2.Col = 1
         codlis = 0
         codlis = Val(vaSpread2.text) 'IIf(Not IsNumeric(rsexcel.Fields(0).Value) And Not IsNull(rsexcel.Fields(0).Value), Val(rsexcel.Fields(0).Value), 0)
         vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_listaprecios 3, " & codlis & ", '', ''")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "lista No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar código producto
         vaSpread2.Col = 2
         codpro = ""
         codpro = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_productos 1, '" & codpro & "'")
         If RS.EOF Then
            vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Producto No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar fecha
         vaSpread2.Col = 3
         fecha = ""
         fecha = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(fecha) = "" Or Not IsDate(fecha) Then
            vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato fecha no corresponde"
            Est = False
         End If
         '-------> validar precio
         vaSpread2.Col = 4
         precio = 0
         precio = Val(vaSpread2.text) 'IIf(Not IsNumeric(rsexcel.Fields(2).Value), 0, rsexcel.Fields(2).Value)
         vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Not IsNumeric(precio) Then
            vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Formato precio no corresponde"
            Est = False
         End If
         rsexcel.MoveNext: i = i + 1
      Loop
      rsexcel.Close: Set rsexcel = Nothing
      If Not Est Then Toolbar1.Buttons(1).Enabled = False Else Toolbar1.Buttons(1).Enabled = True
      fg_descarga
   ElseIf opcion = "regneg" Then
      fg_carga ""
      vaSpread2.MaxCols = 5
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Cód. Regla de Neg."
      vaSpread2.Col = 2
      vaSpread2.text = "Cód. Productos"
      vaSpread2.Col = 3
      vaSpread2.text = "Pedido Normal"
      vaSpread2.Col = 4
      vaSpread2.text = "Pedido Adicional"
      vaSpread2.Col = 5
      vaSpread2.text = "Anulación"
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      vaSpread2.TextTip = TextTipFloating
      sheetname = Trim(Combo1.text) & "$"
      filepath = Trim(fpText1.text)
      Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
      Set rsexcel = dbexcel.OpenRecordset(sheetname)
      rsexcel.MoveFirst
      Est = True
      i = 1
      vaSpread2.TextTip = TextTipFloating
      Do While rsexcel.EOF <> True
         If rsexcel.Fields(0).Value = "*" Then Exit Do
         vaSpread2.Row = i
         vaSpread2.Col = 1
         codpne = 0
         codpne = Val(vaSpread2.text) 'IIf(Not IsNumeric(rsexcel.Fields(0).Value) And Not IsNull(rsexcel.Fields(0).Value), Val(rsexcel.Fields(0).Value), 0)
         vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 4, " & codpne & ", ''")
         If RS.EOF Then
            vaSpread2.Col = 1: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Reglas de Negocios No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar código producto
         vaSpread2.Col = 2
         codpro = ""
         codpro = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         Set RS = vg_dbpedweb.Execute("pedweb_s_productos 1, '" & codpro & "'")
         If RS.EOF Then
            vaSpread2.Col = 2: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Producto No existe"
            Est = False
         End If
         RS.Close: Set RS = Nothing
         '-------> validar pedido normal
         vaSpread2.Col = 3
         pn = ""
         pn = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
         vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(pn) = "" Or IsNull(pn) Or (Trim(pn) <> "S" And Trim(pn) <> "N") Then
            vaSpread2.Col = 3: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Pedido normal no corresponde"
            Est = False
         End If
         '-------> validar pedido adicional
         vaSpread2.Col = 4
         pa = ""
         pa = vaSpread2.text 'IIf(Not IsNumeric(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
         vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(pa) = "" Or IsNull(pa) Or (Trim(pa) <> "S" And Trim(pa) <> "N") Then
            vaSpread2.Col = 4: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Pedido adicional no corresponde"
            Est = False
         End If
         '-------> validar pedido anulaciones
         vaSpread2.Col = 5
         a = ""
         a = vaSpread2.text 'IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
         vaSpread2.Col = 5: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndDoNotFireEvent: vaSpread2.CellNote = ""
         If Trim(a) = "" Or IsNull(a) Or (Trim(a) <> "S" And Trim(a) <> "N") Then
            vaSpread2.Col = 5: vaSpread2.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent: vaSpread2.CellNote = "Pedido adicional no corresponde"
            Est = False
         End If
         rsexcel.MoveNext: i = i + 1
      Loop
      rsexcel.Close: Set rsexcel = Nothing
      If Not Est Then Toolbar1.Buttons(1).Enabled = False Else Toolbar1.Buttons(1).Enabled = True
      fg_descarga
   End If
   vaSpread2.Row = -1
   vaSpread2.Col = -1
   vaSpread2.Lock = True
   fg_descarga
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim codTippla As Long, nomTippla As String
fg_centra Me
fg_carga ""
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = True
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Toolbar1.Buttons(1).Enabled = False
fg_descarga
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    If Val(fpLongInteger1(0).Value) < 1 Then fpayuda(0).Caption = "": Exit Sub
    Set RS = vg_dbpedweb.Execute("SELECT recorrido, descripcion FROM s_Recorrido WHERE recorrido = " & Val(fpLongInteger1(0).Value) & "")
    If RS.EOF Then RS.Close: Set RS = Nothing:: fpayuda(0).Caption = "":  Exit Sub
    fpayuda(0).Caption = Trim(RS!descripcion)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_ButtonHit(Button As Integer, NewIndex As Integer)
Dim List() As String
Dim ListCount As Integer
Dim fromRight As Long, i As Long
'Dim handle As Integer
Dim myPath As String
Dim f As Boolean
ReDim List(1)
CD.DialogTitle = "Seleccionar un archivo XLS"
CD.Filter = "Todos los archivos|*.*|Archivos de texto (*.xls)|*.xls"
CD.FilterIndex = 2
CD.Flags = cdlOFNFileMustExist
CD.ShowOpen
fpText1.text = CD.FileName
If Len(fpText1.text) <= 0 Then Exit Sub
fromRight = InStrRev(CD.FileName, "\", , vbTextCompare)
If fromRight > 1 Then
   myPath = Left(CD.FileName, fromRight)
End If
vaSpread2.MaxRows = 0
vaSpread2.Row = -1
vaSpread2.Col = -1
vaSpread2.Lock = True
f = vaSpread2.GetExcelSheetList(CD.FileName, List, ListCount, (myPath & "log.txt"), handle, True)
If (ListCount - 1 > 1) Then
   ReDim List(ListCount - 1)
   f = vaSpread2.GetExcelSheetList(CD.FileName, List, ListCount, (myPath & "log.txt"), handle, False)
End If
Combo1.Clear
For i = 0 To ListCount - 1
    Combo1.AddItem (List(i))
Next i
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    If opcion = "lispre" Then
       B_TabEst.LlenaDatos "s_Lista_Precios", "sub_", "Lista de Precios", "lispreweb"
    ElseIf opcion = "ruta" Then
       B_TabEst.LlenaDatos "s_Recorrido", "sub_", "Ruta", "recorrido"
    ElseIf opcion = "regneg" Then
       B_TabEst.LlenaDatos "s_ReglasDeNegocios", "sub_", "Reglas de Negocios", "regneg"
    End If
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpText1.SetFocus
End Select
End Sub

Sub LlenaDatos(titgen As String, op As String)
'-------> cargar familia productos
Me.Caption = IIf(op = "lispre", titgen, titgen)
Msgtitulo = IIf(op = "lispre", titgen, titgen)
Label1(4).Caption = IIf(op = "lispre", "List. Precio", IIf(op = "regneg", "Reglas de Neg.", "Ruta"))
Frame2.Visible = IIf(op = "lispre" Or op = "regneg" Or op = "acprod", False, True)
vaSpread2.ColWidth(1) = 10
If op = "acprod" Then
   Label1(4).Visible = False
   fpLongInteger1(0).Visible = False
   Image1(0).Visible = False
   fpayuda(0).Visible = False
   sombra(0).Visible = False
   vaSpread2.MaxRows = 0
   vaSpread2.MaxCols = 4
   vaSpread2.Row = 0
   vaSpread2.Col = 1
   vaSpread2.text = "Código Producto"
   vaSpread2.Col = 2
   vaSpread2.text = "Central de Compras"
   vaSpread2.Col = 3
   vaSpread2.text = "Cantidad"
   vaSpread2.Col = 4
   vaSpread2.text = "Fecha Inicio"
ElseIf op = "ruta" Then
   vaSpread2.MaxRows = 0
   vaSpread2.MaxCols = 3
   vaSpread2.Row = 0
   vaSpread2.Col = 1
   vaSpread2.text = "Ruta"
   vaSpread2.Col = 2
   vaSpread2.text = "Código Producto"
   vaSpread2.Col = 3
   vaSpread2.text = "Central de Compras"
ElseIf op = "lispre" Then
   vaSpread2.MaxRows = 0
   vaSpread2.MaxCols = 4
   vaSpread2.Row = 0
   vaSpread2.Col = 1
   vaSpread2.text = "Lista Precio"
   vaSpread2.Col = 2
   vaSpread2.text = "Código Producto"
   vaSpread2.Col = 3
   vaSpread2.text = "Fecha"
   vaSpread2.Col = 4
   vaSpread2.text = "Precio"
ElseIf op = "regneg" Then
   vaSpread2.MaxRows = 0
   vaSpread2.MaxCols = 5
   vaSpread2.Row = 0
   vaSpread2.Col = 1
   vaSpread2.text = "Código Regla de Neg."
   vaSpread2.Col = 2
   vaSpread2.text = "Código Producto"
   vaSpread2.Col = 3
   vaSpread2.text = "Pedido Normal"
   vaSpread2.Col = 4
   vaSpread2.text = "Pedido Adicional"
   vaSpread2.Col = 5
   vaSpread2.text = "Anulación"
End If
opcion = op
End Sub

Private Sub Option1_Click(Index As Integer)
Select Case Index
Case 0
      vaSpread2.MaxCols = 3
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Ruta"
      vaSpread2.Col = 2
      vaSpread2.text = "Codigo Producto"
      vaSpread2.Col = 3
      vaSpread2.text = "Central de Compras"
Case 1
      vaSpread2.MaxCols = 5
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Código Ruta"
      vaSpread2.Col = 2
      vaSpread2.text = "Fecha Despacho"
      vaSpread2.Col = 3
      vaSpread2.text = "Fecha Tope"
      vaSpread2.Col = 4
      vaSpread2.text = "Día"
      vaSpread2.Col = 5
      vaSpread2.text = "Hora"
Case 2
      vaSpread2.MaxCols = 4
      vaSpread2.Row = 0
      vaSpread2.Col = 1
      vaSpread2.text = "Código Ruta"
      vaSpread2.Col = 2
      vaSpread2.text = "Fecha"
      vaSpread2.Col = 3
      vaSpread2.text = "Código Sac"
      vaSpread2.Col = 4
      vaSpread2.text = "Fecha Web"
End Select
Toolbar1.Buttons(1).Enabled = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Est As Boolean, spid As Long
Select Case Button.Index
Case 1
'    If Val(fpLongInteger1(0).Value) = 0 And opcion <> "acprod" Then MsgBox "Debe seleccionar rutas ...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If vaSpread2.MaxRows < 1 Then MsgBox "No existe información a procesar ...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Dim codpro As String, codcco As String, codrut As Long, codlis As Long, codpne  As Long, fecha As String, precio As Double, i As Long, fectop As String, dia As Long, hora As String, pn As String, pa As String, a As String
    Dim cantidad As Double
    Dim filepath As String, FechaWeb As String
    Dim dbexcel As Database, cSpi As Long
    i = 0
    If opcion = "ruta" Then
       If Option1(0).Value = True Then '-------> Importar Productos
          Label2(1).Caption = ""
          Label2(1).Visible = True
          sheetname = Trim(Combo1.text) & "$"
          filepath = Trim(fpText1.text)
          Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
          Set rsexcel = dbexcel.OpenRecordset(sheetname)
          rsexcel.MoveFirst
          Est = False
          '-----> Borrar tabla paso_importacion ruta producto-----
          vg_dbpedweb.Execute "DELETE paso_importacionrutaproducto WHERE irp_spid = @@spid and irp_usr = '" & vg_NUsr & "'"
          Set RS = vg_dbpedweb.Execute("SELECT @@spid spid")
          If Not RS.EOF Then spid = RS!spid
          RS.Close: Set RS = Nothing
          
          Do While rsexcel.EOF <> True
             DoEvents
             If rsexcel.Fields(0).Value = "*" Then Exit Do
             '-------> Mover código ruta
             codrut = 0
             codrut = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, 0)
             '-------> Mover código producto
             codpro = ""
             codpro = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             '-------> Mover código central de compras
             codcco = ""
             codcco = IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
             vg_dbpedweb.Execute ("INSERT INTO paso_importacionrutaproducto VALUES ('" & vg_NUsr & "', " & spid & ", " & codrut & ", '" & codpro & "', '" & codcco & "')")
             Est = True
'             Set RS = vg_dbpedweb.Execute("pedweb_s_rutaproductos 4, " & codrut & ", '" & codpro & "', '" & codcco & "'")
'             If RS.EOF Then
'                 vg_dbpedweb.Execute ("INSERT INTO s_Recorrido_Productos VALUES (" & codrut & ", '" & codpro & "', '" & codcco & "')")
'             End If
'             RS.Close: Set RS = Nothing
             Label2(1).Caption = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             rsexcel.MoveNext: i = i + 1
          Loop
          rsexcel.Close: Set rsexcel = Nothing
          If Est Then vg_dbpedweb.Execute ("pedweb_p_importarrutaproducto '" & vg_NUsr & "', " & spid & "")
       ElseIf Option1(1).Value = True Then '-------> Calendario
          fg_carga ""
          Toolbar1.Enabled = False
          Frame1.Enabled = False
          Label2(1).Caption = ""
          Label2(1).Visible = True
          i = 0
          sheetname = Trim(Combo1.text) & "$"
          filepath = Trim(fpText1.text)
          Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
          Set rsexcel = dbexcel.OpenRecordset(sheetname)
          rsexcel.MoveFirst
          Est = False
          '-----> Borrar tabla paso_importacion ruta calendario-----
          vg_dbpedweb.Execute "DELETE paso_importacionrutacalendario WHERE irc_spid = @@spid and irc_usr = '" & vg_NUsr & "'"
          Set RS = vg_dbpedweb.Execute("SELECT @@spid spid")
          If Not RS.EOF Then spid = RS!spid
          RS.Close: Set RS = Nothing
          
          Do While rsexcel.EOF <> True
             DoEvents
             If rsexcel.Fields(0).Value = "*" Then Exit Do
             '-------> Mover código ruta
             codrut = 0
             codrut = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, 0)
             '-------> Mover fecha
             fecha = ""
             fecha = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             '-------> Mover fecha tope
             fectop = ""
             fectop = IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
             '-------> Mover día
             dia = 0
             dia = IIf(Not IsNumeric(rsexcel.Fields(3).Value), 0, rsexcel.Fields(3).Value)
             '-------> Mover hora
             hora = ""
             hora = IIf(Not IsNull(rsexcel.Fields(4).Value), rsexcel.Fields(4).Value, "")
             vaSpread2.SetActiveCell 1, i
             vg_dbpedweb.Execute ("INSERT INTO paso_importacionrutacalendario VALUES ('" & vg_NUsr & "', " & spid & ", " & codrut & ", '" & Format(fecha, "yyyymmdd") & "', '" & Format(fectop, "yyyymmdd") & "', " & dia & ", '" & Format(hora, "hh:mm") & "')")
             Est = True
'             Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendario 3, " & codrut & ", '', '', '" & Format(Fecha, "yyyymmdd") & "'")
'             If RS.EOF Then
'                 vg_dbpedweb.Execute ("pedweb_iu_rutacalendario 'A', " & codrut & ", '" & Format(Fecha, "yyyymmdd") & "', '" & Format(fectop, "yyyymmdd") & "', " & dia & ", '" & Format(hora, "hh:mm") & "'")
'             Else
'                 vg_dbpedweb.Execute ("pedweb_iu_rutacalendario 'M', " & codrut & ", '" & Format(Fecha, "yyyymmdd") & "', '" & Format(fectop, "yyyymmdd") & "', " & dia & ", '" & Format(hora, "hh:mm") & "'")
'             End If
'             RS.Close: Set RS = Nothing
             Label2(1).Caption = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             rsexcel.MoveNext: i = i + 1
          Loop
          rsexcel.Close: Set rsexcel = Nothing
          If Est Then vg_dbpedweb.Execute ("pedweb_p_importarrutacalendario '" & vg_NUsr & "', " & spid & "")
       ElseIf Option1(2).Value = True Then '-------> Casino
          fg_carga ""
          Toolbar1.Enabled = False
          Frame1.Enabled = False
          Label2(1).Caption = ""
          Label2(1).Visible = True
          i = 0
          sheetname = Trim(Combo1.text) & "$"
          filepath = Trim(fpText1.text)
          Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
          Set rsexcel = dbexcel.OpenRecordset(sheetname)
          rsexcel.MoveFirst
          Est = False
          '-----> Borrar tabla paso_importacion ruta casino-----
          vg_dbpedweb.Execute "DELETE paso_importacionrutacasino WHERE irc_spid = @@spid and irc_usr = '" & vg_NUsr & "'"
          Set RS = vg_dbpedweb.Execute("SELECT @@spid spid")
          If Not RS.EOF Then spid = RS!spid
          RS.Close: Set RS = Nothing
          
          Do While rsexcel.EOF <> True
             DoEvents
             If rsexcel.Fields(0).Value = "*" Then Exit Do
             '-------> Mover código ruta
             codrut = 0
             codrut = IIf(IsNumeric(rsexcel.Fields(0).Value) And Not IsNull(rsexcel.Fields(0).Value), Val(rsexcel.Fields(0).Value), 0)
             '-------> Mover fecha
             fecha = ""
             fecha = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             '-------> Mover código casino
             codcco = ""
             codcco = IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
             '-------> Mover fecha Web
             FechaWeb = ""
             FechaWeb = IIf(Not IsNull(rsexcel.Fields(3).Value), rsexcel.Fields(3).Value, "")
             
             vaSpread2.SetActiveCell 1, i
             vg_dbpedweb.Execute ("INSERT INTO paso_importacionrutacasino VALUES ('" & vg_NUsr & "', " & spid & ", " & codrut & ", convert(datetime, '" & Format(fecha, "dd/MM/yyyy") & "',103), '" & codcco & "', convert(datetime,'" & Format(FechaWeb, "dd/MM/yyyy") & "',103))")
             Est = True
'             Set RS = vg_dbpedweb.Execute("pedweb_s_rutacalendariocasino 2, " & codrut & ", '" & Format(Fecha, "yyyymmdd") & "', '" & codcco & "'")
'             If RS.EOF Then
'                vg_dbpedweb.Execute ("pedweb_iu_rutacalendariocasino 'A', " & codrut & ", '" & Format(Fecha, "yyyymmdd") & "', '" & codcco & "'")
'             End If
'             RS.Close: Set RS = Nothing
             Label2(1).Caption = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, "")
             rsexcel.MoveNext: i = i + 1
          Loop
          rsexcel.Close: Set rsexcel = Nothing
          If Est Then vg_dbpedweb.Execute ("pedweb_p_importarrutacasino '" & vg_NUsr & "', " & spid & "")
       End If
    ElseIf opcion = "lispre" Then
          fg_carga ""
          Toolbar1.Enabled = False
          Frame1.Enabled = False
          Label2(1).Caption = ""
          Label2(1).Visible = True
          sheetname = Trim(Combo1.text) & "$"
          filepath = Trim(fpText1.text)
          Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
          Set rsexcel = dbexcel.OpenRecordset(sheetname)
          i = 0
          rsexcel.MoveFirst
          Est = False
          '-----> Borrar tabla paso_importacion lista precio-----
          vg_dbpedweb.Execute "DELETE paso_importacionlistaprecio WHERE ilp_spid = @@spid and ilp_usr = '" & vg_NUsr & "'"
          Set RS = vg_dbpedweb.Execute("SELECT @@spid spid")
          If Not RS.EOF Then spid = RS!spid
          RS.Close: Set RS = Nothing
          
          Do While rsexcel.EOF <> True
             DoEvents
             If rsexcel.Fields(0).Value = "*" Then Exit Do
             codlis = 0
             codlis = IIf(Not IsNumeric(rsexcel.Fields(0).Value), 0, rsexcel.Fields(0).Value)
             codpro = ""
             codpro = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             fecha = ""
             fecha = IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
             precio = 0
             precio = IIf(Not IsNumeric(rsexcel.Fields(3).Value), 0, rsexcel.Fields(3).Value)
             vaSpread2.SetActiveCell 1, i
             vg_dbpedweb.Execute ("INSERT INTO paso_importacionlistaprecio VALUES ('" & vg_NUsr & "', " & spid & ", " & codlis & ", '" & codpro & "', '" & Format(fecha, "yyyymmdd") & "', '" & precio & "')")
             Est = True
'             Set RS = vg_dbpedweb.Execute("pedweb_s_precios 1, " & codlis & ", '" & codpro & "', '" & Format(Fecha, "yyyymmdd") & "'")
'             If RS.EOF Then
'                vg_dbpedweb.Execute ("pedweb_iu_precios 'A', " & codlis & ", '" & codpro & "', '" & Format(Fecha, "yyyymmdd") & "',  '" & precio & "'")
'             Else
'                vg_dbpedweb.Execute ("pedweb_iu_precios 'M', " & codlis & ", '" & codpro & "', '" & Format(Fecha, "yyyymmdd") & "', '" & precio & "'")
'             End If
'             RS.Close: Set RS = Nothing
             Label2(1).Caption = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, "")
             rsexcel.MoveNext: i = i + 1
          Loop
          rsexcel.Close: Set rsexcel = Nothing
          If Est Then vg_dbpedweb.Execute ("pedweb_p_importarlistaprecio '" & vg_NUsr & "', " & spid & "")
    ElseIf opcion = "regneg" Then
          fg_carga ""
          Toolbar1.Enabled = False
          Frame1.Enabled = False
          Label2(1).Caption = ""
          Label2(1).Visible = True
          sheetname = Trim(Combo1.text) & "$"
          filepath = Trim(fpText1.text)
          Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
          Set rsexcel = dbexcel.OpenRecordset(sheetname)
          i = 0
          rsexcel.MoveFirst
          Est = False
          '-----> Borrar tabla paso_importacion regla negocio-----
          vg_dbpedweb.Execute "DELETE paso_importacionreglanegocioproducto WHERE irp_spid = @@spid and irp_usr = '" & vg_NUsr & "'"
          Set RS = vg_dbpedweb.Execute("SELECT @@spid spid")
          If Not RS.EOF Then spid = RS!spid
          RS.Close: Set RS = Nothing
          
          Do While rsexcel.EOF <> True
             DoEvents
             If rsexcel.Fields(0).Value = "*" Then Exit Do
             codpne = 0
             codpne = IIf(Not IsNumeric(rsexcel.Fields(0).Value), 0, rsexcel.Fields(0).Value)
             codpro = ""
             codpro = IIf(Not IsNull(rsexcel.Fields(1).Value), rsexcel.Fields(1).Value, "")
             pn = ""
             pn = IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, "")
             pa = ""
             pa = IIf(Not IsNumeric(rsexcel.Fields(3).Value), rsexcel.Fields(3).Value, "")
             a = ""
             a = IIf(Not IsNull(rsexcel.Fields(4).Value), rsexcel.Fields(4).Value, "")
             vaSpread2.SetActiveCell 1, i
             vg_dbpedweb.Execute ("INSERT INTO paso_importacionreglanegocioproducto VALUES ('" & vg_NUsr & "', " & spid & ", '" & codpro & "', " & codpne & ", '" & pn & "',  '" & pa & "', '" & a & "')")
             Est = True
'             Set RS = vg_dbpedweb.Execute("pedweb_s_reglasdenegociosproducto 3, " & codpne & ", '" & codpro & "', '', 0")
'             If RS.EOF Then
'                vg_dbpedweb.Execute ("pedweb_iu_reglasdenegociosproductos 'A',  '" & codpro & "', " & codpne & ", '" & pn & "',  '" & pa & "', '" & a & "'")
'             Else
'                vg_dbpedweb.Execute ("pedweb_iu_reglasdenegociosproductos 'M',  '" & codpro & "', " & codpne & ", '" & pn & "', '" & pa & "', '" & a & "'")
'             End If
'             RS.Close: Set RS = Nothing
             Label2(1).Caption = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, "")
             rsexcel.MoveNext: i = i + 1
          Loop
          rsexcel.Close: Set rsexcel = Nothing
          If Est Then vg_dbpedweb.Execute ("pedweb_p_importarreglanegocioproducto '" & vg_NUsr & "', " & spid & "")
    ElseIf opcion = "acprod" Then
          fg_carga ""
          Toolbar1.Enabled = False
          Frame1.Enabled = False
          Label2(1).Caption = ""
          Label2(1).Visible = True
          sheetname = Trim(Combo1.text) & "$"
          filepath = Trim(fpText1.text)
          Set dbexcel = OpenDatabase(filepath, False, False, "Excel 8.0;HDR=no;")
          Set rsexcel = dbexcel.OpenRecordset(sheetname)
          i = 0
          rsexcel.MoveFirst
          vaSpread2.TextTip = TextTipFloating
          Est = False
          '-----> Borrar tabla paso_importacion agregarcantidadproducto-----
          vg_dbpedweb.Execute "DELETE paso_importacioncantidadproducto WHERE icp_spid = @@spid and icp_usr = '" & vg_NUsr & "'"
          Set RS = vg_dbpedweb.Execute("SELECT @@spid spid")
          If Not RS.EOF Then spid = RS!spid
          RS.Close: Set RS = Nothing
          
          Do While rsexcel.EOF <> True
             DoEvents
             If rsexcel.Fields(0).Value = "*" Then Exit Do
             '-------> mover codigo producto
             codpro = ""
             codpro = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, "")
             '-------> Mover código central de compras
             codcco = ""
             vaSpread2.Col = 2
             codcco = IIf(Not IsNull(rsexcel.Fields(1).Value), LimpiaDato(Trim(rsexcel.Fields(1).Value)), "")
             '-------> Mover cantidad
             cantidad = 0
             cantidad = IIf(Not IsNull(rsexcel.Fields(2).Value), rsexcel.Fields(2).Value, 0)
             '-------> mover fecha
             fecha = ""
             fecha = IIf(Not IsNull(rsexcel.Fields(3).Value), rsexcel.Fields(3).Value, "")
             vg_dbpedweb.Execute ("INSERT INTO paso_importacioncantidadproducto VALUES ('" & vg_NUsr & "', " & spid & ", '" & codpro & "', '" & codcco & "', " & cantidad & ",  '" & Format(fecha, "yyyymmdd") & "')")
             Est = True
'             Set RS = vg_dbpedweb.Execute("pedweb_s_listaproductossincantidad 2, '" & codpro & "', '" & codcco & "'")
'             If RS.EOF Then
'                vg_dbpedweb.Execute ("pedweb_iu_productoscantidad 'A',  '" & codpro & "', '" & codcco & "', " & cantidad & ",  '" & Format(Fecha, "yyyymmdd") & "'")
'             Else
'                vg_dbpedweb.Execute ("pedweb_iu_productoscantidad 'M',  '" & codpro & "', '" & codcco & "', " & cantidad & ", '" & Format(Fecha, "yyyymmdd") & "'")
'             End If
'             RS.Close: Set RS = Nothing
             vaSpread2.SetActiveCell 1, i
             Label2(1).Caption = IIf(Not IsNull(rsexcel.Fields(0).Value), rsexcel.Fields(0).Value, "")
             rsexcel.MoveNext: i = i + 1
          Loop
          rsexcel.Close: Set rsexcel = Nothing
          If Est Then vg_dbpedweb.Execute ("pedweb_p_importarcantidadproducto '" & vg_NUsr & "', " & spid & "")
    End If
    Toolbar1.Enabled = True
    fg_descarga
    Label2(1).Visible = False
    Frame1.Enabled = True
    If Est Then MsgBox "Generación importación finalizado sin problema " & VgLinea & Space(10) & i & " Registro fueron importado", vbInformation + vbOKOnly, Msgtitulo
    vg_codigo = "X"
    fg_descarga
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)
vaSpread2.Row = Row
vaSpread2.Col = Col
vaSpread2.Lock = True
End Sub

Private Sub vaSpread2_ScriptTextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Variant, TipWidth As Variant, TipText As Variant, ShowTip As Variant)
a = a
End Sub
