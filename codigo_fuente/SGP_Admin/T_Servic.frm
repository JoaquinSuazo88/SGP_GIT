VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_Servic 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicio"
   ClientHeight    =   6570
   ClientLeft      =   3675
   ClientTop       =   2595
   ClientWidth     =   18285
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   18285
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   18285
      _ExtentX        =   32253
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6045
      Left            =   120
      TabIndex        =   1
      Top             =   420
      Width           =   18090
      _ExtentX        =   31909
      _ExtentY        =   10663
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Servicio"
      TabPicture(0)   =   "T_Servic.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "vaSpread1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Estructura de Servicio..."
      TabPicture(1)   =   "T_Servic.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lblNOMBRE(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblNOMBRE(2)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "vaSpread2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command3"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Command4"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Comensales Estimados"
      TabPicture(2)   =   "T_Servic.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vaSpread3"
      Tab(2).Control(1)=   "lblNOMBRE(1)"
      Tab(2).Control(2)=   "Label1(2)"
      Tab(2).Control(3)=   "Label1(3)"
      Tab(2).Control(4)=   "Label1(5)"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton Command4 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   10940
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   8240
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   14595
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1980
         Visible         =   0   'False
         Width           =   315
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4110
         Left            =   90
         TabIndex        =   20
         Top             =   1065
         Width           =   18000
         _Version        =   393216
         _ExtentX        =   31750
         _ExtentY        =   7250
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         FormulaSync     =   0   'False
         MaxCols         =   13
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_Servic.frx":0054
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Left            =   330
         TabIndex        =   18
         Top             =   5400
         Width           =   915
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   19
            Top             =   135
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   1290
         TabIndex        =   16
         Top             =   5400
         Width           =   4845
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   17
            Top             =   135
            Width           =   4740
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -72150
         TabIndex        =   2
         Top             =   510
         Width           =   6750
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "T_Servic.frx":1DB6
            Left            =   2175
            List            =   "T_Servic.frx":1DC0
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Left            =   2175
            TabIndex        =   4
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
            Left            =   4755
            TabIndex        =   7
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
            Left            =   660
            TabIndex        =   6
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
            Left            =   660
            TabIndex        =   5
            Top             =   300
            Width           =   1380
         End
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   1245
         Left            =   -73560
         TabIndex        =   8
         Top             =   960
         Width           =   7005
         _Version        =   393216
         _ExtentX        =   12356
         _ExtentY        =   2196
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         MaxRows         =   3
         SpreadDesigner  =   "T_Servic.frx":1DD4
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3615
         Left            =   -74790
         TabIndex        =   9
         Top             =   1620
         Width           =   15020
         _Version        =   393216
         _ExtentX        =   26494
         _ExtentY        =   6376
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         ButtonDrawMode  =   1
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
         FormulaSync     =   0   'False
         MaxCols         =   12
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_Servic.frx":23B5
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   1
         Left            =   -73590
         TabIndex        =   14
         Top             =   600
         Width           =   5280
      End
      Begin VB.Label lblNOMBRE 
         Alignment       =   2  'Center
         Caption         =   "fff"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   240
         Index           =   0
         Left            =   780
         TabIndex        =   13
         Top             =   600
         Width           =   5280
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clientes"
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
         Left            =   -74430
         TabIndex        =   12
         Top             =   1230
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Personal"
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
         Left            =   -74430
         TabIndex        =   11
         Top             =   1470
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Totales"
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
         Index           =   5
         Left            =   -74430
         TabIndex        =   10
         Top             =   1740
         Width           =   645
      End
   End
End
Attribute VB_Name = "T_Servic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long
Dim vTipoSer() As Variant 'en  general
Dim vSerAso() As Variant
Dim EstVect As Boolean
Dim IRow As Long
Dim itop As Long
Public CallForm As String

Private Sub GrabaRegistro(Fila)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim coddet           As Long
Dim codsap           As String
Dim indfac           As String
Dim nomdet           As String
Dim orddet           As Long
Dim codenc           As Long
Dim codser           As Long
Dim nomenc           As String
Dim ordenc           As Long
Dim i                As Long
Dim j                As Long
Dim nrorac           As Long
Dim nummin           As Long
Dim Indicador        As String
Dim LyD              As String
Dim grupoasignacion  As Long
Dim servicioasociado As Long
Dim NombreFantasia   As String
Dim HomEstSer        As Long

Dim varSerActivo     As Long
Dim varEstSerActivo  As Long


OpGr = True

If Command1.Visible = True Then Command1.Visible = False
If Command2.Visible = True Then Command2.Visible = False
If Command3.Visible = True Then Command3.Visible = False
If Command4.Visible = True Then Command4.Visible = False

servicioasociado = 0

vaSpread1.Row = Fila
vaSpread1.Col = 1
codenc = Val(vaSpread1.Value)

vaSpread1.Col = 2
nomenc = Trim(LimpiaDato(vaSpread1.Value))

vaSpread1.Col = 3
ordenc = Val(vaSpread1.Value)

vaSpread1.Col = 4
codsap = vaSpread1.text

vaSpread1.Col = 5
indfac = vaSpread1.text

vaSpread1.Col = 8
LyD = IIf(vaSpread1.text = "Verdadero", "1", "0")

vaSpread1.Col = 10
servicioasociado = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)

vaSpread1.Col = 11
NombreFantasia = Mid(vaSpread1.text, 1, 30)

vaSpread1.Col = 12
varSerActivo = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)


If CallForm = "M_Plami2" Then
    
    codenc = lblNOMBRE(2).Caption
    nomenc = lblNOMBRE(0).Caption

End If

If servicioasociado <= 0 Then MsgBox "Falta información en columna servicio asociado...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 9, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub

If vg_Indppr = "1" Or vg_Indppr = "2" Then
   
   vaSpread1.Col = 6
   If Trim(nomenc) = "" Or Trim(codenc) = "" Or Trim(vaSpread1.text) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
   Indicador = IIf(vg_Indppr = "1", "1", "2")

Else
   
   vaSpread1.Col = 6
   If Trim(nomenc) = "" Or Trim(codenc) = "" Or Trim(vaSpread1.text) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
   Indicador = IIf(vaSpread1.TypeComboBoxCurSel = 0, 1, 2)

End If

If Trim(nomenc) = "" Or ordenc = 0 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub

If Trim(NombreFantasia) = "" Then MsgBox "Nombre Fantasía debe ser ingresada...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 11: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub


If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If modo = "A" And SSTab1.Tab = 0 Then
    
    '-------> ENCABEZADO
   MoverDatosGrillas2
   codenc = 0
   Set RS1 = vg_db.Execute("sgpadm_InsUpd_servicio_V02 'A', 0, '" & Trim(Mid(nomenc, 1, 30)) & "', " & ordenc & ", '" & codsap & "', '" & indfac & "', '" & Indicador & "', '" & LyD & "', " & servicioasociado & ", '" & NombreFantasia & "', " & varSerActivo & "")
   If Not RS1.EOF Then
      
      codenc = RS1!indice
      vaSpread1.Col = 1
      vaSpread1.Value = codenc
   
   End If
   RS1.Close
   Set RS1 = Nothing

ElseIf modo = "M" And SSTab1.Tab = 0 Then

    If varSerActivo = 0 Then
        'Valida si es posible inactivar el servicio
        Set RS = vg_db.Execute("EXEC SGPADM_S_ValidaBorradoServicioEstrucutra " & codenc & ", 0")
        
        If Not RS.EOF Then
            
            If RS!Retorno = 1 Then
                
                MsgBox "La estructura de servicio fue utilizado dentro de los últimos " & GetParametro("inacserini") & " meses o " & GetParametro("inacserfin") & " meses posteriores a la fecha actual...", vbExclamation + vbOKOnly, MsgTitulo
                vaSpread1.Row = Fila
                vaSpread1.Col = 12
                vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow
                vaSpread1.SetFocus
                
                OpGr = False:
                
                Exit Sub
            
            End If
        
        End If
        
        If RS.State = 1 Then RS.Close
    
    End If
    
    Set RS1 = vg_db.Execute("SELECT DISTINCT ser_indppr FROM a_servicio with (nolock) WHERE ser_codigo = " & codenc & "")
    If Not RS1.EOF Then
       
       Ind = RS1!ser_indppr
       RS1.Close
       Set RS1 = Nothing
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS1 = vg_db.Execute("SELECT DISTINCT min_codser, min_indppr FROM b_minuta with (nolock) WHERE min_codser = " & codenc & " AND min_indppr = '" & Ind & "'")
       If Not RS1.EOF Then
          
          If RS1!min_indppr <> Indicador Then
             
             RS1.Close: Set RS1 = Nothing
             vaSpread1.Col = 7
             codaux = -1
             
             For z = 0 To vaSpread1.TypeComboBoxCount
                 
                 vaSpread1.TypeComboBoxCurSel = z
                 If vaSpread1.text = Ind Then codaux = z: Exit For
                 codaux = -1
             
             Next z
             
             vaSpread1.Col = 6
             vaSpread1.TypeComboBoxCurSel = codaux
             Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
             Combo1.Enabled = True
             fpText1.Enabled = True
             modo = ""
             Gl_Ac_Botones Me, 1, 1, modo
             MsgBox "No se puede actualizar servicio, ya que existe minuta asociada...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 1: vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False
             Exit Sub
          
          End If
       
       End If
    
    End If
    RS1.Close
    Set RS1 = Nothing
    
    '-------> ENCABEZADO
    vg_db.Execute "sgpadm_InsUpd_servicio_V02 'M', " & codenc & ", '" & Trim(Mid(nomenc, 1, 30)) & "', " & ordenc & ", '" & codsap & "', '" & indfac & "', '" & Indicador & "', '" & LyD & "', " & servicioasociado & ", '" & NombreFantasia & "', '" & varSerActivo & "'"

End If

'------> DETALLE

Dim CatDietetica As Long
Dim TipoPlato As Long

If vaSpread2.MaxRows > 0 And SSTab1.Tab = 1 Then
    
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1
    coddet = Val(vaSpread2.Value)
    
    vaSpread2.Col = 2
    nomdet = Trim(LimpiaDato(vaSpread2.Value))
    
    vaSpread2.Col = 3
    orddet = Val(vaSpread2.Value)
    
    vaSpread2.Col = 4
    grupoasignacion = Val(vaSpread2.Value)
    
    vaSpread2.Col = 6
    CatDietetica = Val(vaSpread2.Value)
    
    vaSpread2.Col = 8
    TipoPlato = Val(vaSpread2.Value)
    
    
    vaSpread2.Col = 10
    racmin = Val(vaSpread2.Value)
   
    vaSpread2.Col = 11
    HomEstSer = Val(vaSpread2.Value)
    
    vaSpread2.Col = 13
    varEstSerActivo = Val(vaSpread2.Value)

    If Trim(nomdet) = "" Then
        
        MsgBox "Favor ingresar descripción, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo
        vaSpread2.Col = 2
        vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow
        vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    
    End If
    
    If orddet = 0 Then
        
        MsgBox "Favor ingresar Orden, dado a que el campo se encuentra en blanco ó bien cero.", vbExclamation + vbOKOnly, MsgTitulo
        vaSpread2.Col = 3
        vaSpread2.SetActiveCell 3, vaSpread2.ActiveRow
        vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    
    End If
    
    If grupoasignacion = 0 And ValidarCampo("vgestmser") Then
        
        MsgBox "Favor ingresar Grupo de Estructura, dado a que el campo se encuentra en valor cero.", vbExclamation + vbOKOnly, MsgTitulo
        vaSpread2.Col = 4
        vaSpread2.SetActiveCell 4, vaSpread2.ActiveRow
        vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    
    End If
    
    If CatDietetica = 0 And ValidarCampo("vcdietmser") Then
        
        MsgBox "Favor ingresar Categoria Dietetica, dado a que el campo se encuentra en valor cero.", vbExclamation + vbOKOnly, MsgTitulo
        vaSpread2.Col = 6
        vaSpread2.SetActiveCell 6, vaSpread2.ActiveRow
        vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    
    End If
    
    If TipoPlato = 0 And ValidarCampo("vtpltomser") Then
        
        MsgBox "Favor ingresar Tipo de Palto, dado a que el campo se encuentra en valor cero.", vbExclamation + vbOKOnly, MsgTitulo
        vaSpread2.Col = 8
        vaSpread2.SetActiveCell 8, vaSpread2.ActiveRow
        vaSpread2.SetFocus
        OpGr = False
        Exit Sub
    
    End If
    
    If varEstSerActivo = 0 And modo <> "A" Then
        
        'Valida si es posible inactivar la estructura de servicio
        Set RS = vg_db.Execute("EXEC SGPADM_S_ValidaBorradoServicioEstrucutra " & codenc & ", " & coddet)
            If Not RS.EOF Then
                If RS!Retorno = 1 Then
                    MsgBox "La estructura de servicio fue utilizado dentro de los últimos " & GetParametro("inacserini") & " meses o " & GetParametro("inacserfin") & " meses posteriores a la fecha actual...", vbExclamation + vbOKOnly, MsgTitulo
                    vaSpread2.Row = Fila
                    vaSpread2.Col = 7
                    vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow
                    vaSpread2.SetFocus
                    OpGr = False:
                    Exit Sub
                End If
            End If
        
        If RS.State = 1 Then RS.Close
    
    End If
    
    Text1(1).Enabled = True
    Text1(2).Enabled = True
    
    If modo = "A" Then
       
       coddet = 0
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
'       Set RS1 = vg_db.Execute("sgpadm_InsUpd_estservicio_V04 'A', " & codenc & ", 0, '" & Trim(Mid(nomdet, 1, 30)) & "', " & orddet & ", " & racmin & ", " & grupoasignacion & ", " & HomEstSer & ", " & varEstSerActivo & ", " & CatDietetica & ", " & TipoPlato & ", '" & vg_NUsr & "'")
       Set RS1 = vg_db.Execute("sgpadm_InsUpd_estservicio_V04 'A', " & codenc & ", 0, '" & Trim(Mid(nomdet, 1, 30)) & "', " & orddet & ", " & racmin & ", " & grupoasignacion & ", " & HomEstSer & ", 1, " & CatDietetica & ", " & TipoPlato & ", '" & vg_NUsr & "'")
       If Not RS1.EOF Then
          
          coddet = RS1!indice
          vaSpread2.Col = 1
          vaSpread2.Value = coddet
       
          vaSpread2.Col = 13
          vaSpread2.Value = "1"
          
       End If
       If CallForm = "M_Plami2" Then
          
          M_Plami2.CargarListaMenu
       
       End If
       RS1.Close
       Set RS1 = Nothing

    Else
       
       vg_db.Execute "sgpadm_InsUpd_estservicio_V04 'M', " & codenc & ", " & coddet & ", '" & Trim(Mid(nomdet, 1, 30)) & "', " & orddet & ", " & racmin & ", " & grupoasignacion & " , " & HomEstSer & ", " & varEstSerActivo & ", " & CatDietetica & ", " & TipoPlato & ", '" & vg_NUsr & "'"

    End If
    
    vaSpread2.Col = 1
    vaSpread2.Value = coddet

End If

'-------> Raciones estimadas
If SSTab1.Tab = 2 Then
   
   For i = 1 To (vaSpread3.MaxRows - 1)
       
       vaSpread3.Row = i
       
       For j = 1 To vaSpread3.maxcols
           
           vaSpread3.Col = j: nrorac = Val(vaSpread3.Value)
              
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
              
           Set RS1 = vg_db.Execute("SELECT * FROM a_serviciorac with (nolock) WHERE sra_codser=" & codenc & " AND sra_coditem=" & i & " AND sra_serdia=" & j & "")
           If Not RS1.EOF Then
                 
              vg_db.Execute "UPDATE a_serviciorac SET sra_raciones=" & nrorac & " WHERE sra_codser=" & codenc & " AND sra_coditem=" & i & " AND sra_serdia=" & j & ""
            
           Else
                 
              vg_db.Execute "INSERT INTO a_serviciorac (sra_codser, sra_coditem, sra_serdia, sra_raciones) VALUES (" & codenc & ", " & i & ", " & j & ", " & nrorac & ")"
             
           End If
           RS1.Close
           Set RS1 = Nothing
             
       Next j
       
   Next i
   
   modo = "M"
   Gl_Ac_Botones Me, 1, 7, modo

Else
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
   Combo1.Enabled = True
   fpText1.Enabled = True
   modo = ""
   Gl_Ac_Botones Me, 1, 1, modo

End If
If CallForm = "M_Plami2" Then
    
   SSTab1.TabEnabled(0) = False
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = False

Else
    
   SSTab1.TabEnabled(0) = True
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True

End If
OpGr = False

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub Command1_Click()

vg_left = Command1.Left - 5801
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "a_homologacionestservicio", "sec_", "homologacion est. servicio", "homestser"
B_TabEst.Show 1
Me.Refresh

With vaSpread2
    If vg_codigo = "" Then
    
      .Col = 11
      .Row = IRow
      .SetActiveCell 11, IRow
      .EditMode = True
      .EditModeReplace = True
      .SetFocus
      Exit Sub
    
    End If
    
    .Row = IRow
    .Col = 11
    .Value = vg_codigo
    .Col = 12
    .Value = vg_nombre

End With
If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(2) = False

End Sub

Private Sub Command2_Click()

vg_left = Command2.Left + 3801
vg_nombre = ""
vg_codigo = ""
B_TabEst.LlenaDatos "a_grupoestructura", "sec_", "Grupo Estructura", "GrpEstru"
B_TabEst.Show 1
Me.Refresh

With vaSpread2
    If vg_codigo = "" Then
    
      .Col = 4
      .Row = IRow
      .SetActiveCell 4, IRow
      .EditMode = True
      .EditModeReplace = True
      .SetFocus
      Exit Sub
    
    End If
    
    .Row = IRow
    .Col = 4
    .Value = vg_codigo
    .Col = 5
    .Value = vg_nombre

End With
If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(2) = False

End Sub

Private Sub Command3_Click()

vg_left = Command3.Left + 3801
vg_nombre = ""
vg_codigo = ""
B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
B_ArbEst.Show 1
Me.Refresh

With vaSpread2
    
    If vg_codigo = "" Then
    
      .Col = 6
      .Row = IRow
      .SetActiveCell 6, IRow
      .EditMode = True
      .EditModeReplace = True
      .SetFocus
      Exit Sub
    
    End If
    
    .Row = IRow
    .Col = 6
    .Value = vg_codigo
    .Col = 7
    .Value = vg_nombre

End With
If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(2) = False

End Sub

Private Sub Command4_Click()

vg_left = Command4.Left + 3801
vg_nombre = "": vg_codigo = ""
B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
B_ArbEst.Show 1
Me.Refresh

With vaSpread2
    If vg_codigo = "" Then
    
      .Col = 8
      .Row = IRow
      .SetActiveCell 8, IRow
      .EditMode = True
      .EditModeReplace = True
      .SetFocus
      Exit Sub
    
    End If
    
    .Row = IRow
    .Col = 8
    .Value = vg_codigo
    .Col = 9
    .Value = vg_nombre

End With

If modo <> "A" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(2) = False

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

Me.HelpContextID = vg_OpcM
Me.Height = 7005
Me.Width = 18375

MsgTitulo = "Servicio"

fg_centra Me
modo = ""
ibusca = 0
itop = 1

'---> Carga vector Tipo Servicio.
ReDim vTipoSer(2, 2)
If vg_Indppr = "1" Or vg_Indppr = "2" Then
  
  vTipoSer(1, 1) = IIf(vg_Indppr = "1", "1", "2")
  vTipoSer(1, 2) = IIf(vg_Indppr = "1", "Real", "Propuesta")

Else
  
  vTipoSer(1, 1) = 1
  vTipoSer(1, 2) = "Real"
  vTipoSer(2, 1) = 2
  vTipoSer(2, 2) = "Propuesta"

End If
'-------> cargar botoneras
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1

MoverDatosGrillas
OpGr = False
SSTab1.Tab = 0

End Sub

Private Sub Form_Resize()
'Frame1.Move IIf(Me.WindowState = 2, 4200, 435), 360, 6015, 971
'vaSpread1.Move IIf(Me.WindowState = 2, 0, 90), vaSpread1.Top, IIf(Me.WindowState = 2, ScaleWidth, 7005), IIf(Me.WindowState = 2, ScaleHeight - vaSpread1.Top - 400, 3375)
'SSTab1.Move SSTab1.Left, SSTab1.Top, IIf(Me.WindowState = 2, ScaleWidth, 7170), IIf(Me.WindowState = 2, ScaleHeight, 5025)
'Toolbar1.Refresh
'Me.Refresh
End Sub

Private Sub fpText1_Change()

Dim RS2 As New ADODB.Recordset
Dim z As Long

If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub

vaSpread1.Visible = False
If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_Indppr = "1" Or vg_Indppr = "2" Then
    
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
        
       Set RS2 = vg_db.Execute("sgpadm_Sel_Servicio_V01 4, '', " & vg_Indppr & ", '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
        
       Set RS2 = vg_db.Execute("sgpadm_Sel_Servicio_V01 5, '', " & vg_Indppr & ", '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
    End If

Else
    
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
        
        Set RS2 = vg_db.Execute("sgpadm_Sel_Servicio_V01 6, '', 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
        
        Set RS2 = vg_db.Execute("sgpadm_Sel_Servicio_V01 7, '', 0, '%" & UCase(LimpiaDato(fpText1.text)) & "%'")
    
    End If

End If

If RS2.EOF Then

   vaSpread1.MaxRows = 0
   
Else
   
   vaSpread1.MaxRows = RS2!nReg

End If
i = 1

If Not RS2.EOF Then
   
   OpGr = True
   
   Do While Not RS2.EOF
      
      vaSpread1.Row = i
      i = i + 1
      vaSpread1.Col = 1
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.Value = RS2!Ser_codigo
      
      vaSpread1.Col = 2
      vaSpread1.Value = Trim(RS2!ser_nombre)
      
      vaSpread1.Col = 3
      vaSpread1.Value = Trim(RS2!ser_orden)
      
      vaSpread1.Col = 4
      vaSpread1.text = IIf(IsNull(RS2!ser_codsap), "", Trim(RS2!ser_codsap))
      
      vaSpread1.Col = 5
      vaSpread1.text = IIf(IsNull(RS2!ser_facturable), "0", Trim(RS2!ser_facturable))
      lisnom = ""
      liscod = ""
      
      For j = 1 To UBound(vTipoSer)
          
          If vTipoSer(j, 1) <> "" Then
             
             lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoSer(j, 2))
             liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoSer(j, 1)
          
          End If
      
      Next j
      vaSpread1.Col = 6
      vaSpread1.TypeComboBoxList = lisnom
      
      vaSpread1.Col = 7
      vaSpread1.TypeComboBoxList = liscod
      
      vaSpread1.Col = 7
      codaux = -1
      For z = 0 To vaSpread1.TypeComboBoxCount
          
          vaSpread1.TypeComboBoxCurSel = z
          If vaSpread1.text = IIf(IsNull(RS2!ser_indppr), 0, RS2!ser_indppr) Then codaux = z: Exit For
          codaux = -1
      
      Next z
      
      vaSpread1.Col = 6
      vaSpread1.TypeComboBoxCurSel = codaux
      
      vaSpread1.Col = 8
      vaSpread1.text = IIf(IsNull(RS2!Ser_LYD) Or RS2!Ser_LYD = False, "0", Trim(RS2!Ser_LYD))
      
      '-------> Mover servicio asociado
      If EstVect Then
         
         lisnom = ""
         liscod = ""
         
         For z = 1 To UBound(vSerAso)
             
             vaSpread1.Col = 9
             lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vSerAso(z, 2))
             
             vaSpread1.Col = 10
             liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vSerAso(z, 1)
             
             vaSpread1.Col = 9
             vaSpread1.TypeComboBoxList = lisnom
             
             vaSpread1.Col = 10
             vaSpread1.TypeComboBoxList = liscod
         
         Next z

         vaSpread1.Col = 10
         codaux = -1
         For z = 0 To vaSpread1.TypeComboBoxCount
             
             vaSpread1.TypeComboBoxCurSel = z
             If vaSpread1.text = IIf(IsNull(RS2!IdServicio), 0, RS2!IdServicio) Then codaux = z: Exit For
             codaux = -1
         
         Next z
         
         vaSpread1.Col = 9
         vaSpread1.TypeComboBoxCurSel = codaux
      
         vaSpread1.Col = 11
         vaSpread1.text = IIf(IsNull(RS2!ser_NombreFantasia), "", Trim(RS2!ser_NombreFantasia))
         
         vaSpread1.Col = 12
         vaSpread1.text = IIf(IsNull(RS2!ser_activo), "", Trim(RS2!ser_activo))

      
      End If
      
      RS2.MoveNext
   
   Loop
   OpGr = False
   SSTab1.TabEnabled(1) = True
   SSTab1.TabEnabled(2) = True
   Gl_Ac_Botones Me, 1, 1, modo

Else
   
   SSTab1.TabEnabled(1) = False
   SSTab1.TabEnabled(2) = False
   Gl_Ac_Botones Me, 1, 2, modo

End If
RS2.Close
Set RS2 = Nothing

If fpText1.text = "" Then
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"

Else
   
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."

End If

vaSpread1.Visible = True

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

Text1(1).text = ""
Text1(2).text = ""
Dim RS1 As New ADODB.Recordset
Select Case SSTab1.Tab

Case 0, 1
    
    itop = 1

    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("SELECT * FROM a_servicio with (nolock)")
    If Not RS1.EOF Then
       
       Gl_Ac_Botones Me, 1, 1, modo
    
    Else
       
       Gl_Ac_Botones Me, 1, 2, modo
    
    End If
    RS1.Close
    Set RS1 = Nothing
    If SSTab1.Tab = 0 Then Exit Sub
    Me.Refresh
    vaSpread1.Col = 2
    vaSpread1.Row = vaSpread1.ActiveRow
    lblNOMBRE(0).Caption = vaSpread1.Value
    Command1.Visible = False
    Command1.Top = 1980
    Command2.Visible = False
    Command2.Top = 1980
    Command3.Visible = False
    Command3.Top = 1980
    Command4.Visible = False
    Command4.Top = 1980

    MoverDatosGrillas2
    

Case 2
    
    Me.Refresh
    vaSpread1.Col = 2
    vaSpread1.Row = vaSpread1.ActiveRow
    lblNOMBRE(1).Caption = vaSpread1.Value
    MoverDatosGrillas3

End Select

End Sub

Private Sub Text1_Change(Index As Integer)

On Error GoTo Man_Error

If LimpiaDato(Trim(Text1(Index).text)) & Chr(KeyAscii) = "" Then Exit Sub

Dim RS     As New ADODB.Recordset
Dim i      As Long
Dim codigo As Long

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1

If Servicio < 1 Then
    
   codigo = Val(vaSpread1.Value)

Else
    
   codigo = Servicio
   lblNOMBRE(0).Caption = G_Proc.fg_TraerNombre("a_servicio", "ser_codigo", Servicio, "ser_nombre")
   lblNOMBRE(2).Caption = Servicio

End If

vaSpread2.Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

Case 1

   Set RS = vg_db.Execute("sgpadm_Sel_BuscarCodigoEstructuraServicio " & codigo & ", '%" & UCase(LimpiaDato(Text1(Index).text)) & "%'")

Case 2

   Set RS = vg_db.Execute("sgpadm_Sel_BuscarNombreEstructuraServicio " & codigo & ", '%" & UCase(LimpiaDato(Text1(Index).text)) & "%'")

End Select

If RS.EOF Then
   
   vaSpread2.MaxRows = 0

Else

   vaSpread2.MaxRows = RS.RecordCount

End If

i = 1

OpGr = True

Do While Not RS.EOF
    
    vaSpread2.Row = i
    
    vaSpread2.Col = 1
    vaSpread2.Value = RS!ess_codigo
    
    vaSpread2.Col = 2
    vaSpread2.Value = Trim(RS!ess_nombre)
    
    vaSpread2.Col = 3
    vaSpread2.Value = RS!ess_orden
    
    vaSpread2.Col = 4
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = RS!ess_agrupacionestructura
    
    vaSpread2.Col = 5
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = RS!NombreGrupoEstructura
    
    vaSpread2.Col = 6
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = IIf(RS!car_codigo = 0, "", RS!car_codigo)
    
    vaSpread2.Col = 7
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = RS!car_nombre
    
    vaSpread2.Col = 8
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = IIf(RS!tip_codigo = 0, "", RS!tip_codigo)
    
    vaSpread2.Col = 9
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = RS!tip_nombre
    
    vaSpread2.Col = 10
    vaSpread2.Value = RS!ess_racmin
    
    vaSpread2.Col = 11
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = IIf(IsNull(RS!ID_HomologacionEstServicio) Or RS!ID_HomologacionEstServicio = 0, "", RS!ID_HomologacionEstServicio)
    
    vaSpread2.Col = 12
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = IIf(IsNull(RS!descripcion), "", RS!descripcion)
    
    vaSpread2.Col = 13
    vaSpread2.Value = IIf(IsNull(RS!ess_activo), "", RS!ess_activo)
       
    RS.MoveNext
    
    i = i + 1
    
Loop

Gl_Ac_Botones Me, 1, 1, modo
RS.Close
Set RS = Nothing

vaSpread2.SetActiveCell 1, vaSpread1.MaxRows
vaSpread2.SetActiveCell 1, 1
vaSpread2.Visible = True

OpGr = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
   

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim RS1 As New ADODB.Recordset
Dim codigo As Long, Nombre As String, orden As String, codser As Long

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    
    Text1(1).Enabled = False
    Text1(2).Enabled = False
    
    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False
    
    If SSTab1.Tab = 0 Then
        
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        vaSpread2.MaxRows = 0
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        lisnom = ""
        liscod = ""
        
        For j = 1 To UBound(vTipoSer)
          
          If vTipoSer(j, 1) <> "" Then
           
           lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoSer(j, 2))
           liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoSer(j, 1)
          
          End If
        
        Next j
        
        If vg_Indppr = 1 Or vg_Indppr = 2 Then
           
           lisnom = "": liscod = ""
           lisnom = IIf(vg_Indppr = "1", "Real", "Propuesta")
           liscod = IIf(vg_Indppr = "1", "1", "2")
           vaSpread1.Col = 6
           vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
           
           vaSpread1.Col = 7
           vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
        
        Else
           
           vaSpread1.Col = 6
           vaSpread1.TypeComboBoxList = lisnom
           
           vaSpread1.Col = 7
           vaSpread1.TypeComboBoxList = liscod
        
        End If
        
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 6
        vaSpread1.Lock = False
        vaSpread1.TypeComboBoxList = lisnom
        
        vaSpread1.Col = 7
        vaSpread1.TypeComboBoxList = liscod
        
        lisnom = ""
        liscod = ""
        '-------> Mover servicio asociado
        If EstVect Then
           
           For z = 1 To UBound(vSerAso)
               
               vaSpread1.Col = 9
               lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vSerAso(z, 2))
               
               vaSpread1.Col = 10
               liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vSerAso(z, 1)
               
               vaSpread1.Col = 9
               vaSpread1.TypeComboBoxList = lisnom
               
               vaSpread1.Col = 10
               vaSpread1.TypeComboBoxList = liscod
           
           Next z
        
        End If
        
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 2
        vaSpread1.SetActiveCell 2, vaSpread1.MaxRows
        vaSpread1.SetFocus
        'vaSpread1.Col = 6: vaSpread1.TypeComboBoxCurSel = 0
    
    ElseIf SSTab1.Tab = 1 Then
        
        If vaSpread1.MaxRows < 1 Then Exit Sub
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(2) = False
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 2
        vaSpread2.SetActiveCell 2, vaSpread2.MaxRows
        vaSpread2.SetFocus
    
    End If
    Command1.Visible = False
    Command1.Top = 1980
    Command2.Visible = False
    Command2.Top = 1980
    Command3.Visible = False
    Command3.Top = 1980
    Command4.Visible = False
    Command4.Top = 1980

Case 3
    
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    
    If SSTab1.Tab = 0 Then
       
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(2) = False
    
    ElseIf SSTab1.Tab = 1 Then
       
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(2) = False
    
    ElseIf SSTab1.Tab = 2 Then
       
       SSTab1.TabEnabled(0) = False
       SSTab1.TabEnabled(1) = False
    
    End If
'    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False

Case 5 ' borrado
    
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If SSTab1.Tab = 0 Then
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        codigo = Val(vaSpread1.Value)
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        'Valida si es posible inactivar el servicio
        Set RS = vg_db.Execute("EXEC SGPADM_S_ValidaBorradoServicioEstrucutra " & codigo & ", 0")
        
        If Not RS.EOF Then
            
            If RS!Retorno = 1 Then
                
                MsgBox "El servicio fue utilizado dentro de los últimos " & GetParametro("inacserini") & " meses o " & GetParametro("inacserfin") & " meses posteriores a la fecha actual...", vbExclamation + vbOKOnly, MsgTitulo
                vaSpread1.Row = Fila
                vaSpread1.Col = 12
                vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow
                vaSpread1.SetFocus
                
                OpGr = False:
                
                Exit Sub
            
            End If
        
        End If
        
        If RS.State = 1 Then RS.Close

        
        vg_db.Execute "DELETE a_servicio FROM a_servicio WHERE ser_codigo=" & codigo & ""
        vg_db.Execute "DELETE a_serviciorac FROM a_serviciorac WHERE sra_codser=" & codigo & ""
        
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread3.MaxRows = 0
        vaSpread3.MaxRows = 1
    
    ElseIf SSTab1.Tab = 1 Then
        
        vaSpread2.Row = vaSpread2.ActiveRow
        vaSpread1.Row = vaSpread1.ActiveRow
        
        vaSpread2.Col = 1
        codigo = Val(vaSpread2.Value)
        vaSpread1.Col = 1
        codser = Val(vaSpread1.Value)
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        'Valida si es posible inactivar la estructura de servicio
        Set RS = vg_db.Execute("EXEC SGPADM_S_ValidaBorradoServicioEstrucutra " & codser & ", " & codigo)
        If Not RS.EOF Then
           
           If RS!Retorno = 1 Then
              
              MsgBox "La estructura de servicio fue utilizado dentro de los últimos " & GetParametro("inacserini") & " meses o " & GetParametro("inacserfin") & " meses posteriores a la fecha actual...", vbExclamation + vbOKOnly, MsgTitulo
              vaSpread2.Row = Fila
              vaSpread2.Col = 7
              vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow
              vaSpread2.SetFocus
              OpGr = False:
              Exit Sub
                
            End If
        
        End If
        
        If RS.State = 1 Then RS.Close
        
        '------- Validar si existen datos en planificación
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS1 = vg_db.Execute("SELECT TOP 1 mid_estser FROM b_minutadet WHERE mid_estser = " & codigo & "")
        If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "El dato esta asociado planificación, no puede eliminar", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        RS1.Close
        Set RS1 = Nothing
        
        '------- fin validar si existen datos en planificación
        vg_db.Execute "DELETE a_estservicio FROM a_estservicio WHERE ess_codser=" & codser & " AND ess_codigo=" & codigo & ""
        vaSpread2.DeleteRows vaSpread2.Row, 1
        vaSpread2.MaxRows = vaSpread2.MaxRows - 1
    
    End If
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo

Case 7
    
    fpText1.text = ""
    If SSTab1.Tab = 0 Then
       
       MoverDatosGrillas
    
    ElseIf SSTab1.Tab = 1 Then
       
       Text1(1).text = ""
       Text1(2).text = ""
       MoverDatosGrillas2
    
    ElseIf SSTab1.Tab = 2 Then
       
       MoverDatosGrillas3
    
    End If

Case 10
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If Trim(CallForm) = "M_Plami2" Then
        
        SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(1) = True: SSTab1.TabEnabled(2) = False
        modo = "Cancel"
    
    Else
        
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.TabEnabled(2) = True
      
        Text1(1).Enabled = True
        Text1(2).Enabled = True
        Text1(1).text = ""
        Text1(2).text = ""
        
    End If
    
    If modo = "A" Then
            
       If SSTab1.Tab = 0 Then
          
          MoverDatosGrillas
       
       ElseIf SSTab1.Tab = 1 Then
                
          MoverDatosGrillas2
       
       ElseIf SSTab1.Tab = 2 Then
                
          MoverDatosGrillas3
       
       End If
    
    ElseIf modo = "Cancel" Then
            
        MoverDatosGrillas2
        modo = ""
        Cancela
    
    Else
        
        Cancela
    
    End If

Case 12
    
    If Trim(lblNOMBRE(2).Caption) <> "" And Trim(CallForm) = "M_Plami2" Then
        
        vaSpread1.Col = 1
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            
            If vaSpread1.text = lblNOMBRE(2).Caption Then
                
                vaSpread1.SetActiveCell 1, vaSpread1.Row
                Exit For
            
            End If
        
        Next i
        
    End If
    GrabaRegistro vaSpread1.ActiveRow
    
Case 15
    
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If SSTab1.Tab = 0 Then
       
       I_Servic
    
    ElseIf SSTab1.Tab = 1 Then
       
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1
       I_EstructuraServicio Val(vaSpread1.Value)
    
    ElseIf SSTab1.Tab = 2 Then
       
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1
       I_ComensalesEstimados Val(vaSpread1.Value)
    
    End If

Case 18
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

If vaSpread1.MaxRows > 0 And modo <> "A" Then MoverDatosGrillas2

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

If vg_Indppr = 1 Or vg_Indppr = 2 Then
  
  vaSpread1.Col = 6
  vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
  
  vaSpread1.Col = 7
  vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")

End If

SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False

End Sub

Private Sub MoverDatosGrillas()

Dim RS1 As New ADODB.Recordset
Dim i As Long
Dim z As Long
Dim codaux As Long

Dim lisnom As String
Dim liscod As String
Dim cParam As String
Dim encuentra As Boolean

'-------> Mover lista servicio aosciado
EstVect = False
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioAsociado")
i = 1
If Not RS1.EOF Then
   
   ReDim vSerAso(RS1!nReg, 2)
   
   Do While Not RS1.EOF
      
      EstVect = True
      vSerAso(i, 1) = RS1!IdServicio
      vSerAso(i, 2) = RS1!Servicio
      i = i + 1
      RS1.MoveNext
   
   Loop

End If
RS1.Close
Set RS1 = Nothing

vaSpread1.Visible = False
vaSpread1.MaxRows = 0
OpGr = True

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_Indppr = "1" Or vg_Indppr = "2" Then
   
   Set RS1 = vg_db.Execute("sgpadm_Sel_Servicio_V01 5, '', " & vg_Indppr & ", '%%'")

Else
   
   Set RS1 = vg_db.Execute("sgpadm_Sel_Servicio_V01 7, '', 0, '%%'")

End If
   
   Do While Not RS1.EOF
      
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
      If vaSpread1.Row = 1 Then
      
         MoverDatosGrillas2
         OpGr = True
      
      End If
      
      vaSpread1.Col = 1
      vaSpread1.Value = RS1!Ser_codigo
      
      vaSpread1.Col = 2
      vaSpread1.Value = Trim(RS1!ser_nombre)
      
      vaSpread1.Col = 3
      vaSpread1.Value = Trim(RS1!ser_orden)
      
      vaSpread1.Col = 4
      vaSpread1.text = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
      
      vaSpread1.Col = 5
      vaSpread1.text = IIf(IsNull(RS1!ser_facturable), "0", Trim(RS1!ser_facturable))
      
      lisnom = ""
      liscod = ""
      cParam = ""
      encuentra = False
      
      For j = 1 To UBound(vTipoSer)
          
          If vTipoSer(j, 1) <> "" Then
             
             lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vTipoSer(j, 2))
             liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vTipoSer(j, 1)
          
          End If
      
      Next j
      vaSpread1.Col = 6
      vaSpread1.TypeComboBoxList = lisnom
      
      vaSpread1.Col = 7
      vaSpread1.TypeComboBoxList = liscod
    
      vaSpread1.Col = 7
      codaux = -1
      For z = 0 To vaSpread1.TypeComboBoxCount
          
          vaSpread1.TypeComboBoxCurSel = z
          If vaSpread1.text = IIf(RS1!ser_indppr = "1", "1", "2") Then codaux = z: Exit For
          codaux = -1
      
      Next z
      vaSpread1.Col = 6
      vaSpread1.TypeComboBoxCurSel = codaux
      
      vaSpread1.Col = 8
      vaSpread1.text = IIf(IsNull(RS1!Ser_LYD) Or RS1!Ser_LYD = False, "0", Trim(RS1!Ser_LYD))
      
      '-------> Mover servicio asociado
      If EstVect Then
         
         lisnom = ""
         liscod = ""
         For z = 1 To UBound(vSerAso)
             
             vaSpread1.Col = 9
             lisnom = lisnom & IIf(lisnom <> "", Chr$(9), "") & Trim(vSerAso(z, 2))
             
             vaSpread1.Col = 10
             liscod = liscod & IIf(liscod <> "", Chr$(9), "") & vSerAso(z, 1)
             
             vaSpread1.Col = 9
             vaSpread1.TypeComboBoxList = lisnom
             
             vaSpread1.Col = 10
             vaSpread1.TypeComboBoxList = liscod
         
         Next z

         vaSpread1.Col = 10
         codaux = -1
         For z = 0 To vaSpread1.TypeComboBoxCount
             
             vaSpread1.TypeComboBoxCurSel = z
             If vaSpread1.text = IIf(IsNull(RS1!IdServicio), 0, RS1!IdServicio) Then codaux = z: Exit For
             codaux = -1
         
         Next z
         
         vaSpread1.Col = 9
         vaSpread1.TypeComboBoxCurSel = codaux
      
         vaSpread1.Col = 11
         vaSpread1.text = IIf(IsNull(RS1!ser_NombreFantasia), "", Trim(RS1!ser_NombreFantasia))
         
         vaSpread1.Col = 12
         vaSpread1.text = IIf(IsNull(RS1!ser_activo), "", Trim(RS1!ser_activo))

      
      End If
      
      RS1.MoveNext
   
   Loop
   
   RS1.Close
   Set RS1 = Nothing
   
   OpGr = False
   
   vaSpread1.Visible = True
   Gl_Ac_Botones Me, 1, 1, modo
   Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
   
End Sub

Public Sub MoverDatosGrillas2(Optional ByVal Servicio As Long)

Dim RS2 As New ADODB.Recordset
Dim codigo As Long

Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False

OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1

If Servicio < 1 Then
    
   codigo = Val(vaSpread1.Value)

Else
    
   codigo = Servicio
   lblNOMBRE(0).Caption = G_Proc.fg_TraerNombre("a_servicio", "ser_codigo", Servicio, "ser_nombre")
   lblNOMBRE(2).Caption = Servicio

End If

vaSpread2.MaxRows = 0
If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS2 = vg_db.Execute("sgpadm_Sel_EstServicio " & codigo & "")

Do While Not RS2.EOF
    
    vaSpread2.MaxRows = vaSpread2.MaxRows + 1
    vaSpread2.Row = vaSpread2.MaxRows
    
    vaSpread2.Col = 1
    vaSpread2.Value = RS2!ess_codigo
    
    vaSpread2.Col = 2
    vaSpread2.Value = Trim(RS2!ess_nombre)
    
    vaSpread2.Col = 3
    vaSpread2.Value = RS2!ess_orden
    
    vaSpread2.Col = 4
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = RS2!ess_agrupacionestructura
    
    vaSpread2.Col = 5
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = RS2!NombreGrupoEstructura
    
    vaSpread2.Col = 6
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = IIf(RS2!car_codigo = 0, "", RS2!car_codigo)
    
    vaSpread2.Col = 7
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = RS2!car_nombre
    
    vaSpread2.Col = 8
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = IIf(RS2!tip_codigo = 0, "", RS2!tip_codigo)
    
    vaSpread2.Col = 9
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = RS2!tip_nombre
    
    vaSpread2.Col = 10
    vaSpread2.Value = RS2!ess_racmin
    
    vaSpread2.Col = 11
    vaSpread2.CellType = CellTypeEdit
    vaSpread2.TypeHAlign = TypeHAlignLeft
    vaSpread2.Value = IIf(IsNull(RS2!ID_HomologacionEstServicio) Or RS2!ID_HomologacionEstServicio = 0, "", RS2!ID_HomologacionEstServicio)
    
    vaSpread2.Col = 12
    vaSpread2.CellType = CellTypeStaticText
    vaSpread2.Value = IIf(IsNull(RS2!descripcion), "", RS2!descripcion)
    
    vaSpread2.Col = 13
    vaSpread2.Value = IIf(IsNull(RS2!ess_activo), "", RS2!ess_activo)
       
    RS2.MoveNext

Loop
Gl_Ac_Botones Me, 1, 1, modo
RS2.Close
Set RS2 = Nothing

vaSpread2.SetActiveCell 1, vaSpread1.MaxRows
vaSpread2.SetActiveCell 1, 1

OpGr = False

End Sub

Private Sub MoverDatosGrillas3()

Dim RS2 As New ADODB.Recordset
Dim codigo As Long

vaSpread3.Row = -1
vaSpread3.Col = -1:
vaSpread3.BackColor = &H80000018
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = Val(vaSpread1.Value)
vaSpread3.MaxRows = 0
vaSpread3.MaxRows = 2

If RS2.State = 1 Then RS2.Close
RS2.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS2 = vg_db.Execute("SELECT * FROM a_serviciorac WHERE sra_codser = " & codigo & " ORDER BY sra_coditem, sra_serdia")
If Not RS2.EOF Then
   
   Do While Not RS2.EOF
      
      vaSpread3.Row = RS2!sra_coditem
      
      vaSpread3.Col = RS2!sra_serdia
      vaSpread3.text = IIf(RS2!sra_raciones = 0, "", RS2!sra_raciones)
      
      RS2.MoveNext
   
   Loop

End If

vaSpread3.MaxRows = (vaSpread3.MaxRows + 1)
vaSpread3.Row = vaSpread3.MaxRows

vaSpread3.Col = 1
vaSpread3.col2 = vaSpread3.maxcols
vaSpread3.row2 = vaSpread3.MaxRows
vaSpread3.Lock = True
vaSpread3.BlockMode = True
' Lock cells
vaSpread3.Lock = True
' Protect the cells from being edited
vaSpread3.Protect = True
' Turn block mode off
vaSpread3.BlockMode = False
vaSpread3.Col = -1: vaSpread3.BackColor = &HE0E0E0
SumarTotales

RS2.Close
Set RS2 = Nothing
modo = "M"
Gl_Ac_Botones Me, 1, 7, modo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If vg_Indppr = 1 Or vg_Indppr = 2 Then
  
  vaSpread1.Col = 6
  vaSpread1.TypeComboBoxList = ""
  vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
  
  vaSpread1.Col = 7
  vaSpread1.TypeComboBoxList = ""
  vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")

End If

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    
    GrabaRegistro Row

ElseIf Toolbar1.Buttons(12).Visible = False Then
    
    Cancela

End If

End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If (Col <> 5 And Col <> 8 And Col <> 12) Or Row = 0 Or OpGr Then Exit Sub

If modo = "" Then modo = "M"

If vg_Indppr = 1 Or vg_Indppr = 2 Then
  
  vaSpread1.Col = 6
  vaSpread1.TypeComboBoxList = ""
  vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
  
  vaSpread1.Col = 7
  vaSpread1.TypeComboBoxList = ""
  vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")

End If

Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False

End Sub

Private Sub vaSpread1_ComboSelChange(ByVal Col As Long, ByVal Row As Long)
Dim indice As Long

Select Case Col
Case 6

    vaSpread1.Row = Row
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      
      vaSpread1.Col = 6
      vaSpread1.TypeComboBoxList = ""
      vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "Real", "Propuesta")
      
      vaSpread1.Col = 7
      vaSpread1.TypeComboBoxList = ""
      vaSpread1.TypeComboBoxList = IIf(vg_Indppr = "1", "1", "2")
    
    End If
    vaSpread1.Col = 6
    indice = vaSpread1.TypeComboBoxCurSel
    
    vaSpread1.Col = 7
    vaSpread1.TypeComboBoxCurSel = indice
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.EditEnterAction = EditEnterActionNone

Case 9

    vaSpread1.Row = Row
    vaSpread1.Col = 9
    indice = vaSpread1.TypeComboBoxCurSel
    
    vaSpread1.Col = 10
    vaSpread1.TypeComboBoxCurSel = indice
    
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    vaSpread1.EditEnterAction = EditEnterActionNone

End Select

End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If (Col <> 13) Or Row = 0 Or OpGr Then Exit Sub

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False


End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)

Select Case Col

Case Is <> 13
    
    Command1.Visible = False

Case Is <> 4

    Command2.Visible = False

Case Is <> 6

    Command3.Visible = False

Case Is <> 8

    Command4.Visible = False

End Select

Select Case Col

Case 11
    
    Command1.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 11
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 4
    
    Command2.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
    Command2.Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 4
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 6
    
    Command3.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
    Command3.Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 6
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 8
    
    Command4.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
    Command4.Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 8
    vaSpread2.TypeHAlign = TypeHAlignLeft

End Select

End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False

End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

Dim RS2 As New ADODB.Recordset

If SSTab1.Tab = 0 Then Exit Sub

IRow = Row

Select Case Col

    Case 11
   
     Command1.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
     Command1.Visible = True

    Case 4
   
     Command2.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
     Command2.Visible = True
    
    Case 6
   
     Command3.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
     Command3.Visible = True
    
    Case 8
   
     Command4.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
     Command4.Visible = True

End Select

If ChangeMade = False And Col <> 13 Then

   If Col <> 11 Then Command1.Visible = False
   Exit Sub

End If

If ChangeMade = False And Col <> 4 Then

   If Col <> 11 Then Command2.Visible = False
   Exit Sub

End If

If ChangeMade = False And Col <> 6 Then

   If Col <> 11 Then Command3.Visible = False
   Exit Sub

End If

If ChangeMade = False And Col <> 8 Then

   If Col <> 11 Then Command4.Visible = False
   Exit Sub

End If

'If modo = "" Then modo = "M"
'Gl_Ac_Botones Me, 1, 0, modo
'SSTab1.TabEnabled(0) = False
'SSTab1.TabEnabled(2) = False

Select Case Col

'    Case Is <> 11
'
'        Command1.Visible = False
'
'    Case Is <> 4
'
'        Command2.Visible = False
'
'    Case Is <> 6
'
'        Command3.Visible = False
'
'    Case Is <> 8
'
'        Command4.Visible = False
'
    
    Case 11
        
        Command1.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
        Command1.Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS2 = vg_db.Execute("SELECT Descripcion FROM a_homologacionestservicio WHERE Id_HomologacionEstServicio=" & Val(vaSpread2.Value) & "")
        If RS2.EOF Then
           
           RS2.Close
           Set RS2 = Nothing
           vaSpread2.text = ""
           vaSpread2.Col = 12
           vaSpread2.text = ""
           Exit Sub
        
        End If
        
        vaSpread2.Col = 12
        vaSpread2.text = Trim(RS2!descripcion)
        RS2.Close: Set RS2 = Nothing
        Command1.Visible = False
       
    Case 4
        
        Command2.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
        Command2.Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS2 = vg_db.Execute("sgpadm_Sel_CodigoGrupoEstructura " & Val(vaSpread2.Value) & "")
        If RS2.EOF Then
           
           RS2.Close
           Set RS2 = Nothing
           vaSpread2.text = ""
           
           vaSpread2.Col = 5
           vaSpread2.text = ""
           
           Exit Sub
        
        Else
           
           If RS2!Activo = "0" Then
           
              RS2.Close
              Set RS2 = Nothing
   
              MsgBox "Grupo estructura esta desactivado...", vbExclamation + vbOKOnly, MsgTitulo
              
              vaSpread1.text = ""
              
              vaSpread1.Col = 5
              vaSpread1.text = ""
              
              vaSpread1.Row = Row
              vaSpread1.Col = 4
              vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
              vaSpread1.SetFocus
              
              Exit Sub
           
           End If
        
        End If
        
        vaSpread2.Col = 5
        vaSpread2.text = Trim(RS2!Nombre)
        RS2.Close
        Set RS2 = Nothing
        Command2.Visible = False
    
    Case 6
        
        Command3.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
        Command3.Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS2 = vg_db.Execute("sgpadm_Sel_CodigoCategoriaDiateticaTServic " & Val(vaSpread2.Value) & "")
        If RS2.EOF Then
           
           RS2.Close
           Set RS2 = Nothing
           vaSpread2.text = ""
           vaSpread2.Col = 7
           vaSpread2.text = ""
           Exit Sub
        
        Else
        
           If RS2!Activo = "0" Then
           
              RS2.Close
              Set RS2 = Nothing
   
              MsgBox "C. Dietetica esta desactivado...", vbExclamation + vbOKOnly, MsgTitulo
              
              vaSpread1.text = ""
              
              vaSpread1.Col = 7
              vaSpread1.text = ""
              
              vaSpread1.Row = Row
              vaSpread1.Col = 6
              vaSpread1.SetActiveCell 6, vaSpread1.MaxRows
              vaSpread1.SetFocus
              
              Exit Sub
           
           End If
        
        End If
        
        vaSpread2.Col = 7
        vaSpread2.text = Trim(RS2!car_nombre)
        RS2.Close
        Set RS2 = Nothing
        Command3.Visible = False
    
    Case 8
        
        Command4.Top = IIf(Row = 1, 1980, 1980 + (240 * (Row - itop)))
        Command4.Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        If RS2.State = 1 Then RS2.Close
        RS2.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS2 = vg_db.Execute("sgpadm_Sel_CodigoTipoPlatoTServic " & Val(vaSpread2.Value) & "")
        If RS2.EOF Then
           
           RS2.Close
           Set RS2 = Nothing
           vaSpread2.text = ""
           vaSpread2.Col = 8
           vaSpread2.text = ""
           Exit Sub
        
        Else
        
           If RS2!Activo = "0" Then
           
              RS2.Close
              Set RS2 = Nothing
   
              MsgBox "Tipo Plato esta desactivado...", vbExclamation + vbOKOnly, MsgTitulo
              
              vaSpread1.text = ""
              
              vaSpread1.Col = 9
              vaSpread1.text = ""
              
              vaSpread1.Row = Row
              vaSpread1.Col = 8
              
              vaSpread1.SetActiveCell 8, vaSpread1.MaxRows
              vaSpread1.SetFocus
              
              Exit Sub
           
           End If
        
        End If
        
        vaSpread2.Col = 9
        vaSpread2.text = Trim(RS2!tip_nombre)
        RS2.Close
        Set RS2 = Nothing
        Command4.Visible = False
    
    End Select

End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 11 Then Command1.Visible = False
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 4 Then Command2.Visible = False
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 6 Then Command3.Visible = False
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 8 Then Command4.Visible = False
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    
    If Trim(lblNOMBRE(2).Caption) <> "" And Trim(CallForm) = "M_Plami2" Then
        
        vaSpread1.Col = 1
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            
            If vaSpread1.text = lblNOMBRE(2).Caption Then
                
                vaSpread1.SetActiveCell 1, vaSpread1.Row
                Exit For
            
            End If
        
        Next i
        
    End If
    GrabaRegistro vaSpread1.ActiveRow
    
ElseIf Toolbar1.Buttons(12).Visible = False Then
    
    Cancela

End If
End Sub

Private Sub vaSpread2_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)

itop = NewTop
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False

End Sub

Private Sub Cancela()
Dim RS1 As New ADODB.Recordset
Dim z As Long
Dim codaux As Long

If SSTab1.Tab = 0 Then
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = Val(vaSpread1.Value)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute("sgpadm_Sel_Servicio_V01 1, '', " & codigo & ", ''")
    If Not RS1.EOF Then
       
       vaSpread1.Col = 2
       vaSpread1.Value = Trim(RS1!ser_nombre)
       
       vaSpread1.Col = 3
       vaSpread1.Value = Trim(RS1!ser_orden)
       
       vaSpread1.Col = 4
       vaSpread1.text = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
       
       vaSpread1.Col = 5
       vaSpread1.text = IIf(IsNull(RS1!ser_facturable), "0", Trim(RS1!ser_facturable))
       
       vaSpread1.Col = 8
       vaSpread1.text = IIf(IsNull(RS1!Ser_LYD) Or RS1!Ser_LYD = False, "0", Trim(RS1!Ser_LYD))
       
       vaSpread1.Col = 10
       codaux = -1
       For z = 0 To vaSpread1.TypeComboBoxCount
           
           vaSpread1.TypeComboBoxCurSel = z
           If vaSpread1.text = IIf(IsNull(RS1!IdServicio), 0, RS1!IdServicio) Then codaux = z: Exit For
           codaux = -1
       
       Next z
       
       vaSpread1.Col = 9
       vaSpread1.TypeComboBoxCurSel = codaux
   
       vaSpread1.Col = 11
       vaSpread1.text = IIf(IsNull(RS1!ser_NombreFantasia), "", Trim(RS1!ser_NombreFantasia))
       
       vaSpread1.Col = 12
       vaSpread1.text = IIf(IsNull(RS1!ser_activo), "", Trim(RS1!ser_activo))

    
    End If
    RS1.Close
    Set RS1 = Nothing
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True
    fpText1.Enabled = True
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = True

ElseIf SSTab1.Tab = 1 Then
    
    OpGr = True
    vaSpread2.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codser = Val(vaSpread1.Value)
    
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = 1
    codigo = Val(vaSpread2.Value)
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
'    Set RS1 = vg_db.Execute("SELECT a.*, isnull(b.Descripcion,'') Descripcion FROM a_estservicio as a with (nolock) left join a_homologacionestservicio as b with (nolock) on a.ID_HomologacionEstServicio = b.ID_HomologacionEstServicio WHERE ess_codser = " & codser & " AND ess_codigo = " & codigo & "")
    Set RS1 = vg_db.Execute("sgpadm_Sel_CodigoEstServicio " & codser & ", " & codigo & "")
    
    If Not RS1.EOF Then
       
       vaSpread2.Col = 2
       vaSpread2.Value = Trim(RS1!ess_nombre)
       
       vaSpread2.Col = 3
       vaSpread2.Value = RS1!ess_orden
             
       vaSpread2.Col = 4
       'vaSpread2.CellType = SS_CELL_TYPE_NUMBER
       'vaSpread2.TypeHAlign = SS_CELL_H_ALIGN_RIGHT
       vaSpread2.Value = RS1!ess_agrupacionestructura
    
       vaSpread2.Col = 5
       vaSpread2.Value = RS1!NombreGrupoEstructura
    
       vaSpread2.Col = 6
       vaSpread2.Value = IIf(RS1!car_codigo = 0, "", RS1!car_codigo)
    
       vaSpread2.Col = 7
       vaSpread2.Value = RS1!car_nombre
    
       vaSpread2.Col = 8
       vaSpread2.Value = IIf(RS1!tip_codigo = 0, "", RS1!tip_codigo)
    
       vaSpread2.Col = 9
       vaSpread2.Value = RS1!tip_nombre
       
       vaSpread2.Col = 10
       vaSpread2.Value = RS1!ess_racmin
       
       vaSpread2.Col = 11
       vaSpread2.Value = IIf(IsNull(RS1!ID_HomologacionEstServicio) Or RS1!ID_HomologacionEstServicio = 0, "", RS1!ID_HomologacionEstServicio)
    
       vaSpread2.Col = 12
       vaSpread2.Value = IIf(IsNull(RS1!descripcion), "", RS1!descripcion)
       
       vaSpread2.Col = 13
       vaSpread2.Value = IIf(IsNull(RS1!ess_activo), "", RS1!ess_activo)
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    modo = ""
    Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True
    fpText1.Enabled = True
    OpGr = False
    
ElseIf SSTab1.Tab = 2 Then
    
    Me.Refresh
    vaSpread1.Col = 2
    vaSpread1.Row = vaSpread1.ActiveRow
    lblNOMBRE(1).Caption = vaSpread1.Value
    MoverDatosGrillas3

End If
End Sub

Private Sub vaSpread3_EditChange(ByVal Col As Long, ByVal Row As Long)

If vaSpread3.MaxRows < 1 Then Exit Sub
vaSpread3.Row = Row
vaSpread3.Col = Col
If Val(vaSpread3.text) = 0 Then vaSpread3.text = ""
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(1) = False
SumarTotales

End Sub

Private Sub vaSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If vaSpread3.MaxRows < 1 Then Exit Sub
vaSpread3.Row = Row
vaSpread3.Col = Col
If Val(vaSpread3.text) = 0 Then vaSpread3.text = ""

End Sub

Sub SumarTotales()

Dim i As Long, j As Long, nrorac As Long
For j = 1 To vaSpread3.maxcols
    
    vaSpread3.Row = vaSpread3.MaxRows
    vaSpread3.Col = j
    vaSpread3.text = ""

Next j

For i = 1 To (vaSpread3.MaxRows - 1)
    
    nrorac = 0
    
    For j = 1 To vaSpread3.maxcols
        
        vaSpread3.Row = i
        vaSpread3.Col = j
        nrorac = Val(vaSpread3.Value)
        vaSpread3.Row = vaSpread3.MaxRows
        If nrorac > 0 Then vaSpread3.Col = j: vaSpread3.Value = (Val(vaSpread3.Value) + nrorac)
    
    Next j

Next i

End Sub


