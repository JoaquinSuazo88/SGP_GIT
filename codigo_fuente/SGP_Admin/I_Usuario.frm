VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form I_TransUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Usuarios"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8115
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   10335
      Begin VB.Frame Frame1 
         Height          =   645
         Index           =   1
         Left            =   1695
         TabIndex        =   11
         Top             =   240
         Width           =   6690
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   0
            Left            =   705
            TabIndex        =   12
            Top             =   225
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
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
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Date1 
            Height          =   315
            Index           =   1
            Left            =   4170
            TabIndex        =   13
            Top             =   225
            Width           =   1320
            _Version        =   196608
            _ExtentX        =   2328
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
            ButtonStyle     =   3
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
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
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "dd/mm/yyyy"
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Timer1 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   14
            Top             =   225
            Width           =   1065
            _Version        =   196608
            _ExtentX        =   1879
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
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
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "hh:nn:ss"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   2
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime Timer1 
            Height          =   315
            Index           =   1
            Left            =   5505
            TabIndex        =   15
            Top             =   225
            Width           =   1065
            _Version        =   196608
            _ExtentX        =   1879
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
            ButtonStyle     =   1
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
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
            OnFocusPosition =   1
            ControlType     =   0
            Text            =   ""
            DateCalcMethod  =   0
            DateTimeFormat  =   5
            UserDefinedFormat=   "hh:nn:ss"
            DateMax         =   "00000000"
            DateMin         =   "00000000"
            TimeMax         =   "000000"
            TimeMin         =   "000000"
            TimeString1159  =   ""
            TimeString2359  =   ""
            DateDefault     =   "00000000"
            TimeDefault     =   "000000"
            TimeStyle       =   2
            BorderGrayAreaColor=   -2147483637
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   17
            Top             =   270
            Width           =   465
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   1
            Left            =   3675
            TabIndex        =   16
            Top             =   270
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Función del sistema"
         Height          =   915
         Index           =   2
         Left            =   1680
         TabIndex        =   7
         Top             =   1065
         Width           =   6705
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   1
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   465
            Width           =   6465
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Uno"
            Height          =   240
            Index           =   2
            Left            =   165
            TabIndex        =   9
            Top             =   210
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   240
            Index           =   3
            Left            =   2850
            TabIndex        =   8
            Top             =   225
            Width           =   750
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo de operación"
         Height          =   915
         Index           =   3
         Left            =   1680
         TabIndex        =   3
         Top             =   2160
         Width           =   6705
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            Height          =   240
            Index           =   4
            Left            =   2850
            TabIndex        =   6
            Top             =   225
            Width           =   750
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Uno"
            Height          =   240
            Index           =   5
            Left            =   165
            TabIndex        =   5
            Top             =   210
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   2
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   465
            Width           =   6465
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4455
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   3480
      Width           =   10335
      Begin VB.CommandButton Boton 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   8565
         TabIndex        =   19
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton Boton 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   7080
         TabIndex        =   18
         Top             =   3960
         Width           =   1455
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   9855
         _Version        =   393216
         _ExtentX        =   17383
         _ExtentY        =   5953
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
         MaxCols         =   3
         SpreadDesigner  =   "I_Usuario.frx":0000
      End
   End
End
Attribute VB_Name = "I_TransUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_AdoDato As ADODB.Recordset
Dim MsgTitulo As String

Private Sub Combo1_Click(Index As Integer)

Select Case Index

    Case 0
    
        Combo1(3).ListIndex = Combo1(0).ListIndex
        
    Case 3
    
        Combo1(0).ListIndex = Combo1(3).ListIndex
        
End Select

End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

'Me.Height = 4635
'Me.Width = 6900
fg_centra Me
'Option1(0).Value = True
Option1(3).Value = True
Option1(4).Value = True

'Función del sistema
Combo1(1).Clear
Set RS_Dato = vg_db.Execute("select * from a_opcsistema order by opc_codigo")
Do While Not RS_Dato.EOF
    
    Combo1(1).AddItem RS_Dato!opc_nombre & Space(150) & "(" & fg_pone_rchar(RS_Dato!opc_codigo, 14, " ") & ")"
    RS_Dato.MoveNext

Loop
RS_Dato.Close: Set RS_Dato = Nothing

'Tipo de operación
Combo1(2).Clear
Set RS_Dato = vg_db.Execute("select * from log_conceptos order by loc_descripcion")
Do While Not RS_Dato.EOF
    
    Combo1(2).AddItem RS_Dato!loc_descripcion & Space(150) & "(" & fg_pone_cero(RS_Dato!loc_codigo, 3) & ")"
    RS_Dato.MoveNext

Loop
RS_Dato.Close: Set RS_Dato = Nothing

'ControlAccesoGen Boton, "", "", "", "", "0"

MsgTitulo = Me.Caption

End Sub

Private Sub Boton_Click(Index As Integer)

Dim cDateIni As String, cDateFin As String, cUsuario As String, cFunSis As String, cTipOpe As Long, CambiaPass As String

Select Case Index

    Case 0
        If Trim(Date1(0).text) = "" Or Trim(Date1(1).text) = "" Or Trim(Timer1(0).text) = "" Or Trim(Timer1(1).text) = "" Then MsgBox "Debe ingresar período.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
        cDateIni = Format(Date1(0).text, "yyyy-mm-dd") & " " & Format(Timer1(0).text, "hh:nn:ss")
        cDateFin = Format(Date1(1).text, "yyyy-mm-dd") & " " & Format(Timer1(1).text, "hh:nn:ss")
        If CDate(cDateIni) > CDate(cDateFin) Then MsgBox "Período de fechas no válida.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
        If Option1(1).Value And Combo1(0).ListIndex = -1 Then MsgBox "Debe seleccionar usuario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Option1(2).Value And Combo1(1).ListIndex = -1 Then MsgBox "Debe seleccionar Función del sistema.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Option1(5).Value And Combo1(2).ListIndex = -1 Then MsgBox "Debe seleccionar Tipo de operación.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
        fg_carga ""
    
        cUsuario = Trim(Combo1(0).List(Combo1(0).ListIndex))
        cFunSis = Trim(fg_codigocbo(Combo1, 1, 14, ""))
        cTipOpe = fg_codigocbo(Combo1, 2, 3, 0)
        CambiaPass = fg_TraeLogConcepto("vg_logsis_CambiaPass")
        Call fg_GrabaLogSistema(vg_NUsu, fg_TraeLogConcepto("vg_logsis_Aceptar"), vg_OpSis_I_LogSis, "", "")
    
        Set RS_Dato = dbAdo.Execute("san_logMuestraTrans '" & cDateIni & "', '" & cDateFin & "', '" & cUsuario & "', '" & cFunSis & "', " & cTipOpe & "")
        If Not RS_Dato.EOF Then
            
            vg_ArchTxt = fg_ArchivoTxt
            
            Open vg_ArchTxt For Output As #1
            
            Print #1, MsgTitulo
            Print #1, "Fecha Emisión :|" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:nn:ss")
            Print #1, ""
            Print #1, "Fecha|Usuario|Nombre Usuario|Función del sistema|Tipo de operación|Dato anterior|Dato nuevo"
            
            Do While Not RS_Dato.EOF
                'fecha, usuario, usr_nomape, opcionsist, isnull(opt_nombre, '') funsis, loc_descripcion tipope
                Print #1, " " & Format(RS_Dato!Fecha, "dd/mm/yyyy hh:mm:ss") & "|" & RS_Dato!usuario & "|" & RS_Dato!usr_nomape & "|" & IIf(RS_Dato!funsis <> "", RS_Dato!funsis, TipoDato(RS_Dato!opcionsist, "")) & "|" & RS_Dato!tipope & "|" & IIf(RS_Dato!tiporegistro = CambiaPass, "", RS_Dato!datoanterior) & "|" & IIf(RS_Dato!tiporegistro = CambiaPass, "", RS_Dato!datonuevo)
                RS_Dato.MoveNext
            Loop
            Close #1
            vg_opimp = 0
            
            CargaExcel
        
        Else
            
            MsgBox "No existen datos para los filtros. ", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        RS_Dato.Close: Set RS_Dato = Nothing
        fg_descarga
    
    Case 1
        
        Me.Hide
        Unload Me

End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), CStr(Me.HelpContextID), "", "")

'      Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), CStr(Me.HelpContextID), "", "")

End Sub

Private Sub Option1_Click(Index As Integer)

Select Case Index

Case 0
    
    Combo1(0).Enabled = False: Combo1(3).Enabled = False
    Combo1(0).ListIndex = -1: Combo1(3).ListIndex = -1

Case 1
    
    Combo1(0).Enabled = True: Combo1(3).Enabled = True
    If Combo1(0).ListCount > 0 Then Combo1(0).ListIndex = 0
    If Combo1(3).ListCount > 0 Then Combo1(3).ListIndex = 0

Case 3
    
    Combo1(1).Enabled = False
    Combo1(1).ListIndex = -1

Case 2
    
    Combo1(1).Enabled = True
    If Combo1(1).ListCount > 0 Then Combo1(1).ListIndex = 0
    
Case 4
    
    Combo1(2).Enabled = False
    Combo1(2).ListIndex = -1

Case 5
    
    Combo1(2).Enabled = True
    If Combo1(2).ListCount > 0 Then Combo1(2).ListIndex = 0

End Select

End Sub

Private Sub Timer1_KeyPress(Index As Integer, KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

End Sub

