VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_BorrarCecoMasivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Borrar Sitios Masivo"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7665
   ScaleWidth      =   14970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7455
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14865
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   5
         Left            =   11520
         TabIndex        =   23
         Top             =   6120
         Width           =   2955
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   2850
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   4
         Left            =   10440
         TabIndex        =   22
         Top             =   6120
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   7560
         TabIndex        =   21
         Top             =   6120
         Width           =   2835
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   2730
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   6480
         TabIndex        =   20
         Top             =   6120
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   930
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   3360
         TabIndex        =   19
         Top             =   6120
         Width           =   3075
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   2970
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   2280
         TabIndex        =   18
         Top             =   6120
         Width           =   1035
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   930
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
         Height          =   375
         Index           =   1
         Left            =   12840
         TabIndex        =   12
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   11040
         TabIndex        =   11
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Frame Frame2 
         Height          =   1170
         Left            =   1590
         TabIndex        =   14
         Top             =   210
         Width           =   11775
         Begin EditLib.fpDateTime FpFecDesde 
            Height          =   315
            Left            =   2115
            TabIndex        =   1
            Top             =   705
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Text            =   "01/09/2013"
            DateCalcMethod  =   4
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
         Begin EditLib.fpDateTime FpFecHasta 
            Height          =   315
            Left            =   9300
            TabIndex        =   2
            Top             =   705
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            Text            =   "28/09/2013"
            DateCalcMethod  =   4
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Left            =   2130
            TabIndex        =   0
            Top             =   285
            Width           =   1275
            _Version        =   196608
            _ExtentX        =   2249
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
            AutoAdvance     =   -1  'True
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
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   10
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
            ButtonAlign     =   1
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   11160
            TabIndex        =   3
            Top             =   720
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Org. Compras"
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
            Left            =   855
            TabIndex        =   17
            Top             =   390
            Width           =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde"
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
            Left            =   840
            TabIndex        =   16
            Top             =   795
            Width           =   1140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta"
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
            Left            =   7995
            TabIndex        =   15
            Top             =   795
            Width           =   1095
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4515
         Left            =   195
         TabIndex        =   4
         Top             =   1575
         Width           =   14535
         _Version        =   393216
         _ExtentX        =   25638
         _ExtentY        =   7964
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
         MaxCols         =   9
         SpreadDesigner  =   "M_BorrarCecoMasivo.frx":0000
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
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
            Picture         =   "M_BorrarCecoMasivo.frx":1A2B
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "M_BorrarCecoMasivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lc_Aux As String
Dim MsgTitulo As String

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim seleccion As Integer
Dim i         As Long
Dim IdBloque  As Double
Dim Ceco      As String
Dim Regimen   As Long
Dim Servicio  As Long
Dim MyBuffer  As String
Dim Sql       As String
Dim RS        As New ADODB.Recordset

Select Case Index

    Case 0

        '-------> Validar org. compras
        If Trim(fpText.text) = "" Then
           
           MsgBox "Debe ingresar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------> Validar fechas
        If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
           
           MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
          
        If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
           
           MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
        
        End If
        
        '-------> Validar que exista un dato seleccionado
        seleccion = 0
        For i = 1 To vaSpread1.MaxRows
               
            vaSpread1.Row = i
            vaSpread1.Col = 1 'Seleccion
            seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
            
            If seleccion = 1 And vaSpread1.RowHidden = False Then
               
               Exit For
            
            End If
          
        Next i
          
        If seleccion = 0 Then
             
           MsgBox " Se debe seleccionar un item por lo menos", vbExclamation + vbOKOnly, MsgTitulo
           
           Exit Sub
          
        End If
        
        If MsgBox("Esta Seguro ?...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        Let MyBuffer = MyBuffer & "<Min>"
        
        For i = 1 To vaSpread1.MaxRows
               
            vaSpread1.Row = i
            vaSpread1.Col = 1 'Seleccion
            seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
            
            If seleccion = 1 And vaSpread1.RowHidden = False Then
        
               vaSpread1.Col = 2
               IdBloque = vaSpread1.text
              
               vaSpread1.Col = 3
               Ceco = vaSpread1.text
              
               vaSpread1.Col = 5
               Regimen = vaSpread1.text
        
               vaSpread1.Col = 7
               Servicio = vaSpread1.text
        
               IdBloque = Replace(Trim(IdBloque), Chr(34), "&quot;")
               IdBloque = Replace(Trim(IdBloque), Chr(38), "&amp;")
               IdBloque = Replace(Trim(IdBloque), Chr(39), "&apos;")
               IdBloque = Replace(Trim(IdBloque), Chr(60), "&lt;")
               IdBloque = Replace(Trim(IdBloque), Chr(62), "&gt;")
        
               Ceco = Replace(Trim(Ceco), Chr(34), "&quot;")
               Ceco = Replace(Trim(Ceco), Chr(38), "&amp;")
               Ceco = Replace(Trim(Ceco), Chr(39), "&apos;")
               Ceco = Replace(Trim(Ceco), Chr(60), "&lt;")
               Ceco = Replace(Trim(Ceco), Chr(62), "&gt;")
        
               Regimen = Replace(Trim(Regimen), Chr(34), "&quot;")
               Regimen = Replace(Trim(Regimen), Chr(38), "&amp;")
               Regimen = Replace(Trim(Regimen), Chr(39), "&apos;")
               Regimen = Replace(Trim(Regimen), Chr(60), "&lt;")
               Regimen = Replace(Trim(Regimen), Chr(62), "&gt;")
        
               Servicio = Replace(Trim(Servicio), Chr(34), "&quot;")
               Servicio = Replace(Trim(Servicio), Chr(38), "&amp;")
               Servicio = Replace(Trim(Servicio), Chr(39), "&apos;")
               Servicio = Replace(Trim(Servicio), Chr(60), "&lt;")
               Servicio = Replace(Trim(Servicio), Chr(62), "&gt;")
        
               MyBuffer = MyBuffer & " <MinD"
               MyBuffer = MyBuffer & " IdB = " & Chr(34) & IdBloque & Chr(34)
               MyBuffer = MyBuffer & " Cec = " & Chr(34) & Ceco & Chr(34)
               MyBuffer = MyBuffer & " Reg = " & Chr(34) & Regimen & Chr(34)
               MyBuffer = MyBuffer & " Ser = " & Chr(34) & Servicio & Chr(34)
              
               MyBuffer = MyBuffer & "/>"
        
            End If
          
        Next i

        MyBuffer = MyBuffer & "</Min>"
    
       'registrar Log sistema eliminación
       Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Prepara_Eliminar"), CStr(Me.HelpContextID), "", "", "")

        Sql = ""
        Sql = Sql & "'" & MyBuffer & "', "
        Sql = Sql & " " & Format(FpFecDesde, "yyyymmdd") & ", "
        Sql = Sql & " " & Format(FpFecHasta, "yyyymmdd") & ""
        Set RS = vg_db.Execute("sgpadm_Del_XmlMinutaBloqueMasivos_V02 " & Sql & "")
                   
        If Not RS.EOF Then
                      
           If Trim(RS(1)) = "" Then
                         
              MsgBox "Registro eliminado exitosamente", vbInformation + vbOKOnly, MsgTitulo
              
              'registrar Log sistema Eliminar
              Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminado"), Me.HelpContextID, "", "", "")

           Else
            
              MsgBox "Registro finalizo con error " & RS(0), vbInformation + vbOKOnly, MsgTitulo
            
              'registrar Log sistema error Eliminacion
              Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
            
           End If
                   
        End If
        RS.Close
        Set RS = Nothing
        
        'Refrescar grilla
        Toolbar2_ButtonClick Toolbar2.Buttons(1)
    Case 1 '-------> Salir de la opción

       Me.Hide
       Unload Me
        
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    'registrar Log sistema error Eliminacion
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")


End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
fg_centra Me
MsgTitulo = "Borrar Sitios Masivo"

FpFecHasta.text = Format(Date, "dd/mm/yyyy")
FpFecDesde.text = Format(Date, "dd/mm/yyyy")
vaSpread1.MaxRows = 0

Text1(3).text = ""
Text1(4).text = ""
Text1(5).text = ""
Text1(6).text = ""
Text1(7).text = ""
Text1(8).text = ""

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Unload(Cancel As Integer)

'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

Text1(3).text = ""
Text1(4).text = ""
Text1(5).text = ""
Text1(6).text = ""
Text1(7).text = ""
Text1(8).text = ""

vaSpread1.MaxRows = 0
If IsDate(FpFecDesde.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_Change()

On Error GoTo Man_Error

Text1(3).text = ""
Text1(4).text = ""
Text1(5).text = ""
Text1(6).text = ""
Text1(7).text = ""
Text1(8).text = ""

vaSpread1.MaxRows = 0
If IsDate(FpFecHasta.text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub FpFecHasta_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Text1(4).text = ""
Text1(5).text = ""
Text1(6).text = ""
Text1(7).text = ""
Text1(8).text = ""
   
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(Text1(Index).text, ",")

If Index = 3 Then
   
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 4 Then
   
   Text1(3).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 5 Then
   
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(6).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 6 Then
   
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(7).text = ""
   Text1(8).text = ""

ElseIf Index = 7 Then
   
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(8).text = ""

ElseIf Index = 8 Then
   
   Text1(3).text = ""
   Text1(4).text = ""
   Text1(5).text = ""
   Text1(6).text = ""
   Text1(7).text = ""

End If

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    vaSpread1.Col = 9
    vaSpread1.text = 0

Next

Select Case Index

Case 3, 4, 5, 6, 7, 8
    
    vaSpread1.Visible = False
    
    If Trim(Text1(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(wvarArr8020(X)) & "*"
           vaSpread1.Col = 2
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 9
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 9
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 9
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 9
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 9
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
    
    If Trim(Text1(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 9
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS        As New ADODB.Recordset
Dim Sql       As String
Dim i         As Long
Dim xmlceco   As String
Dim seleccion As String
Dim codCeco   As String

Select Case Button.Index
Case 1

  '-------> Validar org. compras
  If Trim(fpText.text) = "" Then
     
     MsgBox "Debe ingresar Org. Compras...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
  
  '-------> Validar fechas
  If Trim(FpFecHasta.text) = "" Or Trim(FpFecDesde.text) = "" Then
     
     MsgBox "Unas de las fecha esta nula...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If
    
  If Format(FpFecHasta, ("YYYYMMDD")) < Format(FpFecDesde, ("YYYYMMDD")) Then
     
     MsgBox "La fecha de hasta no puede ser menor que la fecha desde...", vbExclamation + vbOKOnly, MsgTitulo
     Exit Sub
  
  End If

  vaSpread1.MaxRows = 0
  vaSpread1.Row = -1: vaSpread1.Col = -1
  vaSpread1.BackColor = &HC0FFFF
   
  Text1(3).text = ""
  Text1(4).text = ""
  Text1(5).text = ""
  Text1(6).text = ""
  Text1(7).text = ""
  Text1(8).text = ""
   
  Sql = ""
  Sql = Sql & Trim(LimpiaDato(fpText.text))
  Set RS = vg_db.Execute("sgpadm_Sel_OrgComprasCecoMinutaBloque '" & Sql & "', '" & Format(FpFecDesde, ("YYYYmmdd")) & "', '" & Format(FpFecHasta, ("YYYYmmdd")) & "'")
  If Not RS.EOF Then
  
    Do While Not RS.EOF
        
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        
        vaSpread1.Col = 1
        vaSpread1.text = "0"
        
        vaSpread1.Col = 2
        vaSpread1.text = RS!Id_Bloque
        
        vaSpread1.Col = 3
        vaSpread1.text = RS!Ceco
        
        vaSpread1.Col = 4
        vaSpread1.text = Trim(RS!Cli_nombre)
        
        vaSpread1.Col = 5
        vaSpread1.text = RS!Regimen
        
        vaSpread1.Col = 6
        vaSpread1.text = Trim(RS!reg_nombre)
        
        vaSpread1.Col = 7
        vaSpread1.text = RS!Servicio
        
        vaSpread1.Col = 8
        vaSpread1.text = Trim(RS!ser_nombre)
        
        vaSpread1.Col = 9
        vaSpread1.text = 0
           
        RS.MoveNext
        
    Loop
  
  Else
     
     vaSpread1.MaxRows = 0
     MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo
  
  End If
  RS.Close
  Set RS = Nothing

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

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
    
    For i = 1 To vaSpread1.MaxRows
        
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
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
