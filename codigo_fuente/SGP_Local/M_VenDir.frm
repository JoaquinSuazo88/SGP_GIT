VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_VenDir 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta Directa"
   ClientHeight    =   7605
   ClientLeft      =   735
   ClientTop       =   2580
   ClientWidth     =   10830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Height          =   3795
      Left            =   225
      TabIndex        =   22
      Top             =   2865
      Width           =   10395
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3405
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   10155
         _Version        =   393216
         _ExtentX        =   17912
         _ExtentY        =   6006
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         AutoCalc        =   0   'False
         DisplayRowHeaders=   0   'False
         EditEnterAction =   2
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
         MaxCols         =   10
         MaxRows         =   20
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_VenDir.frx":0000
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin VB.Frame Frame2 
      Height          =   750
      Left            =   2190
      TabIndex        =   17
      Top             =   6690
      Width           =   6585
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   720
         Left            =   330
         TabIndex        =   18
         Top             =   240
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   1270
         ButtonWidth     =   3307
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   2
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agregar Producto"
               Description     =   "Agregar Productos"
               Object.ToolTipText     =   "Agregar Producto"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar Producto "
               Description     =   "Eliminar Producto "
               Object.ToolTipText     =   "Eliminar Producto "
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1635
         Top             =   135
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_VenDir.frx":07D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_VenDir.frx":0AF0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad sobrepasa Stock actual"
         Height          =   450
         Index           =   1
         Left            =   4560
         TabIndex        =   19
         Top             =   225
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   4170
         Top             =   345
         Width           =   300
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   450
         Top             =   315
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Productos igresados por el Usuario"
         Height          =   450
         Index           =   0
         Left            =   840
         TabIndex        =   21
         Top             =   195
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   735
         Top             =   300
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Label Label5 
         Caption         =   "Productos traidos de la Minuta Real"
         Height          =   450
         Index           =   2
         Left            =   1095
         TabIndex        =   20
         Top             =   180
         Visible         =   0   'False
         Width           =   1425
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2475
      Left            =   1110
      TabIndex        =   7
      Top             =   375
      Width           =   8580
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1890
         TabIndex        =   0
         Top             =   585
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         AlignTextH      =   2
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
         NullColor       =   -2147483643
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483648"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   180
         Width           =   3195
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1635
         Width           =   3195
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   0
         Left            =   1890
         TabIndex        =   1
         Top             =   945
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1890
         TabIndex        =   2
         Top             =   1290
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         ButtonStyle     =   3
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         ButtonColor     =   -2147483633
         AutoMenu        =   -1  'True
         StartMonth      =   4
         ButtonAlign     =   0
         BoundDataType   =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText1 
         Height          =   315
         Index           =   1
         Left            =   1890
         TabIndex        =   4
         Top             =   2040
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento"
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
         Left            =   330
         TabIndex        =   27
         Top             =   255
         Width           =   975
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   1950
         TabIndex        =   26
         Top             =   240
         Width           =   3180
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3720
         TabIndex        =   24
         Top             =   2040
         Width           =   4545
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   3270
         Picture         =   "M_VenDir.frx":0E0A
         Top             =   1935
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
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
         Left            =   330
         TabIndex        =   23
         Top             =   2070
         Width           =   600
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   1950
         TabIndex        =   14
         Top             =   1695
         Width           =   3180
      End
      Begin VB.Label Label1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   210
         Left            =   3435
         TabIndex        =   16
         Top             =   645
         Width           =   2055
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Bodega"
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
         Index           =   8
         Left            =   330
         TabIndex        =   13
         Top             =   1695
         Width           =   660
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Contrato"
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
         Index           =   6
         Left            =   330
         TabIndex        =   12
         Top             =   975
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3270
         Picture         =   "M_VenDir.frx":1114
         Top             =   840
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
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
         Left            =   330
         TabIndex        =   11
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión"
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
         Index           =   7
         Left            =   330
         TabIndex        =   10
         Top             =   1320
         Width           =   1245
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3720
         TabIndex        =   8
         Top             =   945
         Width           =   4545
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   6
         Left            =   3765
         TabIndex        =   9
         Top             =   975
         Width           =   4545
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3765
         TabIndex        =   25
         Top             =   2070
         Width           =   4545
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_VenDir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim RS4 As New ADODB.Recordset
Dim est As Boolean

Private Sub Combo1_Click(Index As Integer)
Dim feprod As Long, codser As Long, i As Long
If est Then Exit Sub
If vaSpread1.MaxRows = 0 Then Exit Sub
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i: vaSpread1.Col = 1
    RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
             "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
             "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
    vaSpread1.Col = 9
    If Not RS1.EOF Then vaSpread1.text = RS1!bod_canmer Else vaSpread1.text = 0
    RS1.Close: Set RS1 = Nothing
    '-------> REvisa color
    Dim canrea As Double, canbod As Double
    vaSpread1.Col = 4: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    vaSpread1.Col = 9: canbod = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
    If canbod - canrea < 0 Then
        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
        vaSpread1.Col = 8: vaSpread1.text = "S" 'Bloqueado
    Else
        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(2).FillColor
        vaSpread1.Col = 8: vaSpread1.text = "N"   'No Bloqueado
    End If
Next i
End Sub

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 8010
Me.Width = 10950
fg_centra Me
est = False
EspFecha fpDateTime1(0)
Me.HelpContextID = vg_OpcM
MsgTitulo = "Venta Directa"
Dim X As Boolean
vaSpread1.TextTip = 2
vaSpread1.TextTipDelay = 0
X = vaSpread1.SetTextTipAppearance("Arial", "11", False, False, &HC0FFFF, &H800000)
Gl_Mo_Botones Me, 4
vaSpread1.Row = -1
vaSpread1.Col = 4: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DCa
vaSpread1.Col = 5: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = 2
vaSpread1.Col = 6: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 7: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
vaSpread1.Col = 10: vaSpread1.TypeNumberSeparator = vg_CSep: vaSpread1.TypeNumberDecimal = vg_CDec: vaSpread1.TypeNumberDecPlaces = vg_DPr
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 1, "b_clientes", "cli_", "CliBod", "N"
Combo1(0).AddItem "Factura" & Space(150) & "(FA)"
Combo1(0).AddItem "Guia de Despacho" & Space(150) & "(GD)"
Limpia 2
End Sub

Private Sub Form_Resize()
If Me.WindowState = 2 Then
    Frame3.Left = (Me.Width \ 2) - (vaSpread1.Width \ 2)
    Frame1.Left = (Me.Width \ 2) - (Frame1.Width \ 2)
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
ElseIf Me.WindowState = 0 Then
    Frame3.Left = 255
    Frame1.Left = 1110
    Frame2.Left = 2190
End If
End Sub

Private Sub fpDateTime1_Change(Index As Integer)
If vaSpread1.MaxRows < 1 Then Exit Sub
Dim Cantidad As Double, valfin As Double, porsob As Double
For i = 1 To vaSpread1.MaxRows
    vaSpread1.Row = i: vaSpread1.Col = 1
    RS1.Open "SELECT TOP 1 b.ppd_propon, Max(b.ppd_fecdia) AS ppd_fecdia FROM b_productos a, b_productospmpdia b " & _
             "WHERE a.pro_codigo = b.ppd_codpro " & _
             "AND   b.ppd_cencos = '" & MuestraCasino(1) & "' " & _
             "AND   b.ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND b.ppd_fecdia <= " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & " " & _
             "AND   a.pro_codigo = '" & Trim(LimpiaDato(vaSpread1.text)) & "' " & _
             "AND   a.pro_ctrsto = 1 GROUP BY b.ppd_propon HAVING (b.ppd_propon)>0 ORDER BY Max(b.ppd_fecdia) DESC", vg_db, adOpenStatic
    If Not RS1.EOF Then
       vaSpread1.Col = 4: Cantidad = vaSpread1.text
       vaSpread1.Col = 5: porsob = vaSpread1.text
       vaSpread1.Col = 6: vaSpread1.Lock = True: vaSpread1.text = RS1!ppd_propon
       vaSpread1.Col = 10: vaSpread1.Lock = True: vaSpread1.text = RS1!ppd_propon
       valfin = Round(RS1!ppd_propon + ((RS1!ppd_propon * porsob) / 100), 0)
       vaSpread1.Col = 7: vaSpread1.Lock = True: vaSpread1.text = Format(Format(Cantidad, fg_Pict(9, 0)) * valfin, fg_Pict(9, 2))
    End If
    RS1.Close: Set RS1 = Nothing
Next
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpLongInteger1_LostFocus(Index As Integer)
If Trim(fpLongInteger1(0).text) <> "" Then
    BuscaDoc fpLongInteger1(0).Value
End If
End Sub

Private Sub fpText1_Change(Index As Integer)
fpayuda(Index).Caption = ""
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
Case 1
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Select
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
If fpText1(Index).text = "" Then Exit Sub
Select Case Index
Case 0
    RS1.Open "SELECT cli_nombre FROM b_clientes WHERE cli_codigo = '" & fpText1(0).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          fpayuda(Index).Caption = RS1!cli_nombre
          RS1.MoveNext
       Loop
    Else
       RS1.Close: Set RS1 = Nothing
       fpText1(0).text = ""
       MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo
       If fpText1(0).Enabled = True Then fpText1(0).SetFocus
       Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
Case 1
    RS3.Open "SELECT cli_nombre, cli_codigo FROM b_clientes WHERE cli_codigo = '" & fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))) & "' AND cli_tipo = 1 AND cli_activo = '1'", vg_db, adOpenStatic
    If Not RS3.EOF Then
       fpText1(1).text = fg_PintaRut(RS3!cli_codigo)
       fpayuda(Index).Caption = RS3!cli_nombre
    Else
       RS3.Close: Set RS3 = Nothing
       fpText1(1).text = ""
       MsgBox "Cliente no existe...", vbExclamation + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    RS3.Close: Set RS3 = Nothing
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
Dim Variable As String
vg_codigo = 0
vg_left = fpayuda(Index).Left + 1920
Variable = IIf(Index = 0, "Contrato", "Cliente")
B_TabEst.LlenaDatos "b_clientes", "cli_", Variable, Variable
B_TabEst.Show 1
Me.Refresh
If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
fpText1(Index) = Trim(vg_codigo)
fpayuda(Index).Caption = vg_nombre
Select Case Index
Case 0
    If Trim(fpText1(Index).text) = "" Then Exit Sub
    If fpDateTime1(Index).Enabled = True Then fpDateTime1(Index).SetFocus
Case 1
    If fpText1(1).Enabled = True Then fpText1(1).SetFocus
    fpText1_LostFocus 1
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim rutcli As String, codcas As String, tipdoc As String, NumDoc As Long, codbod  As Long, i As Long, canact As Double
Dim numlin As Long, codmer As String, canmer As Double, porcen As Double, predoc As Double
Dim ptotal As Double, descri As String, fecpro As Date, total As Double, precos As Double, diablq As Date
On Error GoTo Man_Error
fecpro = Format(fpDateTime1(0).Value, "dd/mm/yyyy")
codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Now, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
TraerFechaCierre
Select Case Button.Index
Case 1, 6 '-------> Nuevo
    If Button.Index = 6 And vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    Limpia IIf(Button.Index = 1, 6, 2)
    If fpLongInteger1(0).Enabled = True Then fpLongInteger1(0).SetFocus
Case 8 '-------> Graba
    If Trim(fpText1(0).text) = "" Or Trim(fpLongInteger1(0).text) = "" Or Trim(fpText1(1).text) = "" Or Trim(Combo1(0).text) = "" _
    Or Trim(Combo1(1).text) = "" Or Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe ingresar dato importante...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    '-------> Validar si el contrato tiene asignado inventario rotativo
    If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
       Format(fpDateTime1(0).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then
    
       MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then
    
       MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    'Validar inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 38) Then
        
       MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
        
    'Validar ingreso documento inventario calendarizado 20201001
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 40) Then
        
       MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
    End If
    
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 8) Then
    
       MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then
    
       MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
    End If
    
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 8
        If Left(vaSpread1.text, 1) = "S" Then MsgBox "Existe una cantidad que exede el Stock...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    total = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 7: ptotal = LimpiaDato(vaSpread1.text)
        total = total + ptotal
        vaSpread1_Change 6, i
    Next i
    If total = 0 Then MsgBox "El total del documento debe ser mayor a 0...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Desea grabar...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    rutcli = Trim(LimpiaDato(fpText1(0).text))
    codcas = fg_DespintaRut(Trim(LimpiaDato(fpText1(1).text)))
    tipdoc = fg_codigocbo(Combo1, 0, 2, "")
    NumDoc = fpLongInteger1(0).text
    codbod = Val(fg_codigocbo(Combo1, 1, 10, ""))
    '-------> Encabezado
    If vg_tipbase = "1" Then
       vg_db.Execute "insert into b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                     "values ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", CDate('" & _
                     Format(fpDateTime1(0).text, "dd/mm/yyyy") & "'), 0, 0, 0, '', '" & codcas & "', 0)"
    Else
       vg_db.Execute "INSERT INTO b_totventas (tov_rutcli , tov_tipdoc, tov_numdoc, tov_codbod, tov_fecemi, tov_codreg, tov_codser, tov_otrimp, tov_estdoc, tov_codcas, tov_numinf) " & _
                     "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & codbod & ", '" & _
                     Format(fpDateTime1(0).text, "yyyymmdd") & "', 0, 0, 0, '', '" & codcas & "', 0)"
    End If
    '-------> Detalle
    total = 0
    numlin = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 2: descri = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 4: canmer = LimpiaDato(vaSpread1.text)
        vaSpread1.Col = 5: porcen = LimpiaDato(vaSpread1.text)
        vaSpread1.Col = 6: predoc = LimpiaDato(vaSpread1.text)
        vaSpread1.Col = 7: ptotal = LimpiaDato(vaSpread1.text)
        vaSpread1.Col = 10: precos = LimpiaDato(vaSpread1.text)
        If canmer > 0 Then
            total = total + ptotal
            vg_db.Execute "INSERT INTO b_detventas (dev_rutcli, dev_tipdoc, dev_numdoc, dev_numlin, dev_codmer, dev_canmin, dev_canmer, dev_predoc, dev_ptotal, dev_descri, dev_mueinv, dev_porcen, dev_precos, dev_coding) " & _
                          "VALUES ('" & rutcli & "', '" & tipdoc & "', " & NumDoc & ", " & numlin & ", '" & codmer & "', " & "0" & ", " & canmer & ", " & predoc & ", " & ptotal & ", '" & descri & "', 'S', " & porcen & ", " & precos & ", '')"
            '-------> Control de Stock
            canact = 0
            RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro='" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod=" & vg_codbod, vg_db, adOpenStatic
            If Not RS1.EOF Then
                Do While Not RS1.EOF
                    canact = RS1!bod_canmer - canmer ' Entregado
                    RS1.MoveNext
                Loop
                vg_db.Execute "UPDATE b_bodegas SET bod_canmer=" & canact & " WHERE bod_codpro='" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod=" & vg_codbod
            End If
            RS1.Close: Set RS1 = Nothing: numlin = numlin + 1
        End If
    Next i
    '-------> Total
    vg_db.Execute "UPDATE b_totventas SET tov_totdoc=" & total & " WHERE tov_rutcli='" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                  "AND tov_tipdoc='" & tipdoc & "' AND tov_numdoc=" & fpLongInteger1(0).Value & " AND tov_codbod=" & vg_codbod & ""
    vg_db.CommitTrans
    Gl_Ac_Botones Me, 4, 3, ""
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    '-------> Revisa Stock
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        RS1.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND bod.bod_codpro=pro.pro_codigo AND pro.pro_codigo='" & Trim(LimpiaDato(vaSpread1.text)) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS1.EOF Then vaSpread1.text = Format(RS1!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS1.Close: Set RS1 = Nothing
    Next i
    I_VentaDir Me, fg_codigocbo(Combo1, 0, 2, "")
Case 3 '-------> Eliminar
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 0) Then MsgBox "Periodo cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CierrePeriodo(Format(fpDateTime1(0).text, "yyyymmdd"), codbod, 6) Then MsgBox "No puede ingresar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If CDate(fpDateTime1(0).text) < CDate(vg_ciedia) Then MsgBox "No puede eliminar documento, día esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina documento...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    tipdoc = fg_codigocbo(Combo1, 0, 2, "")
    vg_db.BeginTrans
    '-------> Detalle
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: numlin = i
        vaSpread1.Col = 1: codmer = Trim(LimpiaDato(vaSpread1.text))
        vaSpread1.Col = 4: canmer = LimpiaDato(vaSpread1.text)
        '-------> Control de Stock
        canact = 0
        RS1.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codpro='" & codmer & "' AND bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, "")), vg_db, adOpenStatic
        If Not RS1.EOF Then
            Do While Not RS1.EOF
                    canact = RS1!bod_canmer + canmer ' Anula Entregado
                RS1.MoveNext
            Loop
            vg_db.Execute "UPDATE b_bodegas SET bod_canmer=" & canact & " " & _
                          "WHERE bod_codpro='" & Trim(LimpiaDato(codmer)) & "' AND bod_codbod=" & Val(fg_codigocbo(Combo1, 1, 10, ""))
        End If
        RS1.Close: Set RS1 = Nothing
    Next i
    vg_db.Execute "DELETE b_detventas FROM b_detventas WHERE dev_rutcli = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                  "AND dev_tipdoc = '" & tipdoc & "' AND dev_numdoc = " & fpLongInteger1(0).Value
    vg_db.Execute "DELETE b_detventasimp FROM b_detventasimp WHERE imd_rutdoc = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                  "AND imd_tipdoc = '" & tipdoc & "' AND imd_numdoc = " & fpLongInteger1(0).Value
    vg_db.Execute "DELETE b_totventas FROM b_totventas WHERE tov_rutcli = '" & Trim(LimpiaDato(fpText1(0).text)) & "' " & _
                  "AND tov_tipdoc = '" & tipdoc & "' AND tov_numdoc = " & fpLongInteger1(0).Value & " AND tov_codbod = " & vg_codbod & ""
    vg_db.CommitTrans
    Limpia 2
Case 11 '-------> Busqueda
    vg_codigo = Trim(fpText1(0).text)
    vg_nombre = fg_codigocbo(Combo1, 0, 2, "")
    B_SalBod.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    BuscaDoc Val(vg_codigo)
Case 12 '-------> Imprimir
    I_VentaDir Me, fg_codigocbo(Combo1, 0, 2, "")
Case 15 '-------> Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Sub BuscaDoc(codigo As Long)
'-------> Encabezado
RS2.Open "SELECT tov.tov_numdoc, tov.tov_codbod, tov.tov_fecemi, tov.tov_codser, tov.tov_estdoc, tov.tov_codcas, tov.tov_rutcli " & _
         "FROM b_totventas tov, b_clientes cli " & _
         "WHERE tov.tov_tipdoc = '" & fg_codigocbo(Combo1, 0, 2, "") & "' " & _
         "AND   tov.tov_numdoc = " & Val(codigo) & " AND tov.tov_codbod = " & vg_codbod & " " & _
         "AND   tov.tov_rutcli = cli.cli_codigo", vg_db, adOpenStatic
If Not RS2.EOF Then
    Frame1.Enabled = False
    Frame2.Enabled = False
    vaSpread1.Col = -1: vaSpread1.Row = -1
    vaSpread1.Lock = True
    vaSpread1.MaxRows = 0
    Do While Not RS2.EOF
        est = True
        fpLongInteger1(0).text = RS2!tov_numdoc
        Combo1(1).ListIndex = fg_buscacbo(Combo1, 1, 10, fg_pone_cero(Str(RS2!tov_codbod), 10))
        fpDateTime1(0).text = RS2!tov_fecemi
        Label1.Caption = IIf(RS2!tov_estdoc = "A", "ANULADA", "")
        fpText1(0).text = RS2!tov_rutcli
        fpText1_LostFocus 0
        fpText1(1).text = fg_PintaRut(RS2!tov_codcas)
        fpText1_LostFocus 1
        est = False
        RS2.MoveNext
    Loop
Else
    Gl_Ac_Botones Me, 4, 6, ""
    fpLongInteger1(0).Enabled = False
    Combo1(0).Enabled = False
    RS2.Close: Set RS2 = Nothing
    Exit Sub
End If
RS2.Close: Set RS2 = Nothing
'-------> Detalle
RS3.Open "SELECT dev.dev_codmer, dev.dev_canmer, dev.dev_predoc, " & _
         "dev.dev_ptotal, dev.dev_descri, uni.uni_nombre, dev_precos, dev_porcen " & _
         "FROM b_detventas dev, b_productos pro ,a_unidad uni " & _
         "WHERE dev.dev_rutcli = '" & LimpiaDato(Trim(fpText1(0).text)) & "'" & _
         "AND   dev.dev_tipdoc = '" & fg_codigocbo(Combo1, 0, 2, "") & "' " & _
         "AND   dev.dev_numdoc = " & Val(codigo) & " " & _
         "AND   dev.dev_codmer = pro.pro_codigo " & _
         "AND   pro.pro_coduni = uni.uni_codigo ORDER BY dev.dev_numlin", vg_db, adOpenStatic
If Not RS3.EOF Then
    i = 1
    Do While Not RS3.EOF
        vaSpread1.MaxRows = i
        vaSpread1.Row = i
        vaSpread1.Col = 1: vaSpread1.text = RS3!dev_codmer
        vaSpread1.Col = 2: vaSpread1.text = RS3!dev_descri
        vaSpread1.Col = 3: vaSpread1.text = RS3!uni_nombre
        vaSpread1.Col = 4: vaSpread1.text = RS3!dev_canmer
        vaSpread1.Col = 5: vaSpread1.text = RS3!dev_porcen
        vaSpread1.Col = 6: vaSpread1.text = RS3!dev_predoc
        vaSpread1.Col = 7: vaSpread1.text = RS3!dev_ptotal
        vaSpread1.Col = 10: vaSpread1.text = RS3!dev_precos
        'Trae Stock
        RS4.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                 "WHERE bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                 "AND   bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(RS3!dev_codmer) & "'", vg_db, adOpenStatic
        vaSpread1.Col = 9
        If Not RS4.EOF Then vaSpread1.text = Format(RS4!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
        RS4.Close: Set RS4 = Nothing
        RS3.MoveNext: i = i + 1
    Loop
End If
RS3.Close: Set RS3 = Nothing
vg_codigo = ""
Gl_Ac_Botones Me, 4, IIf(Label1.Caption = "ANULADA", 4, 3), ""
End Sub

Sub Limpia(op As Integer)
est = True
Label1.Caption = ""
Frame1.Enabled = True
fpLongInteger1(0).Enabled = True
Combo1(0).Enabled = True
Frame2.Enabled = True
fpLongInteger1(0).text = ""
fpText1(1).text = ""
fpayuda(1).Caption = ""
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Combo1(0).ListIndex = 0
Combo1(1).ListIndex = IIf(Combo1(1).listcount = 1, 0, -1)
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
vaSpread1.Col = 4
vaSpread1.Lock = False
vaSpread1.Col = 5
vaSpread1.Lock = False
vaSpread1.Col = 6
vaSpread1.Lock = False
fpText1(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(0).text = MuestraCasino(1)
fpayuda(0).Caption = MuestraCasino(2)
Gl_Ac_Botones Me, 4, op, ""
est = False
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, propon As Double
Select Case Button.Index
Case 1
    vg_nombre = "": vg_codigo = "": vg_bodega = 0
    vg_bodega = Val(fg_codigocbo(Combo1, 1, 10, ""))
    vg_left = fpayuda(1).Left + 1920
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pbo"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Col = 1: vaSpread1.Row = i
        If Trim(vaSpread1.text) = Trim(vg_codigo) Then MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    vaSpread1.Row = vaSpread1.ActiveRow
    RS1.Open "SELECT a.pro_codigo, a.pro_nombre, b.uni_nombre " & _
             "FROM b_productos a, a_unidad b " & _
             "WHERE  a.pro_coduni = b.uni_codigo " & _
             "AND    a.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
        i = vaSpread1.MaxRows + 1
        Do While Not RS1.EOF
            vaSpread1.MaxRows = i
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = 1: vaSpread1.text = RS1!pro_codigo
            vaSpread1.Col = 2: vaSpread1.text = RS1!pro_nombre
            vaSpread1.Col = 3: vaSpread1.text = RS1!uni_nombre
            vaSpread1.Col = 4: vaSpread1.text = 0
            vaSpread1.Col = 5: vaSpread1.text = 0
            '-------> Trae pmp
            propon = 0
            RS2.Open "SELECT TOP 1 ppd_cencos, ppd_codpro, ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "FROM b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_codpro = '" & RS1!pro_codigo & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(fpDateTime1(0).text), "yyyymmdd") & " " & _
                     "GROUP BY ppd_cencos, ppd_codpro, ppd_propon " & _
                     "HAVING (ppd_propon)>0 ORDER BY Max(ppd_fecdia) DESC", vg_db, adOpenStatic
            If Not RS2.EOF Then propon = RS2!ppd_propon
            RS2.Close: Set RS2 = Nothing
            
            vaSpread1.Col = 6: vaSpread1.text = propon
            vaSpread1.Col = 7: vaSpread1.text = 0
            vaSpread1.Col = 8: vaSpread1.text = "N" 'No bloquedo
            vaSpread1.Col = 10: vaSpread1.text = propon
            '-------> Trae Stock
            RS2.Open "SELECT bod.bod_canmer FROM b_productos AS pro, b_bodegas AS bod " & _
                     "WHERE  bod.bod_codbod = " & Val(fg_codigocbo(Combo1, 1, 10, "")) & " " & _
                     "AND    bod.bod_codpro = pro.pro_codigo AND pro.pro_codigo = '" & Trim(RS1!pro_codigo) & "'", vg_db, adOpenStatic
            vaSpread1.Col = 9
            If Not RS2.EOF Then vaSpread1.text = Format(RS2!bod_canmer, fg_Pict(9, vg_DCa)) Else vaSpread1.text = 0
            RS2.Close: Set RS2 = Nothing
            RS1.MoveNext
            i = i + 1
        Loop
    End If
    RS1.Close: Set RS1 = Nothing
    If vaSpread1.MaxRows = 1 Then Gl_Ac_Botones Me, 4, 6, ""
    vaSpread1.Col = 4: vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
Case 2
    If vaSpread1.MaxRows = 0 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    If vaSpread1.MaxRows = 0 Then Gl_Ac_Botones Me, 4, 6, ""
    If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
End Select
End Sub

Private Sub vaSpread1_Change(ByVal Col As Long, ByVal Row As Long)
Dim canrea As Double, predoc As Double, precos As Double, porsob As Double
vaSpread1.Col = 4: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 6: predoc = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
vaSpread1.Col = 10: precos = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
Select Case Col
Case 5, 6
    If precos = 0 Then vaSpread1.Col = 5: vaSpread1.text = Format(0, fg_Pict(9, 2)): Exit Sub
    porsob = Format(((predoc - precos) * 100) / precos, fg_Pict(9, 2))
    vaSpread1.Col = 5
    If porsob < 0 Then
        vaSpread1.text = 0
        vaSpread1.Col = 6: vaSpread1.text = Format(precos, fg_Pict(9, vg_DPr))
        vaSpread1.Col = 7: vaSpread1.text = Format(canrea * precos, fg_Pict(9, vg_DPr))
    End If
End Select
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
Dim canrea As Double, predoc As Double, precos As Double, porsob As Double, codmer As String, total As Double, totalgen As Double
vaSpread1.Row = Row
vaSpread1.Col = 1: codmer = vaSpread1.text
vaSpread1.Col = 4: canrea = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 5: porsob = Format(vaSpread1.text, fg_Pict(9, 2))
vaSpread1.Col = 6: predoc = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
vaSpread1.Col = 9: canbod = Format(vaSpread1.text, fg_Pict(9, vg_DCa))
vaSpread1.Col = 10: precos = Format(vaSpread1.text, fg_Pict(9, vg_DPr))
Select Case Col
Case 4 '-------> Cantidad
    vaSpread1.Col = 7: vaSpread1.text = Format(canrea * predoc, fg_Pict(9, vg_DPr))
    If canbod - canrea >= 0 Then
        vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(2).FillColor
        vaSpread1.Col = 8: vaSpread1.text = "N"  'No Bloqueado
        Exit Sub
    End If
    vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
    vaSpread1.Col = 8: vaSpread1.text = "S"  'Bloqueado
Case 5 '-------> %Sobre Costo
    If precos = 0 Then vaSpread1.Col = 5: vaSpread1.text = Format(0, fg_Pict(9, 2)): Exit Sub
    predoc = Round(precos + ((precos * porsob) / 100), 0)
    vaSpread1.Col = 6: vaSpread1.text = Format(predoc, fg_Pict(9, vg_DPr))
    vaSpread1.Col = 7: vaSpread1.text = Format(canrea * predoc, fg_Pict(9, vg_DPr))
Case 6 '-------> Precio
    If precos = 0 Then Exit Sub
    porsob = Format(((predoc - precos) * 100) / precos, fg_Pict(9, 2))
    vaSpread1.Col = 5: vaSpread1.text = Format(porsob, fg_Pict(9, 2))
    vaSpread1.Col = 7: vaSpread1.text = Format(canrea * predoc, fg_Pict(9, vg_DPr))
End Select
End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Or vaSpread1.MaxRows = 0 Then Exit Sub
If vaSpread1.ActiveCol = 4 Then vaSpread1.SetActiveCell 5, vaSpread1.ActiveRow - 1: Exit Sub
If vaSpread1.ActiveCol = 5 Then vaSpread1.SetActiveCell 6, vaSpread1.ActiveRow - 1: Exit Sub
If vaSpread1.ActiveCol = 6 And vaSpread1.ActiveRow - 1 <> vaSpread1.MaxRows Then vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
If Row = 0 Then Exit Sub
Dim Stock As String, Nombre As String
TipWidth = 4000
ShowTip = True
MultiLine = 2
vaSpread1.Row = Row: vaSpread1.Col = 9: Stock = vaSpread1.text
vaSpread1.Row = Row: vaSpread1.Col = 2: Nombre = vaSpread1.text
TipText = "Bodega   : " & Trim(Left(Combo1(1).text, 50)) & vbCrLf & _
          "Producto : " & Trim(Nombre) & vbCrLf & _
          "Stock       : " & Format(Trim(Stock), fg_Pict(9, vg_DCa))
End Sub
