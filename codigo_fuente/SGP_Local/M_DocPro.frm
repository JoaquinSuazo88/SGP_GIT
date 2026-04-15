VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_DocPro 
   Caption         =   "Documento Proveedor"
   ClientHeight    =   8565
   ClientLeft      =   3060
   ClientTop       =   1185
   ClientWidth     =   15630
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   15630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5115
      Left            =   45
      TabIndex        =   48
      Top             =   3300
      Width           =   15525
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   13560
         TabIndex        =   55
         Text            =   "Text2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   6720
         TabIndex        =   54
         Text            =   "Text2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3680
         TabIndex        =   53
         Text            =   "Text2"
         Top             =   3480
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Frame Frame3 
         Caption         =   "Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   3675
         TabIndex        =   51
         Top             =   3750
         Width           =   7530
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            Height          =   1020
            Index           =   0
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   52
            Top             =   225
            Width           =   7320
         End
      End
      Begin VB.Frame Frame6 
         Height          =   825
         Left            =   11595
         TabIndex        =   50
         Top             =   3810
         Width           =   3690
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   390
            Left            =   600
            TabIndex        =   21
            Top             =   330
            Width           =   2880
            _ExtentX        =   5080
            _ExtentY        =   688
            ButtonWidth     =   2302
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agr. Prod."
                  Description     =   "Agregar Productos"
                  Object.ToolTipText     =   "Agregar Producto"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Elim. Prod. "
                  Description     =   "Eliminar Producto "
                  Object.ToolTipText     =   "Eliminar Producto "
                  ImageIndex      =   2
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Impuesto del Producto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         TabIndex        =   49
         Top             =   3750
         Width           =   3495
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   1005
            Left            =   90
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   270
            Width           =   3315
            _Version        =   393216
            _ExtentX        =   5847
            _ExtentY        =   1773
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
            MaxCols         =   8
            MaxRows         =   3
            ProcessTab      =   -1  'True
            ScrollBars      =   2
            SpreadDesigner  =   "M_DocPro.frx":0000
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3195
         Left            =   135
         TabIndex        =   19
         Top             =   225
         Width           =   15255
         _Version        =   393216
         _ExtentX        =   26908
         _ExtentY        =   5636
         _StockProps     =   64
         AutoClipboard   =   0   'False
         ButtonDrawMode  =   1
         ColsFrozen      =   2
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
         MaxCols         =   31
         MaxRows         =   50
         ProcessTab      =   -1  'True
         SelectBlockOptions=   1
         SpreadDesigner  =   "M_DocPro.frx":0598
         VisibleCols     =   11
         TextTip         =   2
         TextTipDelay    =   0
         ScrollBarTrack  =   3
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   105
         Top             =   3585
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
               Picture         =   "M_DocPro.frx":1991
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_DocPro.frx":1CAB
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   6075
         Picture         =   "M_DocPro.frx":1FC5
         ToolTipText     =   "Calcular Totales"
         Top             =   3360
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   45
      TabIndex        =   17
      Top             =   360
      Width           =   13605
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "M_DocPro.frx":22CF
         Left            =   7995
         List            =   "M_DocPro.frx":22D1
         Style           =   2  'Dropdown List
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   465
         Width           =   3525
      End
      Begin VB.ComboBox Combo2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         ItemData        =   "M_DocPro.frx":22D3
         Left            =   3420
         List            =   "M_DocPro.frx":22D5
         Style           =   2  'Dropdown List
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1320
         Width           =   2670
      End
      Begin VB.Frame Frame5 
         Caption         =   "Tipo de Documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   10395
         TabIndex        =   25
         Top             =   855
         Width           =   3105
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "FOFI"
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
            Left            =   135
            TabIndex        =   18
            Top             =   585
            Width           =   1110
         End
         Begin VB.OptionButton Option1 
            Alignment       =   1  'Right Justify
            Caption         =   "CFC"
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
            Index           =   1
            Left            =   135
            TabIndex        =   8
            Top             =   285
            Width           =   1110
         End
         Begin EditLib.fpDoubleSingle Double1 
            Height          =   330
            Index           =   6
            Left            =   1770
            TabIndex        =   9
            Top             =   450
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   582
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
            AllowNull       =   0   'False
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
            Text            =   "0"
            DecimalPlaces   =   0
            DecimalPoint    =   ""
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ""
            UseSeparator    =   0   'False
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
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
            Caption         =   "Folio Nº"
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
            Left            =   1890
            TabIndex        =   26
            Top             =   210
            Width           =   690
         End
      End
      Begin VB.Frame Frame7 
         Height          =   90
         Left            =   870
         TabIndex        =   24
         Top             =   1890
         Width           =   12705
      End
      Begin VB.Frame Frame8 
         Height          =   90
         Left            =   30
         TabIndex        =   23
         Top             =   1890
         Width           =   180
      End
      Begin VB.Frame Frame9 
         Caption         =   "Fletes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10425
         TabIndex        =   22
         Top             =   2100
         Width           =   3105
         Begin EditLib.fpDoubleSingle Double1 
            Height          =   330
            Index           =   17
            Left            =   1770
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   195
            Width           =   1215
            _Version        =   196608
            _ExtentX        =   2143
            _ExtentY        =   582
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
            AllowNull       =   0   'False
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
            Text            =   "0"
            DecimalPlaces   =   0
            DecimalPoint    =   "."
            FixedPoint      =   0   'False
            LeadZero        =   0
            MaxValue        =   "9000000000"
            MinValue        =   "-9000000000"
            NegFormat       =   1
            NegToggle       =   0   'False
            Separator       =   ","
            UseSeparator    =   -1  'True
            IncInt          =   1
            IncDec          =   1
            BorderGrayAreaColor=   -2147483637
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
      End
      Begin EditLib.fpText fpText 
         Height          =   330
         Index           =   0
         Left            =   285
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   465
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   582
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
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
         ControlType     =   0
         Text            =   ""
         CharValidationText=   ""
         MaxLength       =   20
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   5
         Left            =   11805
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   465
         Width           =   1200
         _Version        =   196608
         _ExtentX        =   2117
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
         AllowNull       =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "0"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   0
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   1
         Left            =   1845
         TabIndex        =   16
         Top             =   1320
         Visible         =   0   'False
         Width           =   1440
         _Version        =   196608
         _ExtentX        =   2540
         _ExtentY        =   609
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   285
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1470
         _Version        =   196608
         _ExtentX        =   2593
         _ExtentY        =   609
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
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
         Height          =   345
         Index           =   1
         Left            =   6210
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   609
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
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
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   13
         Left            =   1650
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
         AllowNull       =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   14
         Left            =   3000
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
         AllowNull       =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   15
         Left            =   4365
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
         AllowNull       =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   16
         Left            =   5760
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
         AllowNull       =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   ""
         FixedPoint      =   0   'False
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpDoubleSingle Double1 
         Height          =   315
         Index           =   12
         Left            =   285
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1215
         _Version        =   196608
         _ExtentX        =   2143
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
         AllowNull       =   0   'False
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
         Text            =   "0"
         DecimalPlaces   =   0
         DecimalPoint    =   "."
         FixedPoint      =   -1  'True
         LeadZero        =   0
         MaxValue        =   "9000000000"
         MinValue        =   "-9000000000"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ","
         UseSeparator    =   -1  'True
         IncInt          =   1
         IncDec          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin EditLib.fpText fpText 
         Height          =   330
         Index           =   2
         Left            =   285
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   480
         Visible         =   0   'False
         Width           =   1500
         _Version        =   196608
         _ExtentX        =   2646
         _ExtentY        =   582
         Enabled         =   0   'False
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
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
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   2
         Left            =   1845
         TabIndex        =   5
         Top             =   1320
         Width           =   1440
         _Version        =   196608
         _ExtentX        =   2540
         _ExtentY        =   609
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
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483643
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   7800
         Picture         =   "M_DocPro.frx":22D7
         ToolTipText     =   "Calcular Totales"
         Top             =   1230
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   2190
         TabIndex        =   1
         Top             =   465
         Width           =   5415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rut"
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
         Left            =   285
         TabIndex        =   44
         Top             =   225
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
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
         Left            =   8010
         TabIndex        =   43
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "N° Documento"
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
         Index           =   11
         Left            =   11790
         TabIndex        =   42
         Top             =   225
         Width           =   1245
      End
      Begin VB.Label Label1 
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
         Index           =   16
         Left            =   3435
         TabIndex        =   41
         Top             =   1065
         Width           =   660
      End
      Begin VB.Label Label1 
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
         Index           =   14
         Left            =   285
         TabIndex        =   40
         Top             =   1065
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Recep. Merc."
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
         Index           =   13
         Left            =   1845
         TabIndex        =   39
         Top             =   1065
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1755
         Picture         =   "M_DocPro.frx":25E1
         Top             =   405
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   13050
         Picture         =   "M_DocPro.frx":28EB
         Top             =   345
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   8070
         TabIndex        =   38
         Top             =   510
         Width           =   3495
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   3495
         TabIndex        =   37
         Top             =   1365
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Compra"
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
         Left            =   6195
         TabIndex        =   36
         Top             =   1065
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total"
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
         Left            =   5775
         TabIndex        =   35
         Top             =   2160
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Otr. Imp."
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
         Left            =   4350
         TabIndex        =   34
         Top             =   2160
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "I.V.A"
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
         Left            =   3015
         TabIndex        =   33
         Top             =   2160
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Neto"
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
         Left            =   1650
         TabIndex        =   32
         Top             =   2160
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Exento"
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
         Left            =   285
         TabIndex        =   31
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label Label2 
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
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   210
         TabIndex        =   30
         Top             =   1860
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   7110
         Picture         =   "M_DocPro.frx":2BF5
         ToolTipText     =   "Calcular Totales"
         Top             =   2295
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   285
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   1750
         Picture         =   "M_DocPro.frx":2EFF
         Top             =   405
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   2190
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   3
         Left            =   2235
         TabIndex        =   46
         Top             =   525
         Visible         =   0   'False
         Width           =   5415
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   5
         Left            =   2235
         TabIndex        =   45
         Top             =   510
         Width           =   5415
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   15630
      _ExtentX        =   27570
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_DocPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim i As Long
Dim vecdatos(7) As String
Dim Encontrado As Boolean 'Variable para saber si encontro o no el registro
Dim Impuestos() As Variant
Dim indice, FolioSn As Long
Dim IndiceOpt As Double, est As Boolean, est1 As Boolean, est2 As Boolean
Dim vTotExe As Double, vTotNet As Double, vTotIva As Double, vTotOtr As Double, vTotTot As Double
Dim MsgTitulo As String, modo As String

Private Function Fg_Puntocoma(ByVal Parentesis As String) As String

On Error GoTo Man_Error

Dim X%
Dim ValLcntH$
ValLcntH = ""
For X = 1 To Len(Parentesis)
    If Asc(Mid(Parentesis, X, 1)) <> 59 Then
       ValLcntH = ValLcntH + Mid(Parentesis, X, 1)
    End If
Next X
Fg_Puntocoma = ValLcntH

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Private Function Valida_DatosGrilla() As Boolean

On Error GoTo Man_Error

Valida_DatosGrilla = False
With vaSpread1
    
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 1
        If Trim(.text) = "" Then
           
           .DeleteRows .Row, 1
           .MaxRows = .MaxRows - 1
        
        End If
    
    Next i
    
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 4
        If Val(.text) = 0 And vg_FDC <> "OC" Then Valida_DatosGrilla = True: Exit For
        
        .Row = i
        .Col = 5
        If Val(.text) = 0 Then Valida_DatosGrilla = True: Exit For
        
        .Row = i
        .Col = 10
        If Val(.text) = 0 And vg_FDC <> "SN" Then Valida_DatosGrilla = True: Exit For
    
    Next i

End With

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Private Function Nuevo_Registro()

On Error GoTo Nuevo

Dim RS As New ADODB.Recordset
est = True
est2 = False

fpText(0).text = ""
fpayuda(0).Caption = ""

Double1(12).Enabled = True
Double1(13).Enabled = True
Double1(14).Enabled = True
Double1(15).Enabled = True
Double1(17).Enabled = True
Double1(6).Enabled = False
Double1(5).text = ""
Double1(6).text = ""
Double1(12).text = ""
Double1(13).text = ""
Double1(14).text = ""
Double1(15).text = ""
Double1(16).text = ""
Double1(17).text = ""

Date1(0).text = ""
Date1(1).text = ""
Date1(2).text = ""

Image1(5).Visible = False
'Date1(0).text = Date: Date1(1).text = Date: Date1(2).text = Date
If Combo2(1).listcount > 0 Then Combo2(1).ListIndex = 0
vaSpread1.MaxRows = 0
vaSpread1.MaxRows = 1

'-------> Bloquear columna de ordenes de compras
Text2(0).text = ""
Text2(1).text = ""
Text2(2).text = ""
Text2(0).Visible = False
Text2(1).Visible = False
Text2(2).Visible = False

vaSpread1.Col = 10
vaSpread1.ColHidden = False

vaSpread1.Col = 20
vaSpread1.ColHidden = True

vaSpread1.Col = 21
vaSpread1.ColHidden = True

vaSpread1.Col = 22
vaSpread1.ColHidden = True

vaSpread1.Col = 23
vaSpread1.ColHidden = True
vaSpread1.ColWidth(2) = 55.13 '33.13
vaSpread1.ColWidth(4) = 10
vaSpread1.ColWidth(5) = 9.75
vaSpread1.ColWidth(8) = 11.13
vaSpread1.ColWidth(9) = 9.75
vaSpread1.ColWidth(10) = 10.63

'vaSpread2.MaxRows = 0
Frame2.Enabled = True
Frame3.Enabled = True
Frame5.Enabled = True
Frame6.Enabled = True
Combo2(0).Enabled = True
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = False
' jpaz 01.08.2005 vaSpread1.Row = -1: vaSpread1.Col = 8: vaSpread1.Lock = True
'vaSpread2.Row = -1: vaSpread2.Col = -1: vaSpread2.Lock = False
'-------> Mover impuestos
vaSpread2.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT COUNT(imp_codigo) AS mayor FROM a_impuesto", vg_db, adOpenStatic
If Not RS.EOF Then indice = RS!mayor
RS.Close
Set RS = Nothing

ReDim Impuestos(indice, 5)
indice = 1
est1 = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT * FROM a_impuesto", vg_db, adOpenStatic

Do While Not RS.EOF
    
    Impuestos(indice, 1) = RS!imp_codigo
    Impuestos(indice, 2) = RS!imp_nombre
    Impuestos(indice, 3) = IIf(IsNull(RS!imp_pctimp), 0, RS!imp_pctimp)
    Impuestos(indice, 4) = 0
    Impuestos(indice, 5) = IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos)
    
    If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!imp_codigo), SearchFlagsNone) = -1 Then
       
       vaSpread2.MaxRows = vaSpread2.MaxRows + 1
       vaSpread2.Row = vaSpread2.MaxRows
       vaSpread2.Col = 1
       vaSpread2.text = CStr(RS!imp_codigo)
       vaSpread2.Lock = True
       
       vaSpread2.Col = 2
       vaSpread2.text = "1"
       vaSpread2.Lock = False
       
       vaSpread2.Col = 3
       vaSpread2.text = Trim(RS!imp_nombre)
       vaSpread2.Lock = True
       
       vaSpread2.Col = 4
       vaSpread2.text = IIf(IsNull(RS!imp_pctimp), Format(0, fg_Pict(3, 2)), Format(RS!imp_pctimp, fg_Pict(3, 2))) & " %": vaSpread2.Lock = True
       
       vaSpread2.Col = 5
       vaSpread2.text = Format(0, fg_Pict(6, 0))
       vaSpread2.ForeColor = IIf(RS!imp_indmod = "S", &HC00000, &H80000008): vaSpread2.Lock = IIf(RS!imp_indmod = "S", False, True)
       
       vaSpread2.Col = 6
       vaSpread2.text = IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos)
       vaSpread2.Lock = True
       
       vaSpread2.Col = 7
       vaSpread2.text = IIf(IsNull(RS!imp_pctimp), 0, RS!imp_pctimp)
       vaSpread2.Lock = True
       
       vaSpread2.Col = 8
       vaSpread2.text = IIf(RS!imp_indmod = "S", 1, 0)
       vaSpread2.Lock = True
       vaSpread2.RowHidden = True
    
    End If
    
    RS.MoveNext
    
    indice = indice + 1

Loop
RS.Close
Set RS = Nothing

Option1(0).Value = False
Option1(1).Value = False

Image1(1).Visible = False

Encontrado = False
Image1(4).Visible = False
vaSpread1.Enabled = True
Text1(0).Enabled = True

Double1(12).Enabled = True
Double1(13).Enabled = True
Double1(14).Enabled = True
Double1(15).Enabled = True
Double1(16).Enabled = True
Double1(17).Enabled = True

Combo2(0).ListIndex = -1
vaSpread1.SetActiveCell 1, 1
vg_Guias = ""
vg_GuiasTipo = ""
vg_FDC = ""
est = False

Exit Function
Nuevo:
    MsgBox Err.Description, vbOKOnly, MsgTitulo

End Function

Private Function SumaDiferencias(FilaDif As Long, CantDif As Double, PreDif As Double, PorDes As Double)

Dim i As Long, TotE As Double, TotN As Double, TotI As Double, TotO As Double, StrImp As String, StrImpb As String
Dim codi As Long, PctI As Double, CosI As Long, aPos As Long, Cant As Double
Dim PreC As Double, MonD As Double, MonI As Double, Cdif As Long

On Error GoTo Error_Suma

For i = 1 To UBound(Impuestos)
    
    Impuestos(i, 4) = 0

Next i

TotE = 0
TotN = 0
TotI = 0
TotO = 0
Cant = 0
PreC = 0
MonD = 0
MonI = 0
Cdif = 0
StrImp = ""
StrImpb = ""

For i = FilaDif To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    codpro = Trim(vaSpread1.Value)
    
    vaSpread1.Col = 4
    Cant = Val(vaSpread1.Value)
    
    vaSpread1.Col = 16
    Cdif = Val(vaSpread1.Value)
    
    If codpro <> "" And Cdif > 0 Then
       
       PreC = PreDif
       MonD = (CantDif * PreDif) * (PorDes / 100)
       monpro = (CantDif * PreDif) - MonD
       vaSpread1.Col = 14
       StrImp = Trim(vaSpread1.text)
       MonI = 0
       
       If Len(StrImp) <> 0 Then
          
          Do While InStr(StrImp, ";") <> 0
             
             StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
             StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
             codi = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
             PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
             CosI = Val(Mid(StrImpb, 1))
             
             If (ValidarImpuestoAdicional(codi) And codi <> GetParametro("parivacig")) Or codi = GetParametro("parretfue") Or codi = GetParametro("parretica") Or codi = GetParametro("parrethorf") Then
                
                If (codi <> GetParametro("parretfue") Or IsNull(GetParametro("parretfue"))) And (codi <> GetParametro("parretica") Or IsNull(GetParametro("parretfue"))) And (codi <> GetParametro("parrethorf") Or IsNull(GetParametro("parretfue"))) Then
                   
                   If PctI > 0 Then
                      
                      TotN = TotN + Format(monpro, fg_Pict(9, 0))
                   
                   Else
                      
                      TotE = TotE + monpro
                   
                   End If
                
                End If
                
                TotI = TotI + Format((monpro) * (PctI / 100), fg_Pict(9, 0))
             
             Else
                
                TotO = TotO + Format(monpro * (PctI / 100), fg_Pict(9, 0))
             
             End If
             
             If CosI = 1 Then MonI = MonI + Format((PreC - MonD) * (PctI / 100), fg_Pict(9, 0))
          
          Loop
       
       Else
          
          TotE = TotE + monpro
       
       End If
       
       vaSpread1.Col = 15
       vaSpread1.Value = (PreC - MonD) + MonI
    
    End If

Next i

For i = 1 To UBound(Impuestos)
    
    If ValidarImpuestoAdicional(Impuestos(i, 1)) And Impuestos(i, 1) <> Val(GetParametro("parivacig")) Then
       
       TotI = TotI + Format(Impuestos(i, 4) * (Impuestos(i, 3) / 100), fg_Pict(9, 0))
    
    End If

Next i

vTotExe = 0
vTotNet = 0
vTotIva = 0
vTotOtr = 0
vTotTot = 0

If fg_codigocbo(Combo2, 0, 2, "") <> "GD" Then
    
    vTotExe = Round(TotE, 0)
    vTotNet = Round(TotN, 0)
    vTotIva = Round(TotI, 0)
    vTotOtr = Round(TotO, 0)
    vTotTot = Format(TotE + TotN + TotI + TotO, fg_Pict(9, 0))

Else
    
    vTotExe = 0
    vTotNet = 0
    vTotIva = 0
    vTotOtr = 0
    vTotTot = Format(TotE + TotN, fg_Pict(9, 0))

End If

Exit Function
Error_Suma:
MsgBox "Error : " & Err.Number & " - " & Err.Description, vbExclamation, MsgTitulo
Resume Next

End Function

Public Function SumarTotales()

Dim i As Long, TotE As Double, TotN As Double, TotI As Double, TotO As Double, StrImp As String, StrImpb As String
Dim codi As Long, PctI As Double, CosI As Long, aPos As Long, Cant As Double, DifUni As Double, MonTot As Double
Dim PreC As Double, MonD As Double, MonI As Double, prefun As Double

On Error GoTo Error_Suma

TotE = 0
TotN = 0
TotI = 0
TotO = 0
PctI = 0
CosI = 0
Cant = 0
DifUni = 0
MonTot = 0
PreC = 0
MonD = 0
monpro = 0
MonI = 0
StrImp = ""
StrImpb = ""

For i = 1 To UBound(Impuestos)
    
    Impuestos(i, 4) = 0

Next i

For i = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = i
    vaSpread1.Col = 1
    codpro = Trim(vaSpread1.Value)
    
    If codpro <> "" Or IsNull(codpro) = False Then
        
        vaSpread1.Col = 4
        Cant = Val(vaSpread1.Value)
        
        vaSpread1.Col = 5
        PreC = Val(vaSpread1.Value)
        
        vaSpread1.Col = 7
        MonD = Val(vaSpread1.Value)
        
        vaSpread1.Col = 8
        monpro = Val(vaSpread1.Value)
        
        vaSpread1.Col = 14
        DifUni = 0
        
        If Val(Cant) > 0 And Val(MonD) > 0 Then
        
           DifUni = MonD / Cant
        
        End If
        
        StrImp = Trim(vaSpread1.text)
        MonI = 0
        
        '------- calcula otros impuestos
        If Len(StrImp) <> 0 Then
            
            Do While InStr(StrImp, ";") <> 0
                
                StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                codi = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                CosI = Val(Mid(StrImpb, 1))
                
                If (ValidarImpuestoAdicional(codi) And codi <> GetParametro("parivacig")) Or codi = GetParametro("parretfue") Or codi = GetParametro("parretica") Or codi = GetParametro("parrethorf") Then
                   
                   If (codi <> GetParametro("parretfue") Or IsNull(GetParametro("parretfue"))) And (codi <> GetParametro("parretica") Or IsNull(GetParametro("parretica"))) And (codi <> GetParametro("parrethorf") Or IsNull(GetParametro("parrethorf"))) Then
                      
                      If PctI > 0 Then
                         
                         TotN = TotN + Format(monpro, fg_Pict(9, 0))
                      
                      Else
                         
                         TotE = TotE + monpro
                      
                      End If
                   
                   End If
                   TotI = TotI + Round((monpro) * (PctI / 100), 2)
                
                Else
                   
                   If codi = GetParametro("parivacig") And (fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "FA" Or fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "FE" Or fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "NC" Or fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "CE" Or fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "ND" Or fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "DE") Then
                      
                      TotE = TotE + monpro
                   
                   ElseIf codi = GetParametro("parivacig") And fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "GD" Then
                      
                      TotN = TotN + Format(monpro, fg_Pict(9, 0))
                   
                   End If
                   
                   If PctI > 0 Then
                      
                      TotO = TotO + IIf(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "BH", Round(((monpro / ((100 - PctI) / 100))) / PctI), Round(monpro * (PctI / 100), 2))
                   
                   Else
                      
                      TotO = TotO + 0
                   
                   End If
                
                End If
                
                If CosI = 1 Then MonI = MonI + Format((PreC - DifUni) * (PctI / 100), fg_Pict(9, 0))
            
            Loop
        
        Else
            
            TotE = TotE + monpro
        
        End If
        
        vaSpread1.Col = 15
        vaSpread1.Value = (PreC - DifUni) + MonI

'------- fin calcula otros impuestos
    End If

Next i

For i = 1 To UBound(Impuestos)
    
    If ValidarImpuestoAdicional(Impuestos(i, 1)) And Impuestos(i, 1) <> Val(GetParametro("parivacig")) Then
       
       TotI = TotI + Round((Impuestos(i, 4) + IIf(Double1(17).Value > 0, Double1(17).Value, 0)) * (Impuestos(i, 3) / 100), 2)
    
    End If

Next i
vTotExe = 0
vTotNet = 0
vTotIva = 0
vTotOtr = 0
vTotTot = 0

If fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) <> "GD" And fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) <> "BO" And fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) <> "CG" Then
   
   If Double1(12).Enabled = False Or Double1(13).Enabled = False Then
      
      Double1(12).Value = Format(TotE, fg_Pict(9, 0))
      Double1(13).Value = Format(TotN, fg_Pict(9, 0))
      Double1(14).Value = Format(TotI, fg_Pict(9, 0))
      Double1(15).Value = Format(TotO, fg_Pict(9, 0))
      Double1(16).Value = Format(TotE + TotN + TotI + TotO + IIf(Double1(17).Value > 0, Double1(17).Value, 0), fg_Pict(9, 0))
      
      TotE = Format(TotE, fg_Pict(9, 0))
      TotN = Format(TotN, fg_Pict(9, 0))
      TotI = Format(TotI, fg_Pict(9, 0))
      TotO = Format(TotO, fg_Pict(9, 0))
   
   End If
   
   vTotExe = Format(TotE, fg_Pict(9, 0))
   vTotNet = Format(TotN, fg_Pict(9, 0))
   vTotIva = Round(IIf(fg_codigocbo(Combo2, 0, 2, "") = "BH", 0, TotI), 0)
   vTotOtr = Format(TotO, fg_Pict(9, 0))
   vTotTot = Format(TotE + TotN + IIf(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "BH", 0, Round(TotI, 0)) + Round(TotO, 0) + IIf(Double1(17).Value > 0, Double1(17).Value, 0), fg_Pict(9, 0))

Else
   
   vTotExe = Round(IIf(Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "GD", TotE, 0), 0)
   vTotNet = Round(IIf(Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "GD", TotN, 0), 0)
   vTotIva = 0
   vTotOtr = 0
   vTotTot = Format(TotE + TotN, fg_Pict(9, 0))

End If

Exit Function
Error_Suma:
MsgBox "Error : " & Err.Number & " " & Err.Description, vbExclamation, MsgTitulo
Resume Next

End Function

Private Sub Combo2_Click(Index As Integer)

On Error GoTo Man_Error

If est Or Frame6.Enabled = False Then Exit Sub
Dim RS   As New ADODB.Recordset
Dim sql1 As String, sql2 As String, sql3 As String

vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = False

Double1(12).Enabled = True
Double1(13).Enabled = True
Double1(14).Enabled = True
Double1(15).Enabled = True
Double1(17).Enabled = True
Double1_Change 17 '5

Frame3.Enabled = True
Frame6.Enabled = True
Frame5.Enabled = True
Image1(1).Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FA" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Then
    
    If Trim(vg_FDC) = "" Then vg_FDC = "GD"
    Set RS = vg_db.Execute("sgp_Sel_ValidaCfCGuia '" & fg_DespintaRut(fpText(0).text) & "', " & vg_codbod & ", '" & vg_FDC & "'")
    Image1(1).Visible = IIf(RS.EOF Or RS!nreg = 0 Or IsNull(RS!nreg), False, True)
    RS.Close
    Set RS = Nothing
    ChequearOCompras

ElseIf Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "NC" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "CE" Then
   
   If Trim(vg_FDC) = "" Then vg_FDC = "SN"
    
'   RS.Open "SELECT COUNT(toc_numdoc) AS nreg FROM b_totcompras WHERE ltrim(toc_docaso) IN (SELECT ltrim(toc_numdoc) FROM b_totcompras WHERE (toc_tipdoc = 'FA' OR toc_tipdoc = 'FE') AND (ltrim(toc_docsnc) = '' OR toc_docaso IS NULL)) AND  toc_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' AND toc_tipdoc = '" & Trim(vg_FDC) & "' AND (ltrim(toc_docsnc) = '' OR (toc_docsnc) IS NULL) AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
   Set RS = vg_db.Execute("sgp_Sel_ValidaCfCSolicitudNota '" & fg_DespintaRut(fpText(0).text) & "', " & vg_codbod & ", '" & vg_FDC & "'")
   Image1(1).Visible = IIf(RS.EOF Or RS!nreg = 0 Or IsNull(RS!nreg), False, True)
   RS.Close
   Set RS = Nothing

End If

Select Case fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))

    Case Is = "FA", "FE", "NC", "CE", "ND", "DE" '-------> FACTURA(FA), NOTA DE CREDITO(NC), NOTA DE DEBITO(ND)
        
        Combo2(1).Enabled = False
        fpText(1).Enabled = True
    
    Case Is = "GD" '-------> GUIA DE DEPACHO(GD)
        
        'Double1(12).Enabled = False:
        Double1(12).Enabled = True
        Double1(13).Enabled = True
        Double1(14).Enabled = False
        Double1(15).Enabled = False
        Double1(17).Enabled = False
        Double1(12).text = ""
        Double1(14).text = ""
        Double1(15).text = ""
        Double1(17).text = ""
        Combo2(1).Enabled = False
        fpText(1).Enabled = True
    
    Case Is = "BH" '-------> BOLETA DE HONORARIOS(BH)
        
        Combo2(1).Enabled = False: fpText(1).Enabled = True
    
    Case Is = "BO", "CG" ' Boleta(BO), Comprobante de Gasto(CG)
        
        Combo2(1).Enabled = False
        fpText(1).Enabled = False
        Double1(12).Enabled = False
        Double1(13).Enabled = False
        Double1(14).Enabled = False
        Double1(15).Enabled = False
        Double1(17).Enabled = False
        Double1(12).text = ""
        Double1(13).text = ""
        Double1(14).text = ""
        Double1(15).text = ""
        Double1(17).text = ""
    
    End Select
    
    Select Case fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))
    
        Case Is = "FA", "FE", "BO", "BH"
            
            Option1(0).Enabled = True
            Option1(1).Enabled = True
            Option1(1).Value = True
        
        Case Is = "GD", "NC", "CE", "ND", "DE"
            
            Option1(0).Enabled = False
            Option1(0).Value = False
            Option1(1).Enabled = True
            Option1(1).Value = True
        
        Case Is = "CG"
            
            Option1(0).Enabled = True
            Option1(0).Value = True
            Option1(1).Enabled = False
            Option1(1).Value = False

End Select

ValidarCfc
fg_descarga
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub ChequearOCompras()

On Error GoTo Man_Error

Dim RS   As New ADODB.Recordset
Dim RS1  As New ADODB.Recordset
Dim sql1 As String, sql2 As String, sql3 As String

'-------> Activar o desactivar ordenes de compras
Image1(4).Enabled = True
sql1 = IIf(vg_tipbase = "1", " val(format(a.solite_dtent, 'yyyymm')) ", " substring(CONVERT(varchar(10), a.solite_dtent,112),1,6) ")
sql2 = IIf(vg_tipbase = "1", " '" & Format(Date1(0), "yyyymm") & "' ", " '" & Format(Date1(0), "yyyymm") & "' ")
sql3 = IIf(vg_tipbase = "1", " SUM(IIF(a.tipsol_idsol = 4,(-1 * a.pedite_qtcpa), a.pedite_qtcpa)) AS difer ", " SUM(CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE  a.pedite_qtcpa END) AS difer ")

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open "SELECT " & sql3 & " " & _
        "FROM b_ocsac a " & _
        "WHERE a.cadfor_nrcgc = '" & fg_DespintaRut(fpText(0).text) & "' " & _
        "AND   a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
        "AND   " & sql1 & "   = " & sql2 & " AND a.pedite_flafo = 1", vg_db, adOpenStatic

If Not RS.EOF And Not IsNull(RS!difer) Then
   
   sql3 = IIf(vg_tipbase = "1", " SUM(b.ocr_cancom - (IIF(a.tipsol_idsol = 4,(-1 * a.pedite_qtcpa),a.pedite_qtcpa)) ) AS difer ", " SUM(b.ocr_cancom - (CASE WHEN a.tipsol_idsol = 4 THEN (-1 * a.pedite_qtcpa) ELSE a.pedite_qtcpa END)) AS difer ")
   
   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS1.Open "SELECT " & sql3 & " " & _
            "FROM   b_ocsac a, b_ocsacrecibido b " & _
            "WHERE  a.cadfor_nrcgc = b.ocr_rutpro " & _
            "AND    a.solite_dtent = b.ocr_fecoc  " & _
            "AND    a.cadfil_cdfil = '" & MuestraCasino(1) & "' " & _
            "AND    " & sql1 & "   = " & sql2 & " " & _
            "AND    a.cadfor_nrcgc = '" & fg_DespintaRut(fpText(0).text) & "' AND a.cpopro_cdpro = b.ocr_codprodsac AND a.pedite_flafo = 1", vg_db, adOpenStatic
   Image1(4).Visible = IIf(RS1.EOF Or RS1!difer <> 0 Or IsNull(RS1!difer) Or RS!difer > 0, True, False)
   RS1.Close
   Set RS1 = Nothing
   
End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Date1_Change(Index As Integer)

On Error GoTo Man_Error

'If Not IsDate(Date1(0).text) Or Not IsDate(Date1(1).text) Or Not IsDate(Date1(2).text) Then Exit Sub
If Not IsDate(Date1(0).text) Or Not IsDate(Date1(2).text) Then Exit Sub
Date1(1).text = Date1(0).text
ChequearOCompras

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Date1_GotFocus(Index As Integer)

On Error GoTo Man_Error

StopObjeto True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Date1_LostFocus(Index As Integer)

On Error GoTo Man_Error

'-------> Valida primer campo de fecha
If Not IsDate(Date1(0).text) Or Not IsDate(Date1(1).text) Or Not IsDate(Date1(2).text) Then Exit Sub
'If Date1(0).text = "" Then Date1(0).text = Date
'-------> Valida segundo campo fecha
'If Date1(1).text = "" Then Date1(1).text = Date
'-------> Valida tercera campo fecha
'If Date1(2).text = "" Then Date1(2).text = Date

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Double1_Change(Index As Integer)

On Error GoTo Man_Error

If Index = 17 Then SumarTotales
'amorgado 20100205 If index = 13 And Trim(fg_codigocbo(Combo2, 0, 2, "")) = "GD" Then Double1(16).Value = Double1(13).Value
'amorgado 20100205 If index = 16 And Trim(fg_codigocbo(Combo2, 0, 2, "")) = "GD" Then Double1(13).Value = Double1(16).Value

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Double1_GotFocus(Index As Integer)

On Error GoTo Man_Error

StopObjeto True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Double1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Double1_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim tipord As String
If est Then Exit Sub

Dim RS  As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

If Index = 17 Then SumarTotales
If Index <> 5 Then Exit Sub
If Index = 5 And Val(Double1(5).text) = 0 Then Exit Sub
est = True
tipord = ""

'mod 20130401 sacar la bodega        "AND toc_codbod = " & vg_codbod & " AND toc_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' " & _

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
RS.Open "SELECT * FROM  b_totcompras WHERE toc_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' " & _
        "AND toc_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' " & _
        "AND toc_numdoc = " & Val(Double1(5).Value) & "", vg_db, adOpenStatic
If RS.EOF Then '-------> Si no existe el documento

    If modo <> "A" Then modo = "A"

    vg_RDC = fg_DespintaRut(fpText(0).text)
    '-------> Validar si proveedor puede ingresar documento
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS1 = vg_db.Execute("select isnull(prv_permiteingdoc,0) AS prv_permiteingdoc FROM b_proveedor where prv_codigo = '" & vg_RDC & "'")
    If Not RS1.EOF Then
       
       If RS1(0) = False Then
          
          RS1.Close
          Set RS1 = Nothing
          MsgBox "Proveedor esta bloqueado para el ingreso documento...", vbCritical, MsgTitulo
          Nuevo_Registro
          Gl_Ac_Botones Me, 7, 2, ""
          Exit Sub
       
       End If
    
    End If
    RS1.Close
    Set RS1 = Nothing
    Gl_Ac_Botones Me, 7, 3, ""

Else
    
    Encontrado = True
    modo = "M":
    Gl_Ac_Botones Me, 7, 2, ""
    '-------> Deshabilitar Botones
    Date1(0).text = RS!toc_fecemi
    Date1(1).text = RS!toc_fecven
    Date1(2).text = RS!toc_fecrem
    Double1(12).text = RS!toc_exedoc
    Double1(13).text = RS!toc_netdoc
    Double1(14).text = RS!toc_ivadoc
    Double1(15).text = RS!toc_otrimp
    Double1(16).text = RS!toc_totdoc
    Double1(17).text = RS!toc_fledoc
    
    If RS!toc_tipinf = "C" Or RS!toc_tipinf = "P" Then
        
        Double1(6).text = RS!toc_numinf
        Option1(1).Value = True
    
    ElseIf RS!toc_tipinf = "F" Then
        
        Double1(6).text = RS!toc_numinf
        Option1(0).Value = True
    
    End If
    Combo2(0).ListIndex = fg_buscacbostring(Combo2, 0, 2, (RS!toc_tipdoc))
    Combo2(1).ListIndex = fg_buscacbo(Combo2, 1, 10, fg_pone_cero(Str(RS!toc_codbod), 10))
    Frame2.Enabled = False
    Frame3.Enabled = False
    Frame6.Enabled = False
    vaSpread1.Row = -1
    vaSpread1.Col = -1
    vaSpread1.Lock = True
    
    vaSpread2.Row = -1
    vaSpread2.Col = -1
    vaSpread2.Lock = True
    Image1(4).Enabled = False
    
    '-------> Detalle de Documento
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    RS1.Open "SELECT DISTINCT ocr_rutpro FROM b_ocsacrecibido WHERE ocr_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' AND ocr_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' AND ocr_numdoc = " & Val(Double1(5).Value) & "", vg_db, adOpenStatic
    If Not RS1.EOF Then
       
       RS1.Close
       Set RS1 = Nothing
       '-------> Validar si los datos bienen desde la ordene de compras
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       RS1.Open "SELECT DISTINCT ocr_rutpro FROM b_ocsacrecibido WHERE ocr_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' AND ocr_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' AND ocr_numdoc = " & Val(Double1(5).Value) & " AND ocr_canoc > 0 ", vg_db, adOpenStatic
       If Not RS1.EOF Then
          
          vaSpread1.ColWidth(2) = 42.13 '41.13
          vaSpread1.ColWidth(4) = 9 '8
          vaSpread1.ColWidth(5) = 8.75 '7.75
          vaSpread1.ColWidth(8) = 10.13 '9.13
          vaSpread1.ColWidth(9) = 8.75 '7.75
          vaSpread1.ColWidth(10) = 9.63 '8.63
          vaSpread1.ColWidth(21) = 9.63
          vaSpread1.ColWidth(22) = 9.63
          vaSpread1.ColWidth(23) = 7.99
          vaSpread1.Col = 20: vaSpread1.ColHidden = True
          vaSpread1.Col = 21: vaSpread1.ColHidden = False
          vaSpread1.Col = 22: vaSpread1.ColHidden = False
          vaSpread1.Col = 23: vaSpread1.ColHidden = False
          vaSpread1.Col = 10: vaSpread1.ColHidden = True
       
       Else
          
          vaSpread1.Col = 10: vaSpread1.ColHidden = False
          vaSpread1.Col = 20: vaSpread1.ColHidden = True
          vaSpread1.Col = 21: vaSpread1.ColHidden = True
          vaSpread1.Col = 22: vaSpread1.ColHidden = True
          vaSpread1.Col = 23: vaSpread1.ColHidden = True
          vaSpread1.ColWidth(2) = 55.13 '33.13
          vaSpread1.ColWidth(4) = 10
          vaSpread1.ColWidth(5) = 9.75
          vaSpread1.ColWidth(8) = 11.13
          vaSpread1.ColWidth(9) = 9.75
          vaSpread1.ColWidth(10) = 10.63
       
       End If
       RS1.Close
       Set RS1 = Nothing
       Text2(0).Visible = True
       Text2(1).Visible = True
       Text2(2).Visible = True
       tipord = "OC"
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       RS1.Open "SELECT DISTINCT a.*, b.*, c.uni_nomcor, " & _
                "(SELECT DISTINCT ocr_canoc FROM b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer) AS ocr_canoc, " & _
                "(SELECT DISTINCT ocr_preoc FROM b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer) AS ocr_preoc, " & _
                "(SELECT DISTINCT ocr_fecoc FROM b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer) AS ocr_fecoc, " & _
                "(SELECT DISTINCT ocr_codprodsac FROM b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer) AS ocr_codprodsac, " & _
                "(SELECT DISTINCT e.foc_nomsac FROM b_formatocompras e, b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer AND ocr_codprodsac = e.foc_codsac) AS foc_nomsac, " & _
                "(SELECT DISTINCT e.foc_unisac FROM b_formatocompras e, b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer AND ocr_codprodsac = e.foc_codsac) AS foc_unisac, " & _
                "(SELECT DISTINCT e.foc_faccon FROM b_formatocompras e, b_ocsacrecibido WHERE ocr_rutpro = a.dec_rutpro AND ocr_tipdoc = a.dec_tipdoc AND ocr_numdoc = a.dec_numdoc AND ocr_numlin = a.dec_numlin AND ocr_codprodsgp = a.dec_codmer AND ocr_codprodsac = e.foc_codsac) AS foc_faccon " & _
                "FROM b_detcompras a, b_productos b, a_unidad c " & _
                "WHERE a.dec_codmer = b.pro_codigo " & _
                "AND   b.pro_coduni = c.uni_codigo " & _
                "AND   a.dec_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' " & _
                "AND   a.dec_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' " & _
                "AND   a.dec_numdoc = " & Val(Double1(5).Value) & " " & _
                "ORDER BY a.dec_numlin", vg_db, adOpenStatic
                
    Else
    
       RS1.Close
       Set RS1 = Nothing
       
       vaSpread1.Col = 10
       vaSpread1.ColHidden = False
       
       vaSpread1.Col = 20
       vaSpread1.ColHidden = True
       
       vaSpread1.Col = 21
       vaSpread1.ColHidden = True
       vaSpread1.Col = 22
       vaSpread1.ColHidden = True
       
       vaSpread1.Col = 23
       vaSpread1.ColHidden = True
       
       vaSpread1.ColWidth(2) = 55.13 '33.13
       vaSpread1.ColWidth(4) = 10
       vaSpread1.ColWidth(5) = 9.75
       vaSpread1.ColWidth(8) = 11.13
       vaSpread1.ColWidth(9) = 9.75
       vaSpread1.ColWidth(10) = 10.63
       tipord = ""
       
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       RS1.Open "SELECT a.*, b.*, c.uni_nomcor, '' AS ocr_codprodsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon FROM b_detcompras a, b_productos b, a_unidad c WHERE a.dec_codmer = b.pro_codigo AND b.pro_coduni = c.uni_codigo AND a.dec_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' AND a.dec_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' AND a.dec_numdoc = " & Val(Double1(5).Value) & " ORDER BY a.dec_numlin", vg_db, adOpenStatic
    
    End If
    With vaSpread1
         
         .Visible = False
         .MaxRows = 0
         
         Do While Not RS1.EOF
            
            .MaxRows = .MaxRows + 1
            .Col = 1
            .Row = .MaxRows: .Value = IIf(vg_pais = "CO" And tipord = "OC", RS1!ocr_codprodsac, RS1!dec_codmer)
            
            .Col = 2
            .Value = IIf(vg_pais = "CO" And tipord = "OC", Trim(RS1!foc_nomsac), Trim(RS1!pro_nombre))
            
            .Col = 3
            .Value = IIf(vg_pais = "CO" And tipord = "OC", Trim(RS1!foc_unisac), Trim(RS1!uni_nomcor))
            
            .Col = 4
            .Value = IIf(vg_pais <> "CO", RS1!dec_canmer, RS1!dec_cmefac)
            
            .Col = 5
            .Value = IIf(vg_pais <> "CO", RS1!dec_precom, RS1!dec_pmefac)
            
            .Col = 6
            .Value = RS1!dec_pctdes
            
            .Col = 7
            .Value = RS1!dec_valdes
            
            .Col = 8
            .Value = RS1!dec_ptotal
            
            .Col = 9
            .Value = IIf(vg_pais <> "CO", RS1!dec_canrec, RS1!dec_crefac)
            
            .Col = 10
            .Value = IIf(vg_pais <> "CO", RS1!dec_prerec, RS1!dec_prefac)
            
            .Col = 11
            .Value = RS1!dec_descri
            '20080417
            Text1(0).text = RS1!dec_descri
            .Col = 13
            .Value = RS1!dec_mueinv
            
            .Col = 31
            .Value = RS1!dec_canrec
            
            If tipord = "OC" Then
               
               .Col = 21
               .Value = Format(IIf(IsNull(RS1!ocr_canoc), 0, RS1!ocr_canoc), fg_Pict(9, vg_DCa))
               
               .Col = 22
               .Value = Format(IIf(IsNull(RS1!ocr_preoc), 0, RS1!ocr_preoc), fg_Pict(9, 2))
               
               .Col = 23
               .Value = IIf(IsNull(RS1!ocr_fecoc), "", RS1!ocr_fecoc)
               
               .Col = 24
               .Value = IIf(vg_pais = "CO", IIf(IsNull(RS1!dec_codmer), "", RS1!dec_codmer), IIf(IsNull(RS1!ocr_codprodsac), "", RS1!ocr_codprodsac))
               
               .Col = 25
               .Value = IIf(vg_pais = "CO", IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre)), IIf(IsNull(RS1!foc_nomsac), "", Trim(RS1!foc_nomsac)))
               
               .Col = 29
               .Value = IIf(IsNull(RS1!foc_faccon), 0, (RS1!foc_faccon))
               
               If .MaxRows = 1 Then
                  
                  Text2(0).text = IIf(vg_pais = "CO", IIf(IsNull(RS1!dec_codmer), "", RS1!dec_codmer), IIf(IsNull(RS1!ocr_codprodsac), "", RS1!ocr_codprodsac))
                  Text2(1).text = IIf(vg_pais = "CO", IIf(IsNull(RS1!pro_nombre), "", Trim(RS1!pro_nombre)), IIf(IsNull(RS1!foc_nomsac), "", Trim(RS1!foc_nomsac)))
                  
                  Text2(2).text = IIf(vg_pais = "CO", IIf(IsNull(RS1!foc_faccon), 0, (RS1!foc_faccon)), 1)
               
               End If
            
            End If
            
            RS1.MoveNext
         
         Loop
         .Visible = True
         
    End With
    RS1.Close
    Set RS1 = Nothing
    Dim codpro As String
    codpro = ""
    est1 = True
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    RS1.Open "SELECT a.*, b.imp_inccos FROM b_detcomprasimp a, a_impuesto b WHERE a.imd_codimp = b.imp_codigo AND a.imd_rutdoc = '" & fg_DespintaRut(fpText(0).text) & "' AND a.imd_tipdoc = '" & fg_codigocbo(Combo2, 0, 2, "") & "' AND a.imd_numdoc = " & Val(Double1(5).Value) & " ORDER BY a.imd_codpro, a.imd_numlin", vg_db, adOpenStatic
    If Not RS1.EOF Then
        
        Do While Not RS1.EOF
           
           If vaSpread1.SearchCol(IIf(vg_pais = "CO", 24, 1), 0, vaSpread1.MaxRows, RS1!imd_codpro, SearchFlagsNone) <> -1 Then
              
              vaSpread1.Row = RS1!imd_numlin
              vaSpread1.Col = 14
              vaSpread1.text = vaSpread1.text & Trim(Str(RS1!imd_codimp)) & "&"
              
              vaSpread1.Col = 14
              vaSpread1.text = vaSpread1.text & Trim(Str(RS1!imd_pctimp)) & "&"
              
              vaSpread1.Col = 14
              vaSpread1.text = vaSpread1.text & Trim(Str(RS1!imp_inccos)) & ";"
           
           End If
           
           RS1.MoveNext
        
        Loop
    
    End If
    
    RS1.Close
    Set RS1 = Nothing
    
    '******Fin detalle de Documento
    If vaSpread1.MaxRows > 0 Then
       
       vaSpread1.Row = 1
       vaSpread1.Col = 24
       Text2(0).text = Trim(vaSpread1.text)
       vaSpread1.Row = 1: vaSpread1.Col = 25
       Text2(1).text = Trim(vaSpread1.text)
       vaSpread1.Row = 1: vaSpread1.Col = 29
       Text2(2).text = Trim(vaSpread1.text)
    
    End If

End If
RS.Close
Set RS = Nothing
est = False
est1 = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

Dim v_rut As String, v_bodega  As Long
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

'*** Creador : MSP
'*** Fecha   : 05-08-2004
Me.Width = 15750
Me.Height = 9078
Me.HelpContextID = vg_OpcM

EspFecha Date1(0)
EspFecha Date1(1)
EspFecha Date1(2)

MsgTitulo = "Documento de Proveedor"

fpText(2).Enabled = ModCasino
Image1(3).Enabled = ModCasino
fpText(2).text = MuestraCasino(1)
fpayuda(2).Caption = MuestraCasino(2)
'vaSpread1.Col = 8
'vaSpread1.Row = -1
'vaSpread1.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
'vaSpread1.CellNote = "Para modificar Total presione click derecho"
est = False
est1 = False
est2 = False

StopObjeto True
fg_centra Me

vaSpread1.Row = -1
vaSpread1.Col = 4
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread1.Col = 5
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

vaSpread1.Col = 6
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

vaSpread1.Col = 7
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

vaSpread1.Col = 8
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

vaSpread1.Col = 9
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread1.Col = 10
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = 2 'vg_DCa

vaSpread1.Col = 30
vaSpread1.TypeNumberSeparator = vg_CSep
vaSpread1.TypeNumberDecimal = vg_CDec
vaSpread1.TypeNumberDecPlaces = vg_DCa

vaSpread2.MaxRows = 0

Double1(12).UseSeparator = True
Double1(13).UseSeparator = True
Double1(14).UseSeparator = True
Double1(15).UseSeparator = True
Double1(16).UseSeparator = True
Double1(17).UseSeparator = True
Double1(12).DecimalPoint = vg_CDec
Double1(12).Separator = vg_CSep
Double1(12).DecimalPlaces = vg_DPr

Double1(13).DecimalPoint = vg_CDec
Double1(13).Separator = vg_CSep
Double1(13).DecimalPlaces = vg_DPr

Double1(14).DecimalPoint = vg_CDec
Double1(14).Separator = vg_CSep
Double1(14).DecimalPlaces = vg_DPr

Double1(15).DecimalPoint = vg_CDec
Double1(15).Separator = vg_CSep
Double1(15).DecimalPlaces = vg_DPr

Double1(16).DecimalPoint = vg_CDec
Double1(16).Separator = vg_CSep
Double1(16).DecimalPlaces = vg_DPr

Double1(17).DecimalPoint = vg_CDec
Double1(17).Separator = vg_CSep
Double1(17).DecimalPlaces = vg_DPr

'-------> Cargar Combo tipo documento
CargarDatoCombo Combo2, 0, "a_tipodocumento", "tdo_", "Gen", "A"

'-------> Cargar Combo Bodega
CargarDatoCombo Combo2, 1, "b_clientes", "cli_", "CliBod", "N"

vaSpread1.Refresh
Gl_Mo_Botones Me, 7
Gl_Ac_Botones Me, 7, 1, ""
vg_Guias = ""
vg_GuiasTipo = ""
Nuevo_Registro
modo = "N"
modo = "A"
indice = 0

'-------> Traer fecha cierre día
TraerFechaCierre

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState = 2 Then
    
    Frame2.Left = (Me.Width \ 2) - (Frame2.Width \ 2)
    Frame4.Left = (Me.Width \ 2) - (Frame4.Width \ 2)

Else
    
    Frame2.Left = 45
    Frame4.Left = 45

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Select Case Index

    Case 0
        
        fpayuda(Index).Caption = ""
    
    Case 2
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(2).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
        fpayuda(0).Caption = Trim(RS!cli_nombre)
        RS.Close
        Set RS = Nothing

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_GotFocus(Index As Integer)

On Error GoTo Man_Error

StopObjeto True

Select Case Index
    
    Case 0
        
        If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
        fpText(0).text = fg_DespintaRut(fpText(0).text)
        fpText(0).text = Mid(fpText(0).text, 1, Len(Trim(fpText(0).text)) - 1)

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
If Index = 1 And Frame5.Enabled = True Then Option1(1).SetFocus
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText_LostFocus(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Select Case Index
    
    Case 0
        
        If fpText(0).text = "" Then Exit Sub
        fpText(0).text = fg_RutDig(Trim(fpText(0).text))
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        RS.Open "SELECT * FROM b_proveedor WHERE prv_codigo = '" & Trim(fpText(0).text) & "'", vg_db, adOpenStatic
        
        If Not RS.EOF Then
            
            fpText(0).text = fg_PintaRut(fpText(0).text)
            fpayuda(0).Caption = RS!prv_nombre
            Double1_Change 5
            Image1(1).Visible = False
            
            If Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FA" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Then
               
               If Trim(vg_FDC) = "" Then vg_FDC = "GD"
               
               If RS1.State = 1 Then RS1.Close
               RS1.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient

               Set RS1 = vg_db.Execute("sgp_Sel_ValidaCfCGuia '" & fg_DespintaRut(fpText(0).text) & "', " & vg_codbod & ", '" & vg_FDC & "'")
               Image1(1).Visible = IIf(RS1.EOF Or RS1!nreg = 0 Or IsNull(RS1!nreg), False, True)
               RS1.Close
               Set RS1 = Nothing
            
            ElseIf Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "NC" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "CE" Then
               
               If Trim(vg_FDC) = "" Then vg_FDC = "SN"
               
               If RS1.State = 1 Then RS1.Close
               RS1.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
                
               Set RS1 = vg_db.Execute("sgp_Sel_ValidaCfCSolicitudNota '" & fg_DespintaRut(fpText(0).text) & "', " & vg_codbod & ", '" & vg_FDC & "'")
                              
               Image1(1).Visible = IIf(RS1.EOF Or RS1!nreg = 0 Or IsNull(RS1!nreg), False, True)
               RS1.Close
               Set RS1 = Nothing
            
            End If
        Else
           
           RS.Close
           Set RS = Nothing
           MsgBox "Proveedor no existe...", vbCritical, MsgTitulo
           fpText(0).text = ""
           fpayuda(0).Caption = ""
           Exit Sub
        
        End If
        RS.Close
        Set RS = Nothing
       
       '------> Revisar si el usuario cambia de proveedor
        If vaSpread1.MaxRows < 1 Then Exit Sub
        Dim cPro As String
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            cPro = vaSpread1.text
            
            If Trim(vaSpread1.text) <> "" Then
               
               Revisa cPro, vaSpread1.Row
               vaSpread1_EditChange i, 4
            
            End If
        
        Next i
        vaSpread2.Refresh
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error Resume Next

Dim i As Long, sql1 As String
Dim RS As New ADODB.Recordset
vg_codigo = 0

Select Case Index

    Case 0
        
        vg_left = fpayuda(Index).Left + 1920
        B_TabEst.LlenaDatos "b_proveedor", "prv_", "Proveedor", "Proveedor"
        B_TabEst.Show 1, Me
        Me.Refresh
        If Trim(vg_codigo) = "" Or Val(vg_codigo) = 0 Then Exit Sub
        fpText(Index).text = fg_PintaRut(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        Double1_Change 5
        Image1(1).Visible = False
        sql1 = IIf(vg_tipbase = "1", "  ", "  ")
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        If Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FA" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Then
           
           vg_FDC = "GD"
           'RS.Open "SELECT COUNT(toc_numdoc) AS nreg FROM b_totcompras WHERE toc_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' AND (toc_docaso = '' OR toc_docaso IS NULL) AND toc_tipdoc = '" & vg_FDC & "' AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
           Set RS = vg_db.Execute("sgp_Sel_ValidaCfCGuia '" & fg_DespintaRut(fpText(0).text) & "', " & vg_codbod & ", '" & vg_FDC & "'")
           Image1(1).Visible = IIf(RS.EOF Or RS!nreg = 0 Or IsNull(RS!nreg), False, True)
           RS.Close
           Set RS = Nothing
        
        ElseIf Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "NC" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "CE" Then
           
           vg_FDC = "SN"
           'RS.Open "SELECT COUNT(toc_numdoc) AS nreg FROM b_totcompras WHERE trim(toc_docaso) IN (SELECT trim(toc_numdoc) FROM b_totcompras WHERE (toc_tipdoc = 'FA' OR toc_tipdoc = 'FE') AND (toc_docsnc = '' OR toc_docsnc IS NULL) AND toc_codbod = " & vg_codbod & ") AND toc_rutpro = '" & fg_DespintaRut(fpText(0).text) & "' AND toc_tipdoc = '" & Trim(vg_FDC) & "' AND (toc_docsnc = '' OR (toc_docsnc) IS NULL) AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
           Set RS = vg_db.Execute("sgp_Sel_ValidaCfCSolicitudNota '" & fg_DespintaRut(fpText(0).text) & "', " & vg_codbod & ", '" & vg_FDC & "'")
           Image1(1).Visible = IIf(RS.EOF Or RS1!nreg = 0 Or IsNull(RS!nreg), False, True)
           RS.Close
           Set RS = Nothing
        
        End If
        
        If Combo2(0).Enabled = True Then Combo2(0).SetFocus
        
        '------> Revisar si el usuario cambia de proveedor
        If vaSpread1.MaxRows < 1 Then Exit Sub
        Dim cPro As String
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            cPro = vaSpread1.text
            
            If Trim(vaSpread1.text) <> "" Then
               
               Revisa cPro, vaSpread1.Row
               vaSpread1_EditChange i, 4
            
            End If
        
        Next i
        vaSpread2.Refresh
        
    Case 1
        
        If Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FA" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Then
            
            vg_FDC = "GD"
            B_Guias.Cargar_DoctoGrilla Me, "GD", "Guía de Despacho", fg_DespintaRut(fpText(0).text), "", 0
        
        ElseIf Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "NC" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "CE" Then
            
            vg_FDC = "SN"
            B_Guias.Cargar_DoctoGrilla Me, "SN", "Solicitud NC", fg_DespintaRut(fpText(0).text), "", 0
        
        End If
        
        vg_RDC = fg_DespintaRut(fpText(0).text)
        B_Guias.Show 1
        
        If Trim(vg_Guias) <> "" Then
            
           Combo2(0).Enabled = False
           SumarTotales
           Double1(12).text = vTotExe
           Double1(13).text = vTotNet
           Double1(14).text = vTotIva
           Double1(15).text = vTotOtr
           Double1(16).text = vTotTot
           
           Text1(0).Enabled = False
           
           Double1(12).Enabled = False
           Double1(13).Enabled = False
           Double1(14).Enabled = False
           Double1(15).Enabled = False
           Double1(16).Enabled = False
           
           Text2(0).Visible = IIf(vg_pais = "CO", True, False)
           Text2(1).Visible = IIf(vg_pais = "CO", True, False)
           Text2(2).Visible = IIf(vg_pais = "CO", True, False)
    '       Text2(0).Visible = False: Text2(1).Visible = False
           Image1(5).Visible = False
           
           vaSpread1.Col = 10
           vaSpread1.ColHidden = False
           
           vaSpread1.Col = 20
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 21
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 22
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 23
           vaSpread1.ColHidden = True
           
           vaSpread1.ColWidth(2) = 55.13 '33.13
           vaSpread1.ColWidth(4) = 10
           vaSpread1.ColWidth(5) = 9.75
           vaSpread1.ColWidth(8) = 11.13
           vaSpread1.ColWidth(9) = 9.75
           vaSpread1.ColWidth(10) = 10.63
           vaSpread1.SetActiveCell 1, 1
           vaSpread1.Refresh
           If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
        
        End If
    
    Case 2
        
        SumarTotales
        MsgBox "Exento   : " & Format(vTotExe, fg_Pict(9, 0)) & VgLinea & _
               "Neto      : " & Format(vTotNet, fg_Pict(9, 0)) & VgLinea & _
               "Iva         : " & Format(vTotIva, fg_Pict(9, 0)) & VgLinea & _
               "Otr.Imp.  : " & Format(vTotOtr, fg_Pict(9, 0)) & VgLinea & _
               "Total      : " & Format(vTotTot, fg_Pict(9, 0)), vbInformation, MsgTitulo & " - Totales calculados"
    
    Case 3
        
        vg_left = fpayuda(0).Left + 2300
        vg_nombre = ""
        vg_codigo = ""
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpText(2).text = vg_codigo
        fpayuda(2).Caption = vg_nombre
    
    Case 4
        
        If Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FA" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Then
            
            vg_FDC = "OC"
            B_Guias.Cargar_DoctoGrilla Me, "OC", "Ordenes de Compras", fg_DespintaRut(fpText(0).text), "docpro", 1
        
        End If
        
        vg_RDC = fg_DespintaRut(fpText(0).text)
        B_Guias.Show 1
        If Trim(vg_Guias) <> "" Then
           
           Combo2(0).Enabled = False
           SumarTotales
           Double1(12).text = vTotExe
           Double1(13).text = vTotNet
           Double1(14).text = vTotIva
           Double1(15).text = vTotOtr
           Double1(16).text = vTotTot
           Text1(0).Enabled = False
           Double1(12).Enabled = True
           Double1(13).Enabled = True
           Double1(14).Enabled = True
           Double1(15).Enabled = True
           Double1(16).Enabled = True
           
           Frame6.Enabled = True
           Toolbar2.Enabled = True
           
           Text2(0).Visible = True
           Text2(1).Visible = True
           Text2(2).Visible = True
           
           vaSpread1.ColWidth(2) = 42.13 '41.13
           vaSpread1.ColWidth(4) = 9 '8
           vaSpread1.ColWidth(5) = 8.75 '7.75
           vaSpread1.ColWidth(8) = 10.13 '9.13
           vaSpread1.ColWidth(9) = 8.75 '7.75
           vaSpread1.ColWidth(10) = 9.63 '8.63
           vaSpread1.ColWidth(21) = 9.63
           vaSpread1.ColWidth(22) = 9.63
           vaSpread1.ColWidth(23) = 7.99
           
           vaSpread1.Col = 20
           vaSpread1.ColHidden = True
           
           vaSpread1.Col = 21
           vaSpread1.ColHidden = False
           
           vaSpread1.Col = 22
           vaSpread1.ColHidden = False
           
           vaSpread1.Col = 23
           vaSpread1.ColHidden = False
           
           vaSpread1.Col = 10
           vaSpread1.ColHidden = True
           
           vaSpread1.SetActiveCell 1, 1
           vaSpread1.Refresh
           If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
           
        End If
        
    Case 5 '-------> Mostrar para colombia los productos asociados a formato de compras
        
        If vg_pais = "CO" Then Exit Sub 'mientras tantos que va estar en procesos
        vg_left = fpayuda(Index).Left + 1290
        Me.Refresh
        vg_codigo = ""
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        vg_codigo = vaSpread1.text
        
        vaSpread1.Col = 24
        vg_nombre = Trim(vaSpread1.text)
        B_TabEst.LlenaDatos vg_nombre, vg_codigo, "Productos SAC", "CamPSAC"
    '       B_TabEst.LlenaDatos "b_productos", "pro_", "Productos SGP", "ProVigSac" '"PSGP"
        B_TabEst.Show 1, Me
        Me.Refresh
        If Trim(vg_codigo) = "" Or Val(vg_codigo) = 0 Then Exit Sub
        Text2(0).text = Trim(vg_codigo)
        Text2(1).text = Trim(vg_nombre)
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 24
        vaSpread1.text = vg_codigo
        
        vaSpread1.Col = 25
        vaSpread1.text = vg_nombre
        
End Select

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

If Encontrado = False Then
   
   ValidarCfc
   IndiceOpt = Index

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_GotFocus(Index As Integer)
'StopObjeto True
End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
IndiceOpt = Index
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 11

If Trim(fg_codigocbo(Combo2, 0, 2, "")) = "CG" Then vaSpread1.text = Text1(0).text: Exit Sub

NumEnter = Len(Text1(0).text) - InStr(1, Text1(0).text, Chr(13)) + 1

If Text1(0).text = "" Then Text1(0) = "  "

Glosa = Text1(0).text ' Mid$(Text1(0).Text, 1, Len(Text1(0).Text) - NumEnter)
vaSpread1.text = Glosa
vaSpread1.SetFocus
vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim v_rut As String, v_tipo As String, v_bodega As Long, v_CtaCon As String, v_Rebaja As String, v_can As Double, v_precio As Double, acepre As String, ingcfc As Boolean
Dim v_pctdes As Double, v_valdes As Double, v_total As Double, v_descrip As String, v_canrec As Double, v_prerec As Double, prefto As Double, prefun As Double
Dim StoA As Double, PPPA As Double, PreC As Double, Cdif As Long, v_casino As String, vTotRec As Double, vPreDec As Double, periodo As Long, activo As String, docele As String
Dim sql1 As String, codpro As String, codsac As String, v_fecoc As Date, v_canoc As Double, v_preoc As Double

If Date1(0).text <> "" Then Date1(1).text = Date1(0).text
v_casino = TipoDato(GetParametro("casino"), "")
v_rut = fg_DespintaRut(fpText(0).text)
v_tipo = fg_codigocbo(Combo2, 0, 2, "")
v_bodega = fg_codigocbo(Combo2, 1, 10, 0)
v_fecemi = Format(Date1(0).text, "dd/mm/yyyy")
v_fecven = Format(Date1(1).text, "dd/mm/yyyy")
v_fecrem = Format(Date1(2).text, "dd/mm/yyyy")

TraerFechaCierre

Select Case Button.Index
    
    Case 1 '-------> Agregar
        
        Gl_Ac_Botones Me, 7, 3, ""
        Nuevo_Registro
        modo = "A"
    
    Case 6 '-------> Cancelar
        
        If vaSpread1.MaxRows > 0 Then If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        Gl_Ac_Botones Me, 7, 1, ""
        Nuevo_Registro
    
    Case 8 '-------> Grabar
        
        '-------> Validar si proveedor esta esta activo
        If Date1(0).text = "" Or Date1(2).text = "" Then MsgBox "Fecha esta en blanco...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        '-------> Validar si el contrato tiene asignado inventario rotativo
        If ValidarInventarioRotativo(MuestraCasino(1)) And ValidarActividadesDiariaInvRotativo(MuestraCasino(1)) And _
           Format(Date1(2).Value, "dd/mm/yyyy") > CDate(vg_ciedia) Then MsgBox "Tiene que realizar cierre diario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        activo = ""
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT prv_activo FROM b_proveedor WHERE prv_codigo = '" & v_rut & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then activo = Trim(RS1!prv_activo)
        RS1.Close
        Set RS1 = Nothing
        
        If (activo = "1" Or activo = "2") And Len(Trim(vg_Guias)) < 1 Then MsgBox "Proveedor esta en estado : (Inactivo ó bien Eliminado), No puede ingresar documento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        '-------> traer proveedor
        ingcfc = True
        
        '-------> Validar si proveedor es local
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT prv_origen, prv_docele FROM b_proveedor WHERE prv_codigo='" & v_rut & "'", vg_db, adOpenStatic
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
        If RS1!prv_origen = "0" And Option1(1).Value = True And ingcfc Then RS1.Close: Set RS1 = Nothing: MsgBox "Proveedor es local, solamente puede ingresar documento fofi...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        If RS1!prv_docele = "S" And (fg_TraerRelacionTipoDocumento(v_tipo) = "FA" Or fg_TraerRelacionTipoDocumento(v_tipo) = "NC" Or fg_TraerRelacionTipoDocumento(v_tipo) = "ND") Then
           
           If MsgBox("El tipo de documento predeterminado para este proveedor es tipo ELECTRONICO" & VgLinea & VgLinea & "                Está seguro de que el documento ingresado es tipo MANUAL...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS1.Close: Set RS1 = Nothing: Exit Sub
        
        ElseIf RS1!prv_docele = "N" And (fg_TraerRelacionTipoDocumento(v_tipo) = "FE" Or fg_TraerRelacionTipoDocumento(v_tipo) = "CE" Or fg_TraerRelacionTipoDocumento(v_tipo) = "DE") Then
           
           If MsgBox("El tipo de documento predeterminado para este proveedor es tipo MANUAL" & VgLinea & VgLinea & "    Está seguro de que el documento ingresado es tipo ELECTRONICO...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then RS1.Close: Set RS1 = Nothing: Exit Sub
        
        End If
        RS1.Close
        Set RS1 = Nothing
        
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 0) And Len(Trim(vg_Guias)) < 1 Then
           
           MsgBox "Documento no corresponde al periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 6) And Format(CDate(vg_ciedia), "mm/yyyy") = Format(CDate(Date1(2).Value), "mm/yyyy") Then
        
           MsgBox "No puede ingresar documentos anteriores a la última toma de inventario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 12) Then
           
           MsgBox "La fecha del documento es mayor a la fecha del periodo : " & VgLinea & VgLinea & CierreFecha, vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        'Validar inventario calendarizado 20201001
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 38) Then
        
           MsgBox "Se esta realizando la toma de inventario en estos momento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
        End If
        
        'Validar ingreso documento inventario calendarizado 20201001
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 40) Then
        
           MsgBox "No puede ingresar documento, antes de un inventario calendarizado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
        End If
        
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 8) Then
        
           MsgBox "No ha realizado el ajuste correspondiente a la última toma de inventario...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
        End If
        
        If CDate(Date1(2).Value) < CDate(vg_ciedia) And Format(CDate(vg_ciedia), "mm/yyyy") = Format(CDate(Date1(2).Value), "mm/yyyy") Then
        
           MsgBox "Día se encuentra cerrado, no es posible ingresar...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        '-------> Validar documento si existe
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT b.bod_nombre FROM  b_totcompras a, a_bodega b WHERE a.toc_codbod=b.bod_codigo AND a.toc_rutpro='" & fg_DespintaRut(fpText(0).text) & "' " & _
                 "AND a.toc_tipdoc='" & fg_codigocbo(Combo2, 0, 2, "") & "' AND a.toc_numdoc=" & Val(Double1(5).Value), vg_db, adOpenStatic
        If Not RS1.EOF Then MsgBox "Documento ya existe en la bodega : " & Trim(RS1!bod_nombre) & VgLinea & VgLinea, vbExclamation + vbOKOnly, MsgTitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
        RS1.Close
        Set RS1 = Nothing
        
        '-------> Validar Numero folio
    '20080311    RS1.Open "SELECT DISTINCT toc_numinf, FORMAT(toc_fecemi, 'mm/yyyy') AS fecemi FROM b_totcompras WHERE toc_numinf=" & Val(Double1(6).text) & " AND  toc_tipinf='" & IIf(Option1(0).Value = True, "F", "C") & "' AND toc_codbod=" & vg_codbod & "", vg_db, adOpenStatic
    '20080311    If Not RS1.EOF Then If RS1!fecemi <> Format(Date1(2).text, "mm/yyyy") Then MsgBox "Nº folio corresponde al periodo : " & RS1!fecemi & " " & VgLinea & VgLinea & "Tiene que generar un nuevo folio", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
    '20080311    RS1.Close: Set RS1 = Nothing
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        If vg_tipbase = "1" Then
           
           RS1.Open "SELECT DISTINCT toc_numinf, FORMAT(toc_fecrem, 'mm/yyyy') AS fecemi FROM b_totcompras WHERE toc_numinf = " & Val(Double1(6).text) & " AND  toc_tipinf = '" & IIf(Option1(0).Value = True, "F", "C") & "' AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        
        Else
           
           RS1.Open "SELECT DISTINCT toc_numinf, CASE WHEN convert(varchar(2),datepart(mm,toc_fecrem)) < 10 THEN '0' + convert(varchar(2),datepart(mm,toc_fecrem)) ELSE convert(varchar(2),datepart(mm,toc_fecrem)) END + '/' + convert(varchar(4),datepart(year, toc_fecrem)) AS fecemi FROM b_totcompras WHERE toc_numinf = " & Val(Double1(6).text) & " AND  toc_tipinf = '" & IIf(Option1(0).Value = True, "F", "C") & "' AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        
        End If
        If Not RS1.EOF Then If RS1!fecemi <> Format(Date1(2).text, "mm/yyyy") Then MsgBox "Nº folio corresponde al periodo : " & RS1!fecemi & " " & VgLinea & VgLinea & "Tiene que generar un nuevo folio", vbExclamation + vbOKOnly, MsgTitulo: RS1.Close: Set RS1 = Nothing: Exit Sub
        RS1.Close
        Set RS1 = Nothing
        
        '-------> Validar cantidad documento en un folio
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT COUNT(*) AS nreg FROM b_totcompras a, a_tipodocumento b WHERE a.toc_tipdoc = b.tdo_codigo AND (b.tdo_cladoc) IS NOT NULL AND b.tdo_cladoc <> '' AND a.toc_numinf = " & Val(Double1(6).text) & " AND a.toc_tipinf = '" & IIf(Option1(0).Value = True, "F", "C") & "' AND a.toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        If Not RS1.EOF Then If RS1!nreg > 20 Then RS1.Close: Set RS1 = Nothing: MsgBox "Folio excede los 20 documento, genero un nuevo folio...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        RS1.Close
        Set RS1 = Nothing
        
        If modo <> "A" Then MsgBox "Insuficiencia de datos. Documento no grabado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Trim(fpayuda(2).Caption) = "" Then MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Trim(fpText(0).text) = "" Or Combo2(0).text = "" Then MsgBox "No hay datos proveedor...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Val(Double1(5).Value) = 0 Then MsgBox "Debe ingresar Nº de documento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Val(vTotTot) = 0 And v_tipo <> "GD" Then MsgBox "Total documento no puede ser cero...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Date1(0).text = "" Or Date1(1).text = "" Or Date1(2).text = "" Then MsgBox "Debe seleccionar fechas...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Option1(IndiceOpt).Value = False Then MsgBox "Tipo de documento no valido...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Not fg_Check_Rut(fg_DespintaRut(fpText(0).text)) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If vaSpread1.MaxRows = 0 Then MsgBox "Documento sin detalle...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If Valida_DatosGrilla Then MsgBox "La cantidad o el precio de un producto es cero...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If vTotExe <> Trim(Double1(12).text) Or vTotNet <> Trim(Double1(13).text) _
           Or vTotIva <> Trim(Double1(14).text) Or vTotOtr <> Trim(Double1(15).text) _
           Or vTotTot <> Trim(Double1(16).text) Then MsgBox "Los totales del documento no conciden con el detalle ingresado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        '-------> Validar precio
        Dim Precio As Double, prerea As Double, porpre As Double, porpar As Double, estpre As Boolean
        porpar = 0
        estpre = False
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT par_valor FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'porprepro'", vg_db, adOpenStatic
        If Not RS1.EOF Then porpar = RS1!par_valor
        RS1.Close
        Set RS1 = Nothing
        
        If porpar > 0 Then
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Row = i
               vaSpread1.Col = 1
               
               vaSpread1.Col = 5
               Precio = vaSpread1.text
               
               vaSpread1.Col = 18
               prerea = vaSpread1.text
               
               vaSpread1.Col = 19
               vaSpread1.text = "N"
               
               vaSpread1.Col = 13
               If prerea > 0 And vaSpread1.text = "S" Then
                  
                  porpre = (IIf(Precio > prerea, Round(Precio / prerea, 1), Round(prerea / Precio, 1)) * porpar)
                  If porpre > porpar Then estpre = True: vaSpread1.Col = 19: vaSpread1.text = "S"
               
               End If
           
           Next i
        
        End If
        If MsgBox(IIf(estpre = True, "Existen precios ingresados, que excede al ultimo precio registrado" & VgLinea & VgLinea & "               Graba documento...", "Graba documento..."), vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        
        '-------> suma los descuentos
        v_totdesc = 0
        For cont = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = cont
            vaSpread1.Col = 7
            v_totdesc = v_totdesc + Val(vaSpread1.Value)
        
        Next cont
        
        '-------> Obtiene el parametro si el documento es FIFO o CFC
        opcion = IIf(Option1(0).Value = True, "F", "C")
        
        '-------> Comienza Transaccion
        periodo = 0
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT cie_periodo FROM b_cierreperiodo WHERE cie_cencos = '" & MuestraCasino(1) & "' AND cie_estado = 1", vg_db, adOpenStatic
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "No existe periodo proceso cancelado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        periodo = RS1!cie_periodo
        RS1.Close
        Set RS1 = Nothing
        
        Toolbar1.Enabled = False
        Image1(5).Visible = False
        
        vg_db.BeginTrans
        
        If vg_FDC = "OC" Then vg_Guias = ""
        
        If vg_tipbase = "1" Then
           
           vg_db.Execute "INSERT INTO b_totcompras (toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem, EnvioDocSGPADM, toc_docasotipo) " & _
                         "VALUES ('" & v_rut & "', '" & v_tipo & "'," & Val(Double1(5).text) & "," & v_bodega & ",'" & v_fecemi & "','" & v_fecven & "'," & v_totdesc & "," & vTotNet & "," & vTotExe & "," & vTotIva & "," & vTotOtr & "," & vTotTot & "," & vTotTot & ",'" & opcion & "'," & Double1(6).Value & ", '" & IIf(vg_FDC = "OC", "", Trim(vg_Guias)) & "', '" & Trim(fpText(1).text) & "'," & Double1(17).Value & ", '', '0', '" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "h:m:s") & "', " & periodo & ", '" & v_fecrem & "', '0', '" & Trim(vg_GuiasTipo) & "')"
        
        Else
           
           vg_db.Execute "INSERT INTO b_totcompras (toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc, toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem, EnvioDocSGPADM, toc_docasotipo) " & _
                         "VALUES ('" & v_rut & "', '" & v_tipo & "'," & Val(Double1(5).text) & "," & v_bodega & ", '" & Format(v_fecemi, "yyyymmdd") & "','" & Format(v_fecven, "yyyymmdd") & "'," & v_totdesc & "," & vTotNet & "," & vTotExe & "," & vTotIva & "," & vTotOtr & "," & vTotTot & "," & vTotTot & ",'" & opcion & "'," & Double1(6).Value & ", '" & IIf(vg_FDC = "OC", "", Trim(vg_Guias)) & "', '" & Trim(fpText(1).text) & "'," & Double1(17).Value & ", '', '0', '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "h:m:s") & "', " & periodo & ", '" & Format(v_fecrem, "yyyymmdd") & "', '0', '" & Trim(vg_GuiasTipo) & "')"
        
        End If
        
        Dim v_faccon As Double
        Dim v_cmefac As Double
        Dim v_pmefac As Double
        Dim v_crefac As Double
        Dim v_prefac As Double
        
        For cont = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = cont
            vaSpread1.Col = 1
            
            v_faccon = 0
            v_faccan = 0
            v_prefac = 0
            
            If Trim(vaSpread1.text) <> "" And IsNull((vaSpread1.text)) = False Then
                
                '-------> Mover codigo de producto SGP Chile = 1, Colombia = 24
                vaSpread1.Col = IIf(vg_pais = "CO", 24, 1)
                codpro = Trim(vaSpread1.text)
                
                vaSpread1.Col = 29
                v_faccon = Val(vaSpread1.Value)
                
                If vg_pais = "CO" Then
                   
                   vaSpread1.Col = 4
                   v_cmefac = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 5
                   v_pmefac = Val(vaSpread1.Value)
                   v_can = (v_cmefac * v_faccon)
                   
                   If v_can > 0 Then
                      
                      v_precio = Round(((v_cmefac * v_pmefac) / v_can), 2)
                   
                   Else
                      
                      v_precio = 0
                   
                   End If
                
                Else
                   
                   vaSpread1.Col = 4
                   v_can = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 5
                   v_precio = Val(vaSpread1.Value)
                
                End If
                
                vaSpread1.Col = 6
                v_pctdes = Val(vaSpread1.Value)
                
                vaSpread1.Col = 7
                v_valdes = Val(vaSpread1.Value)
                
                vaSpread1.Col = 8
                v_total = Val(vaSpread1.Value)
                
                If vg_pais = "CO" Then
                   
                   vaSpread1.Col = 9
                   v_crefac = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 10
                   v_prefac = Val(vaSpread1.Value)
                   
                   v_canrec = (v_crefac * v_faccon)
                   
                   If v_canrec > 0 Then
                      
                      v_prerec = Round(((v_crefac * v_prefac) / v_canrec), 2)
                   
                   Else
                      
                      v_prerec = Round(((v_cmefac * v_prefac) / v_can), 2)
                   
                   End If
                   
                   vaSpread1.Col = 31
                   vaSpread1.Value = v_canrec
                
                Else
                   
                   vaSpread1.Col = 9
                   v_canrec = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 10
                   v_prerec = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 31
                   vaSpread1.Value = v_canrec
                
                End If
    
                vaSpread1.Col = 11
                v_descrip = LimpiaDato(Trim(Text1(0).text)) '20080417Trim(vaSpread1.text)
                
                vaSpread1.Col = 12
                v_CtaCon = Trim(vaSpread1.text)
                
                vaSpread1.Col = 13
                v_Rebaja = Trim(vaSpread1.text)
                
                If ((fg_TraerRelacionTipoDocumento(v_tipo) = "FA" Or fg_TraerRelacionTipoDocumento(v_tipo) = "FE") And Len(Trim(vg_Guias)) > 0) Or (fg_TraerRelacionTipoDocumento(v_tipo) = "NC" Or fg_TraerRelacionTipoDocumento(v_tipo) = "CE") Then v_Rebaja = "N"
                
                If vg_pais = "CO" Then
                   
                   vaSpread1.Col = 15
                   PreC = Val(vaSpread1.text)
                   PreC = Round(((v_cmefac * PreC) / v_can), 2)
                   
                   vaSpread1.Col = 17
                   vPreDec = Val(vaSpread1.Value)
                   vPreDec = Round(((v_cmefac * vPreDec) / v_can), 3)
                
                Else
                   
                   vaSpread1.Col = 15
                   PreC = Val(vaSpread1.text)
                   
                   vaSpread1.Col = 17
                   vPreDec = Val(vaSpread1.Value)
                
                End If
                vaSpread1.Col = 19
                acepre = Trim(vaSpread1.text)
                
                vaSpread1.Col = 21
                v_canoc = IIf(Trim(vaSpread1.text) = "" Or Trim(vaSpread1.text) = "0", 0, (vaSpread1.text))
                
                vaSpread1.Col = 22
                v_preoc = IIf(Trim(vaSpread1.text) = "" Or Trim(vaSpread1.text) = "0", 0, (vaSpread1.text))
                
                vaSpread1.Col = 23
                v_fecoc = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
                
                vaSpread1.Col = IIf(vg_pais = "CO", 1, 24)
                codsac = vaSpread1.text
                
                vTotRec = IIf(v_can = v_canrec And v_precio = v_prerec, v_total, Round(v_canrec * vPreDec, 0))
                
                '-------> jpaz calcular precio del flete al productos
                prefto = 0
                prefun = 0
                If Double1(17).Value > 0 And v_can > 0 Then prefun = Round(((Double1(17).Value / Double1(13).Value) * v_total) / v_can): prefto = Round((Double1(17).Value / Double1(13).Value) * v_total, 2) 'Round((Double1(17).Value / Double1(13).Value) * v_total)
                PreC = PreC + prefun
                If vg_FDC = "OC" And (v_can > 0 Or v_canrec > 0) Then
                   
                   vg_db.Execute "INSERT INTO b_detcompras (dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon) " & _
                                 "VALUES ('" & v_rut & "', '" & v_tipo & "', " & Double1(5).Value & ", " & cont & ", '" & codpro & "', " & v_can & ", " & v_precio & ", " & v_pctdes & ", " & v_valdes & ", " & v_total & ", '" & TipoDato(v_descrip, "") & "', " & v_canrec & ", " & v_prerec & ", '" & v_Rebaja & "', " & prefto & ", " & vTotRec & ", '" & acepre & "', " & v_cmefac & ", " & v_pmefac & ", " & v_crefac & ", " & v_prefac & ", " & v_faccon & ")"
                
                ElseIf vg_FDC <> "OC" Then
                   
                   vg_db.Execute "INSERT INTO b_detcompras (dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_acepre, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon) " & _
                                 "VALUES ('" & v_rut & "', '" & v_tipo & "', " & Double1(5).Value & ", " & cont & ", '" & codpro & "', " & v_can & ", " & v_precio & ", " & v_pctdes & ", " & v_valdes & ", " & v_total & ", '" & TipoDato(v_descrip, "") & "', " & v_canrec & ", " & v_prerec & ", '" & v_Rebaja & "', " & prefto & ", " & vTotRec & ", '" & acepre & "', " & v_cmefac & ", " & v_pmefac & ", " & v_crefac & ", " & v_prefac & ", " & v_faccon & ")"
                
                End If
                
                If vg_pais = "CO" Then
                   
                   '-------> Actualizar tabla formato compras sgp el campo predeterminado
                   vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_cenpre = 0 WHERE fcs_codsgp = '" & codpro & "'"
                   vg_db.Execute "UPDATE b_formatocomprassgp SET fcs_cenpre = 1 WHERE fcs_codsac = '" & codsac & "' AND fcs_codsgp = '" & codpro & "'"
                   
                   '-------> detalle orden de compras sac recibido
                   vg_db.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
                                 "VALUES ('" & v_rut & "', '" & v_tipo & "', " & Double1(5).Value & ", " & cont & ", '" & codpro & "', '" & codsac & "', " & v_can & ", " & v_precio & ", " & v_canrec & ", '" & IIf(vg_tipbase = "1", v_fecoc, Format(v_fecoc, "yyyymmdd")) & "', " & v_canoc & ", " & v_preoc & ")"
                   
                   '-------> Actualizar campo ocr_fecoc si registro no pertenece a la orden de compras
                   If Format(CDate(v_fecoc), "dd/mm/yyyy") = Format(CDate("00000000"), "dd/mm/yyyy") Then
                      
                      If vg_tipbase = "1" Then
                         
                         vg_db.Execute "UPDATE b_ocsacrecibido SET ocr_fecoc = null WHERE ocr_rutpro = '" & v_rut & "' AND ocr_tipdoc = '" & v_tipo & "' AND ocr_numdoc = " & Double1(5).Value & " AND ocr_numlin = " & cont & " AND ocr_codprodsgp = '" & codpro & "'"
                      
                      Else
                         
                         vg_db.Execute "UPDATE b_ocsacrecibido SET ocr_fecoc = Null WHERE ocr_rutpro = '" & v_rut & "' AND ocr_tipdoc = '" & v_tipo & "' AND ocr_numdoc = " & Double1(5).Value & " AND ocr_numlin = " & cont & " AND ocr_codprodsgp = '" & codpro & "'"
                      
                      End If
                   
                   End If
                   
                   If vg_FDC = "OC" And v_can = 0 Then
                      
                      '-------> Esto es para que no ingrese a realizar rebaja stock
                      v_Rebaja = "N"
                   
                   End If
                
                Else
    '20100728               If vg_FDC = "OC" And v_canoc > 0 And (v_can > 0 Or v_canrec > 0) Then
                   If vg_FDC = "OC" Then
                      
                      '-------> detalle orden de compras sac recibido
                      vg_db.Execute "INSERT INTO b_ocsacrecibido (ocr_rutpro, ocr_tipdoc, ocr_numdoc, ocr_numlin, ocr_codprodsgp, ocr_codprodsac, ocr_cancom, ocr_precom, ocr_canrec, ocr_fecoc, ocr_canoc, ocr_preoc) " & _
                                    "VALUES ('" & v_rut & "', '" & v_tipo & "', " & Double1(5).Value & ", " & cont & ", '" & codpro & "', '" & codsac & "', " & v_can & ", " & v_precio & ", " & v_canrec & ", '" & IIf(vg_tipbase = "1", v_fecoc, Format(v_fecoc, "yyyymmdd")) & "', " & v_canoc & ", " & v_preoc & ")"
                      
                      '-------> Actualizar campo ocr_fecoc si registro no pertenece a la orden de compras
                      If Format(CDate(v_fecoc), "dd/mm/yyyy") = Format(CDate("00000000"), "dd/mm/yyyy") Then
                         If vg_tipbase = "1" Then
                            
                            vg_db.Execute "UPDATE b_ocsacrecibido SET ocr_fecoc = '' WHERE ocr_rutpro = '" & v_rut & "' AND " & _
                                          "ocr_tipdoc = '" & v_tipo & "' AND ocr_numdoc = " & Double1(5).Value & " AND " & _
                                          "ocr_numlin = " & cont & " AND ocr_codprodsgp = '" & codpro & "'"
                         
                         Else
                            
                            vg_db.Execute "UPDATE b_ocsacrecibido SET ocr_fecoc = Null WHERE ocr_rutpro = '" & v_rut & "' AND " & _
                                          "ocr_tipdoc = '" & v_tipo & "' AND ocr_numdoc = " & Double1(5).Value & " AND " & _
                                          "ocr_numlin = " & cont & " AND ocr_codprodsgp = '" & codpro & "'"
                         
                         End If
                      End If
                   
                   End If
                   
                   If vg_FDC = "OC" And v_can = 0 Then
                      
                      '-------> Esto es para que no ingrese a realizar rebaja stock
                      v_Rebaja = "N"
                   
                   End If
                   
    '20100728               ElseIf vg_FDC = "OC" And v_can = 0 Then
    '20100728                  '-------> Esto es para que no ingrese a realizar rebaja stock
    '20100728                  v_Rebaja = "N"
    '20100728               End If
                
                End If
                
                If v_Rebaja = "S" And Len(Trim(vg_Guias)) = 0 And (fg_TraerRelacionTipoDocumento(v_tipo) <> "NC" Or fg_TraerRelacionTipoDocumento(v_tipo) <> "CE") Then 'si el producto rebaja stock
                    
                    ValidaBod v_bodega, codpro
                    '-------> Proceso calculo pmp
                    Dim pmp As Double, auxCanmer As Double, auxPropon As Double, feccos As Long
                    pmp = 0
                    pmp = Cal_PMP(MuestraCasino(1), v_bodega, codpro, Format(Date1(2).text, "dd/mm/yyyy"), PreC, v_canrec)
    
                    If RS1.State = 1 Then RS1.Close
                    RS1.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    RS1.Open "SELECT DISTINCT a.ppd_propon FROM b_productospmpdia a, b_productos b WHERE a.ppd_codpro = b.pro_codigo AND " & _
                             "b.pro_ctrsto = 1 AND a.ppd_cencos = '" & MuestraCasino(1) & "' AND a.ppd_codpro = '" & codpro & "' AND " & _
                             "a.ppd_fecdia = " & Format(CDate(v_fecrem), "yyyymmdd") & "", vg_db, adOpenStatic
                    
                    If RS1.EOF Then
                       
                       RS1.Close
                       Set RS1 = Nothing
                       
                       If vg_tipbase = "1" Then
                          
                          vg_db.Execute "INSERT INTO b_productospmpdia VALUES ('" & MuestraCasino(1) & "', '" & codpro & "', " & Format(v_fecrem, "yyyymmdd") & ", " & pmp & ", 0, " & PreC & ", cdate('" & v_fecrem & "'))"
                       
                       Else
                          
                          vg_db.Execute "INSERT INTO b_productospmpdia VALUES ('" & MuestraCasino(1) & "', '" & codpro & "', " & Format(v_fecrem, "yyyymmdd") & ", " & pmp & ", 0, " & PreC & ", '" & Format(v_fecrem, "yyyymmdd") & "')"
                       
                       End If
                    
                    ElseIf Not RS1.EOF Then
                       
                       '-------> Actualizar pmp si es menor que cero
                       RS1.Close
                       Set RS1 = Nothing
                       
                       If RS1.State = 1 Then RS1.Close
                       RS1.CursorLocation = adUseClient
                       vg_db.CursorLocation = adUseClient
                       
                       RS1.Open "SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = '" & codpro & "' AND ppd_fecdia = " & Format(CDate(v_fecrem), "yyyymmdd") & "", vg_db, adOpenStatic
                       
                       If Not RS1.EOF Then
                          
                          '-------> Actualizar precio ultima compra, fecha ultima compra y propon
                          If vg_tipbase = "1" Then
                             
                             vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & pmp & ", ppd_upreco = " & PreC & ", ppd_fecuco = cdate('" & v_fecrem & "') WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = '" & codpro & "' AND ppd_fecdia = " & Format(CDate(v_fecrem), "yyyymmdd") & ""
                          
                          Else
                             
                             vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & pmp & ", ppd_upreco = " & PreC & ", ppd_fecuco = '" & Format(v_fecrem, "yyyymmdd") & "' WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = '" & codpro & "' AND ppd_fecdia = " & Format(CDate(v_fecrem), "yyyymmdd") & ""
                          
                          End If
                       Else
                          
                          '-------> Actualizar precio ultima compra y fecha ultima compra
                          sql1 = IIf(vg_tipbase = "1", " cdate('" & v_fecrem & "') ", " '" & Format(v_fecrem, "yyyymmdd") & "' ")
                          vg_db.Execute "UPDATE b_productospmpdia SET ppd_propon = " & pmp & ", ppd_upreco = " & PreC & ", ppd_fecuco = " & sql1 & " WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_codpro = '" & codpro & "' AND ppd_fecdia = " & Format(CDate(v_fecrem), "yyyymmdd") & ""
                       
                       End If
                       
                       RS1.Close
                       Set RS1 = Nothing
                    
                    End If
    
        '------------Fin PMP ------------------------------------------------
        '------------Actuliza Stock de bodega---------------------------
                    vg_db.Execute "UPDATE b_bodegas SET bod_canmer=bod_canmer+" & v_canrec & " WHERE bod_codbod = " & v_bodega & " AND bod_codpro = '" & codpro & "'"
        '------------Fin Actualiza Stock -----------------------------------------
        '------------Actuliza codigo de ultimo producto de compra  compra---------
                    
                    If RS1.State = 1 Then RS1.Close
                    RS1.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    RS1.Open "SELECT pri_coding FROM b_productosing WHERE pri_codpro = '" & codpro & "'", vg_db, adOpenStatic
                    If Not RS1.EOF Then
                       
                       vg_db.Execute "UPDATE b_contlistpreing SET cpi_codcom = '" & codpro & "' WHERE cpi_coding = '" & RS1!pri_coding & "' AND cpi_cencos = '" & MuestraCasino(1) & "'"
                    
                    End If
                    RS1.Close
                    Set RS1 = Nothing
        '------------Fin Actualiza -----------------------------------------------
                
                End If
             
             End If
        
        Next cont
        
        If Len(Trim(vg_Guias)) > 0 And (fg_TraerRelacionTipoDocumento(v_tipo) = "FA" Or fg_TraerRelacionTipoDocumento(v_tipo) = "FE" Or fg_TraerRelacionTipoDocumento(v_tipo) = "NC" Or fg_TraerRelacionTipoDocumento(v_tipo) = "CE") Then
            
            Dim StrImp As String, StrImpb As String
            Dim StrImpTipo As String, StrImpbTipo As String
            StrImp = Trim(vg_Guias)
            StrImpTipo = Trim(vg_GuiasTipo)
            
            Do While InStr(StrImp, ";") <> 0
                
                StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                
                If Trim(vg_GuiasTipo) <> "" Then
                   
                   StrImpbTipo = Mid(StrImpTipo, 1, InStr(StrImpTipo, ";") - 1)
                   StrImpTipo = IIf(Len(StrImpTipo) > InStr(StrImpTipo, ";"), Mid(StrImpTipo, InStr(StrImpTipo, ";") + 1), "")
                
                End If
                
                If vg_FDC <> "OC" And fg_TraerRelacionTipoDocumento(v_tipo) = "FA" Or fg_TraerRelacionTipoDocumento(v_tipo) = "FE" Then
                    
                   If Trim(vg_GuiasTipo) = "" Then

                      vg_db.Execute "UPDATE b_totcompras SET toc_docaso = " & Str(Double1(5).Value) & " WHERE toc_rutpro = '" & v_rut & "' AND toc_codbod = " & vg_codbod & " AND toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'GD')  AND toc_numdoc = " & Val(StrImpb)
                
                   ElseIf Trim(vg_GuiasTipo) <> "" Then
                
                          vg_db.Execute "UPDATE b_totcompras SET toc_docaso = " & Str(Double1(5).Value) & " " & _
                                        "WHERE toc_rutpro = '" & v_rut & "' " & _
                                        "AND toc_codbod = " & vg_codbod & " " & _
                                        "AND toc_tipdoc = '" & StrImpbTipo & "' AND toc_numdoc = " & Val(StrImpb)
                   End If
                
                ElseIf vg_FDC <> "OC" Then
                    
                    vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = " & Str(Double1(5).Value) & " WHERE toc_rutpro = '" & v_rut & "' AND toc_codbod = " & vg_codbod & " AND toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') AND toc_numdoc = " & Val(StrImpb)
                
                End If
            
            Loop
        
        End If
        
        vg_db.Execute "INSERT INTO b_detcomprasimp (imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp)  " & _
                      "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, b.ipr_codimp, c.imp_pctimp, round(a.dec_ptotal*(c.imp_pctimp/100),0) " & _
                      "FROM b_detcompras a, b_productosimp b, a_impuesto c WHERE a.dec_codmer = b.ipr_codpro AND b.ipr_codimp = c.imp_codigo AND a.dec_rutpro = '" & v_rut & "' AND a.dec_tipdoc = '" & v_tipo & "' AND a.dec_numdoc = " & Val(Double1(5).Value)
        
        '-------> Actualizar impuesto real al b_detcomprasimp
        Dim codimp As Long, PctI As Double, CosI As Long
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = IIf(vg_pais = "CO", 24, 1)
            codpro = Trim(vaSpread1.text)
            vaSpread1.Col = 14
            StrImp = Trim(vaSpread1.text)
            
            If Len(StrImp) <> 0 Then
               
               Do While InStr(StrImp, ";") <> 0
                  
                  StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                  StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                  codimp = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                  PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                  CosI = Val(Mid(StrImpb, 1))
                  vg_db.Execute "UPDATE b_detcomprasimp SET imd_pctimp = " & PctI & " WHERE imd_codpro = '" & codpro & "' AND imd_codimp = " & codimp & " AND imd_rutdoc = '" & v_rut & "' AND imd_tipdoc = '" & v_tipo & "' AND imd_numdoc = " & Val(Double1(5).Value) & ""
                  vg_db.Execute "UPDATE b_detcomprasimp SET imd_monimp = 0 WHERE imd_pctimp = 0 AND imd_codpro = '" & codpro & "' AND imd_codimp = " & codimp & " AND imd_rutdoc = '" & v_rut & "' AND imd_tipdoc = '" & v_tipo & "' AND imd_numdoc = " & Val(Double1(5).Value) & ""
                  
                  If vg_tipbase = "1" Then
                     
                     vg_db.Execute "UPDATE b_detcomprasimp INNER JOIN b_detcompras ON (b_detcompras.dec_numlin = b_detcomprasimp.imd_numlin) AND (b_detcompras.dec_numdoc = b_detcomprasimp.imd_numdoc) AND (b_detcompras.dec_tipdoc = b_detcomprasimp.imd_tipdoc) AND (b_detcompras.dec_rutpro = b_detcomprasimp.imd_rutdoc) SET b_detcomprasimp.imd_monimp = (b_detcompras.dec_ptotal*(" & PctI & "/100)) " & _
                                   "WHERE b_detcomprasimp.imd_pctimp > 0 AND ((b_detcomprasimp.imd_monimp) Is Null OR b_detcomprasimp.imd_monimp > 0) AND b_detcomprasimp.imd_codpro = '" & codpro & "' AND b_detcomprasimp.imd_codimp = " & codimp & " AND b_detcomprasimp.imd_rutdoc = '" & v_rut & "' AND b_detcomprasimp.imd_tipdoc = '" & v_tipo & "' AND b_detcomprasimp.imd_numdoc = " & Val(Double1(5).Value) & ""
                  
                  Else
                     
                     vg_db.Execute "UPDATE b_detcomprasimp SET b_detcomprasimp.imd_monimp = round(b_detcompras.dec_ptotal*(convert(float," & PctI & ")/100),0) FROM b_detcompras, b_detcomprasimp WHERE b_detcompras.dec_numlin = b_detcomprasimp.imd_numlin AND b_detcompras.dec_numdoc = b_detcomprasimp.imd_numdoc AND b_detcompras.dec_tipdoc = b_detcomprasimp.imd_tipdoc AND b_detcompras.dec_rutpro = b_detcomprasimp.imd_rutdoc " & _
                                   "AND b_detcomprasimp.imd_pctimp > 0 AND (b_detcomprasimp.imd_monimp Is Null OR b_detcomprasimp.imd_monimp > 0) AND b_detcomprasimp.imd_codpro = '" & codpro & "' AND b_detcomprasimp.imd_codimp = " & codimp & " AND b_detcomprasimp.imd_rutdoc = '" & v_rut & "' AND b_detcomprasimp.imd_tipdoc = '" & v_tipo & "' AND b_detcomprasimp.imd_numdoc = " & Val(Double1(5).Value) & ""
                  
                  End If
               
               Loop
            
            End If
        
        Next i
        '-------> Fin actualizar impuesto real al b_detcomprasimp
    
        '-------> Chequeo de Diferencias
        If fg_TraerRelacionTipoDocumento(v_tipo) = "FA" Or fg_TraerRelacionTipoDocumento(v_tipo) = "FE" Then
            
            Cdif = 0
            
            For i = 1 To vaSpread1.MaxRows
                
                vaSpread1.Row = i
                vaSpread1.Col = 4
                v_can = Val(vaSpread1.Value)
                
                vaSpread1.Col = 5
                v_precio = Val(vaSpread1.Value)
                
                vaSpread1.Col = 6
                v_pctdes = Val(vaSpread1.Value)
                
                vaSpread1.Col = 9
                v_canrec = Val(vaSpread1.Value)
                
                vaSpread1.Col = 10
                v_prerec = Val(vaSpread1.Value)
                
                If v_can > v_canrec Or v_precio > v_prerec Then
                    
                    vaSpread1.Col = 16
                    vaSpread1.Value = 1
                    
                    vaSpread1.Col = 7
                    v_totdesc = v_totdesc + Val(vaSpread1.text)
                    
                    'Sumo solamente las lineas con diferencias (cant =10 ;recib= 9 ;cantcal =1)
                    SumaDiferencias i, (v_can - v_canrec), v_prerec, v_pctdes
                    Cdif = Cdif + 1
                
                End If
            
            Next i
            
            If Cdif > 0 Then 'Si existen diferencias
                
                MsgBox "Documento con diferencias. Se emitira solicitud de nota de crédito... ", vbInformation + vbOKOnly, MsgTitulo
            
            Else
                
                modo = "A"
                Gl_Ac_Botones Me, 7, 2, ""
                Frame2.Enabled = False
                Frame3.Enabled = False
                Frame6.Enabled = False
                
                vaSpread1.Row = -1
                vaSpread1.Col = -1
                vaSpread1.Lock = True
                
                vaSpread2.Row = -1
                vaSpread2.Col = -1
                vaSpread2.Lock = True
                
                vg_RDC = fg_DespintaRut(fpText(0).text)
                vg_TDC = fg_codigocbo(Combo2, 0, 2, "")
                vg_NDC = Val(Double1(5).Value)
                vg_NSOL = TipoDato(FolioSn, 0)
                
                fg_carga ""
                vg_db.CommitTrans
                I_DocProvee
                Toolbar1.Enabled = True
                Exit Sub
            
            End If

            If RS1.State = 1 Then RS1.Close
            RS1.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            RS1.Open "SELECT toc_numdoc FROM b_totcompras WHERE  toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN') ORDER BY toc_numdoc DESC", vg_db, adOpenStatic
            If Not RS1.EOF Then
               
               FolioSn = RS1!toc_numdoc + 1
            
            Else
            
               FolioSn = 1
            
            End If
            RS1.Close
            Set RS1 = Nothing
            
            If vg_tipbase = "1" Then
               
               vg_db.Execute "INSERT INTO b_totcompras (toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc,toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem, EnvioDocSGPADM, toc_docasotipo) " & _
                             "VALUES ('" & v_rut & "', 'SN'," & FolioSn & "," & v_bodega & ",'" & v_fecemi & "','" & v_fecven & "'," & v_totdesc & "," & vTotNet & "," & vTotExe & "," & vTotNet & "," & vTotOtr & "," & vTotTot & "," & vTotTot & ",'" & opcion & "'," & Str(Double1(6).Value) & ", " & Str(Double1(5).Value) & ", '" & Trim(fpText(1).text) & "', 0, '', '0', '" & Format(Date, "dd/mm/yyyy") & " " & Format(Time, "h:m:s") & "', " & periodo & ", '" & v_fecrem & "', '0', '')"
            
            Else
               
               vg_db.Execute "INSERT INTO b_totcompras (toc_rutpro, toc_tipdoc, toc_numdoc, toc_codbod, toc_fecemi, toc_fecven, toc_desdoc, toc_netdoc, toc_exedoc,toc_ivadoc, toc_otrimp, toc_totdoc, toc_pendoc, toc_tipinf, toc_numinf, toc_docaso, toc_ordcom, toc_fledoc, toc_docsnc, toc_envsap, toc_fecdig, toc_fecper, toc_fecrem, EnvioDocSGPADM, toc_docasotipo) " & _
                             "VALUES ('" & v_rut & "', 'SN'," & FolioSn & "," & v_bodega & ",'" & Format(v_fecemi, "yyyymmdd") & "','" & Format(v_fecven, "yyyymmdd") & "'," & v_totdesc & "," & vTotNet & "," & vTotExe & "," & vTotNet & "," & vTotOtr & "," & vTotTot & "," & vTotTot & ",'" & opcion & "'," & Str(Double1(6).Value) & ", " & Str(Double1(5).Value) & ", '" & Trim(fpText(1).text) & "', 0, '', '0', '" & Format(Date, "yyyymmdd") & " " & Format(Time, "h:m:s") & "', " & periodo & ", '" & Format(v_fecrem, "yyyymmdd") & "', '0', '')"
            
            End If
            
            For i = 1 To vaSpread1.MaxRows
                
                vaSpread1.Row = i
                vaSpread1.Col = IIf(vg_pais = "CO", 24, 1)
                codpro = Trim(vaSpread1.text)
                
                vaSpread1.Col = 29
                v_faccon = Val(vaSpread1.Value)
                
                If vg_pais = "CO" Then
                   
                   vaSpread1.Col = 4
                   v_cmefac = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 5
                   v_pmefac = Val(vaSpread1.Value)
                   
                   v_can = (v_cmefac * v_faccon)
                   
                   If v_can > 0 Then
                      
                      v_precio = Round(((v_cmefac * v_pmefac) / v_can), 2)
                   
                   Else
                      
                      v_precio = 0
                   
                   End If
                
                Else
                   
                   vaSpread1.Col = 4
                   v_can = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 5
                   v_precio = Val(vaSpread1.Value)
                
                End If
                
                vaSpread1.Col = 6
                v_pctdes = Val(vaSpread1.Value)
                
                vaSpread1.Col = 7
                v_valdes = Val(vaSpread1.Value)
                
                vaSpread1.Col = 8
                v_total = Val(vaSpread1.Value)
                
                If vg_pais = "CO" Then
                   
                   vaSpread1.Col = 9
                   v_crefac = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 10
                   v_prefac = Val(vaSpread1.Value)
                   v_canrec = (v_crefac * v_faccon)
                   
                   If v_canrec > 0 Then
                      
                      v_prerec = Round(((v_crefac * v_prefac) / v_canrec), 2)
                   
                   Else
                      
                      v_prerec = Round(((v_cmefac * v_prefac) / v_can), 2)
                   
                   End If
                
                Else
                   
                   vaSpread1.Col = 9
                   v_canrec = Val(vaSpread1.Value)
                   
                   vaSpread1.Col = 10
                   v_prerec = Val(vaSpread1.Value)
                
                End If
                
    '            vaSpread1.Col = 9: v_canrec = Val(vaSpread1.Value)
    '            vaSpread1.Col = 10: v_prerec = Val(vaSpread1.Value)
                
                vaSpread1.Col = 11
                v_descrip = Trim(vaSpread1.text)
                
                vaSpread1.Col = 12
                v_CtaCon = Trim(vaSpread1.text)
                
                If vg_pais = "CO" Then
                   
                   vaSpread1.Col = 15
                   PreC = Val(vaSpread1.text)
                   PreC = Round(((v_cmefac * PreC) / v_can), 2)
                   
                   vaSpread1.Col = 17
                   vPreDec = Val(vaSpread1.Value)
                   vPreDec = Round(((v_cmefac * vPreDec) / v_can), 3)
                
                Else
                   
                   vaSpread1.Col = 15
                   PreC = Val(vaSpread1.text)
                   
                   vaSpread1.Col = 17
                   vPreDec = Val(vaSpread1.Value)
                
                End If
                
                vTotRec = IIf(v_can = v_canrec And v_precio = v_prerec, v_total, Round(v_canrec * v_prerec, 0))
                v_Rebaja = "S"
                
                If v_can > v_canrec Or v_precio > v_prerec Then
                   
                   vg_db.Execute "INSERT INTO b_detcompras (dec_rutpro, dec_tipdoc, dec_numdoc, dec_numlin, dec_codmer, dec_canmer, dec_precom, dec_pctdes, dec_valdes, dec_ptotal, dec_descri, dec_canrec, dec_prerec, dec_mueinv, dec_prefle, dec_ptotrec, dec_cmefac, dec_pmefac, dec_crefac, dec_prefac, dec_faccon) " & _
                                 "VALUES ('" & v_rut & "', '" & "SN" & "', " & FolioSn & ", " & i & ", '" & codpro & "', " & v_can & ", " & v_precio & ", " & v_pctdes & ", " & v_valdes & ", " & v_total & ", '" & TipoDato(v_descrip, "") & "', " & v_canrec & ", " & v_prerec & ", '" & v_Rebaja & "', 0, " & vTotRec & ", " & v_cmefac & ", " & v_pmefac & ", " & v_crefac & ", " & v_prefac & ", " & v_faccon & ")"
                
                End If
             
             Next i
        
             vg_db.Execute "INSERT INTO b_detcomprasimp (imd_rutdoc, imd_tipdoc, imd_numdoc, imd_numlin, imd_codpro, imd_codimp, imd_pctimp, imd_monimp)  " & _
                           "SELECT DISTINCT a.dec_rutpro, a.dec_tipdoc, a.dec_numdoc, a.dec_numlin, a.dec_codmer, b.ipr_codimp, c.imp_pctimp, ((((a.dec_canmer-a.dec_canrec)*a.dec_prerec) - ((a.dec_canmer-a.dec_canrec)*a.dec_prerec)*(a.dec_pctdes / 100))*(c.imp_pctimp/100)) " & _
                           "FROM b_detcompras a, b_productosimp b, a_impuesto c WHERE a.dec_codmer=b.ipr_codpro AND b.ipr_codimp=c.imp_codigo AND a.dec_rutpro='" & v_rut & "' AND a.dec_tipdoc='SN' AND a.dec_numdoc=" & FolioSn
             
             '-------> Actualizar impuesto real al b_detcomprasimp
             For i = 1 To vaSpread1.MaxRows
                 
                 vaSpread1.Row = i
                 vaSpread1.Col = IIf(vg_pais = "CO", 24, 1): codpro = Trim(vaSpread1.text)
                 
                 vaSpread1.Col = 14
                 StrImp = Trim(vaSpread1.text)
                 
                 If Len(StrImp) <> 0 Then
                    
                    Do While InStr(StrImp, ";") <> 0
                       
                       StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                       StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                       codimp = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                       PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                       CosI = Val(Mid(StrImpb, 1))
                       vg_db.Execute "UPDATE b_detcomprasimp SET imd_pctimp = " & PctI & " WHERE imd_codpro = '" & codpro & "' AND imd_codimp = " & codimp & " AND imd_rutdoc = '" & v_rut & "' AND imd_tipdoc = 'SN' AND imd_numdoc = " & FolioSn & ""
                       vg_db.Execute "UPDATE b_detcomprasimp SET imd_monimp = 0 WHERE imd_pctimp = 0 AND imd_codpro = '" & codpro & "' AND imd_codimp = " & codimp & " AND imd_rutdoc = '" & v_rut & "' AND imd_tipdoc = 'SN' AND imd_numdoc = " & FolioSn & ""
                    
                    Loop
                 
                 End If
             
             Next i
        
        End If
        
        vg_db.CommitTrans
        '----Fin Chequeo de Diferencias---
        'MsgBox "Grabado ok...", vbInformation + vbOKOnly, MsgTitulo
        modo = "A"
        Gl_Ac_Botones Me, 7, 2, ""
    '    vaSpread2.MaxRows = 0
        Frame2.Enabled = False
        Frame3.Enabled = False
        Frame6.Enabled = False
        
        vaSpread1.Row = -1
        vaSpread1.Col = -1
        vaSpread1.Lock = True
        
        vaSpread2.Row = -1
        vaSpread2.Col = -1
        vaSpread2.Lock = True
        
        vg_RDC = fg_DespintaRut(fpText(0).text)
        vg_TDC = fg_codigocbo(Combo2, 0, 2, "")
        vg_NDC = Val(Double1(5).Value)
        vg_NSOL = TipoDato(FolioSn, 0)
        
        fg_carga ""
        
        If Trim(fg_codigocbo(Combo2, 0, 2, "")) <> "CG" Then I_DocProvee Else I_ComprobanteGasto
        Toolbar1.Enabled = True
        
    Case 3 '-------> Borrar
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT DISTINCT toc_fecemi, toc_fecper FROM b_totcompras WHERE toc_rutpro = '" & v_rut & "' AND toc_tipdoc = '" & v_tipo & "' AND toc_numdoc = " & Val(Double1(5).text) & " AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: Exit Sub
        periodo = RS1!toc_fecper
        RS1.Close
        Set RS1 = Nothing
        
        If CierrePeriodo(periodo, v_bodega, 13) Then MsgBox "Periodo esta cerrado...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If CierrePeriodo(Format(Date1(2).text, "yyyymmdd"), v_bodega, 6) And Format(CDate(vg_ciedia) - 1, "mm/yyyy") = Format(CDate(Date1(2).Value), "mm/yyyy") Then MsgBox "No puede eliminar documentos anteriores a la última toma de inventario.", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If CDate(Date1(2).Value) < CDate(vg_ciedia) And Format(CDate(vg_ciedia) - 1, "mm/yyyy") = Format(CDate(Date1(2).Value), "mm/yyyy") Then MsgBox "No puede eliminar documentos, día esta cerrado...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        '-------> Validar si Documento fue enviado sap
        If Option1(1).Value = True And ValidarDocumentoSap(v_rut, v_tipo, Val(Double1(5).text), vg_codbod, MuestraCasino(1)) Then MsgBox "Documento no puede ser borrado, fue enviado CFC a SAP...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        Borra_Datos
        Frame2.Enabled = True
        Frame3.Enabled = True
        Frame6.Enabled = True
        vaSpread1.Row = -1
        vaSpread1.Col = -1
        vaSpread1.Lock = False
        
        Nuevo_Registro
        Gl_Ac_Botones Me, 7, 1, ""
        
    Case 11 '-------> Imprimir
        
        If Trim(fpText(0).text) = "" Or Combo2(0).ListIndex < 0 Or Val(Double1(5).Value) = 0 Then MsgBox "No existe documento...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        vg_RDC = fg_DespintaRut(fpText(0).text)
        vg_TDC = fg_codigocbo(Combo2, 0, 2, "")
        vg_NDC = Val(Double1(5).Value)
        vg_NSOL = TipoDato(FolioSn, 0)
        fg_carga ""
        If Trim(fg_codigocbo(Combo2, 0, 2, "")) <> "CG" Then I_DocProvee Else I_ComprobanteGasto
        
    Case 14 '-------> Salir
        
        Me.Hide
        Unload Me
        
End Select

Exit Sub
Man_Error:
Toolbar1.Enabled = True
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
If Err.Number = -2147467259 Then vg_db.RollbackTrans: MsgBox "Documento ya existe...", vbExclamation + vbOKOnly, MsgTitulo: Resume Next: Exit Sub
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
vg_db.RollbackTrans
fg_descarga

End Sub

Private Sub Borra_Datos()

On Error GoTo Man_Error

Dim RS            As New ADODB.Recordset
Dim codigo        As String
Dim v_bodega      As Long
Dim v_cant        As Double
Dim actbod        As Boolean
Dim tipaux        As String
Dim rut           As String
Dim codpro        As String
Dim i             As Long
Dim VecGuia()     As String
Dim VecGuiaTipo() As String


'-------> Obtiene rut, tipo de docto y  Nº de Docto.
rut = fg_DespintaRut(fpText(0).text)
codigo = fg_codigocbo(Combo2, 0, 2, "")
Num = Val(Double1(5).Value)
If Num = 0 Then Exit Sub

'-------> Fin Obtiene Documento
If MsgBox("¿Desea Eliminar Documento?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    v_bodega = fg_codigocbo(Combo2, 1, 10, 0)
    actbod = True
    vg_db.BeginTrans
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("SELECT isnull(btc.toc_docaso,'') as toc_docaso, isnull(btc.toc_docasoTipo,'') as toc_docasoTipo " & _
                           "FROM b_totcompras as btc with (nolock) " & _
                           "WHERE btc.toc_rutpro = '" & rut & "' " & _
                           "AND btc.toc_tipdoc = '" & codigo & "' " & _
                           "AND btc.toc_numdoc = " & Num & " " & _
                           "AND btc.toc_codbod = " & vg_codbod & "")
    If Not RS.EOF Then
        
        tipaux = ""
        If fg_TraerRelacionTipoDocumento(codigo) = "FA" Or fg_TraerRelacionTipoDocumento(codigo) = "FE" Then tipaux = "GD"
        If fg_TraerRelacionTipoDocumento(codigo) = "NC" Or fg_TraerRelacionTipoDocumento(codigo) = "CE" Then tipaux = "SN"
        
        If tipaux = "SN" Then
           
           vg_db.Execute "UPDATE b_totcompras SET toc_docsnc = '' WHERE toc_rutpro = '" & rut & "' AND toc_tipdoc in (select tdo_Codigo from a_tipodocumento where tdo_IdCodigo = '" & tipaux & "' ) AND toc_docsnc = '" & Trim(Str(Num)) & "' AND toc_codbod = " & vg_codbod & ""
        
        ElseIf tipaux = "GD" Then
           
           If Trim(RS!toc_docaso) <> "" Then
              
              VecGuia = Split(Trim(RS!toc_docaso), ";")
           
              If Trim(RS!toc_docasotipo) <> "" Then
              
                 VecGuiaTipo = Split(Trim(RS!toc_docasotipo), ";")
              
              End If
              
              For i = 0 To UBound(VecGuia) - 1
              
'                  vg_db.Execute "UPDATE b_totcompras SET toc_docaso = '' WHERE toc_rutpro = '" & rut & "' " & _
'                                "AND toc_tipdoc in (select tdo_Codigo from a_tipodocumento as a with (nolock) where tdo_IdCodigo = '" & tipaux & "') " & _
'                                "AND toc_docaso = '" & Trim(Str(Num)) & "' AND toc_codbod = " & vg_codbod & ""

                  If Trim(VecGuia(i)) <> "" And Trim(RS!toc_docasotipo) <> "" Then
                     
                     vg_db.Execute "UPDATE b_totcompras SET toc_docaso = '' WHERE toc_rutpro = '" & rut & "' " & _
                                   "AND toc_tipdoc = '" & Trim(VecGuiaTipo(i)) & "' " & _
                                   "AND toc_docaso = '" & Trim(Str(Num)) & "' AND toc_codbod = " & vg_codbod & " " & _
                                   "AND toc_numdoc = " & Str(VecGuia(i)) & ""
                                   
                  ElseIf Trim(VecGuia(i)) <> "" And Trim(RS!toc_docasotipo) = "" Then
                     
                     vg_db.Execute "UPDATE b_totcompras SET toc_docaso = '' WHERE toc_rutpro = '" & rut & "' " & _
                                   "AND toc_tipdoc in (select tdo_Codigo from a_tipodocumento as a with (nolock) where tdo_IdCodigo = '" & tipaux & "') " & _
                                   "AND toc_docaso = '" & Trim(Str(Num)) & "' AND toc_codbod = " & vg_codbod & " " & _
                                   "AND toc_numdoc = " & Str(VecGuia(i)) & ""
                                  
                  End If
           
              Next i
              
           End If
           
        End If
        actbod = IIf(Len(Trim(RS!toc_docaso)) > 0, False, True)
    
    End If
    RS.Close
    Set RS = Nothing
    
    vg_db.Execute "DELETE FROM b_detcomprasimp WHERE imd_rutdoc = '" & rut & "' AND imd_tipdoc = '" & codigo & "' AND imd_numdoc = " & Num
    vg_db.Execute "DELETE FROM b_detcompras WHERE dec_rutpro = '" & rut & "' AND dec_tipdoc = '" & codigo & "' AND dec_numdoc = " & Num
    vg_db.Execute "DELETE FROM b_totcompras WHERE toc_rutpro = '" & rut & "' AND toc_tipdoc = '" & codigo & "' AND toc_numdoc = " & Num & " AND toc_codbod = " & vg_codbod & ""
    vg_db.Execute "DELETE FROM b_ocsacrecibido WHERE ocr_rutpro = '" & rut & "' AND ocr_tipdoc = '" & codigo & "' AND ocr_numdoc = " & Num & ""
    
    '------> Borrar solicitud nota credito
    If fg_TraerRelacionTipoDocumento(codigo) = "FA" Or fg_TraerRelacionTipoDocumento(codigo) = "FE" Then
    
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
        
        RS.Open "SELECT toc_numdoc, toc_docsnc FROM b_totcompras WHERE toc_rutpro = '" & rut & "' AND toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo = 'SN')  " & _
                "AND toc_docaso = '" & Trim(Str(Num)) & "' AND toc_codbod = " & vg_codbod & "", vg_db, adOpenStatic
    
        If Not RS.EOF Then
       
            If Trim(RS!toc_docsnc) <> "" And Not IsNull(RS!toc_docsnc) Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "Documento esta asociado Solicitud de Nota Credito... ", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
            vg_db.Execute "DELETE FROM b_detcomprasimp WHERE imd_rutdoc = '" & rut & "' AND imd_tipdoc = 'SN' AND imd_numdoc = " & RS!toc_numdoc
            vg_db.Execute "DELETE FROM b_detcompras WHERE dec_rutpro = '" & rut & "' AND dec_tipdoc = 'SN' AND dec_numdoc = " & RS!toc_numdoc
            vg_db.Execute "DELETE FROM b_totcompras WHERE toc_rutpro = '" & rut & "' AND toc_tipdoc = 'SN' AND toc_numdoc = " & RS!toc_numdoc & " AND toc_codbod = " & vg_codbod & ""
    
        End If
        RS.Close
        Set RS = Nothing
    
        '------> Fin borrar solicitud nota credito
    
    End If
    
    codpro = ""
    If actbod Then
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 13
            
            If Trim(vaSpread1.text) = "S" Then
                
                vaSpread1.Row = i
                vaSpread1.Col = IIf(vg_pais = "CO", 24, 1)
                codpro = Trim(vaSpread1.text)

                vaSpread1.Col = IIf(vg_pais <> "CO", 9, 31)
                v_cant = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
                
                If Trim(fg_TraerRelacionTipoDocumento(codigo)) <> "NC" And Trim(fg_TraerRelacionTipoDocumento(codigo)) <> "CE" Then
                    
                    '-------> Validar si existen diferencia en bodega
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    RS.Open "SELECT bod_canmer FROM b_bodegas WHERE bod_codbod = " & v_bodega & " " & _
                            "AND bod_codpro = '" & codpro & "'", vg_db, adOpenStatic
                    If Not RS.EOF Then If (RS!bod_canmer - IIf(v_cant < 0, (v_cant * -1), v_cant)) <= -1 Then RS.Close: Set RS = Nothing: vg_db.RollbackTrans: MsgBox "Documento no puede ser eliminado. Existen diferencia ...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
                    RS.Close
                    Set RS = Nothing
                    
                    '-------> Fin validar si existen diferencia en bodega
                    vg_db.Execute "UPDATE b_bodegas SET bod_canmer = bod_canmer - " & v_cant & " WHERE bod_codbod = " & v_bodega & " AND bod_codpro = '" & codpro & "'"
                    vg_db.Execute "UPDATE b_totcompras SET toc_docaso = '' WHERE toc_rutpro = '" & rut & "' AND toc_tipdoc = '" & codigo & "' AND toc_docaso = '" & Trim(Str(Num)) & "' AND toc_codbod = " & vg_codbod & ""
                
                End If
                
            End If
            
        Next i
        
    End If
    
    vg_db.CommitTrans
'*    '-------> rutinar de recalculo de precio
'*    If Trim(CodPro) <> "" Then RecalPrecioDoc Format(Date1(2).text, "yyyymmdd"), fg_codigocbo(Combo2, 1, 1, 0), CodPro
    fpText(0).text = ""
    fpayuda(0).Caption = ""
    
    Double1(5).text = ""
    Double1(6).text = ""
    Double1(12).text = ""
    Double1(13).text = ""
    Double1(14).text = ""
    Double1(15).text = ""
    Double1(16).text = ""
    Double1(17).text = ""
    
    Date1(0).text = ""
    Date1(1).text = ""
    Date1(2).text = ""
    vaSpread1.MaxRows = 0
    vaSpread1.Col = 1
    vaSpread1.Row = 1
    modo = "N"
    
Exit Sub
Man_Error:
    If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, MsgTitulo: Exit Sub
    MsgBox "¡Datos no Eliminados!", vbCritical, MsgTitulo
    vg_db.RollbackTrans
    fg_descarga
    
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim i As Long
Dim cCta As String, cPro As String, fecoc As String, sql1 As String
Dim RS As New ADODB.Recordset
Select Case Button.Index
    
    Case 1
        
        Toolbar2.Enabled = False
        StopObjeto False
        vg_nombre = ""
        vg_codigo = ""
        vg_left = fpayuda(0).Left + 1920
        Text2(0).text = ""
        Text2(1).text = ""
        Text2(2).text = ""
        
        If vg_pais = "CO" Then
           
           B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Productos SAC", "PSAC"
        
        Else
           
           B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", IIf(Trim(fg_codigocbo(Combo2, 0, 2, "")) = "CG", "ProGrl", "ProVig")
        
        End If
        
        B_TabEst.Show 1
        If vg_codigo = "" Then Toolbar2.Enabled = True: vaSpread1.ProcessTab = True: Exit Sub
        Toolbar2.Enabled = True
        vaSpread1.ProcessTab = False
    '    SendKeys "+{BREAK}"
        If vg_pais <> "CL" Or vg_FDC <> "OC" Then
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Col = 1
               vaSpread1.Row = i
               If Trim(vaSpread1.text) = Trim(vg_codigo) Then Frame5.Enabled = False: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
           
           Next i
        
        End If
        sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        If vg_pais = "CO" Then
           
           RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, " & _
                   "a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                   "c.foc_codsac, c.foc_nomsac, c.foc_unisac, c.foc_faccon " & _
                   "FROM b_productos a, a_unidad b, b_formatocompras c, b_formatocomprassgp d " & _
                   "WHERE c.foc_codsac = d.fcs_codsac " & _
                   "AND   d.fcs_codsgp = a.pro_codigo " & _
                   "AND   a.pro_coduni = b.uni_codigo " & _
                   "AND   c.foc_codsac = '" & vg_codigo & "'", vg_db, adOpenStatic
           Text2(0).Visible = True: Text2(1).Visible = True: Text2(2).Visible = True
        
        Else
    '               "'' AS foc_codsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon "
           RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, " & _
                   "a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                   "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
                   "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
                   "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
                   "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
                   "FROM  b_productos a, a_unidad b " & _
                   "WHERE a.pro_coduni = b.uni_codigo " & _
                   "AND   a.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
        
        End If
        Text2(0).text = ""
        Text2(1).text = ""
        Text2(2).text = ""
        
        If Not RS.EOF Then
            
            If IsNull(RS!pro_ctrsto) Then Frame5.Enabled = False: MsgBox "Producto no tiene asignado, el Movimiento...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
            If RS!pro_facing = 0 Or RS!pro_facsto = 0 Then RS.Close: Set RS = Nothing: MsgBox "Factor del producto en cero...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
            cPro = RS!pro_codigo
            
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = 1
            
            If vaSpread1.Lock = True Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
            
            '-------> desbloquear celda
            For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = False: Next i
            
            vaSpread1.Row = vaSpread1.MaxRows
            vaSpread1.Col = 1
            vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac), IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo))
            
            vaSpread1.Col = 2
            vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac), IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre))
            
            vaSpread1.Col = 3
            vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_unisac), "No existe descripción U.M. SAC", RS!foc_unisac), IIf(IsNull(RS!uni_nomcor), "No existe descripción U.M. SGP", RS!uni_nomcor))
            
            vaSpread1.Col = 12
            vaSpread1.text = RS!pro_ctacon
            
            vaSpread1.Col = 11
            vaSpread1.text = ""
            
            vaSpread1.Col = 18
            vaSpread1.text = IIf(IsNull(RS!ppd_propon), 0, RS!ppd_propon)
            
            vaSpread1.Col = 24
            vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo), IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac))
            
            vaSpread1.Col = 25
            vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre), IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac))
            
            vaSpread1.Col = 26
            vaSpread1.text = IIf(IsNull(RS!pro_codref), 0, RS!pro_codref)
            
            vaSpread1.Col = 27
            vaSpread1.text = IIf(IsNull(RS!pro_codrei), 0, RS!pro_codrei)
            
            vaSpread1.Col = 29
            vaSpread1.text = IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon)
            
            vaSpread1.Col = 24
            Text2(0).text = Trim(vaSpread1.text)
            
            vaSpread1.Col = 25
            Text2(1).text = Trim(vaSpread1.text)
            
            vaSpread1.Col = 29
            Text2(2).text = vaSpread1.text
            
            For i = 4 To 10: vaSpread1.Col = i: vaSpread1.text = Format(0, fg_Pict(9, IIf(i = 4 Or i = 9, vg_DCa, 2))): Next i
            
            If Trim(RS!pro_ctacon) = "" Or IsNull(RS!pro_ctacon) Then RS.Close: Set RS = Nothing: MsgBox "El producto no tiene asosiada una cuenta contable...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
            vaSpread1.Col = 13
            vaSpread1.text = IIf(RS!pro_ctrsto = 1, "S", "N")
        
        Else
           
           '-------> bloquear celda
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.Col = 1: vaSpread1.text = ""
           vaSpread1.Col = 2: vaSpread1.text = "": Text1(0).text = ""
           vaSpread1.Col = 3: vaSpread1.text = ""
           vaSpread1.Col = 24: vaSpread1.text = ""
           vaSpread1.Col = 25: vaSpread1.text = ""
           vaSpread1.Col = 26: vaSpread1.text = ""
           vaSpread1.Col = 27: vaSpread1.text = ""
           vaSpread1.Col = 29: vaSpread1.text = ""
           vaSpread1.Col = 24: Text2(0).text = Trim(vaSpread1.text)
           vaSpread1.Col = 25: Text2(1).text = Trim(vaSpread1.text)
           vaSpread1.Col = 29: Text2(2).text = vaSpread1.text
           For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = True: Next i
        
        End If
        RS.Close
        Set RS = Nothing
        est2 = False
        Revisa cPro, vaSpread1.ActiveRow
        vaSpread1.EditModePermanent = True
        vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
        vaSpread1.Refresh
        If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
    
    Case 2
        
        If vaSpread1.MaxRows = 0 Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 23
        fecoc = ""
        fecoc = Trim(vaSpread1.text)
        vaSpread1.Col = 1
        If vaSpread1.Lock = False Or fecoc <> "" Then Exit Sub
        If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        If vaSpread1.MaxRows = 0 Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: Text1(0).text = ""
        If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
        SumarTotales

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Revisa(codpro As String, Row As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim v_rut As String, estrut As Boolean
Dim regimp As String, autoret As String, cuohor As String
Dim codmun As Long

vaSpread1.Col = 14
vaSpread1.text = ""
v_rut = fg_DespintaRut(fpText(0).text)
estrut = False
estrut = ValidarRetProveedor(v_rut)

If estrut Then
   
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   '-------> sacar datos de parametros
   RS.Open "SELECT prv_regimp, prv_autret, prv_cuohor, prv_codmun " & _
           "FROM b_proveedor " & _
           "WHERE prv_codigo = '" & v_rut & "' AND prv_regimp IS NOT NULL", vg_db, adOpenStatic
   If Not RS.EOF Then
      
      regimp = IIf(IsNull(RS!prv_regimp), "0", RS!prv_regimp)
      autret = IIf(IsNull(RS!prv_autret), "S", RS!prv_autret)
      cuohor = IIf(IsNull(RS!prv_cuohor), "N", RS!prv_cuohor)
      codmun = IIf(IsNull(RS!prv_codmun), 0, RS!prv_codmun)
   
   End If
   RS.Close
   Set RS = Nothing

End If
estrut = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
RS.Open "SELECT a.*, b.*, c.pro_codref, c.pro_codrei, c.pro_cuohor " & _
        "FROM   b_productosimp a, a_impuesto b, b_productos c " & _
        "WHERE  a.ipr_codimp = b.imp_codigo " & _
        "AND    a.ipr_codpro = c.pro_codigo " & _
        "AND    a.ipr_codpro = '" & codpro & "'", vg_db, adOpenStatic
Do While Not RS.EOF
   
   vaSpread1.Col = vaSpread1.ActiveCol
   vaSpread1.Row = Row
'   If estrut And (CStr(RS!ipr_codimp) = GetParametro("parretfue") Or CStr(RS!ipr_codimp) = GetParametro("parretica")) Then
'      '-------> Colocar impuesto retención fuente
'      If CStr(RS!ipr_codimp) = GetParametro("parretfue") And (regimp = "1" Or (regimp = "2" And autret = "N") Or (regimp = "3" And autret = "N")) Then
'         '-------> Colocar impuesto retención Fuente
'         If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone) <> -1 Then
'            vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone)
'            vaSpread2.Col = 4: vaSpread2.text = IIf(IsNull(RetencionFuente(RS!pro_codref)), Format(0, fg_Pict(3, 2)), Format(RetencionFuente(RS!pro_codref), fg_Pict(3, 2))) & " %"
'            vaSpread2.Col = 7: vaSpread2.text = IIf(IsNull(RetencionFuente(RS!pro_codref)), 0, RetencionFuente(RS!pro_codref))
'         End If
'         vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(RS!ipr_codimp)) & "&"
'         vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RetencionFuente(RS!pro_codref)), 0, RetencionFuente(RS!pro_codref)))) & "&"
'         vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos))) & ";"
'      ElseIf CStr(RS!ipr_codimp) = GetParametro("parretica") And regimp <> "3" Then
'         '-------> Colocar impuesto retención ica
'         If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone) <> -1 Then
'            vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone)
'            vaSpread2.Col = 4: vaSpread2.text = IIf(IsNull(RetencionIca(v_rut, RS!pro_codrei)), Format(0, fg_Pict(3, 2)), Format(RetencionIca(v_rut, RS!pro_codrei), fg_Pict(3, 2))) & " %"
'            vaSpread2.Col = 7: vaSpread2.text = IIf(IsNull(RetencionIca(v_rut, RS!pro_codrei)), 0, RetencionIca(v_rut, RS!pro_codrei))
'         End If
'         vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(RS!ipr_codimp)) & "&"
'         vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RetencionIca(v_rut, RS!pro_codrei)), 0, RetencionIca(v_rut, RS!pro_codrei)))) & "&"
'         vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos))) & ";"
'      End If
'   Else
'      If estrut Then
'         If regimp = "1" Or regimp = "2" Then
'            If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone) <> -1 Then
'               vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone)
'               vaSpread2.Col = 4: vaSpread2.text = IIf(IsNull(RS!imp_pctimp), Format(0, fg_Pict(3, 2)), Format(((RS!imp_pctimp * 50) / 100), fg_Pict(3, 2))) & " %"
'               vaSpread2.Col = 7: vaSpread2.text = IIf(IsNull(RS!imp_pctimp), 0, ((RS!imp_pctimp * 50) / 100))
'            End If
'            vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(RS!ipr_codimp)) & "&"
'            vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_pctimp), 0, ((RS!imp_pctimp * 50) / 100)))) & "&"
'            vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos))) & ";"
'         Else
'            vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(RS!ipr_codimp)) & "&"
'            vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_pctimp), 0, RS!imp_pctimp))) & "&"
'            vaSpread1.Col = 14: vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos))) & ";"
'         End If
'      Else
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(RS!ipr_codimp)) & "&"
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_pctimp), 0, RS!imp_pctimp))) & "&"
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos))) & ";"

'      End If
'   End If
   
   '-------> Validar si aplica cuota hortofruticola
   If RS!pro_cuohor = "S" And cuohor = "S" Then
      
      If CStr(RS!ipr_codimp) = GetParametro("parrethorf") Then
         
         '-------> Colocar impuesto retención hortofruticola
         If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone) <> -1 Then
            
            vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(RS!ipr_codimp), SearchFlagsNone)
            vaSpread2.Col = 4
            vaSpread2.text = IIf(IsNull(GetParametro("parhorfru")), Format(0, fg_Pict(3, 2)), Format(GetParametro("parhorfru"), fg_Pict(3, 2))) & " %"
            
            vaSpread2.Col = 7
            vaSpread2.text = IIf(IsNull(GetParametro("parhorfru")), 0, GetParametro("parhorfru"))
         
         End If
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(RS!ipr_codimp)) & "&"
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(GetParametro("parhorfru")), 0, GetParametro("parhorfru")))) & "&"
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(IIf(IsNull(RS!imp_inccos), 0, RS!imp_inccos))) & ";"
      
      End If
   
   End If
   
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Advance(ByVal AdvanceNext As Boolean)

On Error GoTo Man_Error

Dim RS     As New ADODB.Recordset
Dim codigo As String
Dim Nombre As String
Dim i      As Long

If vaSpread1.MaxRows < 1 Or Frame6.Enabled = False Then Exit Sub

vaSpread1.Row = vaSpread1.ActiveRow

vaSpread1.Col = 1
codigo = Trim(vaSpread1.text)

vaSpread1.Col = 2
Nombre = Trim(vaSpread1.text)

If AdvanceNext = True And codigo <> "" And Nombre <> "" Then
   
   vaSpread1.Col = 1
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   If vg_pais = "CO" Then
      
      RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, " & _
              "a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
              "c.foc_codsac, c.foc_nomsac, c.foc_unisac, c.foc_faccon " & _
              "FROM b_productos a, a_unidad b, b_formatocompras c, b_formatocomprassgp d " & _
              "WHERE c.foc_codsac = d.fcs_codsac " & _
              "AND   d.fcs_codsgp = a.pro_codigo " & _
              "AND   a.pro_coduni = b.uni_codigo " & _
              "AND   c.foc_codsac = '" & LimpiaDato(Trim(vaSpread1.text)) & "' " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0)", vg_db, adOpenStatic
   
   Else
      
      RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, " & _
              "(SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND " & _
              "ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon " & _
              "FROM  b_productos a, a_unidad b " & _
              "WHERE a.pro_coduni = b.uni_codigo " & _
              "AND   a.pro_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "' " & _
              "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0)", vg_db, adOpenStatic
   
   End If
   
   If RS.EOF Then
      
      RS.Close
      Set RS = Nothing
      vaSpread1.Row = vaSpread1.ActiveRow
      
      vaSpread1.Col = 1
      vaSpread1.text = ""
      
      vaSpread1.Col = 2
      vaSpread1.text = ""
      Text1(0).text = ""
      
      vaSpread1.Col = 3
      vaSpread1.text = ""
      
      For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = True: Next i
      
      vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
      MsgBox "producto no existe...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
   
   End If
   RS.Close
   Set RS = Nothing
   vaSpread1.Col = 1
   vaSpread1.Lock = True
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
   vaSpread1.Row = vaSpread1.MaxRows
   
   For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = False: Next i

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim codsgp As String
If vaSpread1.MaxRows < 1 Then Exit Sub
If Encontrado = True Then
   
   vaSpread1.Col = IIf(Trim(fg_codigocbo(Combo2, 0, 2, "")) <> "CG", 2, 11): vaSpread1.Row = vaSpread1.ActiveRow
'20080417   Text1(0).text = vaSpread1.text

End If

vaSpread1.Row = vaSpread1.ActiveRow

vaSpread1.Col = 24
Text2(0).text = Trim(vaSpread1.text)

vaSpread1.Col = 25
Text2(1).text = Trim(vaSpread1.text)

vaSpread1.Col = 29
Text2(2).text = vaSpread1.text

vaSpread1.Col = 1
codsgp = Trim(vaSpread1.text)

If vg_pais = "CL" And vg_FDC = "OC" And Trim(codsgp) <> "" Then
   
   Image1(5).Visible = IIf(ValidarProductosSgpSac(Trim(Text2(0).text), Trim(codsgp)), True, False)

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim fecoc As String, i As Long
Dim codsac As String, cPro As String, sql1 As String

If vaSpread1.MaxRows < 1 Or Col <> 1 Or vg_FDC <> "OC" Then Exit Sub

vaSpread1.Row = Row
vaSpread1.Col = 23
fecoc = ""
fecoc = Trim(vaSpread1.text)

vaSpread1.Col = 24
codsac = ""
codsac = Trim(vaSpread1.text)
vaSpread1.Col = 1

If vaSpread1.Lock = True Or fecoc <> "" Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   If vg_pais = "CO" Then
      
      RS.Open "SELECT COUNT(pro_codigo) AS nreg " & _
              "FROM b_productos WHERE pro_codigo = '" & codsac & "'", vg_db, adOpenStatic
   
   Else
      
      RS.Open "SELECT COUNT(b.fcs_codsgp) AS nreg " & _
              "FROM  b_formatocompras a, b_formatocomprassgp b " & _
              "WHERE a.foc_codsac = b.fcs_codsac " & _
              "AND   b.fcs_codsac = '" & codsac & "'", vg_db, adOpenStatic
   End If
   
   If Not RS.EOF And Not IsNull(RS!nreg) And RS!nreg > 1 Then
      
      RS.Close
      Set RS = Nothing
      StopObjeto False
      vg_nombre = "": vg_codigo = "": vg_codigo = codsac
      vg_left = fpayuda(0).Left + 1920
      Text2(0).text = "": Text2(1).text = "": Text2(2).text = ""
      
      If vg_pais = "CO" Then
         
         B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Productos SAC", "PSAC"
      
      Else
         
         B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "ProVigSac"
      
      End If
      Text2(0).text = ""
      Text2(1).text = ""
      Text2(2).text = ""
      B_TabEst.Show 1
      
      If vg_codigo = "" Then vaSpread1.ProcessTab = True: Exit Sub
      
      If vg_pais <> "CL" Or vg_FDC <> "OC" Then
         
         For i = 1 To vaSpread1.MaxRows
             
             vaSpread1.Col = 1: vaSpread1.Row = i
             If Trim(vaSpread1.text) = Trim(vg_codigo) Then Frame5.Enabled = False: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
         
         Next i
      
      End If
      sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      If vg_pais = "CO" Then
         
         RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, " & _
                 "a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                 "c.foc_codsac, c.foc_nomsac, c.foc_unisac, c.foc_faccon " & _
                 "FROM b_productos a, a_unidad b, b_formatocompras c, b_formatocomprassgp d " & _
                 "WHERE c.foc_codsac = d.fcs_codsac " & _
                 "AND   d.fcs_codsgp = a.pro_codigo " & _
                 "AND   a.pro_coduni = b.uni_codigo " & _
                 "AND   c.foc_codsac = '" & vg_codigo & "'", vg_db, adOpenStatic
      
      Else

'                 "'' AS foc_codsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon "
         
         RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo  AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon " & _
                 "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
                 "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
                 "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
                 "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
                 "FROM  b_productos a, a_unidad b " & _
                 "WHERE a.pro_coduni = b.uni_codigo " & _
                 "AND   a.pro_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
      
      End If
      Text2(0).text = ""
      Text2(1).text = ""
      Text2(2).text = ""
      
      If Not RS.EOF Then
         
         If IsNull(RS!pro_ctrsto) Then RS.Close: Set RS = Nothing: Frame5.Enabled = False: MsgBox "Producto no tiene asignado, el Movimiento...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
         If RS!pro_facing = 0 Or RS!pro_facsto = 0 Then RS.Close: Set RS = Nothing: MsgBox "Factor del producto en cero...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.acitverow: Exit Sub
         
         cPro = RS!pro_codigo
         vaSpread1.Row = Row
         vaSpread1.Col = 1
'         If vaSpread1.Lock = True Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1
         
         vaSpread1.Row = Row
         vaSpread1.SetActiveCell 1, Row
         
         '-------> desbloquear celda
         For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = False: Next i
         
         vaSpread1.Col = 1
         vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac), IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo))
         
         vaSpread1.Col = 2
         vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac), IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre))
         
         vaSpread1.Col = 3
         vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_unisac), "No existe unidad SAC", RS!foc_unisac), IIf(IsNull(RS!uni_nomcor), "No existe unidad SGP", RS!uni_nomcor))
         
         vaSpread1.Col = 12
         vaSpread1.text = RS!pro_ctacon
         
         vaSpread1.Col = 11
         vaSpread1.text = ""
         
         vaSpread1.Col = 24
         vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo), IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac))
         
         vaSpread1.Col = 25
         vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre), IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac))
         
         vaSpread1.Col = 29
         vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon), IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon))
         
         vaSpread1.Col = 24
         Text2(0).text = Trim(vaSpread1.text)
         
         vaSpread1.Col = 25
         Text2(1).text = Trim(vaSpread1.text)
         
         vaSpread1.Col = 29
         Text2(2).text = vaSpread1.text
         
         If Trim(RS!pro_ctacon) = "" Or IsNull(RS!pro_ctacon) Then RS.Close: Set RS = Nothing: MsgBox "El producto no tiene asosiada una cuenta contable...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
         
         vaSpread1.Col = 13
         vaSpread1.text = IIf(RS!pro_ctrsto = 1, "S", "N")
     
     Else
        
        '-------> bloquear celda
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        vaSpread1.text = ""
        
        vaSpread1.Col = 2
        vaSpread1.text = "": Text1(0).text = ""
        
        vaSpread1.Col = 3
        vaSpread1.text = ""
        
        vaSpread1.Col = 24
        vaSpread1.text = ""
        
        vaSpread1.Col = 25
        vaSpread1.text = ""
        
        vaSpread1.Col = 29
        vaSpread1.text = ""
        
        vaSpread1.Col = 24
        Text2(0).text = Trim(vaSpread1.text)
        
        vaSpread1.Col = 25
        Text2(1).text = Trim(vaSpread1.text)
        
        vaSpread1.Col = 29
        Text2(1).text = vaSpread1.text
        
        For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = True: Next i
     
     End If
     RS.Close
     Set RS = Nothing
     est2 = False
     Revisa cPro, vaSpread1.ActiveRow
'     vaSpread1.EditModePermanent = True
'     vaSpread1.SetActiveCell 4, vaSpread1.MaxRows
     vaSpread1.Refresh
     If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
   
   Else
      
      RS.Close: Set RS = Nothing
   
   End If

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Error_Celda

Dim RS As New ADODB.Recordset
Dim Precio As Double, Cantidad As Double, Descto As Double, Producto As String, DesctoFinal As Double, subtot As Double, totitem As Double
Dim canrec As Double, faccon As Double, canfco As Double
Dim Switch As Boolean, encuentra As Boolean
Dim codigo As String, cPro As String, Nombre As String, sql1 As String, sql2 As String, sql3 As String

Frame5.Enabled = False
vaSpread1.ProcessTab = True
With vaSpread1
    
    .Row = Row
    .Col = Col
    'If .Lock = True Then Exit Sub
    
    .Col = 1
    codigo = .Value
    
    .Col = 2
    Nombre = .Value
    
    .Col = 4
    Cantidad = Val(.Value)
    
    .Col = 5
    Precio = Val(.Value)
    
    .Col = 6
    Descto = Val(.Value)
    
    .Col = 7
    DesctoFinal = Val(.Value)
    
    .Col = 8
    totitem = Val(.Value)
    
    .Col = 9
    canrec = Val(.Value)
    
    .Col = 29
    faccon = Val(.Value)
    
    subtot = Precio * Cantidad
    .EditModePermanent = False
    
    Select Case Col
    
        Case 1
            
            .Row = Row
            .Col = 1
            
            If LimpiaDato(Trim(.text)) <> "" And .Lock = False Then
               
               sql2 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
               sql3 = IIf(vg_tipbase = "1", " AND cdate(c.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), c.foc_vigfin,101) >  '" & Date & "'")
    
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               If vg_pais = "CO" Then
                  
                  Set RS = vg_db.Execute("SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, " & _
                           "a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                           "c.foc_codsac, c.foc_nomsac, c.foc_unisac, c.foc_faccon " & _
                           "FROM b_productos a, a_unidad b, b_formatocompras c, b_formatocomprassgp d " & _
                           "WHERE c.foc_codsac = d.fcs_codsac " & _
                           "AND   d.fcs_codsgp = a.pro_codigo " & _
                           "AND   a.pro_coduni = b.uni_codigo " & _
                           "AND  (c.foc_flexec = 0 OR (c.foc_flexec = -1 " & sql3 & "))  " & _
                           "AND   c.foc_codsac = '" & LimpiaDato(Trim(.text)) & "'")
                  Text2(0).Visible = True: Text2(1).Visible = True: Text2(2).Visible = True
               
               Else
    
    '                       "'' AS foc_codsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon "
                  Set RS = vg_db.Execute("SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                           "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql2 & "))) AS foc_codsac, " & _
                           "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql2 & "))) AS foc_nomsac, " & _
                           "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql2 & "))) AS foc_unisac, " & _
                           "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql2 & "))) AS foc_faccon " & _
                           "FROM  b_productos a, a_unidad b " & _
                           "WHERE a.pro_coduni = b.uni_codigo " & _
                           "AND   a.pro_codigo = '" & LimpiaDato(Trim(.text)) & "'")
               
               End If
               
               If Not RS.EOF Then
                  
                  If IsNull(RS!pro_ctrsto) Then RS.Close: Set RS = Nothing: Frame5.Enabled = False: MsgBox "Producto no tiene asignado, el Movimiento...", vbExclamation + vbOKOnly, MsgTitulo: .SetActiveCell 1, Row: Frame5.Enabled = True: Exit Sub
                  If RS!pro_facing = 0 Or RS!pro_facsto = 0 Then RS.Close: Set RS = Nothing: MsgBox "Factor del producto en cero...", vbExclamation + vbOKOnly, MsgTitulo: .SetActiveCell 1, Row: Exit Sub
                  cPro = RS!pro_codigo
                  .Col = 1
                  
                  '-------> desbloquear celda
    '              .Col = 1: .text = RS!pro_codigo
                  .Col = 2
                  .text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_nomsac), "No existe descripción producto SAC", RS!foc_nomsac), IIf(IsNull(RS!pro_nombre), "No existe descripción producto SGP", RS!pro_nombre))
                  
                  .Col = 3
                  .text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_unisac), "No existe descripción unidad SAC", RS!foc_unisac), IIf(IsNull(RS!uni_nomcor), "No existe descripción unidad SGP", RS!uni_nomcor))
                  
                  .Col = 12
                  .text = RS!pro_ctacon
                  
                  .Col = 11
                  .text = ""
                  
                  .Col = 18
                  .text = IIf(IsNull(RS!ppd_propon), 0, RS!ppd_propon)
                  
                  For i = 4 To 10: .Col = i: .text = Format(0, fg_Pict(9, IIf(i = 4 Or i = 9, vg_DCa, 2))): Next i
                  If Trim(RS!pro_ctacon) = "" Or IsNull(RS!pro_ctacon) Then RS.Close: Set RS = Nothing: MsgBox "El producto no tiene asosiada una cuenta contable...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
                  
                  .Row = Row
                  
                  .Col = 13
                  .text = IIf(RS!pro_ctrsto = 1, "S", "N")
                  
                  .Col = 24
                  .text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo), IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac))
                  
                  .Col = 25
                  .text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre), IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac))
                  
                  .Col = 26
                  .text = IIf(IsNull(RS!pro_codref), 0, RS!pro_codref)
                  
                  .Col = 27
                  .text = IIf(IsNull(RS!pro_codrei), 0, RS!pro_codrei)
                  
                  .Col = 29
                  .text = IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon)
               
               Else
                  
                  '-------> bloquear celda
                  .Row = Row
    '              .Col = 1: .text = ""
                  .Col = 2
                  .text = ""
                  Text1(0).text = ""
                  
                  .Col = 3
                  .text = ""
                  
                  .Col = 24
                  .text = ""
                  
                  .Col = 25
                  .text = ""
                  
                  .Col = 26
                  .text = ""
                  
                  .Col = 27
                  .text = ""
                  
                  .Col = 29
                  .text = ""
               
               End If
               RS.Close
               Set RS = Nothing
               
               .Col = 24
               Text2(0).text = Trim(.text)
               
               .Col = 25
               Text2(1).text = Trim(.text)
               
               .Col = 29
               Text2(2).text = .text
               est2 = False
               Revisa cPro, Row
               .Refresh
    '           If .Enabled = True Then On Error Resume Next: .SetFocus
    
            End If
            
        Case 4
            
            If vg_pais = "CO" Then
               
               MarcaPredeterminadoFormatoCompras
               codigo = ""
               .Col = 1
               codigo = vaSpread1.Value
               'sin utilizar sql1 = IIf(vg_tipbase = "1", " AND cdate(a.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), a.foc_vigfin,101) >  '" & Date & "'")
               
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               RS.Open "SELECT top 1 a.foc_faccon FROM b_formatocompras a, b_formatocomprassgp b " & _
                       "WHERE a.foc_codsac = b.fcs_codsac " & _
                        "AND   a.foc_faccon > 0 " & _
                        "AND   a.foc_codsac = '" & codigo & "'", vg_db, adOpenStatic
               If RS.EOF Then
                  
                  RS.Close
                  Set RS = Nothing
                  '-------> bloquear celda
                  .Row = Row
                  .Col = 4
                  .text = 0
                  
                  .Col = 9
                  .text = 0
                  
                  .Col = 29
                  .text = 0
                  MsgBox "Producto no tiene asignado, factor converción...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 4, vaSpread1.ActiveRow
                  Exit Sub
                  
               Else
               
                  .Col = 29
                  .Value = RS!foc_faccon
                  faccon = RS!foc_faccon
                  
               End If
               RS.Close
               Set RS = Nothing
               
               .Col = 30
               .Value = (Cantidad * faccon) ' Mover cantidad unidad colombia
               
            End If
            
            .Col = 4
            canfac = Val(.Value)
            
            .Col = 9
            .Value = canfac
            
            If Cantidad > 0 And totitem > 0 And Precio = 0 Then
               
               .Col = 5
               .Value = Round((totitem / Cantidad), 2): Precio = Round((totitem / Cantidad), 2)
               
               .Col = 10
               .Value = Round((totitem / Cantidad), 2)
               
               .Col = 17
               .Value = Round((totitem / Cantidad), 3)
               
               subtot = Precio * Cantidad
            
            End If
            
        Case 5
        
            .Col = 5
            valfac = Val(.Value)
            
            .Col = 10
            .Value = valfac
            
            .Col = 17
            .Value = valfac
            
        Case 6
        
            .Col = 7
            .Value = subtot * (Descto / 100)
            
            .Col = 8
            .Value = Round(subtot - ((Cantidad * Precio) * (Descto / 100)), 0)
            
        Case 7
        
            .Col = 6
            .Value = (DesctoFinal * 100) / subtot
            
            .Col = 6
            Descto = Val(.Value)
            
            If Descto >= 99.99 Then
                
                Descto = 99.99
                .Col = 6
                .Value = Descto
                
                .Col = 7
                .Value = subtot * (Descto / 100)
            
            End If
            
            .Col = 8
            .Value = Round(subtot - (subtot * (Descto / 100)), 0)
            
        Case 8
        
            If Cantidad <= 0 Then Exit Sub
            .Col = 5
            .Value = Round((totitem / Cantidad), 2)
            
            .Col = 10
            .Value = Round((totitem / Cantidad), 2)
            
            .Col = 17
            .Value = Round((totitem / Cantidad), 3)
            
        Case 9
        
            If vg_pais = "CO" Then
               
               MarcaPredeterminadoFormatoCompras
               codigo = ""
               .Col = 1: codigo = vaSpread1.Value
               'sin utilizar sql1 = IIf(vg_tipbase = "1", " AND cdate(a.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), a.foc_vigfin,101) >  '" & Date & "'")
               
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               RS.Open "SELECT top 1 a.foc_faccon FROM b_formatocompras a, b_formatocomprassgp b " & _
                       "WHERE a.foc_codsac = b.fcs_codsac " & _
                        "AND   a.foc_faccon > 0 " & _
                        "AND   a.foc_codsac = '" & codigo & "'", vg_db, adOpenStatic
               If RS.EOF Then
                  
                  RS.Close
                  Set RS = Nothing
                  '-------> bloquear celda
                  .Row = Row
                  .Col = 9
                  .text = 0
                  MsgBox "Producto no tiene asignado, factor converción...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 9, vaSpread1.ActiveRow
                  Exit Sub
                  
               End If
               RS.Close
               Set RS = Nothing
               
            End If
            .Col = 9
            canrec = .Value
            If canrec > Cantidad Then MsgBox "La cantidad recibida excede de la cantidad es menor...", vbCritical, MsgTitulo
    
    End Select
    
    If Col = 4 Or Col = 5 Then
        
        .Col = 7
        .Value = subtot * (Descto / 100)
        
        .Col = 8
        .Value = Round(subtot - (subtot * (Descto / 100)), 0)
    
    End If
    SumarTotales
    
End With

est1 = False: MostrarImpuesto Row
Frame5.Enabled = True

Exit Sub
Error_Celda:
    MsgBox "Error: " & Err.Number & " " & Err.Description, vbExclamation, MsgTitulo
    
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

If ChangeMade Then
   
   Text1(0).Enabled = True

ElseIf Not ChangeMade Then
   
   Text1(0).Enabled = False

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_GotFocus()

On Error GoTo Man_Error

If vaSpread1.MaxRows < 0 Then Exit Sub
est1 = False: MostrarImpuesto vaSpread1.ActiveRow

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Or Frame6.Enabled = False Then Exit Sub
Select Case KeyCode

    Case 46 And Shift = 1
        
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        If vaSpread1.Lock = False Then Exit Sub
        If MsgBox("Elimina producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        If vaSpread1.MaxRows = 0 Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: Text1(0).text = ""
        If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
        SumarTotales

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

StopObjeto False
Dim RS As New ADODB.Recordset
Dim sql1 As String, codsgp As String, codsac As String, sql2 As String
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = NewRow
vaSpread1.Col = 1
codsgp = Trim(vaSpread1.text)

vaSpread1.Col = 24
Text2(0).text = Trim(vaSpread1.text)

vaSpread1.Col = 24
codsac = Trim(vaSpread1.text)

vaSpread1.Col = 25
Text2(1).text = Trim(vaSpread1.text)

vaSpread1.Col = 29
Text2(2).text = vaSpread1.text

If vg_pais = "CL" And vg_FDC = "OC" And Trim(codsgp) <> "" Then
   
   Image1(5).Visible = IIf(ValidarProductosSgpSac(Trim(Text2(0).text), Trim(codsgp)), True, False)

End If

If NewRow < 1 Or vaSpread1.MaxRows < 1 Or NewCol > 2 Then Exit Sub

If Row <> NewRow Then
   
   vaSpread1.Row = Row
   vaSpread1.Col = 1
   
   If Trim(vaSpread1.text) = "" Then
      
      vaSpread1.MaxRows = vaSpread1.MaxRows - 1
      vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
      Exit Sub
   
   ElseIf Col <> 1 Then
      
      vaSpread1.Col = 1: vaSpread1.Lock = True
   
   End If

End If

vaSpread1.Row = Row
vaSpread1.Col = Col
Select Case Col
    
    Case 1 And vaSpread1.Lock = False
        
        Dim cPro As String
        vaSpread1.Row = Row
        vaSpread1.Col = 1
        If vaSpread1.Lock = True Then Exit Sub
        
        If vaSpread1.text = "" Then
           
           vg_nombre = ""
           vg_codigo = ""
           
           vg_left = fpayuda(0).Left + 1920
           
           Text2(0).text = ""
           Text2(1).text = ""
           Text2(2).text = ""
           
           If vg_pais = "CO" Then
              
              B_TabEst.LlenaDatos "b_formatocompras", "foc_", "Productos SAC", "PSAC"
           
           Else
              
              B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", IIf(Trim(fg_codigocbo(Combo2, 0, 2, "")) = "CG", "ProGrl", "ProVig")
           
           End If
           
           Text2(0).text = ""
           Text2(1).text = ""
           Text2(2).text = ""
           B_TabEst.Show 1
           vaSpread1.Row = Row
           vaSpread1.Col = 1
           
           If vg_codigo = "" Then vaSpread1.Refresh: vaSpread1.ProcessTab = True: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
           
           codigo = vg_codigo
           vaSpread1.text = vg_codigo
        
        End If
        
        codigo = vaSpread1.text
        
        If vg_pais <> "CL" Or vg_FDC <> "OC" Then
           
           For i = 1 To vaSpread1.MaxRows
               
               vaSpread1.Col = 1: vaSpread1.Row = i
               If Trim(vaSpread1.text) = Trim(codigo) And Row <> i And Trim(vaSpread1.text) <> "" Then Frame5.Enabled = False: MsgBox "El producto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
           
           Next i
        
        End If
        sql1 = IIf(vg_tipbase = "1", " AND CDATE(x.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), x.foc_vigfin,101) >  '" & Date & "'")
        sql2 = IIf(vg_tipbase = "1", " AND cdate(c.foc_vigfin) >  cdate('" & Date & "') ", " AND convert(varchar(10), c.foc_vigfin,101) >  '" & Date & "'")
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        If vg_pais = "CO" Then
           
           RS.Open "SELECT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, " & _
                   "a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                   "c.foc_codsac, c.foc_nomsac, c.foc_unisac, c.foc_faccon " & _
                   "FROM b_productos a, a_unidad b, b_formatocompras c, b_formatocomprassgp d " & _
                   "WHERE c.foc_codsac = d.fcs_codsac " & _
                   "AND   d.fcs_codsgp = a.pro_codigo " & _
                   "AND   a.pro_coduni = b.uni_codigo " & _
                   "AND  (c.foc_flexec = 0 OR (c.foc_flexec = -1 " & sql2 & ")) " & _
                   "AND   c.foc_codsac = '" & codigo & "'", vg_db, adOpenStatic
           
           Text2(0).text = ""
           Text2(1).text = ""
           Text2(0).Visible = True
           Text2(1).Visible = True
           Text2(2).Visible = True
        
        Else
    '               "'' AS foc_codsac, '' AS foc_nomsac, '' AS foc_unisac, 1 AS foc_faccon "
           RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, a.pro_ctacon, a.pro_ctrsto, b.uni_nomcor, a.pro_facing, a.pro_facsto, a.pro_codref, a.pro_codrei, (SELECT DISTINCT ppd_propon FROM b_productospmpdia WHERE ppd_codpro = a.pro_codigo AND ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ") AS ppd_propon, " & _
                   "(SELECT TOP 1 x.foc_codsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_codsac, " & _
                   "(SELECT TOP 1 x.foc_nomsac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_nomsac, " & _
                   "(SELECT TOP 1 x.foc_unisac FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_unisac, " & _
                   "(SELECT TOP 1 x.foc_faccon FROM b_formatocompras x, b_formatocomprassgp z WHERE x.foc_codsac = z.fcs_codsac AND z.fcs_codsgp = a.pro_codigo AND (z.fcs_sgppre = 1 OR z.fcs_sgppre = 0) AND (x.foc_flexec  = 0 OR (x.foc_flexec = -1  " & sql1 & "))) AS foc_faccon " & _
                   "FROM  b_productos a, a_unidad b, a_tiposervicio d, b_clientes e " & _
                   "WHERE (d.tis_codigo = e.cli_codtis OR a.pro_maepro < 1) AND e.cli_codigo = '" & MuestraCasino(1) & "' AND (d.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_coduni = b.uni_codigo " & _
                   "AND   a.pro_codigo = '" & codigo & "' " & _
                   "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven = 0)", vg_db, adOpenStatic
        End If
        
        Text2(0).text = ""
        Text2(1).text = ""
        Text2(2).text = ""
        
        If Not RS.EOF Then
           
           If IsNull(RS!pro_ctrsto) Then RS.Close: Set RS = Nothing: Frame5.Enabled = False: MsgBox "Producto no tiene asignado, el Movimiento...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Frame5.Enabled = True: Exit Sub
           If RS!pro_facing = 0 Or RS!pro_facsto = 0 Then RS.Close: Set RS = Nothing: MsgBox "Factor del producto en cero...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
           
           vaSpread1.Row = Row
           cPro = RS!pro_codigo
           vaSpread1.SetActiveCell 1, vaSpread1.Row
           
           For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = False: Next i
           
           vaSpread1.Row = Row
           vaSpread1.Col = 1
           vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac), IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo))
           
           vaSpread1.Col = 2
           vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac), IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre))
           
           vaSpread1.Col = 3
           vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!foc_unisac), "No existe Unidad SAC", RS!foc_unisac), IIf(IsNull(RS!uni_nomcor), "No existe Unidad SGP", RS!uni_nomcor))
           
           vaSpread1.Col = 12
           vaSpread1.text = RS!pro_ctacon
           
           vaSpread1.Col = 11
           vaSpread1.text = ""
           
           vaSpread1.Col = 18
           vaSpread1.text = IIf(IsNull(RS!ppd_propon), 0, RS!ppd_propon)
           
           vaSpread1.Col = 24
           vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_codigo), "", RS!pro_codigo), IIf(IsNull(RS!foc_codsac), "", RS!foc_codsac))
           
           vaSpread1.Col = 25
           vaSpread1.text = IIf(vg_pais = "CO", IIf(IsNull(RS!pro_nombre), "No existe descripción SGP", RS!pro_nombre), IIf(IsNull(RS!foc_nomsac), "No existe descripción SAC", RS!foc_nomsac))
           
           vaSpread1.Col = 29
           vaSpread1.text = IIf(IsNull(RS!foc_faccon), 0, RS!foc_faccon)
           
           vaSpread1.Col = 24
           Text2(0).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 25
           Text2(1).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 29
           Text2(2).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 26
           vaSpread1.text = IIf(IsNull(RS!pro_codref), 0, RS!pro_codref)
           
           vaSpread1.Col = 27
           vaSpread1.text = IIf(IsNull(RS!pro_codrei), 0, RS!pro_codrei)
           
           vaSpread1.Col = 4
           If Trim(vaSpread1.text) = "" Then For i = 4 To 10: vaSpread1.Col = i: vaSpread1.text = Format(0, fg_Pict(9, IIf(i = 4 Or i = 9, vg_DCa, 2))): Next i
           If Trim(RS!pro_ctacon) = "" Or IsNull(RS!pro_ctacon) Then RS.Close: Set RS = Nothing: MsgBox "El producto no tiene asosiada una cuenta contable...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow: Exit Sub
           
           vaSpread1.Col = 13
           vaSpread1.text = IIf(RS!pro_ctrsto = 1, "S", "N")
           
           i = 4
           If Row <> NewRow Then vaSpread1.Col = 1: vaSpread1.Lock = True
        
        Else
           
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.Col = 1
           vaSpread1.text = ""
           
           vaSpread1.Col = 2
           vaSpread1.text = ""
           Text1(0).text = ""
           
           vaSpread1.Col = 3
           vaSpread1.text = ""
           
           vaSpread1.Col = 24
           vaSpread1.text = ""
           
           vaSpread1.Col = 25
           vaSpread1.text = ""
           
           vaSpread1.Col = 26
           vaSpread1.text = ""
           
           vaSpread1.Col = 27
           vaSpread1.text = ""
           
           vaSpread1.Col = 29
           vaSpread1.text = 0
           
           vaSpread1.Col = 24
           Text2(0).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 25
           Text2(1).text = Trim(vaSpread1.text)
           
           vaSpread1.Col = 29
           Text2(2).text = Trim(vaSpread1.text)
           For i = 4 To 10: vaSpread1.Col = i: vaSpread1.Lock = True: Next i
           vaSpread1.SetActiveCell 1, vaSpread1.ActiveRow
           i = 1
           MsgBox "producto no existe o bien descontinuado ...", vbExclamation + vbOKOnly, MsgTitulo
        
        End If
        RS.Close
        Set RS = Nothing
        est2 = False
        Revisa cPro, vaSpread1.ActiveRow
        If i = 4 Then vaSpread1.EditModePermanent = True
        vaSpread1.SetActiveCell i, vaSpread1.ActiveRow 'vaSpread1.Row
        vaSpread1.Refresh
        If vaSpread1.Enabled = True Then On Error Resume Next: vaSpread1.SetFocus
        
End Select
vaSpread1.ProcessTab = True
est1 = False
MostrarImpuesto NewRow

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_TextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)
'MultiLine = 2
End Sub

Sub MostrarImpuesto(Fila As Long)

On Error GoTo Man_Error

Dim StrImp As String, StrImpb As String, v_rut As String
Dim codimp As Long, PctI As Double, CosI As Long
Dim CosTot As Double, valimp As Double
Dim regimp As String, autoret As String, cuohor As String
Dim codmun As Long

vaSpread2.Visible = False
For i = 1 To vaSpread2.MaxRows
    
    vaSpread2.Row = i
    vaSpread2.RowHidden = True

Next i

CosTot = 0
valimp = 0
StrImp = ""
StrImpb = ""
v_rut = fg_DespintaRut(fpText(0).text)

vaSpread1.Row = Fila
vaSpread1.Col = IIf(Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) <> "CG", 2, 11) '20080417: Text1(0).text = vaSpread1.text
vaSpread1.Col = 8
CosTot = Val(vaSpread1.Value)

vaSpread1.Col = 2
Text1(0).text = vaSpread1.text

vaSpread1.Col = 14
StrImp = Trim(vaSpread1.text)

If Len(StrImp) <> 0 Then
   
   Do While InStr(StrImp, ";") <> 0
      
      valimp = 0
      StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
      StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
      codimp = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
      PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
      CosI = Val(Mid(StrImpb, 1))
      
      If vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(codimp), SearchFlagsNone) <> -1 Then
         
         vaSpread2.Row = vaSpread2.SearchCol(1, 0, vaSpread2.MaxRows, CStr(codimp), SearchFlagsNone)
         
         If PctI > 0 Then
            
            est1 = True: est2 = True
            vaSpread2.Col = 2
            vaSpread2.text = "1"
            
            vaSpread2.Col = 4
            vaSpread2.text = Format(PctI, fg_Pict(3, 2)) & " %"
            
            vaSpread2.Col = 7
            vaSpread2.text = PctI
            
            est1 = False
            est2 = False
         
         Else
            
            est1 = True
            est2 = True
            
            vaSpread2.Col = 2
            vaSpread2.text = "0"
            
            vaSpread2.Col = 4
            vaSpread2.text = "0 %"
            vaSpread2.Col = 7
            vaSpread2.text = 0
            
            vaSpread2.Col = 4
            vaSpread2.text = "0 %"
            
            vaSpread2.Col = 7
            vaSpread2.text = 0
            
            est1 = False
            est2 = False
         
         End If
         
         vaSpread2.Col = 7
         valimp = IIf(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "BH" And codimp <> 1, Val(vaSpread2.Value), (Val(vaSpread2.Value) / 100))
         
         vaSpread2.Col = 5
         If valimp > 0 Then
            
            vaSpread2.text = IIf(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, "")) = "BH" And codimp <> 1, Format(Round(((CosTot / ((100 - valimp) / 100))) / valimp), fg_Pict(6, 0)), Format(valimp * CosTot, fg_Pict(6, 0)))
         
         Else
            
            vaSpread2.text = 0
         
         End If
         vaSpread2.SetActiveCell 1, 1
         vaSpread2.RowHidden = False
      
      End If
   
   Loop

End If

vaSpread2.Visible = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

On Error GoTo Man_Error

est2 = True
If vaSpread2.MaxRows < 1 Or est1 Then Exit Sub

Dim StrImp As String, StrImpb As String
Dim codimp As Long, PctI As Double, CosI As Long

vaSpread2.Row = Row
vaSpread2.Col = 1

If vaSpread2.Value = 1 Or vaSpread2.Value = 11 Or vaSpread2.Value = GetParametro("parretfue") Or vaSpread2.Value = GetParametro("parretica") Then vaSpread2.Col = 2: vaSpread2.text = "1": Exit Sub
vaSpread2.Col = 2

If ButtonDown = 0 Then
   
   vaSpread2.Col = 4
   est2 = True
   vaSpread2.text = "0 %"
   
   vaSpread2.Col = 5
   est2 = True
   vaSpread2.text = Format(0, fg_Pict(6, 0))
   
   vaSpread2.Col = 7
   est2 = True
   vaSpread2.text = 0
   
   vaSpread1.Row = vaSpread1.ActiveRow
   
   vaSpread1.Col = 14
   StrImp = Trim(vaSpread1.text)
   vaSpread1.text = ""
   
   If Len(StrImp) <> 0 Then
      
      vaSpread2.Col = 1
      
      Do While InStr(StrImp, ";") <> 0
         
         StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
         StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
         codimp = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
         
         If codimp = vaSpread2.text Then
            
            PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
            PctI = 0
         
         Else
            
            PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
         
         End If
         
         CosI = Val(Mid(StrImpb, 1))
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(codimp) & "&"
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(PctI)) & "&"
         
         vaSpread1.Col = 14
         vaSpread1.text = vaSpread1.text & Trim(Str(CosI)) & ";"
      
      Loop
      
      SumarTotales
   
   End If

ElseIf ButtonDown = 1 Then
   
   vaSpread2.Row = Row
   vaSpread2.Col = 1
   vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 14
   
   For i = 1 To UBound(Impuestos)
       
       If Impuestos(i, 1) = vaSpread2.text Then
          
          vaSpread2.Col = 4
          est2 = True
          vaSpread2.text = Impuestos(i, 3) & " %"
          
          vaSpread2.Col = 5
          est2 = True
          vaSpread2.text = Format(0, fg_Pict(6, 0))
          
          vaSpread2.Col = 7
          est2 = True
          vaSpread2.text = Impuestos(i, 3)
          
          StrImp = Trim(vaSpread1.text): vaSpread1.text = ""
          
          If Len(StrImp) <> 0 Then
             
             Do While InStr(StrImp, ";") <> 0
                
                StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
                StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
                codimp = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                vaSpread2.Col = 1
                
                If codimp = vaSpread2.text Then
                   
                   PctI = Impuestos(i, 3)
                   CosI = Impuestos(i, 5)
                
                Else
                   
                   PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1)): StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
                   CosI = Val(Mid(StrImpb, 1))
                
                End If
                
                vaSpread1.Col = 14
                vaSpread1.text = vaSpread1.text & Trim(codimp) & "&"
                
                vaSpread1.Col = 14
                vaSpread1.text = vaSpread1.text & Trim(Str(PctI)) & "&"
                
                vaSpread1.Col = 14
                vaSpread1.text = vaSpread1.text & Trim(Str(CosI)) & ";"
             
             Loop
             
             SumarTotales
          
          End If
          Exit For
       
       End If
   
   Next i
   
   est1 = True: MostrarImpuesto vaSpread1.ActiveRow

End If
est2 = False

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

If ChangeMade = False Or vaSpread2.MaxRows < 1 Or est2 Then Exit Sub

Dim codimp As Long
Dim valimp As Double
Dim totpro As Double
Dim porimp As Double

codimp = 0
valimp = 0
totpro = 0
porimp = 0

'-------> Devolver impuesto si es modificable
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 8
totpro = vaSpread1.text

vaSpread1.Col = 14
StrImp = Trim(vaSpread1.text)
vaSpread1.text = ""

'-------> Calcular & impuesto
vaSpread2.Row = Row
vaSpread2.Col = 5
valimp = vaSpread2.text

If totpro > 0 Then vaSpread2.Col = 4: vaSpread2.text = Format(Round((valimp / totpro) * 100, 2), fg_Pict(3, 2)) & " %": porimp = ((valimp / totpro) * 100)

If Len(StrImp) <> 0 Then
   
   Do While InStr(StrImp, ";") <> 0
      
      StrImpb = Mid(StrImp, 1, InStr(StrImp, ";") - 1)
      StrImp = IIf(Len(StrImp) > InStr(StrImp, ";"), Mid(StrImp, InStr(StrImp, ";") + 1), "")
      codimp = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1))
      StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
      vaSpread2.Col = 1
      
      If codimp = vaSpread2.text Then
         
         PctI = porimp
         vaSpread2.Col = 6
         CosI = vaSpread2.text
      
      Else
         
         PctI = Val(Mid(StrImpb, 1, InStr(StrImpb, "&") - 1))
         StrImpb = Mid(StrImpb, InStr(StrImpb, "&") + 1)
         CosI = Val(Mid(StrImpb, 1))
      
      End If
      
      vaSpread1.Col = 14
      vaSpread1.text = vaSpread1.text & Trim(codimp) & "&"
      
      vaSpread1.Col = 14
      vaSpread1.text = vaSpread1.text & Trim(Str(PctI)) & "&"
      
      vaSpread1.Col = 14
      vaSpread1.text = vaSpread1.text & Trim(Str(CosI)) & ";"
   
   Loop

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Sub StopObjeto(est As Boolean)

On Error GoTo Man_Error

fpText(0).TabStop = est
fpText(1).TabStop = est
Combo2(0).TabStop = est
Combo2(1).TabStop = est
Double1(5).TabStop = est
Double1(6).TabStop = est
Double1(12).TabStop = est
Double1(13).TabStop = est
Double1(14).TabStop = est
Double1(15).TabStop = est
Double1(16).TabStop = est
Double1(17).TabStop = est
Date1(0).TabStop = est
Date1(2).TabStop = est
vaSpread2.TabStop = est
'Option1(0).TabStop = est: Option1(1).TabStop = est

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Function ValidarRetProveedor(rut As String) As Boolean

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
ValidarRetProveedor = False
'-------> Validar si proveedor posee impuestos adicionales
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT prv_regimp, prv_autret, prv_cuohor, prv_codmun FROM b_proveedor WHERE prv_codigo = '" & rut & "' AND prv_regimp IS NOT NULL", vg_db, adOpenStatic
If Not RS.EOF Then
   
   regimp = RS!prv_regimp
   autret = RS!prv_autret
   cuohor = RS!prv_cuohor
   codmun = RS!prv_codmun
   ValidarRetProveedor = True

End If
RS.Close
Set RS = Nothing

Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Function

Sub ValidarCfc()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim sql1 As String, sql2 As String
Dim numcfc As Long
If Combo2(0).ListIndex = -1 Then Exit Sub

'-------> Rutina separar Folio
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
RS1.Open "SELECT MAX(inf_numero) AS Mayor FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(fpText(2).text) & "' AND inf_tipo = '" & IIf(Option1(0).Value = True, "F", "C") & "'", vg_db, adOpenStatic
numcfc = TipoDato(RS1!mayor, 0)
RS1.Close
Set RS1 = Nothing

SepararFolioDocumento Trim(fpText(2).text), vg_codbod, numcfc
numcfc = 0

If Encontrado = False Then

    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    RS.Open "SELECT * FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(fpText(2).text) & "' AND inf_tipo = '" & IIf(Option1(0).Value = True, "F", "C") & "' AND inf_feccie = 0", vg_db, adOpenStatic
    
    If Not RS.EOF Then
       
       '-------> Validar cantidad documento en un folio
       '-------> 1º Validar si existen datos compras con numero folio
'       sql1 = IIf(vg_tipbase = "1", " IIF(toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE','FE', 'FA') AS toc_tipdoc ", " (CASE WHEN toc_tipdoc = 'FE' OR toc_tipdoc = 'DE' OR toc_tipdoc = 'CE' THEN 'FE' ELSE 'FA' END) AS toc_tipdoc ")
'       sql2 = IIf(Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "CE" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "DE", " IN ('FE', 'CE', 'DE')", " NOT IN ('FE', 'CE', 'DE')")
    
       If RS1.State = 1 Then RS1.Close
       RS1.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient

'       RS1.Open "SELECT DISTINCT " & sql1 & ", toc_numinf " & _
'                "FROM b_totcompras " & _
'                "WHERE  toc_codbod = " & vg_codbod & " AND toc_tipdoc NOT IN ('SN') AND toc_numinf IN (SELECT DISTINCT inf_numero FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(fpText(2).text) & "' AND inf_tipo = '" & IIf(Option1(0).Value = True, "F", "C") & "' AND (inf_feccie = 0 OR (inf_feccie) IS NULL)) AND toc_tipdoc " & sql2 & "", vg_db, adOpenStatic
       
       If Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "FE" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "CE" Or Trim(fg_TraerRelacionTipoDocumento(fg_codigocbo(Combo2, 0, 2, ""))) = "DE" Then
          
          Set RS1 = vg_db.Execute("sgp_Sel_ValidaCfCElectronico '" & Trim(fpText(2).text) & "', " & vg_codbod & ", '" & IIf(Option1(0).Value = True, "F", "C") & "'")
       
       Else
       
          Set RS1 = vg_db.Execute("sgp_Sel_ValidaCfCNormal '" & Trim(fpText(2).text) & "', " & vg_codbod & ", '" & IIf(Option1(0).Value = True, "F", "C") & "'")
       
       End If
       
       If Not RS1.EOF Then
          
          Do While Not RS1.EOF
             
             Double1(6).Value = IIf(IsNull(RS1!toc_numinf), 0, RS1!toc_numinf)
             RS1.MoveNext
          
          Loop
          RS1.Close
          Set RS1 = Nothing
       
       Else
          
          numcfc = 0
          RS1.Close
          Set RS1 = Nothing
          
          If RS1.State = 1 Then RS1.Close
          RS1.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          RS1.Open "SELECT MAX(inf_numero) AS Mayor FROM a_infcfcfofi WHERE inf_cencos = '" & Trim(fpText(2).text) & "' AND inf_tipo = '" & IIf(Option1(0).Value = True, "F", "C") & "'", vg_db, adOpenStatic
          numcfc = TipoDato(RS1!mayor, 0)
          RS1.Close
          Set RS1 = Nothing
          
          If RS1.State = 1 Then RS1.Close
          RS1.CursorLocation = adUseClient
          vg_db.CursorLocation = adUseClient
          RS1.Open "SELECT TOP 1 toc_tipdoc FROM b_totcompras WHERE toc_codbod = " & vg_codbod & " AND toc_numinf = " & numcfc & "", vg_db, adOpenStatic
          If Not RS1.EOF Then
             
             Double1(6).Value = GenerarFolioCFC(Trim(fpText(2).text), IIf(Option1(0).Value = True, "F", "C"))
          
          Else
             
             Double1(6).Value = IIf(IsNull(numcfc), 0, numcfc)
          
          End If
          RS1.Close
          Set RS1 = Nothing
       
       End If
    
    Else
        
        Double1(6).Value = GenerarFolioCFC(Trim(fpText(2).text), IIf(Option1(0).Value = True, "F", "C"))
    
    End If
    RS.Close
    Set RS = Nothing

End If

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub
