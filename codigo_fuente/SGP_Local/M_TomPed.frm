VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form M_TomPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Toma Pedido Paciente"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Preview Pedido"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   12300
      Begin MSComctlLib.Toolbar tlbPreview 
         Height          =   390
         Left            =   7080
         TabIndex        =   29
         Top             =   150
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   688
         ButtonWidth     =   688
         ButtonHeight    =   688
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageListPreviewPedido"
         DisabledImageList=   "ImageListPreviewPedido"
         HotImageList    =   "ImageListPreviewPedido"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Agregar Cantidad"
               Object.Tag             =   "Add"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Eliminar Cantidad"
               Object.Tag             =   "Delete"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageListPreviewPedido 
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   19
         ImageHeight     =   20
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomPed.frx":0000
               Key             =   "Add"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomPed.frx":05E2
               Key             =   "Del"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomPed.frx":0C44
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_TomPed.frx":0FDE
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VSFlex8LCtl.VSFlexGrid grC 
         Height          =   4575
         Left            =   120
         TabIndex        =   30
         Top             =   600
         Width           =   12015
         _cx             =   21193
         _cy             =   8070
         Appearance      =   0
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   16757683
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"M_TomPed.frx":18B8
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Pedido"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2040
      Width           =   12300
      Begin VB.CheckBox Check1 
         Caption         =   "Con filtro"
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
         Left            =   8880
         TabIndex        =   8
         Top             =   480
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ControlType     =   3
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
      Begin EditLib.fpDateTime Date1 
         Height          =   345
         Index           =   0
         Left            =   1560
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   1710
         _Version        =   196608
         _ExtentX        =   3016
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   3480
         TabIndex        =   7
         Top             =   495
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
         Caption         =   "Régimen :"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   240
         Width           =   870
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   4320
         Picture         =   "M_TomPed.frx":19DB
         Top             =   420
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   4740
         TabIndex        =   26
         Top             =   495
         Width           =   3855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Ingreso :"
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
         Index           =   9
         Left            =   1560
         TabIndex        =   18
         Top             =   240
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Pedido:"
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
         Top             =   240
         Width           =   930
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   4785
         TabIndex        =   28
         Top             =   540
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Paciente [Activo]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   360
      Width           =   12300
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   3
         Left            =   9960
         TabIndex        =   4
         Top             =   1080
         Width           =   1380
         _Version        =   196608
         _ExtentX        =   2434
         _ExtentY        =   556
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
         ThreeDInsideHighlightColor=   -2147483633
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
         ThreeDTextHighlightColor=   -2147483633
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
         ControlType     =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   6120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   480
         Width           =   1740
         _Version        =   196608
         _ExtentX        =   3069
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   6120
         TabIndex        =   3
         Top             =   1080
         Width           =   3540
         _Version        =   196608
         _ExtentX        =   6244
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
         ThreeDTextHighlightColor=   -2147483633
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
         MaxLength       =   20
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   0
         Top             =   480
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
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   1080
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
         Caption         =   "Régimen :"
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
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Width           =   870
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   960
         Picture         =   "M_TomPed.frx":1CE5
         Top             =   975
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   1380
         TabIndex        =   23
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1380
         TabIndex        =   21
         Top             =   480
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   960
         Picture         =   "M_TomPed.frx":1FEF
         Top             =   390
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
         Index           =   10
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Grupo Paciente :"
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
         Left            =   6120
         TabIndex        =   15
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   8220
         TabIndex        =   13
         Top             =   480
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   7800
         Picture         =   "M_TomPed.frx":22F9
         Top             =   360
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Rut :"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cama :"
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
         Index           =   15
         Left            =   9960
         TabIndex        =   11
         Top             =   840
         Width           =   870
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   8250
         TabIndex        =   14
         Top             =   525
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   1425
         TabIndex        =   22
         Top             =   525
         Width           =   3855
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   1425
         TabIndex        =   25
         Top             =   1125
         Width           =   3855
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Menu MenuDetalle 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu OpGrilla 
         Caption         =   "Ingresa Receta"
         Index           =   10
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Ingresa Producto"
         Index           =   20
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Borra Línea"
         Index           =   30
      End
   End
End
Attribute VB_Name = "M_TomPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim modo As String, codreg As Long, indcol As Long, indfil As Long
Dim est As Boolean, estval As Boolean, estedi As Boolean
Const DRAGTOL = 100         ' mouse movement before dragging starts
Private Type DRAGINFO
    bDragging As Boolean    ' currently dragging
    bCheckDrag As Boolean   ' currently checking mouse to start dragging
    lSrc As Long            ' row being dragged
    xDown As Long           ' mouse down position
    yDown As Long           ' mouse down position
End Type
Dim g_DragInfo As DRAGINFO

Private Sub Check1_Click(Index As Integer)
Select Case Index
Case 0
    fpLongInteger1(1).Enabled = IIf(Check1(Index).Value = 1, False, True)
    Image1(1).Enabled = IIf(Check1(Index).Value = 1, False, True)
End Select
End Sub

Private Sub Date1_Change(Index As Integer)
If est Or Trim(Date1(0).text) = "" Then Exit Sub
MostrarTomaPedido Date1(0).text
End Sub

Private Sub Date1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9150
Me.Width = 12615
Msgtitulo = "Toma Pedido Paciente"
fg_centra Me
modo = "": est = True: estval = True: estedi = False
codreg = 0
Gl_Mo_Botones Me, 13
Gl_Ac_Botones Me, 13, 2, modo
est = True: OpGr = False
fpText(4).Enabled = ModCasino
Image1(4).Enabled = ModCasino
fpText(4).text = MuestraCasino(1)
fpayuda(4).Caption = MuestraCasino(2)
'------- Mover concepto aporte nutricionales grilla receta
Dim indnut As Long
indnut = 0: indcol = 9
RS.Open "SELECT COUNT(*) as nreg FROM a_nutriente", vg_db, adOpenStatic
If RS.EOF Or RS!nreg < 1 Or IsNull(RS!nreg) Then RS.Close: Set RS = Nothing: MsgBox "No existen nutrientes, proceso cancelado...", vbCritical, Msgtitulo
indcol = indcol + RS!nreg
grC.Cols = 10 + RS!nreg: indnut = 10
grC.Rows = 1
grC.ColWidth(0) = 1000
grC.ColFormat(3) = fg_Pict(6, 2): grC.ColDataType(3) = flexDTDouble
grC.MergeCells = flexMergeOutline
grC.FixedCols = 0
grC.ExtendLastCol = True
RS.Close: Set RS = Nothing
RS.Open "SELECT * FROM a_nutriente ORDER BY nut_secnro", vg_db, adOpenStatic
Do While Not RS.EOF
   'Formatear grilla aportes
   grC.TextMatrix(0, indnut) = RS!nut_codigo & "-" & Trim(RS!nut_nombre)
   grC.ColWidth(indnut) = 900: grC.ColAlignment(indnut) = flexAlignRightCenter ': grC.ColFormat(indnut) = fg_Pict(4, 2)
   RS.MoveNext: indnut = indnut + 1
Loop
RS.Close: Set RS = Nothing

grC.ExtendLastCol = False
'grC.Editable = flexEDKbdMouse
grC.OutlineCol = 0
grC.OutlineBar = flexOutlineBarSimpleLeaf
grC.MergeCells = flexMergeNever
grC.AllowUserResizing = flexResizeColumns '= flexResizeNone
grC.AllowSelection = True
grC.Gridlines = flexGridFlatVert
grC.FixedCols = 1 'Fijar columna
est = False
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 1
    vg_codregimen = 0
    vg_Aux = fg_DespintaRut(fpText(0).text)
    RS1.Open "SELECT DISTINCT pac_codreg FROM b_pacientes WHERE pac_codigo = '" & vg_Aux & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then vg_codregimen = RS1!pac_codreg
    RS1.Close: Set RS1 = Nothing
    If Check1(0).Value = 1 And vg_codregimen > 0 Then
       RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_codigo = " & vg_codregimen & "", vg_db, adOpenStatic
    Else
       RS1.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    End If
    vg_codregimen = 0
    vg_Aux = ""
    fpayuda(1).Caption = ""
    fpayuda(1).Tag = ""
    If Not RS.EOF Then fpayuda(1).Caption = Trim(RS1!reg_nombre): fpayuda(1).Tag = Trim(RS1!reg_nombre)
    RS1.Close: Set RS1 = Nothing
    If Trim(Date1(0).text) = "" Then Exit Sub
    MostrarTomaPedido Date1(0).text
Case 2
    If (Date1(0).text > Date) Then Exit Sub
    If fpLongInteger1(1).Value = fpLongInteger1(2).Value Then fpayuda(3).Caption = "": MsgBox "Régimen deben ser distinto", vbCritical, Msgtitulo: Exit Sub
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(2).Value) & "", vg_db, adOpenStatic
    fpayuda(3).Caption = ""
    If Not RS.EOF Then fpayuda(3).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    If Trim(fpayuda(1).Caption) <> "" Then
        If (grC.Rows < 1) Then
            MsgBox "Debe seleccionar un Servicio previamente.", vbCritical, Msgtitulo
            Exit Sub
        End If
        Dim lngServicioSelected As Long
        If grC.Rows <= 1 Then Exit Sub
'        lngServicioSelected = GetItem(TvwDir(0).SelectedItem.Key, 2)
        MostrarTomaPedido Date1(0).text, Val(fpLongInteger1(2).Value)
    End If
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    fpayuda(0).Caption = ""
Case 4
    RS.Open "SELECT * FROM b_clientes WHERE cli_codigo = '" & fpText(4).text & "' AND cli_tipo = 0", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!cli_nombre)
    RS.Close: Set RS = Nothing
End Select
End Sub

Private Sub fpText_GotFocus(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    If Trim(fpText(0).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText(0).text = fg_DespintaRut(fpText(0).text)
    fpText(0).text = Mid(fpText(0).text, 1, Len(Trim(fpText(0).text)) - 1)
End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_LostFocus(Index As Integer)
If est Then Exit Sub
Select Case Index
Case 0
    If fpText(Index).text = "" Or est Then Exit Sub
    fpText(Index).text = fg_RutDig(Trim(fpText(0).text))
    RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
            "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo = b_pacientes.pac_codgrp) " & _
            "LEFT JOIN a_regimen ON b_pacientes.pac_codreg = a_regimen.reg_codigo WHERE b_pacientes.pac_codigo =  '" & Trim(fpText(Index).text) & "'", vg_db, adOpenStatic
    codreg = 0
    If Not RS.EOF Then
       fpText(0).text = fg_PintaRut(fpText(0).text)
       fpayuda(0).Caption = Trim(RS!pac_nombre) & " " & Trim(RS!pac_appaterno) & " " & Trim(RS!pac_apmaterno)
       Limpiar 1
       fpText(1).text = Trim(RS!grp_nombre)
       Image1(1).Enabled = IIf(IsNull(RS!pac_codreg), False, True)
       fpLongInteger1(1).Value = IIf(IsNull(RS!pac_codreg), "", RS!pac_codreg)
       codreg = IIf(IsNull(RS!pac_codreg), 0, RS!pac_codreg)
       fpayuda(1).Caption = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
       fpText(3).text = IIf(IsNull(RS!pac_nrocam), "", RS!pac_nrocam)
       If modo = "A" And Not IsNull(RS!reg_nombre) Then
           For i = 1 To 2
               Frame1(i).Enabled = True
           Next i
           fpLongInteger1(0).Enabled = False: fpLongInteger1(1).Enabled = False
        Else
           Check1(0).Value = 0
        End If
        Toolbar1.Buttons(15).Enabled = True: Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
    Else
        Limpiar 1
        For i = 1 To 2
            Frame1(i).Enabled = False
        Next i
        RS.Close: Set RS = Nothing: MsgBox "Pacientes no existe...", vbCritical, Msgtitulo
        fpText(0).text = "": fpayuda(0).Caption = ""
        Toolbar1.Buttons(15).Enabled = False: Toolbar1.Buttons(15).ToolTipText = ""
        Exit Sub
    End If
    fpLongInteger1(0).Value = ""
    RS.Close: Set RS = Nothing
End Select
End Sub

Sub Limpiar(indtop As Long)
est = True
grC.Rows = 1
For i = indtop To 3
    If Trim(fpayuda(0).Caption) = "" And i <> 2 Then fpText(i).text = ""
    If i <= 3 And i <> 2 And i <> 1 And i <> 0 Then fpayuda(i).Caption = ""
    If i < 3 And i <> 1 Then fpLongInteger1(i).Value = ""
    If i < 1 Then Check1(i).Value = 1
Next i
Date1(0).text = ""
est = False
End Sub

Sub MostrarTomaPedido(Fecha As Date, Optional codreg)
Dim auxser As Long, auxess As Long, auxreg As Long, auxrec As Long, auxing As String, indadi As Long, auxhor As String, nomser As String
Dim lngIndex As Long, sql1 As String
Dim lngIndexDeleteItem As Long
If IsNull(Fecha) Then Exit Sub
'------- Validar si existe una toma pedido de un cliente
vg_Aux = fg_DespintaRut(fpText(0).text)
sql1 = IIf(vg_tipbase = "1", " cdate('" & Fecha & "') ", " '" & Format(Fecha, "yyyymmdd") & "' ")
RS.Open "SELECT DISTINCT top_codigo " & _
        "FROM b_tomapedido " & _
        "WHERE top_codpac = '" & vg_Aux & "' " & _
        "AND   top_fecped = " & sql1 & "", vg_db, adOpenStatic
If Not RS.EOF Then
   est = True
   grC.Rows = 1
   tlbPreview.Buttons(1).Enabled = False
   tlbPreview.Buttons(2).Enabled = False
   Frame1(2).Enabled = False
   RS.Close: Set RS = Nothing
   MsgBox "No puede crear más de pedido normal, para la misma fecha", vbCritical, Msgtitulo
   Date1(0).text = ""
   est = False
   Exit Sub
Else
   tlbPreview.Buttons(1).Enabled = True
   tlbPreview.Buttons(2).Enabled = True
End If
RS.Close: Set RS = Nothing
vg_Aux = ""
est = True
RS.Open "SELECT DISTINCT a.min_codigo, b.mid_tiprec, b.mid_numlin, i.reg_codigo, i.reg_nombre, c.ser_codigo, c.ser_nombre, c.ser_orden, c.ser_horent, " & _
        "d.ess_codigo, d.ess_nombre, d.ess_orden, e.rec_codigo, e.rec_nombre, g.ing_codigo, g.ing_nombre, h.unm_nomcor, f.red_nroite, f.red_canpro, f.red_cospro, " & _
        "f.red_pctapr, f.red_pctcoc, f.red_pctnut, SUM(((f.red_pctapr/100)*f.red_canpro)*(f.red_pctcoc/100)) AS canser, ((f.red_pctnut/100)*(f.red_canpro)) AS cannet " & _
        "FROM b_minuta a, b_minutadet b, a_servicio c, a_estservicio d, b_receta e, b_recetadet f, b_ingrediente g, a_unidadmed h, a_regimen i " & _
        "WHERE a.min_codigo = b.mid_codigo " & _
        "AND   a.min_codreg = i.reg_codigo " & _
        "AND   a.min_codser = c.ser_codigo " & _
        "AND   b.mid_codrec = e.rec_codigo " & _
        "AND   e.rec_codigo = f.red_codigo " & _
        "AND   f.red_codpro = g.ing_codigo " & _
        "AND   g.ing_unimed = h.unm_codigo " & _
        "AND   b.mid_estser = d.ess_codigo AND a.min_cencos = d.ess_cencos " & _
        "AND   b.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
        "AND   a.min_codreg = " & IIf((IsMissing(codreg)), Val(fpLongInteger1(1).Value), codreg) & " " & _
        "AND   a.min_fecmin = " & Format(Fecha, "yyyymmdd") & " " & _
        "AND   b.mid_tipmin = '1' AND CDATE( FORMAT( Now(), 'dd/mm/yyyy' ) & ' ' & c.ser_horent ) >= Now() " & _
        "GROUP BY a.min_codigo, b.mid_tiprec, b.mid_numlin, i.reg_codigo, i.reg_nombre, c.ser_codigo, c.ser_nombre, c.ser_orden, c.ser_horent, " & _
        "d.ess_codigo, d.ess_nombre, d.ess_orden, e.rec_codigo, e.rec_nombre, g.ing_codigo, g.ing_nombre, h.unm_nomcor, f.red_nroite, f.red_canpro, f.red_cospro, " & _
        "f.red_pctapr, f.red_pctcoc, f.red_pctnut ORDER BY i.reg_codigo, c.ser_codigo, b.mid_numlin, c.ser_orden, d.ess_orden, f.red_nroite", vg_db, adOpenStatic
With grC
     indadi = -1
     lngIndexDeleteItem = -1
     If (IsMissing(codreg)) Then
        .Rows = 1
     End If
     If Not RS.EOF Then
        auxser = 0: auxess = 0
        auxhor = RS!ser_horent
        nomser = Trim(RS!ser_nombre)
        Do While Not RS.EOF
           If RS!ser_codigo <> auxser Or RS!reg_codigo <> auxreg Then
              If auxser > 0 Or auxreg > 0 Then
                '------- Mover totales recetas
                .Rows = .Rows + 1: .Row = .Rows - 1
                .IsSubtotal(.Row) = True
                .RowOutlineLevel(.Row) = 4
                .IsCollapsed(.Row) = flexOutlineCollapsed
                .Col = 0: .text = " TOTALES "
                .Col = 1: .text = "T"
                indrow = .Row
                For i = 10 To .Cols - 1
                    .Row = indrow
                    .Col = i
                    .text = Format(0, fg_Pict(6, 2))
                Next i
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
                .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
                 
                 '------- Agregar concepto adicionales
                 .Rows = .Rows + 1: .Row = .Rows - 1
                 .IsSubtotal(.Row) = True
                 .RowOutlineLevel(.Row) = 2
                 .IsCollapsed(.Row) = flexOutlineCollapsed
                 .Col = 0: .text = "Adicionales"
                 .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HC0&
                 .Col = 1: .text = "O;" & auxreg & ";" & auxser & ";" & -99999999
                 .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &HFF00&
                 .Col = 2
                 .text = "" & 0 & "" & _
                        ";" & Format(auxhor, "Hh:Nn") & ";" & _
                        Trim(nomser)
             End If
             .Rows = .Rows + 1: .Row = .Rows - 1
             .IsSubtotal(.Row) = True
             .RowOutlineLevel(.Row) = 1
             .IsCollapsed(.Row) = flexOutlineCollapsed
             .Col = 0: .text = Trim(RS!ser_nombre) & " - " & Trim(RS!reg_nombre)
             .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H800000 '&HFF0000
             .Col = 1: .text = "S;" & RS!reg_codigo & ";" & RS!ser_codigo & ";" & RS!ess_codigo
             .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
             auxser = RS!ser_codigo
             auxreg = RS!reg_codigo
             auxess = 0
             auxrec = 0
         End If
         If RS!ess_codigo <> auxess Then
            If auxrec > 0 Then
               '------- Mover totales recetas
               .Rows = .Rows + 1: .Row = .Rows - 1
               .IsSubtotal(.Row) = True
               .RowOutlineLevel(.Row) = 4
               .IsCollapsed(.Row) = flexOutlineCollapsed
               .Col = 0: .text = " TOTALES "
               .Col = 1: .text = "T"
               indrow = .Row
               For i = 10 To .Cols - 1
                   .Row = indrow
                   .Col = i
                   .text = Format(0, fg_Pict(6, 2))
               Next i
              .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
              .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
              .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
              .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            End If
            
            .Rows = .Rows + 1: .Row = .Rows - 1
            .IsSubtotal(.Row) = True
            .RowOutlineLevel(.Row) = 2
            .IsCollapsed(.Row) = flexOutlineCollapsed
            .Col = 0: .text = Trim(RS!ess_nombre)
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = IIf(RS!ess_codigo = -99999999, &HC0&, &H808080) '&HFF0000
            .Col = 1: .text = "O;" & RS!reg_codigo & ";" & RS!ser_codigo & ";" & RS!ess_codigo
            .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            auxess = RS!ess_codigo
            auxrec = 0
         End If
    
         If auxrec <> RS!rec_codigo Then
            If auxrec > 0 Then
               .Rows = .Rows + 1: .Row = .Rows - 1
               .IsSubtotal(.Row) = True
               .RowOutlineLevel(.Row) = 4
               .IsCollapsed(.Row) = flexOutlineCollapsed
               .Col = 0: .text = " TOTALES "
               .Col = 1: .text = "T"
               indrow = .Row
               For i = 10 To .Cols - 1
                   .Row = indrow
                   .Col = i
                   .text = Format(0, fg_Pict(6, 2))
               Next i
               .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
               .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
               .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
               .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            End If
            
            .Rows = .Rows + 1: .Row = .Rows - 1
            .IsSubtotal(.Row) = True
            .RowOutlineLevel(.Row) = 3
            grC.IsCollapsed(.Row) = flexOutlineCollapsed
            .Col = 0: .text = Trim(RS!rec_nombre)
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012  '&HFF&
            .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            .Col = 1: .text = "R;" & _
                              RS!reg_codigo & ";" & _
                              RS!ser_codigo & ";" & _
                              RS!ess_codigo & ";" & _
                              RS!rec_codigo & ";" & _
                              Trim(RS!ing_codigo) & ";" & _
                              RS!red_nroite & ";" & _
                              RS!min_codigo & ";" & _
                              RS!mid_tiprec
            .Col = 2: .text = "" & 0 & "" & _
                   ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
                   Trim(RS!ser_nombre)
            auxrec = RS!rec_codigo
         End If
        
         .Rows = .Rows + 1: .Row = .Rows - 1
         .IsSubtotal(.Row) = True
         .RowOutlineLevel(.Row) = 4
         .IsCollapsed(.Row) = flexOutlineCollapsed
         .Col = 0: .text = Trim(RS!ing_nombre)
         .Col = 1: .text = "I;" & _
                           RS!reg_codigo & ";" & _
                           RS!ser_codigo & ";" & _
                           RS!ess_codigo & ";" & _
                           RS!rec_codigo & ";" & _
                           Trim(RS!ing_codigo) & ";" & _
                           RS!red_nroite & ";" & _
                           RS!min_codigo & ";" & _
                           RS!mid_tiprec

         
         .Col = 3: .ColFormat(3) = fg_Pict(6, 2): .text = Format(RS!canser, fg_Pict(6, 2))
         .Col = 4: .text = Format(RS!red_canpro, fg_Pict(6, 2))
         .Col = 5: .text = Trim(RS!unm_nomcor)
         .Col = 6: .text = Format(RS!red_pctapr, fg_Pict(6, 2))
         .Col = 7: .text = Format(RS!red_pctcoc, fg_Pict(6, 2))
         .Col = 8: .text = Format(RS!red_pctnut, fg_Pict(6, 2))
         .Col = 9: .text = Format(RS!cannet, fg_Pict(6, 2))
         indrow = .Row
         auxing = RS!ing_codigo
         For i = 10 To .Cols - 1
             .Row = indrow
             .Col = i
             .text = Format(0, fg_Pict(6, 2))
         Next i
         .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
'         .Cell(flexcpForeColor, .Row, 3, .Row, 3) = &HFF0000
         .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
         .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HC0FFFF

'         Set nodNode = .Nodes.Add("O;" & RS!ser_codigo & ";" & RS!reg_codigo & ";" & RS!ess_codigo & ";", _
'                       tvwChild, "P;" & _
'                       RS!ser_codigo & ";" & _
'                       RS!ess_codigo & ";" & _
'                       RS!rec_codigo & ";" & _
'                       RS!reg_codigo & ";" & _
'                       RS!min_codigo & ";" & _
'                       RS!mid_tiprec & ";" & _
'                       RS!canser & ";", _
'                       Trim(RS!rec_nombre))
'         nodNode.Expanded = True
'         nodNode.Tag = "0" & _
'                       ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
'                      Trim(RS!ser_nombre)
         auxhor = RS!ser_horent
         nomser = Trim(RS!ser_nombre)
         RS.MoveNext
       Loop
       If auxser > 0 Then
          '------- Mover totales recetas
          .Rows = .Rows + 1: .Row = .Rows - 1
          .IsSubtotal(.Row) = True
          .RowOutlineLevel(.Row) = 4
          .IsCollapsed(.Row) = flexOutlineCollapsed
          .Col = 0: .text = " TOTALES "
          .Col = 1: .text = "T"
          indrow = .Row
          For i = 10 To .Cols - 1
              .Row = indrow
              .Col = i
              .text = Format(0, fg_Pict(6, 2))
          Next i
          .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
          .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
          .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
          .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
          
          '------- Agregar concepto adicionales
          .Rows = .Rows + 1: .Row = .Rows - 1
          .IsSubtotal(.Row) = True
          .RowOutlineLevel(.Row) = 2
          .IsCollapsed(.Row) = flexOutlineCollapsed
          .Col = 0: .text = "Adicionales"
          .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HC0&
          .Col = 1: .text = "O;" & auxreg & ";" & auxser & ";" & -99999999
          .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
          .Col = 2
          .text = "" & 0 & "" & _
                  ";" & Format(auxhor, "Hh:Nn") & ";" & _
                  Trim(nomser)
       End If
        If .Rows > 1 Then
           .Outline 4: .Outline 3: .Outline 2: .Outline 1
           .AutoSize 0, 0, False
           .Cell(flexcpPictureAlignment, 1, 1, .Rows - 1, 1) = flexPicAlignCenterCenter
        End If
    End If
    RS.Close: Set RS = Nothing
End With
est = False
End Sub

Sub TraerTomaPedido(codigo As Long, fecped As String)
If Trim(fpayuda(0).Caption) = "" Then Exit Sub
Dim sql1 As String, sql2 As String, sql3 As String, sql4 As String, sql5 As String, sql6 As String
Dim auxser As Long, auxess As Long, auxreg As Long, auxrec As Long, auxing As Long, auxhor As String, nomser As String
est = True
fg_carga ""
'------- Encabezado toma pedido paciente
sql5 = IIf(vg_tipbase = "1", " cdate('" & fecped & "') ", " '" & Format(fecped, "yyyymmdd") & "' ")
RS.Open "SELECT DISTINCT a.top_codigo, a.top_fecped, a.top_codreg, b.reg_nombre, a.top_codusu, c.usu_nombre " & _
        "FROM b_tomapedido a, a_regimen b, a_usuarios c " & _
        "WHERE a.top_codreg = b.reg_codigo " & _
        "AND   a.top_codusu = c.usu_codigo " & _
        "AND   a.top_codigo = " & codigo & " " & _
        "AND   a.top_fecped = " & sql5 & "", vg_db, adOpenStatic
If Not RS.EOF Then
   fpLongInteger1(0).Value = RS!top_codigo
   Date1(0).text = RS!top_fecped
   fpLongInteger1(1).Value = RS!top_codreg
   fpayuda(1).Caption = Trim(RS!reg_nombre)
End If
RS.Close: Set RS = Nothing
sql5 = IIf(vg_tipbase = "1", " cdate('" & fecped & "') ", " '" & Format(fecped, "yyyymmdd") & "' ")
sql6 = IIf(vg_tipbase = "1", " ORDER BY f.reg_codigo, g.ser_codigo, g.ser_orden, h.ess_orden ", "")
sql1 = "SELECT DISTINCT f.reg_codigo, f.reg_nombre, g.ser_codigo, g.ser_nombre, g.ser_orden, g.ser_horent, h.ess_codigo, h.ess_nombre, h.ess_orden, " & _
       "c.tdr_codrec, d.rec_nombre, c.tdr_nroite, c.tdr_coding, c.tdr_canpro, c.tdr_cospro, c.tdr_pctapr, c.tdr_pctcoc, c.tdr_pctnut, " & _
       "e.ing_nombre, i.unm_nomcor, b.tpd_codmin, b.tpd_tiprec, b.tpd_cansel, b.tpd_canser, b.tpd_caning, b.tpd_numlin, b.tpd_prorec " & _
       "FROM b_tomapedido a, b_tomapedidodet b, b_tomapedidodetrec c, b_receta d, b_ingrediente e, a_regimen f, a_servicio g, a_estservicio h, a_unidadmed i " & _
       "WHERE a.top_codigo = b.tpd_codigo " & _
       "AND   b.tpd_codigo = c.tdr_codigo " & _
       "AND   b.tpd_numlin = c.tdr_numlin " & _
       "AND   b.tpd_codrec = c.tdr_codrec " & _
       "AND   b.tpd_codreg = f.reg_codigo " & _
       "AND   b.tpd_codser = g.ser_codigo " & _
       "AND   b.tpd_estser = h.ess_codigo AND h.ess_cencos = '" & MuestraCasino(1) & "' " & _
       "AND   c.tdr_codrec = d.rec_codigo " & _
       "AND   c.tdr_coding = e.ing_codigo " & _
       "AND   e.ing_unimed = i.unm_codigo " & _
       "AND   a.top_codigo = " & codigo & " " & _
       "AND   a.top_fecped = " & sql5 & " AND b.tpd_prorec = 'R' " & _
       "" & sql6 & ""

sql6 = IIf(vg_tipbase = "1", " ORDER BY f.reg_codigo, g.ser_codigo, g.ser_orden, ess_orden ", "")
sql2 = "SELECT  DISTINCT f.reg_codigo, f.reg_nombre, g.ser_codigo, g.ser_nombre, g.ser_orden, g.ser_horent, b.tpd_estser AS ess_codigo, " & _
       "'Adicionales' AS ess_nombre, 99999999 AS ess_orden, c.tdr_codrec, d.rec_nombre, c.tdr_nroite, c.tdr_coding, " & _
       "c.tdr_canpro, c.tdr_cospro, c.tdr_pctapr, c.tdr_pctcoc, c.tdr_pctnut, e.ing_nombre, i.unm_nomcor, b.tpd_codmin, " & _
       "b.tpd_tiprec, b.tpd_cansel, b.tpd_canser, b.tpd_caning, b.tpd_numlin, b.tpd_prorec " & _
       "FROM b_tomapedido a, b_tomapedidodet b, b_tomapedidodetrec c, b_receta d, b_ingrediente e, a_regimen f, a_servicio g, a_unidadmed i " & _
       "WHERE a.top_codigo = b.tpd_codigo " & _
       "AND   b.tpd_codigo = c.tdr_codigo " & _
       "AND   b.tpd_numlin = c.tdr_numlin " & _
       "AND   b.tpd_codrec = c.tdr_codrec " & _
       "AND   b.tpd_codreg = f.reg_codigo " & _
       "AND   b.tpd_codser = g.ser_codigo " & _
       "AND   c.tdr_codrec = d.rec_codigo " & _
       "AND   c.tdr_coding = e.ing_codigo " & _
       "AND   e.ing_unimed = i.unm_codigo " & _
       "AND   a.top_codigo = " & codigo & " " & _
       "AND   a.top_fecped = " & sql5 & " AND b.tpd_prorec = 'R' AND b.tpd_estser = -99999999 " & _
       "" & sql6 & ""

sql6 = IIf(vg_tipbase = "1", " ORDER BY f.reg_codigo, g.ser_codigo, g.ser_orden, ess_orden ", "")
sql3 = "SELECT DISTINCT f.reg_codigo, f.reg_nombre, g.ser_codigo, g.ser_nombre, g.ser_orden, g.ser_horent, b.tpd_estser AS ess_codigo, " & _
       "'Adicionales' AS ess_nombre, 99999999 AS ess_orden, c.tdr_codrec, d.pro_nombre, 0 AS tdr_nroite, c.tdr_coding, " & _
       "c.tdr_canpro, c.tdr_cospro, c.tdr_pctapr, c.tdr_pctcoc, c.tdr_pctnut, e.ing_nombre, i.unm_nomcor, b.tpd_codmin, " & _
       "b.tpd_tiprec, b.tpd_cansel, b.tpd_canser, b.tpd_caning, b.tpd_numlin, b.tpd_prorec " & _
       "FROM b_tomapedido a, b_tomapedidodet b, b_tomapedidodetrec c, b_productos d, b_ingrediente e, a_regimen f, a_servicio g, a_unidadmed i " & _
       "WHERE a.top_codigo = b.tpd_codigo " & _
       "AND b.tpd_codigo = c.tdr_codigo " & _
       "AND b.tpd_numlin = c.tdr_numlin " & _
       "AND b.tpd_codrec = c.tdr_codrec " & _
       "AND b.tpd_codreg = f.reg_codigo " & _
       "AND b.tpd_codser = g.ser_codigo " & _
       "AND c.tdr_codrec = val(d.pro_codigo) " & _
       "AND c.tdr_coding = e.ing_codigo " & _
       "AND e.ing_unimed = i.unm_codigo " & _
       "AND a.top_codigo = " & codigo & " " & _
       "AND a.top_fecped = " & sql5 & " AND b.tpd_prorec = 'P' AND b.tpd_estser = -99999999 " & _
       "" & sql6 & ""

sql5 = IIf(vg_tipbase = "1", " CDATE( FORMAT( Now(), 'dd/mm/yyyy' ) ", " " & Format(Now(), "yyyymmdd") & " ")
sql4 = "SELECT DISTINCT d.reg_codigo AS reg_codigo, d.reg_nombre, e.ser_codigo AS ser_codigo, e.ser_nombre, e.ser_orden AS ser_orden, e.ser_horent, " & _
       "f.ess_codigo, f.ess_nombre, f.ess_orden AS ess_orden, c.rec_codigo, c.rec_nombre, g.red_nroite AS tdr_nroite, g.red_codpro AS tdr_coding, " & _
       "g.red_canpro AS tdr_canpro, g.red_cospro AS tdr_cospro, g.red_pctapr AS tdr_pctapr, g.red_pctcoc AS tdr_pctcoc, g.red_pctnut AS tdr_pctnut, " & _
       "h.ing_nombre, i.unm_nomcor, a.min_codigo AS tdp_codmin, b.mid_tiprec AS tpd_tiprec, 0 AS tpd_cansel, 0 AS tpd_canser, 0 AS tpd_caning, 0 AS tpd_numlin, 'R' AS tpd_prorec " & _
       "FROM b_minuta a, b_minutadet b, b_receta c, a_regimen d, a_servicio e, a_estservicio f, b_recetadet g, b_ingrediente h, a_unidadmed i " & _
       "WHERE a.min_codigo = b.mid_codigo " & _
       "AND   a.min_codreg = d.reg_codigo " & _
       "AND   a.min_codser = e.ser_codigo " & _
       "AND   b.mid_codrec = c.rec_codigo " & _
       "AND   c.rec_codigo = g.red_codigo " & _
       "AND   b.mid_tiprec = g.red_tiprec AND ((g.red_tiprec <> 0 AND g.red_cencos = '" & MuestraCasino(1) & "') OR (g.red_tiprec = 0 AND g.red_cencos = '0')) " & _
       "AND   b.mid_estser = f.ess_codigo AND a.min_cencos = f.ess_cencos " & _
       "AND   a.min_codreg IN (SELECT DISTINCT bb.tpd_codreg FROM b_tomapedido aa, b_tomapedidodet bb WHERE aa.top_codigo = bb.tpd_codigo AND aa.top_codigo = " & codigo & ") " & _
       "AND   a.min_codser IN (SELECT DISTINCT bb.tpd_codser FROM b_tomapedido aa, b_tomapedidodet bb WHERE aa.top_codigo = bb.tpd_codigo AND aa.top_codigo = " & codigo & ") " & _
       "AND   b.mid_codrec NOT IN (SELECT DISTINCT bb.tpd_codrec FROM b_tomapedido aa, b_tomapedidodet bb WHERE aa.top_codigo = bb.tpd_codigo AND a.min_codreg = bb.tpd_codreg AND a.min_codser = bb.tpd_codser AND aa.top_codigo = " & codigo & " AND tpd_prorec = 'R') " & _
       "AND   g.red_codpro = h.ing_codigo " & _
       "AND   h.ing_unimed = i.unm_codigo " & _
       "AND   a.min_fecmin = " & Format(fecped, "yyyymmdd") & "  AND a.min_fecmin = " & Format(Date, "yyyymmdd") & " " & _
       "AND   b.mid_tipmin = '1' AND " & sql5 & " & ' ' & e.ser_horent ) >= Now() " & _
       "ORDER BY reg_codigo, ser_codigo, ser_orden, ess_orden"

RS.Open sql1 & " UNION " & sql2 & " UNION " & sql3 & " UNION " & sql4, vg_db, adOpenStatic
With grC
    lngIndexDeleteItem = -1
    grC.Rows = 1
    If Not RS.EOF Then
       auxser = 0: auxess = 0: auxreg = 0
       auxhor = RS!ser_horent
       nomser = Trim(RS!ser_nombre)
       Do While Not RS.EOF
          If RS!ser_codigo <> auxser Or auxreg <> RS!reg_codigo Then
             '------- Agregar concepto adicionales
             If auxser > 0 Or auxreg > 0 Then
                '------- Mover totales recetas
                .Rows = .Rows + 1: .Row = .Rows - 1
                .IsSubtotal(.Row) = True
                .RowOutlineLevel(.Row) = 4
                .IsCollapsed(.Row) = flexOutlineCollapsed
                .Col = 0: .text = " TOTALES "
                .Col = 1: .text = "T"
                indrow = .Row
                For i = 10 To .Cols - 1
                    .Row = indrow
                    .Col = i
                    .text = Format(0, fg_Pict(6, 2))
                Next i
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
                .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
                .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
                .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
             End If
             If auxser > 0 And auxess <> -99999999 Then
                 '------- Agregar concepto adicionales
                 .Rows = .Rows + 1: .Row = .Rows - 1
                 .IsSubtotal(.Row) = True
                 .RowOutlineLevel(.Row) = 2
                 .IsCollapsed(.Row) = flexOutlineCollapsed
                 .Col = 0: .text = "Adicionales"
                 .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HC0&
                 .Col = 1: .text = "O;" & auxreg & ";" & auxser & ";" & -99999999
                 .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &HFF00&
                 .Col = 2
                 .text = "" & 0 & "" & _
                        ";" & Format(auxhor, "Hh:Nn") & ";" & _
                        Trim(nomser)
             End If
             
             .Rows = .Rows + 1: .Row = .Rows - 1
             .IsSubtotal(.Row) = True
             .RowOutlineLevel(.Row) = 1
             .IsCollapsed(.Row) = flexOutlineCollapsed
             .Col = 0: .text = Trim(RS!ser_nombre) & " - " & Trim(RS!reg_nombre)
             .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H800000 '&HFF0000
             .Col = 1: .text = "S;" & RS!reg_codigo & ";" & RS!ser_codigo & ";" & RS!ess_codigo
             .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
             auxser = RS!ser_codigo
             auxreg = RS!reg_codigo
             auxess = 0
             auxrec = 0
         End If
         
         If RS!ess_codigo <> auxess Then
            If auxrec > 0 Then
               '------- Mover totales recetas
               .Rows = .Rows + 1: .Row = .Rows - 1
               .IsSubtotal(.Row) = True
               .RowOutlineLevel(.Row) = 4
               .IsCollapsed(.Row) = flexOutlineCollapsed
               .Col = 0: .text = " TOTALES "
               .Col = 1: .text = "T"
               indrow = .Row
               For i = 10 To .Cols - 1
                   .Row = indrow
                   .Col = i
                   .text = Format(0, fg_Pict(6, 2))
               Next i
              .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
              .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
              .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
              .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            End If
            
            .Rows = .Rows + 1: .Row = .Rows - 1
            .IsSubtotal(.Row) = True
            .RowOutlineLevel(.Row) = 2
            .IsCollapsed(.Row) = flexOutlineCollapsed
            .Col = 0: .text = Trim(RS!ess_nombre)
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = IIf(RS!ess_codigo = -99999999, &HC0&, &H808080) '&HFF0000
            .Col = 1: .text = "O;" & RS!reg_codigo & ";" & RS!ser_codigo & ";" & RS!ess_codigo
            .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            auxess = RS!ess_codigo
            auxrec = 0
         End If
    
'         Set nodNode = .Nodes.Add("O;" & RS!ser_codigo & ";" & RS!reg_codigo & ";" & RS!ess_codigo & ";", _
'                       tvwChild, "P;" & _
'                       RS!ser_codigo & ";" & _
'                       RS!ess_codigo & ";" & _
'                       RS!tdr_codrec & ";" & _
'                       RS!reg_codigo & ";" & _
'                       RS!tpd_codmin & ";" & _
'                       IIf(RS!tpd_prorec = "P", "P", RS!tpd_tiprec) & ";" & _
'                       RS!tpd_cansel & ";", _
'                       IIf(RS!tpd_cansel > 0, Trim(RS!rec_nombre) & " (" & RS!tpd_cansel & ")", Trim(RS!rec_nombre)))
'         nodNode.Expanded = True
'         nodNode.Tag = "" & IIf(RS!tpd_cansel > 0, 1, 0) & "" & _
'                       ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
'                      Trim(RS!ser_nombre)
         
         If auxrec <> RS!tdr_codrec Then
            If auxrec > 0 Then
               .Rows = .Rows + 1: .Row = .Rows - 1
               .IsSubtotal(.Row) = True
               .RowOutlineLevel(.Row) = 4
               .IsCollapsed(.Row) = flexOutlineCollapsed
               .Col = 0: .text = " TOTALES "
               .Col = 1: .text = "T"
               indrow = .Row
               For i = 10 To .Cols - 1
                   .Row = indrow
                   .Col = i
                   .text = Format(0, fg_Pict(6, 2))
               Next i
               .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012
               .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
               .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
               .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
            End If
            
            .Rows = .Rows + 1: .Row = .Rows - 1
            .IsSubtotal(.Row) = True
            .RowOutlineLevel(.Row) = 3
            grC.IsCollapsed(.Row) = flexOutlineCollapsed
            .Col = 0: .text = IIf(RS!tpd_cansel > 0, Trim(RS!rec_nombre) & " (" & RS!tpd_cansel & ")", Trim(RS!rec_nombre))
            .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012  '&HFF&
            .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
'            GrC.Cell(flexcpBackColor, GrC.Rows - 1, 1) = &HF0F0F0
            .Col = 1: .text = Trim(RS!tpd_prorec) & ";" & _
                              RS!reg_codigo & ";" & _
                              RS!ser_codigo & ";" & _
                              RS!ess_codigo & ";" & _
                              RS!tdr_codrec & ";" & _
                              Trim(RS!tdr_coding) & ";" & _
                              RS!tdr_nroite & ";" & _
                              RS!tpd_codmin & ";" & _
                              RS!tpd_tiprec
            .Col = 2: .text = "" & IIf(RS!tpd_cansel > 0, 1, 0) & "" & _
                              ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
                              Trim(RS!ser_nombre)
            auxrec = RS!tdr_codrec
         End If
        
         .Rows = .Rows + 1: .Row = .Rows - 1
         .IsSubtotal(.Row) = True
         .RowOutlineLevel(.Row) = 4
         .IsCollapsed(.Row) = flexOutlineCollapsed
         .Col = 0: .text = Trim(RS!ing_nombre)
         .Col = 1: .text = IIf(Trim(RS!tpd_prorec) = "R", "I;", "I;") & _
                           RS!reg_codigo & ";" & _
                           RS!ser_codigo & ";" & _
                           RS!ess_codigo & ";" & _
                           RS!tdr_codrec & ";" & _
                           Trim(RS!tdr_coding) & ";" & _
                           RS!tdr_nroite & ";" & _
                           RS!tpd_codmin & ";" & _
                           RS!tpd_tiprec

         
         .Col = 3: .ColFormat(3) = fg_Pict(6, 2): .text = Format((((RS!tdr_pctapr / 100) * RS!tdr_canpro) * (RS!tdr_pctcoc / 100)), fg_Pict(6, 2))
         grC.Col = 4: grC.text = Format(RS!tdr_canpro, fg_Pict(6, 2))
         grC.Col = 5: grC.text = Trim(RS!unm_nomcor)
         grC.Col = 6: grC.text = Format(RS!tdr_pctapr, fg_Pict(6, 2))
         grC.Col = 7: grC.text = Format(RS!tdr_pctcoc, fg_Pict(6, 2))
         grC.Col = 8: grC.text = Format(RS!tdr_pctnut, fg_Pict(6, 2))
         grC.Col = 9: grC.text = Format(((RS!tdr_pctnut / 100) * (RS!tdr_canpro)), fg_Pict(6, 2))
         indrow = .Row
         auxing = RS!tdr_coding
         For i = 10 To .Cols - 1
             .Row = indrow
             .Col = i
             .text = Format(0, fg_Pict(6, 2))
         Next i
         .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012  '&HFF&
         .Cell(flexcpForeColor, .Row, 3, .Row, 3) = IIf((Format(Date1(0).text, "yyyymmdd") < Format(Date, "yyyymmdd")), &H80000012, &HFF0000)
         .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
         .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HC0FFFF
         
         auxhor = RS!ser_horent
         nomser = Trim(RS!ser_nombre)
         RS.MoveNext
       Loop
       If auxser > 0 Then
          '------- Mover totales recetas
          .Rows = .Rows + 1: .Row = .Rows - 1
          .IsSubtotal(.Row) = True
          .RowOutlineLevel(.Row) = 4
          .IsCollapsed(.Row) = flexOutlineCollapsed
          .Col = 0: .text = " TOTALES "
          .Col = 1: .text = "T"
          indrow = .Row
          For i = 10 To .Cols - 1
              .Row = indrow
              .Col = i
              .text = Format(0, fg_Pict(6, 2))
          Next i
          .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &H80000012   '&HFF&
          .Cell(flexcpFontBold, .Row, 0, .Row, .Cols - 1) = True
          .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = &HFF&
          .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
       End If
       If auxser > 0 And auxess <> -99999999 Then
          '------- Agregar concepto adicionales
          .Rows = .Rows + 1: .Row = .Rows - 1
          .IsSubtotal(.Row) = True
          .RowOutlineLevel(.Row) = 2
          .IsCollapsed(.Row) = flexOutlineCollapsed
          .Col = 0: .text = "Adicionales"
          .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = &HC0&
          .Col = 1: .text = "O;" & auxreg & ";" & auxser & ";" & -99999999
          .Cell(flexcpBackColor, .Row, 0, .Row, 1) = &H80FF80
          .Col = 2
          .text = "" & 0 & "" & _
                  ";" & Format(auxhor, "Hh:Nn") & ";" & _
                  Trim(nomser)
       End If
       If .Rows > 1 Then
          .Outline 4: .Outline 3: .Outline 2: .Outline 1
          .AutoSize 0, 0, False
          .Cell(flexcpPictureAlignment, 1, 1, .Rows - 1, 1) = flexPicAlignCenterCenter
       End If
       
       '------- Agregar concepto adicionales
       If auxser > 0 And auxess <> -99999999 Then
'          Set nodNode = .Nodes.Add("S;" & auxser & ";" & auxreg & ";", _
'                        tvwChild, "O;" & auxser & ";" & auxreg & ";" & -99999999 & ";", _
'                        "Adicionales")
'          nodNode.Expanded = True
'          nodNode.ForeColor = &HC0&      '&H808080
'          nodNode.Bold = True
'
'          nodNode.Tag = "" & 0 & "" & _
'                        ";" & Format(auxhor, "Hh:Nn") & ";" & _
'                        Trim(nomser)
        End If
    End If
    RS.Close: Set RS = Nothing
End With
est = False
fg_descarga
End Sub

Private Sub grC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
If grC.Rows < 1 Then Exit Sub
grC.Row = Row
grC.Col = Col
If grC.Editable = flexEDKbd Then
   Dim canpro As Double, pctapr As Double, pctcoc As Double, canbru As Double, cannet As Double, pctnet As Double
   modo = "M"
   Gl_Ac_Botones Me, 13, 0, modo
   With grC
        .Row = Row
        'Cantidad servida
        .Col = 3: canser = .text
        'Calcular cantidad bruta
        .Col = 6: pctapr = .text
        .Col = 7: pctcoc = .text
        canbru = (((canser / pctapr) / pctcoc) * 10000)
        .Col = 4: .text = Format(canbru, fg_Pict(6, 2))
        'Calcular cantidad neta
        .Col = 8: pctnet = .text
        cannet = CCur((pctnet / 100) * canbru)
        .Col = 9: .text = Format(cannet, fg_Pict(6, 2))
        .Col = 1
        CalcularAporte GetItem(grC.text, 6), canbru, pctnet, .Row
        .Row = Row
        .Col = 4
'        Case 7, 8 '%Aprovechamiento
'            .Col = 5: canbru = .text
'            .Col = 7: pctapr = .text
'            .Col = 8: pctcoc = .text
'            '------- Calcular cantidad servir
'            .Col = 4
'            .text = Format(CCur(((pctapr / 100) * canbru) * (pctcoc / 100)), fg_Pict(6, 2))
'        Case 9
'           'Calcular cantidad neta
'            .Col = 5: canbru = .text
'            .Col = 9: pctnet = .text
'            cannet = CCur((pctnet / 100) * canbru)
'            .Col = 10: .text = Format(cannet, fg_Pict(6, 2))
'            .Col = 2
'            CalcularAporte .text, canbru, pctnet, .Row
'        End Select
    End With

End If

End Sub

Private Sub grC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
If grC.Rows < 1 Or est Then Exit Sub
indfil = NewRow
grC.Editable = IIf(NewCol = 3 And grC.CellBackColor = &HC0FFFF And grC.CellForeColor = &HFF0000, flexEDKbd, flexEDNone)
End Sub

Private Sub grC_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
If est Then Exit Sub
indfil = Row
grC.Row = Row
grC.Col = 1
If GetItem(grC.text, 1) <> "R" And GetItem(grC.text, 1) <> "P" Then Exit Sub
grC.Col = 2
estedi = IIf(CInt(Trim(GetItem(grC.text, 1))) > 0, True, False)
'------- Calcular aportes nutricionales
Dim canpro As Double, pctnut As Double, i As Long, j As Long, X As Long, indnut  As Long, indfin As Long, totapo As Double, canser As Double, canbru As Double, cannet As Double
grC.Col = 1
If GetItem(grC.text, 1) = "P" Then
   '------- Calcular aportes nutricionales con producto
   RS.Open "SELECT DISTINCT c.ing_codigo, e.nut_nombre, e.nut_codigo, " & _
           "d.pnu_canapo, a.pro_facsto, c.ing_facnut, e.nut_secnro " & _
           "FROM  b_productos a, b_productosing b, b_ingrediente c, b_productonut d , a_nutriente e " & _
           "WHERE a.pro_codigo = b.pri_codpro " & _
           "AND   b.pri_coding = c.ing_codigo " & _
           "AND   c.ing_codigo = d.pnu_codpro " & _
           "AND   d.pnu_codapo = e.nut_codigo " & _
           "AND   a.pro_codigo = '" & Val(GetItem(grC.text, 5)) & "' " & _
           "AND   c.ing_facnut > 0 ORDER BY e.nut_secnro", vg_db, adOpenStatic
Else
   '------- Calcular aportes nutricionales con receta
   RS.Open "SELECT DISTINCT c.ing_codigo, b.red_nroite, e.nut_nombre, e.nut_codigo, " & _
           "d.pnu_canapo, a.rec_basrac, c.ing_facnut, e.nut_secnro " & _
           "FROM  b_receta a, b_recetadet b, b_ingrediente c, b_productonut d , a_nutriente e " & _
           "WHERE b.red_codigo = a.rec_codigo " & _
           "AND   b.red_codpro = c.ing_codigo " & _
           "AND   c.ing_codigo = d.pnu_codpro " & _
           "AND   d.pnu_codapo = e.nut_codigo " & _
           "AND   b.red_codigo = " & Val(GetItem(grC.text, 5)) & " " & _
           "AND   b.red_tiprec = " & Val(GetItem(grC.text, 9)) & " AND ((b.red_tiprec <> 0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) " & _
           "AND   c.ing_facnut > 0 " & _
           "ORDER BY e.nut_secnro, b.red_nroite", vg_db, adOpenStatic
End If
'------- Activar grilla
With grC
     indrow = .Row
     If Not RS.EOF Then
        Do While Not RS.EOF
           For i = indrow + 1 To .Rows - 1
               .Row = i
               .Col = 1
               If GetItem(.text, 1) = "T" Then indrow = .Row: indfin = i: Exit For
               If GetItem(.text, 1) = "I" And RS!ing_codigo = GetItem(.text, 6) And 0 = Val(GetItem(.text, 7)) Then
'                  .Cell(flexcpForeColor, .Row, 3, .Row, 3) = IIf(estedi, &HFF0000, &H80000012)
                  .Col = 4: canpro = .text
                  .Col = 8: pctnut = .text
                   .Row = 0
                    For X = 10 To .Cols - 1
                        .Col = X
                         If Val(GetItem(.text, 1)) = RS!nut_codigo Then
                            .Row = i
                            .text = Format(((((pctnut / 100) * (RS!pnu_canapo * (canpro / 1))) / RS!ing_facnut)), fg_Pict(6, 2))
                            Exit For
                         End If
                     Next X
                    Exit For
               ElseIf GetItem(.text, 1) = "I" And _
                      RS!ing_codigo = Trim(GetItem(.text, 6)) And _
                      RS!red_nroite = Val(GetItem(.text, 7)) Then
'                      .Cell(flexcpForeColor, .Row, 3, .Row, 3) = IIf(estedi, &HFF0000, &H80000012)
                      .Col = 4: canpro = .text
                      .Col = 8: pctnut = .text
                      .Row = 0
                      For X = 10 To .Cols - 1
                          .Col = X
                          If Val(GetItem(.text, 1)) = RS!nut_codigo Then
                             .Row = i
                             .text = Format(((((pctnut / 100) * (RS!pnu_canapo * (canpro / RS!rec_basrac))) / RS!ing_facnut)), fg_Pict(6, 2))
                            Exit For
                         End If
                     Next X
                     Exit For
               End If
           Next i
           RS.MoveNext
        Loop
     End If
     RS.Close: Set RS = Nothing
     '------- Buscar linea totales
     For i = indrow To .Rows - 1
         .Row = i
         .Col = 1
         If GetItem(.text, 1) = "T" Then indfin = i: Exit For
     Next i
     '------ Mover zero totales
     For i = indrow To indfin - 1
         .Row = indfin
         .Col = 3: .text = Format(0, fg_Pict(6, 2))
         .Col = 4: .text = Format(0, fg_Pict(6, 2))
         .Col = 9: .text = Format(0, fg_Pict(6, 2))
         For j = 10 To .Cols - 1
             .Row = indfin
             .Col = j: .text = Format(0, fg_Pict(6, 2))
         Next j
     Next i
     '------- Mover aportes totales
     For i = (indrow + 1) To (indfin - 1)
         .Row = i
         .Col = 3: canser = .text
         .Col = 4: canbru = .text
         .Col = 9: cannet = .text
         .Row = indfin
         .Col = 3: .text = Format(.text + canser, fg_Pict(6, 2))
         .Col = 4: .text = Format(.text + canbru, fg_Pict(6, 2))
         .Col = 9: .text = Format(.text + cannet, fg_Pict(6, 2))
         For j = 10 To .Cols - 1
             .Row = i
             .Col = j
             totapo = .text
             .Row = indfin
             .Col = j
             .text = Format(.text + totapo, fg_Pict(6, 2))
         Next j
     Next i
End With
End Sub

Private Sub grC_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'    ' left button, no shift: start tracking mouse to drag
'    If (Button = 1) And (Shift = 0) And grC.Rows < 1 Then
'        If (chkDrag.Value <> 0) And (g_DragInfo.bDragging = False) Then
'
'            ' save current row and mouse position
'            g_DragInfo.lSrc = grC.Row
'            g_DragInfo.xDown = X
'            g_DragInfo.yDown = y
'
'            ' start checking
'            g_DragInfo.bCheckDrag = True
'        End If
'    End If

indfil = grC.Row
If Button = 1 Or grC.Rows < 1 Then Exit Sub
grC.Row = indfil
grC.Col = 1
If GetItem(grC.text, 4) = -99999999 And (GetItem(grC.text, 1) = "O" Or GetItem(grC.text, 1) = "R") Then
   OpGrilla(10).Visible = IIf(GetItem(grC.text, 1) = "O", True, False)
   OpGrilla(30).Visible = IIf(GetItem(grC.text, 1) = "O", False, True)
   OpGrilla(20).Visible = IIf(GetItem(grC.text, 1) = "O", True, False)

   estval = False
   If (Not blnValidSelectedPreparacion(grC.Row)) Then Exit Sub
    Select Case Button
    Case 2
        PopupMenu MenuDetalle
    End Select
End If
End Sub

Private Sub grC_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
Select Case Col
Case 3
    If Val(grC.EditText) <= 0 Then MsgBox "Campo debe ser numerico", vbCritical, Msgtitulo: Beep: Cancel = True
End Select
End Sub

Private Sub Image1_Click(Index As Integer)
est = True
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_pacientes", "pac_", "Pacientes", "Pac"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(Index).text = fg_PintaRut(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    RS.Open "SELECT b_pacientes.*, a_grupopaciente.grp_nombre, a_regimen.reg_nombre " & _
            "FROM (a_grupopaciente RIGHT JOIN b_pacientes ON a_grupopaciente.grp_codigo = b_pacientes.pac_codgrp) " & _
            "LEFT JOIN a_regimen ON b_pacientes.pac_codreg = a_regimen.reg_codigo WHERE b_pacientes.pac_codigo = '" & vg_codigo & "'", vg_db, adOpenStatic
    codreg = 0
    If Not RS.EOF Then
       fpayuda(0).Caption = Trim(RS!pac_nombre) & " " & Trim(RS!pac_appaterno) & " " & Trim(RS!pac_apmaterno)
       Limpiar 1
       est = True
       fpText(1).text = Trim(RS!grp_nombre)
       Image1(1).Enabled = IIf(IsNull(RS!pac_codreg), False, True)
       fpLongInteger1(1).Value = IIf(IsNull(RS!pac_codreg), "", RS!pac_codreg)
       codreg = IIf(IsNull(RS!pac_codreg), 0, RS!pac_codreg)
       fpayuda(1).Caption = IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre)
       fpText(3).text = IIf(IsNull(RS!pac_nrocam), "", RS!pac_nrocam)
       If modo = "A" And IsNull(RS!reg_nombre) Then
          For i = 1 To 2
              Frame1(i).Enabled = True
          Next i
          fpLongInteger1(0).Enabled = False: fpLongInteger1(1).Enabled = False
        ElseIf modo = "A" And Not IsNull(RS!reg_nombre) Then
           For i = 1 To 2
               Frame1(i).Enabled = True
           Next i
           fpLongInteger1(1).Enabled = True: Image1(1).Enabled = True
           Check1(0).Value = 0
        End If
        est = False
        Toolbar1.Buttons(15).Enabled = True: Toolbar1.Buttons(15).ToolTipText = "Historico Toma Pedido Pacientes "
    End If
    fpLongInteger1(0).Value = ""
    RS.Close: Set RS = Nothing
Case 1
    est = False
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = "": vg_codregimen = 0
    vg_Aux = fg_DespintaRut(fpText(0).text)
    RS.Open "SELECT DISTINCT pac_codreg FROM b_pacientes WHERE pac_codigo = '" & vg_Aux & "'", vg_db, adOpenStatic
    If Not RS.EOF Then vg_codregimen = RS!pac_codreg
    RS.Close: Set RS = Nothing
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", IIf(Check1(0).Value = 1, "Reg", "Gen")
    B_TabEst.Show 1
    Me.Refresh
    vg_codregimen = 0
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = Trim(vg_nombre)
    fpayuda(1).Tag = Trim(vg_nombre)
    If Trim(Date1(0).text) = "" Then Exit Sub
    MostrarTomaPedido Date1(0).text
Case 3
    est = False
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = "": vg_codregimen = 0
    vg_Aux = fg_DespintaRut(fpText(0).text)
    RS.Open "SELECT DISTINCT pac_codreg FROM b_pacientes WHERE pac_codigo = '" & vg_Aux & "'", vg_db, adOpenStatic
    If Not RS.EOF Then vg_codregimen = RS!pac_codreg
    RS.Close: Set RS = Nothing
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", IIf(Check1(0).Value = 1, "Reg", "Gen")
    B_TabEst.Show 1
    Me.Refresh
    vg_codregimen = 0
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(3).Caption = Trim(vg_nombre)
    fpayuda(3).Tag = Trim(vg_nombre)
End Select
End Sub

Private Sub tlbPreview_ButtonClick(ByVal Button As MSComctlLib.Button)
estval = False
If grC.Rows < 1 Then Exit Sub
grC.Row = indfil
grC.Col = 1
If GetItem(grC.text, 1) <> "R" And GetItem(grC.text, 1) <> "P" Then Exit Sub
If (Not blnValidSelectedPreparacion(grC.Row)) Or (GetItem(grC.text, 1) <> "R" And GetItem(grC.text, 1) <> "P") Then Exit Sub
Select Case Button.Index
Case 1
    grC.Row = indfil
    grC.Col = 1
    codess = GetItem(grC.text, 4)
    grC.Col = 2
    If (CInt(GetItem(grC.text, 1)) = 1) And codess > 0 Then Exit Sub
    grC.Col = 1
    Call SetPreviewPedidoSelectedPreparacion(GetItem(grC.text, 4), GetItem(grC.text, 5), GetItem(grC.text, 2))
Case 3
    grC.Col = 2
    If (CInt(GetItem(grC.text, 1)) = 0) Then Exit Sub
    grC.Col = 1
    Call SetPreviewPedidoSelectedPreparacion(GetItem(grC.text, 4), GetItem(grC.text, 5), GetItem(grC.text, 2), False)
End Select
Gl_Ac_Botones Me, 13, 0, modo
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codpac As String, codigo As Long, codser As Long, codmin As Long, codess As Long, codrec As Long, cansel As Long, i As Long, indlin As Long, tiprec As Long, canser As Double, caning As Double
Dim coding As String, nroite As Long, canbru As Double, pctapr As Double, pctcoc As Double, pctnut As Double
On Error GoTo Man_Error
Select Case Button.Index
Case 1 '------- Agregar
    modo = "A"
    fpLongInteger1(0).Value = ""
    Limpiar 1
    Frame1(0).Enabled = True
    If fpayuda(0).Caption <> "" Then
       Frame1(1).Enabled = True
       Frame1(2).Enabled = True
    End If
    Gl_Ac_Botones Me, 13, 0, modo
Case 3 '------- Modificar
    modo = "M"
    Gl_Ac_Botones Me, 13, 0, modo
Case 5 '------- Eliminar
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vg_db.BeginTrans
    vg_db.Execute "DELETE b_tomapedidodetrec FROM b_tomapedidodetrec WHERE tdr_codigo = " & Val(fpLongInteger1(0).Value)
    vg_db.Execute "DELETE b_tomapedidodet FROM b_tomapedidodet WHERE tpd_codigo = " & Val(fpLongInteger1(0).Value)
    vg_db.Execute "DELETE b_tomapedido FROM b_tomapedido WHERE top_codigo = " & Val(fpLongInteger1(0).Value)
    vg_db.CommitTrans
    Limpiar 0
    For i = 1 To 2
        Frame1(i).Enabled = False
    Next i
    modo = "": Gl_Ac_Botones Me, 13, 2, modo
Case 7 '------- Actualizar
   If Format(Date1(0).text, "yyyymmdd") >= Format(Date, "yyyymmdd") And Trim(fpText(0).text) <> "" And Val(fpLongInteger1(0).Value) > 0 And grC.Rows < 1 Then
      MostrarTomaPedido Date1(0).text
   Else
      TraerTomaPedido Val(fpLongInteger1(0).Value), Date1(0).text
   End If
Case 10 '------- Cancelar
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If modo = "A" Then
       Limpiar 0
       For i = 1 To 2
           Frame1(i).Enabled = False
       Next i
    Else
       TraerTomaPedido Val(fpLongInteger1(0).Value), Date1(0).text
    End If
    modo = "": Gl_Ac_Botones Me, 13, IIf(Date1(0).text < Date, 8, 1), modo
Case 12 '------- Grabar
    '------- Validar fecha
    If (Format(Date1(0).text, "yyyymmdd") < Format(Date, "yyyymmdd")) Then
        MsgBox "Fecha pedido debe ser mayor o igual a la fecha actual (" & _
               Format(Date, "dd/mm/yyyy") & ").", vbCritical, Msgtitulo
        Exit Sub
    End If
    '------- Validar rut
    codpac = fg_DespintaRut(fpText(0).text)
    If Not fg_Check_Rut(codpac) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If Trim(fpayuda(0).Caption) = "" Or Trim(Date1(0).text) = "" Or Val(fpLongInteger1(1).Value) < 1 Then MsgBox "Faltan datos importantes para toma pedido paciente...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    est = False
    '------- Validar si existen datos selecionado en la lista
    With grC
         For i = 1 To .Rows - 1
             .Row = i
             .Col = 1
             If GetItem(.text, 1) = "R" Or GetItem(.text, 1) = "P" Then
                .Col = 2
                If (CInt(GetItem(.text, 1)) > 0) Then est = True: Exit For
             End If
         Next
    End With
    If Not est Then MsgBox "Debe seleccionar a lo menos una receta, para su pedido...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    est = False
    If modo = "A" Then
       vg_db.BeginTrans
       RS.Open "SELECT top_codigo FROM b_tomapedido ORDER BY top_codigo DESC", vg_db, adOpenStatic
       If Not RS.EOF Then
          RS.MoveFirst
          codigo = RS!top_codigo + 1
       Else
          codigo = 1
       End If
       RS.Close: Set RS = Nothing
       '------- Grabar encabezado toma pedido
       vg_db.Execute "INSERT INTO b_tomapedido (top_codigo, top_cencos, top_codreg, top_fecped, top_codpac, top_tipmen, top_codusu)  " & _
                     "VALUES (" & codigo & ", '" & Trim(fpText(4).text) & "', " & codreg & ", '" & Date1(0).text & "', '" & codpac & "', 1, '" & vg_NUsr & "')"
       '------- Grabar detalle toma pedido
       With grC
            indlin = 0
            For i = 1 To .Rows - 1
                .Row = i
                .Col = 1
                If (GetItem(.text, 1) = "R" Or GetItem(.text, 1) = "P") Then
                   .Col = 2
                   If (CInt(GetItem(.text, 1)) > 0) Then
                       indlin = indlin + 1
                       .Col = 1
                       codreg = GetItem(.text, 2)
                       codser = GetItem(.text, 3)
                       codess = GetItem(.text, 4)
                       codrec = GetItem(.text, 5)
                       codmin = GetItem(.text, 8)
                       tiprec = IIf(GetItem(.text, 1) = "P", 0, GetItem(.text, 9))
                       .Col = 2
                       canser = GetItem(.text, 1)
                       cansel = GetItem(.text, 1)
                       .Col = 1
                       vg_db.Execute "INSERT INTO b_tomapedidodet (tpd_codigo, tpd_numlin, tpd_codreg, tpd_codser, tpd_codmin, tpd_estser, tpd_codrec, tpd_tiprec, tpd_cansel, tpd_canser, tpd_caning, tpd_prorec) " & _
                                     "VALUES (" & codigo & ", " & indlin & ", " & codreg & ", " & codser & ", " & codmin & ", " & codess & ", " & codrec & ", " & tiprec & ", " & cansel & ", " & canser & ", 0, '" & IIf(GetItem(.text, 1) = "P", "P", "R") & "')"
                       .Col = 2
                       If (CInt(GetItem(.text, 1)) > 0) Then
                           '------- Garbar detalle pedido receta
                            For j = .Row + 1 To .Rows - 1
                                .Row = j
                                .Col = 1
                                If GetItem(.text, 1) = "T" Then i = .Row: Exit For
                                coding = GetItem(.text, 6)
                                nroite = Val(GetItem(.text, 7))
                                .Col = 3: canser = .text
                                .Col = 4: canbru = .text
                                .Col = 6: pctapr = .text
                                .Col = 7: pctcoc = .text
                                .Col = 8: pctnut = .text
                                vg_db.Execute "INSERT INTO b_tomapedidodetrec (tdr_codigo, tdr_numlin, tdr_codrec, tdr_nroite, tdr_coding, tdr_canpro, tdr_cospro, tdr_pctapr, tdr_pctcoc, tdr_pctnut) " & _
                                              "VALUES (" & codigo & ", " & indlin & ", " & codrec & ", " & IIf(nroite = 0, 1, nroite) & ", '" & coding & "', " & canbru & ", 0, " & pctapr & ", " & pctcoc & ", " & pctnut & ")"
                            Next j
                        End If
                    End If
                End If
            Next
       End With
       vg_db.CommitTrans
       fpLongInteger1(0).Value = codigo
    ElseIf modo = "M" Then
       '------- Grabar detalle toma pedido
       vg_db.BeginTrans
       vg_db.Execute "DELETE b_tomapedidodetrec FROM b_tomapedidodetrec WHERE tdr_codigo = " & Val(fpLongInteger1(0).Value) & ""
       vg_db.Execute "DELETE b_tomapedidodet FROM b_tomapedidodet WHERE tpd_codigo = " & Val(fpLongInteger1(0).Value) & ""
       With grC
            indlin = 0
            For i = 1 To .Rows - 1
                .Row = i
                .Col = 1
                If (GetItem(.text, 1) = "R" Or GetItem(.text, 1) = "P") Then
                   .Col = 2
                   If (CInt(GetItem(.text, 1)) > 0) Then
                      indlin = indlin + 1
                      .Col = 1
                      codreg = GetItem(.text, 2)
                      codser = GetItem(.text, 3)
                      codess = GetItem(.text, 4)
                      codrec = GetItem(.text, 5)
                      codmin = GetItem(.text, 8)
                      tiprec = IIf(GetItem(.text, 1) = "P", 0, GetItem(.text, 9))
                      .Col = 2
                      canser = GetItem(.text, 1)
                      cansel = GetItem(.text, 1)
                      .Col = 1
                      vg_db.Execute "INSERT INTO b_tomapedidodet (tpd_codigo, tpd_numlin, tpd_codreg, tpd_codser, tpd_codmin, tpd_estser, tpd_codrec, tpd_tiprec, tpd_cansel, tpd_canser, tpd_caning, tpd_prorec) " & _
                                    "VALUES (" & Val(fpLongInteger1(0).Value) & ", " & indlin & ", " & codreg & ", " & codser & ", " & codmin & ", " & codess & ", " & codrec & ", " & tiprec & ", " & cansel & ", " & canser & ", 0,'" & IIf(GetItem(.text, 1) = "P", "P", "R") & "')"
                      .Col = 2
                      If (CInt(GetItem(.text, 1)) > 0) Then
                         '------- Garbar detalle pedido receta
                         For j = .Row + 1 To .Rows - 1
                             .Row = j
                             .Col = 1
                             If GetItem(.text, 1) = "T" Then i = .Row: Exit For
                             coding = GetItem(.text, 6)
                             nroite = Val(GetItem(.text, 7))
                             .Col = 3: canser = .text
                             .Col = 4: canbru = .text
                             .Col = 6: pctapr = .text
                             .Col = 7: pctcoc = .text
                             .Col = 8: pctnut = .text
                             vg_db.Execute "INSERT INTO b_tomapedidodetrec (tdr_codigo, tdr_numlin, tdr_codrec, tdr_nroite, tdr_coding, tdr_canpro, tdr_cospro, tdr_pctapr, tdr_pctcoc, tdr_pctnut) " & _
                                           "VALUES (" & Val(fpLongInteger1(0).Value) & ", " & indlin & ", " & codrec & ", " & IIf(nroite = 0, 1, nroite) & ", '" & coding & "', " & canbru & ", 0, " & pctapr & ", " & pctcoc & ", " & pctnut & ")"
                         Next j
                      End If
                   End If
                End If
            Next
       End With
       vg_db.CommitTrans
    End If
    modo = "M": Gl_Ac_Botones Me, 13, 1, modo
Case 15 'Busqueda paciente
   est = True
   vg_codigo = "": vg_nombre = ""
   codpac = fg_DespintaRut(fpText(0).text)
   B_PedPac.LlenarPedidoPaciente codpac
   B_PedPac.Show 1
   est = False
   modo = ""
   If vg_codigo = "" Then Exit Sub
   modo = "M"
   TraerTomaPedido Val(vg_codigo), vg_nombre
   Gl_Ac_Botones Me, 13, IIf(Date1(0).text < Date, 8, 1), modo
   Frame1(1).Enabled = IIf(Date1(0).text < Date, False, True)
   tlbPreview.Buttons(1).Enabled = IIf(Date1(0).text < Date, False, True)
   tlbPreview.Buttons(2).Enabled = IIf(Date1(0).text < Date, False, True)
   Frame1(2).Enabled = True
Case 17 'Impirmir
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_TipMer
Case 20
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

'Private Sub TvwDir_NodeClick(Index As Integer, ByVal Node As MSComctlLib.Node)
''20070808 If Frame1(1).Enabled = False Then Exit Sub
'''    If (intActionForm = gcoActionFormSearch) Then Exit Sub
'    estval = True
'    tlbPreview.Buttons(1).Enabled = (GetItem(Node.Key, 1) = "P")
'    tlbPreview.Buttons(3).Enabled = (GetItem(Node.Key, 1) = "P")
'
'    If (Not blnSetDataCtrl) Then
'        TvwDir(0).Tag = GetItem(Node.Key, 2)
'''        Call GetFillMenus(Date1(0).Value, True, GetItem(node.Key, 2))
'        If (GetItem(Node.Key, 1) = "P") Then
'           If (Not blnValidSelectedPreparacion(Node)) Or (CInt(GetItem(Node.Tag, 1)) = 0) Then vaSpread1.Lock = True
'           '------- Calcular aportes nutricionales
'           CalApoNut (TvwDir(0).SelectedItem.Index)
'           SSTab1.TabVisible(1) = True
'''            Call SetPreviewPedidoSelectedPreparacion(GetItem(node.Key, 3), GetItem(node.Key, 4), GetItem(node.Key, 5))
'        Else
'           SSTab1.TabVisible(1) = False
'        End If
'    End If
'End Sub

Private Sub GetFillMenus(dtFchPedido As Date, Optional blnMsgErrorNoData = True, Optional lngIdServicio = -1)
On Error GoTo GetFillMenus_Err
Dim strSQL As String
Dim strMsg As String
Dim intCurrentMousePointer As Integer

    intCurrentMousePointer = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    DoEvents
    
    
    If (lngIdServicio = -1) Then cboPedidoMenu.Clear
    
    cboPedidoMenuDetalle.Clear
    strMsg = "Menues definidos para el tipo y fecha indicada."
    If (Not blnMsgErrorNoData) Then strMsg = ""
    
    blnSetDataCtrl = True

    strSQL = "SELECT "
    strSQL = strSQL & "Men.Id AS Cod, "
    strSQL = strSQL & "Men.Descripcion "
    strSQL = strSQL & "FROM Menues Men "
    strSQL = strSQL & "WHERE "
    If (intActionForm = gcoActionFormNew) Then
        strSQL = strSQL & "Men.Tipo = " & GetItem(cboPedidoTipo.List(cboPedidoTipo.ListIndex), 2) & " AND "
    End If
    strSQL = strSQL & "CDate('" & dtFchPedido & "') BETWEEN Men.FchInicio AND Men.FchTermino "
    strSQL = strSQL & "AND ( ( CDate('" & dtFchPedido & "') - Men.FchInicio ) Mod Ciclo ) = 0 "
    
    If (Trim(txtPacienteRegimen.Tag) <> "") And (chkPedidoFiltroMenu.Value = ValueTrue) Then
        strSQL = strSQL & "AND Men.IdRegimen = " & Trim(txtPacienteRegimen.Tag) & " "
    End If
    
    If (lngIdServicio > 0) Then
        strSQL = strSQL & "AND EXISTS ( SELECT 1 "
        strSQL = strSQL & "FROM DetalleMenues DetMen "
        strSQL = strSQL & "WHERE DetMen.IdMenu = Men.Id "
        strSQL = strSQL & "AND DetMen.IdServicio = " & lngIdServicio & " ) "
    End If
    
    strSQL = strSQL & "ORDER BY Descripcion "
    
'    If (lngIdServicio = -1) Then
'        Call GetFillComboBox(cboPedidoMenu, strSQL, strMsg)
'    End If
    
'    Call GetFillComboBox(cboPedidoMenuDetalle, strSQL, strMsg)
    
    If (lngIdServicio = -1) Then
        tvwMenu.Nodes.Clear
        tvwMenu.Tag = ""
    End If

    Screen.MousePointer = intCurrentMousePointer
    DoEvents
    
    blnSetDataCtrl = False
    
    Exit Sub
    
GetFillMenus_Err:
    Screen.MousePointer = vbDefault
    DoEvents
    blnSetDataCtrl = False
    
    'Call MsgErrorApp(Me.Name, "GetFillMenus", Err.Number & "-" & Err.Description, strSQL)

End Sub

Private Function blnValidSelectedPreparacion(Row As Long) As Boolean
Dim tmeTimeValid As Date
   blnValidSelectedPreparacion = False
   grC.Row = Row
   grC.Col = 2
   tmeTimeValid = Format(GetItem(grC.text, 2), "Hh:Nn")
   If (Format(Date1(0).text, "yyyymmdd") < Format(Date, "yyyymmdd")) Or ((Format(Date1(0).text, "yyyymmdd") <= Format(Date, "yyyymmdd")) And (tmeTimeValid <= Format(Now, "Hh:Nn"))) Then
      If estval Then Exit Function
         MsgBox "Servicio                 :  " & Trim(GetItem(grC.text, 3)) & Chr(13) & _
                "Hr. Tope Entrega  :  " & Format(GetItem(grC.text, 2), "Hh:Nn") & Chr(13) & Chr(13) & _
                "No puede seleccionar Preparación." & Chr(13) & _
                "Hora tope de entrega para el Servicio ha caducado.", vbCritical ', 'App.TITLE
         Exit Function
   End If
   blnValidSelectedPreparacion = True
End Function

Private Sub SetPreviewPedidoSelectedPreparacion(lngIdGrupoOferta As Long, _
                                                lngIdPreparacion As Long, _
                                                lngIdRegimen As Long, _
                                                Optional blnAddCantidad = True)
Dim intIndex As Integer
Dim strTmp As String
Dim intCantidad As Integer
    With grC
    
        ' Válido para PEDIDOS EXTRAS
        .Col = 1
         If (GetItem(.text, 4) = -99999999) Then
'        If (GetItem(cboPedidoTipo.List(cboPedidoTipo.ListIndex), 2) = gcoMenuTypeExtra) Then
            .Col = 2
            strTmp = Trim(GetItem(.text, 1))
            .Col = 0
            If (InStr(.text, "(") > 0) Then
                strTmp = Trim(Mid(.text, 1, InStr(.text, "(") - 1))
            Else
                strTmp = Trim(.text)
            End If
            .Col = 2
            intCantidad = CInt(Trim(GetItem(.text, 1))) + IIf(blnAddCantidad, 1, -1)
            .Col = 0
            .text = strTmp & IIf(intCantidad > 0, "  (" & intCantidad & ")", "")
'            .SelectedItem.Bold = (intCantidad > 0)
            .Col = 2
            .text = IIf(intCantidad > 0, intCantidad, "0") & ";" & _
                    Trim(GetItem(.text, 2)) & ";" & _
                    Trim(GetItem(.text, 3)) & ";"
            Exit Sub
'        End If
        End If
        
        ' Válido para PEDIDOS NORMALES
        For intIndex = 1 To .Rows - 1
            .Row = intIndex
            .Col = 1
            If (GetItem(.text, 1) = "R") Then
                If (GetItem(.text, 4) = lngIdGrupoOferta) And _
                   (GetItem(.text, 5) = lngIdPreparacion) And _
                   (GetItem(.text, 2) = lngIdRegimen) Then
                    .Col = 2
                    If (Trim(GetItem(.text, 1)) = "0") Then
                       .Col = 0
                       .text = .text & "  (1)"
                       .Col = 2
                       .text = "1;" & Trim(GetItem(.text, 2)) & ";" & _
                                      Trim(GetItem(.text, 3)) & ";"
                       For i = .Row + 1 To .Rows - 1
                           .Row = i
                           .Col = 1
                           If GetItem(.text, 1) = "T" Then Exit For
                           .Cell(flexcpForeColor, i, 3, i, 3) = &HFF0000
                       Next i
                    Else
                        .Col = 0
                        .text = Trim(Mid(.text, 1, InStr(.text, "(") - 1))
                        .Col = 2
                        .text = "0;" & Trim(GetItem(.text, 2)) & ";" & _
                                       Trim(GetItem(.text, 3)) & ";"
                       For i = .Row + 1 To .Rows - 1
                           .Row = i
                           .Col = 1
                           If GetItem(.text, 1) = "T" Then Exit For
                           .Cell(flexcpForeColor, i, 3, i, 3) = &H80000012
                       Next i
                    
                    End If
                Else
                    If (GetItem(.text, 4) = lngIdGrupoOferta) And _
                       (GetItem(.text, 2) = lngIdRegimen) Then
                        .Col = 0
                        If (InStr(.text, "(") > 0) Then
                            .text = Trim(Mid(.text, 1, InStr(.text, "(") - 1))
                        End If
                        .Col = 2
                        .text = "0;" & Trim(GetItem(.text, 2)) & ";" & _
                                       Trim(GetItem(.text, 3)) & ";"
                       For i = .Row + 1 To .Rows - 1
                           .Row = i
                           .Col = 1
                           If GetItem(.text, 1) = "T" Then Exit For
                           .Cell(flexcpForeColor, i, 3, i, 3) = &H80000012
                       Next i
                    
                    End If
                End If
            End If
        Next
    End With
End Sub

Sub CalcularAporte(codpro As String, canbru As Double, pctnut As Double, Row As Long)
Dim i As Long, j As Long, candie As Double, indtot As Long, canser As Double, canbru1 As Double, cannet As Double, totapo As Double
Dim sql1 As String, sql2 As String
With grC
     '------- Sumar aporte
     For i = Row To .Rows - 1
         .Row = i
         .Col = 0
         If Trim(.text) = "TOTALES" Then indtot = i: Exit For
     Next i
    If canbru < 0 Then Exit Sub
    sql1 = IIf(vg_tipbase = "1", " (((((" & pctnut & "/100)*(c.pnu_canapo*(" & canbru & "/" & 100 & ")))/a.ing_facnut))) AS  candiet ", " (((((convert(float," & pctnut & ")/100)*(c.pnu_canapo*(convert(float," & canbru & ")/" & 100 & ")))/a.ing_facnut))) AS  candiet ")
    sql2 = IIf(vg_tipbase = "1", " (((" & pctnut & "/100)*" & canbru & ")) AS cangrverneto ", " (((convert(float," & pctnut & ")/100)*convert(float," & canbru & "))) AS cangrverneto ")
    RS.Open "SELECT DISTINCT b.nut_nombre, b.nut_codigo, a.ing_indgrv, a.ing_indpav, b.nut_secnro, " & _
            "" & sql1 & ", " & _
            "" & sql2 & " " & _
            "FROM  b_ingrediente a, a_nutriente b, b_productonut c " & _
            "WHERE a.ing_codigo = c.pnu_codpro " & _
            "AND   c.pnu_codapo = b.nut_codigo " & _
            "AND   a.ing_codigo = '" & codpro & "' " & _
            "ORDER BY b.nut_secnro", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
    i = 10
    Do While Not RS.EOF
       For i = 10 To .Cols - 1
           .Row = 0
           .Col = i
           If Val(GetItem(.text, 1)) = RS!nut_codigo Then
              .Row = Row
              .text = Format(CCur(RS!candiet), fg_Pict(6, 2))
              Exit For
           End If
       Next i
       RS.MoveNext
    Loop
    RS.Close: Set RS = Nothing
    '------- Mover zero a la fila totales del item selecionado
    For j = 10 To .Cols - 1
        .Row = indtot
        .Col = j
        .text = Format(0, fg_Pict(6, 2))
        If j = 10 Then
           .Col = 4: .text = Format(0, fg_Pict(6, 2))
           .Col = 3: .text = Format(0, fg_Pict(6, 2))
           .Col = 9: .text = Format(0, fg_Pict(6, 2))
        End If
    Next j
    '------- Mover aportes totales
    For i = Row To 1 Step -1
        .Row = i
        .Col = 1
        If GetItem(.text, 1) = "R" Or GetItem(.text, 1) = "P" Then Row = i: Exit For
    Next i
    
    For i = Row + 1 To (indtot - 1)
         .Row = i
         .Col = 3: canser = .text
         .Col = 4: canbru1 = .text
         .Col = 9: cannet = .text
         .Row = indtot
         .Col = 3: .text = Format(.text + canser, fg_Pict(6, 2))
         .Col = 4: .text = Format(.text + canbru1, fg_Pict(6, 2))
         .Col = 9: .text = Format(.text + cannet, fg_Pict(6, 2))
        For j = 10 To (.Cols - 1)
            .Row = i
            .Col = j
            totapo = .text
            .Row = indtot
            .Col = j
            .text = Format(.text + totapo, fg_Pict(6, 2))
        Next j
    Next i
End With
End Sub

Private Sub Opgrilla_Click(Index As Integer)
If grC.Rows < 1 Then Exit Sub
Dim auxser As Long, auxreg As Long, indrow As Long, nivel As Long, Fila As Long, sql1 As String, sql2 As String, sql3 As String
Select Case Index
Case 10 'Inserta receta
    vg_codigo = "": vg_nombre = "": vg_tiprec = -2
    indrow = grC.Row
    '------- Validar receta 5 etapa
    B_Receta.Show 1, Me
    If Trim(vg_codigo) = "" Or Trim(vg_nombre) = "" Or vg_tiprec < -1 Then Exit Sub
    '------- Validar si existe receta
    grC.Col = 1
    auxreg = Val(GetItem(grC.text, 2))
    auxser = Val(GetItem(grC.text, 3))
    For i = 1 To grC.Rows - 1
        grC.Row = i
        grC.Col = 1
        If GetItem(grC.text, 1) = "R" And _
           Val(GetItem(grC.text, 2)) = auxreg And _
           Val(GetItem(grC.text, 3)) = auxser Then
           If Val(vg_codigo) = GetItem(grC.text, 5) And GetItem(grC.text, 1) = "R" Then MsgBox "Receta existe en los adicionales", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        End If
    Next i
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo=" & auxser & "", vg_db, adOpenStatic
'    grC.RemoveItem indrow + 1
'    AddNodeGroup indrow + 1, grC.RowOutlineLevel(indrow) + 1
'    RS.Close: Set RS = Nothing
    
    nivel = 2
    CrearNiveles nivel, indrow, vg_nombre
    grC.Col = 1
    indrow = indrow + 1: nivel = 2
    CrearNiveles nivel, indrow, " TOTALES"
    '------- Insertar receta grilla
    sql1 = IIf(vg_tipbase = "1", " CDATE( FORMAT( Now(), 'dd/mm/yyyy' ) ", "  FORMAT( Now(), 'yyyymmdd') ")
    RS1.Open "SELECT DISTINCT a.rec_codigo, c.ing_codigo, c.ing_nombre, d.unm_nomcor, b.red_nroite, " & _
             "b.red_canpro, b.red_cospro, b.red_pctapr, b.red_pctcoc, b.red_pctnut, " & _
             "(((b.red_pctapr/100)*b.red_canpro)*(b.red_pctcoc/100)) AS canser, ((b.red_pctnut/100)*(b.red_canpro)) AS cannet " & _
             "FROM b_receta a, b_recetadet b, b_ingrediente c, a_unidadmed d, a_servicio e " & _
             "WHERE a.rec_codigo = b.red_codigo " & _
             "AND   b.red_codpro = c.ing_codigo " & _
             "AND   c.ing_unimed = d.unm_codigo " & _
             "AND   a.rec_codigo = " & Val(vg_codigo) & " " & _
             "AND   b.red_tiprec = " & Val(vg_tiprec) & " AND ((b.red_tiprec <> 0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) " & _
             "AND   e.ser_codigo = " & auxser & " " & _
             "AND   " & sql1 & " & ' ' & e.ser_horent ) >= Now() " & _
             "ORDER BY b.red_nroite DESC", vg_db, adOpenStatic
    Fila = indrow
    fil = indrow
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          nivel = 2
          CrearNiveles nivel, indrow, Trim(RS1!ing_nombre)
          RS1.MoveNext: fil = fil + 1
       Loop
       RS1.MoveFirst
       est = True
       grC.Row = Fila
       grC.Col = 0
       grC.text = Trim(vg_nombre) & " (1)"
       grC.Col = 1
       grC.text = "R" & ";" & _
                  auxreg & ";" & _
                  auxser & ";" & _
                  -99999999 & ";" & _
                  Val(vg_codigo) & ";" & _
                  0 & ";" & _
                  0 & ";" & _
                  0 & ";" & _
                  vg_tiprec
       grC.Col = 2
       grC.text = "1" & _
                  ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
                  Trim(RS!ser_nombre)

       grC.Cell(flexcpBackColor, grC.Row, 0, grC.Row, 1) = &H80FF80
       Fila = Fila + 1
       Do While Not RS1.EOF
          grC.Row = fil
          grC.Col = 1
          grC.text = "I" & ";" & _
                     auxreg & ";" & _
                     auxser & ";" & _
                     -99999999 & ";" & _
                     Val(vg_codigo) & ";" & _
                     RS1!ing_codigo & ";" & _
                     RS1!red_nroite & ";" & _
                     0 & ";" & _
                     vg_tiprec
          grC.Col = 2
          grC.text = "1" & _
                     ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
                     Trim(RS!ser_nombre)

          grC.Col = 3: grC.text = Format(RS1!canser, fg_Pict(6, 2))
          grC.Col = 4: grC.text = Format(RS1!red_canpro, fg_Pict(6, 2))
          grC.Col = 5: grC.text = Trim(RS1!unm_nomcor)
          grC.Col = 6: grC.text = Format(RS1!red_pctapr, fg_Pict(6, 2))
          grC.Col = 7: grC.text = Format(RS1!red_pctcoc, fg_Pict(6, 2))
          grC.Col = 8: grC.text = Format(RS1!red_pctnut, fg_Pict(6, 2))
          grC.Col = 9: grC.text = Format(RS1!cannet, fg_Pict(6, 2))
          For i = 10 To grC.Cols - 1
              grC.Col = i
              grC.text = Format(0, fg_Pict(6, 2))
          Next i
          grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H80000012  '&HFF&
          grC.Cell(flexcpForeColor, grC.Row, 3, grC.Row, 3) = &HFF0000
          grC.Cell(flexcpBackColor, grC.Row, 0, grC.Row, 1) = &H80FF80
          grC.Cell(flexcpBackColor, grC.Row, 1, grC.Row, grC.Cols - 1) = &HC0FFFF
          RS1.MoveNext: Fila = Fila + 1: fil = fil - 1
       Loop
    End If
    '------- Mover zero columnas totales
    grC.Row = Fila
    grC.Col = 1: grC.text = "T"
    grC.Col = 3: grC.text = Format(0, fg_Pict(6, 2))
    grC.Col = 4: grC.text = Format(0, fg_Pict(6, 2))
    grC.Col = 9: grC.text = Format(0, fg_Pict(6, 2))
    For i = 10 To grC.Cols - 1
        grC.Col = i
        grC.text = Format(0, fg_Pict(6, 2))
    Next i
    grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H80000012
    grC.Cell(flexcpFontBold, grC.Row, 0, grC.Row, grC.Cols - 1) = True
    grC.Cell(flexcpBackColor, grC.Row, 1, grC.Row, grC.Cols - 1) = &HFF&
    grC.Cell(flexcpBackColor, grC.Row, 0, grC.Row, 1) = &H80FF80

    RS1.Close: Set RS1 = Nothing
    RS.Close: Set RS = Nothing
    est = False
Case 20 'Isertar producto
    vg_nombre = "": vg_codigo = ""
    indrow = grC.Row
    B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "ProVig"
    B_TabEst.Show 1
    If vg_codigo = "" Then Exit Sub
    '------- Validar si existe receta
    grC.Col = 1
    auxreg = Val(GetItem(grC.text, 2))
    auxser = Val(GetItem(grC.text, 3))
    For i = 1 To grC.Rows - 1
        grC.Row = i
        grC.Col = 1
        If (GetItem(grC.text, 1) = "P") And _
           Val(GetItem(grC.text, 2)) = auxreg And _
           Val(GetItem(grC.text, 3)) = auxser Then
           If Val(vg_codigo) = GetItem(grC.text, 5) And GetItem(grC.text, 1) = "P" Then MsgBox "Producto existe en los adicionales", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        End If
    Next i
    RS.Open "SELECT * FROM a_servicio WHERE ser_codigo=" & auxser & "", vg_db, adOpenStatic
    nivel = 2
    CrearNiveles nivel, indrow, vg_nombre
    grC.Col = 1
    indrow = indrow + 1: nivel = 2
    CrearNiveles nivel, indrow, " TOTALES"
    '------- Insertar producto grilla
    RS1.Open "SELECT DISTINCT a.pro_codigo, c.ing_codigo, c.ing_nombre, d.unm_nomcor, " & _
             "c.ing_pctapr, c.ing_pctcoc, c.ing_pctnut, a.pro_facsto " & _
             "FROM b_productos a, b_productosing b, b_ingrediente c, a_unidadmed d, a_servicio e " & _
             "WHERE a.pro_codigo = b.pri_codpro " & _
             "AND   b.pri_coding = c.ing_codigo " & _
             "AND   c.ing_unimed = d.unm_codigo " & _
             "AND   a.pro_codigo = '" & vg_codigo & "' " & _
             "AND   e.ser_codigo = " & auxser & " " & _
             "AND   CDATE( FORMAT( Now(), 'dd/mm/yyyy' ) & ' ' & e.ser_horent ) >= Now()", vg_db, adOpenStatic
    Fila = indrow
    fil = indrow
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          nivel = 2
          CrearNiveles nivel, indrow, Trim(RS1!ing_nombre)
          RS1.MoveNext: fil = fil + 1
       Loop
       RS1.MoveFirst
       est = True
       grC.Row = Fila
       grC.Col = 0
       grC.text = Trim(vg_nombre) & " (1)"
       grC.Col = 1
       grC.text = "P" & ";" & _
                  auxreg & ";" & _
                  auxser & ";" & _
                  -99999999 & ";" & _
                  Val(vg_codigo) & ";" & _
                  0 & ";" & _
                  0 & ";" & _
                  0 & ";" & _
                  0
       grC.Col = 2
       grC.text = "1" & _
                  ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
                  Trim(RS!ser_nombre)

       grC.Cell(flexcpBackColor, grC.Row, 0, grC.Row, 1) = &H80FF80
       Fila = Fila + 1
       Do While Not RS1.EOF
          grC.Row = fil
          grC.Col = 1
          grC.text = "I" & ";" & _
                     auxreg & ";" & _
                     auxser & ";" & _
                     -99999999 & ";" & _
                     Val(vg_codigo) & ";" & _
                     RS1!ing_codigo & ";" & _
                     0 & ";" & _
                     0 & ";" & _
                     0
          grC.Col = 2
          grC.text = "1" & _
                     ";" & Format(RS!ser_horent, "Hh:Nn") & ";" & _
                     Trim(RS!ser_nombre)

          grC.Col = 3: grC.text = Format(RS1!pro_facsto, fg_Pict(6, 2))
          grC.Col = 4: grC.text = Format(RS1!pro_facsto, fg_Pict(6, 2))
          grC.Col = 5: grC.text = Trim(RS1!unm_nomcor)
          grC.Col = 6: grC.text = Format(RS1!ing_pctapr, fg_Pict(6, 2))
          grC.Col = 7: grC.text = Format(RS1!ing_pctcoc, fg_Pict(6, 2))
          grC.Col = 8: grC.text = Format(RS1!ing_pctnut, fg_Pict(6, 2))
          grC.Col = 9: grC.text = Format(RS1!pro_facsto, fg_Pict(6, 2))
          For i = 10 To grC.Cols - 1
              grC.Col = i
              grC.text = Format(0, fg_Pict(6, 2))
          Next i
          grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H80000012  '&HFF&
          grC.Cell(flexcpForeColor, grC.Row, 3, grC.Row, 3) = &HFF0000
          grC.Cell(flexcpBackColor, grC.Row, 0, grC.Row, 1) = &H80FF80
          grC.Cell(flexcpBackColor, grC.Row, 1, grC.Row, grC.Cols - 1) = &HC0FFFF
          RS1.MoveNext: Fila = Fila + 1: fil = fil - 1
       Loop
    End If
    '------- Mover zero columnas totales
    grC.Row = Fila
    grC.Col = 1: grC.text = "T"
    grC.Col = 3: grC.text = Format(0, fg_Pict(6, 2))
    grC.Col = 4: grC.text = Format(0, fg_Pict(6, 2))
    grC.Col = 9: grC.text = Format(0, fg_Pict(6, 2))
    For i = 10 To grC.Cols - 1
        grC.Col = i
        grC.text = Format(0, fg_Pict(6, 2))
    Next i
    grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H80000012
    grC.Cell(flexcpFontBold, grC.Row, 0, grC.Row, grC.Cols - 1) = True
    grC.Cell(flexcpBackColor, grC.Row, 1, grC.Row, grC.Cols - 1) = &HFF&
    grC.Cell(flexcpBackColor, grC.Row, 0, grC.Row, 1) = &H80FF80

    RS1.Close: Set RS1 = Nothing
    RS.Close: Set RS = Nothing
    est = False
Case 30 'Borrar linea
   Dim nomtot As String
   est = True
   grC.Col = 1
   indrow = grC.Row
   If (GetItem(grC.text, 4) = -99999999 And (GetItem(grC.text, 1) = "R" Or GetItem(grC.text, 1) = "P")) Then
      'Borrar de grilla
      With grC
           For i = indrow To .Rows - 1
               .Row = i
               .Col = 1
               If GetItem(.text, 1) = "T" Then indfin = i: Exit For
           Next i
           
           For i = indfin To indrow Step -1
               .RemoveItem i
           Next i
      End With
   End If
   est = False
End Select
End Sub

Sub CrearNiveles(nivel As Long, indfil As Long, Nombre As String)
    est = True
    Dim nd As New VSFlexNode
    grC.Row = indfil
    Set nd = grC.GetNode
    ' add relative as requested by user
    ' (could be a child or a sibling)
    nd.AddNode nivel, Nombre
    nd.Expanded = False
'    grC.RowOutlineLevel(indfil) = 1
'    grC.IsCollapsed(indfil) = flexOutlineCollapsed
    est = False
End Sub

' Private Sub Form_Load()
'
'        fg.Rows = 1
'
'        fg.Cols = 1
'
'        fg.FixedCols = 0
'
'        fg.ExtendLastCol = True
'
'        fg.OutlineBar = flexOutlineBarSimpleLeaf
'
'        AddNodeGroup 1, 0
'
'    End Sub
'
'    Private Sub fg_BeforeCollapse(ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
'
'        If Row < 0 Then Cancel = True: Exit Sub
''a = fg.Rows
'        If State = flexOutlineCollapsed Then Exit Sub
'
'        If fg.TextMatrix(Row + 1, 0) <> "Dummy" Then Exit Sub
'
'        fg.RemoveItem Row + 1
'
'        AddNodeGroup Row + 1, fg.RowOutlineLevel(Row) + 1
'
'    End Sub
    
'    Sub AddNodeGroup(r&, level&)
'
'        Dim i%
'
'        For i = 1 To 5
'
'            grC.AddItem "Row " & grC.Rows, r
'
'            grC.AddItem "Dummy", r + 1
'
'            grC.IsSubtotal(r) = True
'
'            grC.RowOutlineLevel(r) = level
'
'            grC.IsCollapsed(r) = flexOutlineCollapsed
'
'            r = r + 2
'
'        Next
'
'    End Sub
