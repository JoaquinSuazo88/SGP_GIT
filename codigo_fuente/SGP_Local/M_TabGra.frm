VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_TabGra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Gramaje"
   ClientHeight    =   9690
   ClientLeft      =   1785
   ClientTop       =   960
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4695
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   4920
      Width           =   10635
      Begin VB.Frame Frame5 
         Height          =   435
         Left            =   1080
         TabIndex        =   22
         Top             =   4080
         Width           =   4245
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   23
            Top             =   135
            Width           =   4150
         End
      End
      Begin VB.Frame Frame4 
         Height          =   435
         Left            =   240
         TabIndex        =   20
         Top             =   4080
         Width           =   790
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   21
            Top             =   135
            Width           =   690
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3765
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   10365
         _Version        =   393216
         _ExtentX        =   18283
         _ExtentY        =   6641
         _StockProps     =   64
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
         MaxCols         =   10
         SpreadDesigner  =   "M_TabGra.frx":0000
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Left            =   5520
         TabIndex        =   13
         Top             =   4320
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4800
      Index           =   0
      Left            =   100
      TabIndex        =   4
      Top             =   10
      Width           =   10575
      Begin VB.Frame Frame6 
         Height          =   3375
         Left            =   240
         TabIndex        =   24
         Top             =   1320
         Width           =   7815
         Begin VB.CheckBox Check2 
            Caption         =   "Activar Filtros Opcionales"
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   240
            Width           =   2175
         End
         Begin VB.Frame Frame2 
            Caption         =   "Filtros Opcionales"
            Enabled         =   0   'False
            Height          =   2625
            Left            =   240
            TabIndex        =   25
            Top             =   600
            Width           =   7365
            Begin VB.ComboBox Combo2 
               Height          =   315
               Index           =   0
               Left            =   1800
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   600
               Width           =   3480
            End
            Begin EditLib.fpLongInteger fpLongInteger1 
               Height          =   315
               Index           =   2
               Left            =   1800
               TabIndex        =   27
               Top             =   240
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
               NoSpecialKeys   =   2
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
            Begin EditLib.fpDateTime fpDateTime1 
               Height          =   315
               Left            =   1800
               TabIndex        =   28
               Top             =   1020
               Width           =   945
               _Version        =   196608
               _ExtentX        =   1667
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
               Text            =   "09/2010"
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
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   8
               Left            =   1800
               TabIndex        =   45
               Top             =   2145
               Width           =   1215
            End
            Begin VB.Label lblSOMBRA 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   1845
               TabIndex        =   44
               Top             =   2190
               Width           =   1230
            End
            Begin VB.Label Label3 
               Caption         =   "Tipo Minuta"
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
               Index           =   22
               Left            =   120
               TabIndex        =   43
               Top             =   2160
               Width           =   1815
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   3
               Left            =   1800
               TabIndex        =   37
               Top             =   1755
               Width           =   5175
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   3
               Left            =   1140
               Picture         =   "M_TabGra.frx":1ACB
               Top             =   1650
               Width           =   480
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   2
               Left            =   1140
               Picture         =   "M_TabGra.frx":1DD5
               Top             =   1260
               Width           =   480
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Plato"
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
               TabIndex        =   36
               Top             =   1800
               Width           =   885
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C. Dietetica"
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
               Left            =   105
               TabIndex        =   35
               Top             =   1410
               Width           =   1020
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1800
               TabIndex        =   34
               Top             =   1365
               Width           =   5175
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   3090
               TabIndex        =   33
               Top             =   240
               Width           =   3975
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Servicio"
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
               TabIndex        =   32
               Top             =   240
               Width           =   705
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   5
               Left            =   2625
               Picture         =   "M_TabGra.frx":20DF
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label3 
               Caption         =   "Zona"
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
               Index           =   0
               Left            =   120
               TabIndex        =   31
               Top             =   600
               Width           =   975
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
               Index           =   2
               Left            =   120
               TabIndex        =   30
               Top             =   1020
               Width           =   1230
            End
            Begin VB.Label sombra 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   5
               Left            =   1845
               TabIndex        =   29
               Top             =   705
               Width           =   3495
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   7
               Left            =   3120
               TabIndex        =   40
               Top             =   255
               Width           =   4005
            End
            Begin VB.Label lblSOMBRA 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   1815
               TabIndex        =   39
               Top             =   1800
               Width           =   5205
            End
            Begin VB.Label lblSOMBRA 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   6
               Left            =   1815
               TabIndex        =   38
               Top             =   1410
               Width           =   5205
            End
         End
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   2640
            TabIndex        =   42
            Top             =   240
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
            Enabled         =   0   'False
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2250
         Left            =   8160
         TabIndex        =   17
         Top             =   120
         Width           =   2295
         Begin MSComctlLib.TreeView TvwZon 
            Height          =   1575
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   2778
            _Version        =   393217
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
            Enabled         =   0   'False
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Activa copiado a otras zonas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            TabIndex        =   18
            Top             =   120
            Width           =   2055
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   225
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1440
         TabIndex        =   2
         Top             =   930
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
         Left            =   1440
         TabIndex        =   1
         Top             =   570
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Regimen"
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
         Left            =   90
         TabIndex        =   15
         Top             =   630
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2400
         Picture         =   "M_TabGra.frx":23E9
         Top             =   480
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   2850
         TabIndex        =   14
         Top             =   570
         Width           =   5175
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2850
         TabIndex        =   11
         Top             =   930
         Width           =   5175
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2850
         TabIndex        =   9
         Top             =   225
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   2400
         Picture         =   "M_TabGra.frx":26F3
         Top             =   840
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   2400
         Picture         =   "M_TabGra.frx":29FD
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ing. Receta"
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
         Left            =   90
         TabIndex        =   8
         Top             =   990
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sub-Segmento"
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
         Left            =   90
         TabIndex        =   7
         Top             =   300
         Width           =   1245
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   99
         Left            =   2865
         TabIndex        =   10
         Top             =   255
         Width           =   5205
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   98
         Left            =   2865
         TabIndex        =   12
         Top             =   945
         Width           =   5205
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   300
         Index           =   5
         Left            =   2865
         TabIndex        =   16
         Top             =   585
         Width           =   5205
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   9690
      Left            =   10815
      TabIndex        =   6
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   17092
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VB.Menu MenuDetalle 
      Caption         =   ""
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu OpGrilla 
         Caption         =   "Copiar"
         Index           =   0
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Pegar"
         Enabled         =   0   'False
         Index           =   1
      End
   End
End
Attribute VB_Name = "M_TabGra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim NomFor As String, Msgtitulo As String
Dim filtippla As Long
Dim iblockcol As Long, iblockrow As Long, iblockcol2 As Long, iblockrow2 As Long, irow As Long
Dim coding As String, noming As String, Est As Boolean, TmpCopiaGramaje As Double
Dim rootNode As Node, nd As Node
Dim filcatdie As Long

Dim FilIni As Variant, FilFin As Variant, Colini As Variant, ColFin As Variant

Private Sub Check1_Click()
If Check1.Value = 0 Then
   TvwZon.Enabled = False
Else
   TvwZon.Enabled = True
End If
End Sub

Private Sub Check2_Click()
    If Check2.Value = 1 Then
        Frame2.Enabled = True
        Toolbar2.Enabled = True
        DataLoad
    Else
        Frame2.Enabled = False
        Toolbar2.Enabled = False
        DataLoad
    End If
End Sub




Private Sub Combo2_Click(Index As Integer)
    DataLoad
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Dim RS As New ADODB.Recordset
fg_carga ""
Me.HelpContextID = vg_OpcM
Me.Height = 10170
Me.Width = 11445
Msgtitulo = "Tabla Gramaje"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Toolbar2.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): btnX.Visible = True: btnX.ToolTipText = "Confirmar ": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): btnX.Visible = True: btnX.ToolTipText = "": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 3, 1) = "0", False, True)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): btnX.Visible = True: btnX.ToolTipText = "Deshacer"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): btnX.Visible = True: btnX.ToolTipText = "Historico Tabla Gramaje"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): btnX.Visible = True: btnX.ToolTipText = "Imprimir": btnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "0", False, True)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): btnX.Visible = True: btnX.ToolTipText = "Copiar Tabla Gramaje"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
'Set btnX = Toolbar2.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar2.Buttons.Add(, "Proceso", , tbrDefault, "Proceso"): btnX.Visible = True: btnX.ToolTipText = "Proceso"

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &H80000018
fpayuda(2).Caption = "Todos": fpayuda(3).Caption = "Todos"
iayuda = 0: Est = True
RS.Open "SELECT par_valor FROM a_param WHERE par_codigo='catdefecto'", vg_db, adOpenStatic
If Not RS.EOF Then filcatdie = RS!par_valor: fpayuda(2).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
RS.Close: Set RS = Nothing
'Cargar zona
TvwZon.Nodes.Clear
Set RS = vg_db.Execute("sgpadm_s_zona 6, 0,''")
Do While Not RS.EOF
   Set rootNode = TvwZon.Nodes.Add(, , "H" & RS!zon_codigo, Trim(RS!Zon_nombre))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing

'------> Llenado combo Zonas
Set RS = vg_db.Execute("SELECT * FROM a_zona")
Combo2(0).Clear
Do While Not RS.EOF
   Combo2(0).AddItem Trim(RS!Zon_nombre) & Space(150) & "(" & Trim(RS!zon_codigo) & ")"
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing

vg_fecha = Format(fpDateTime1.text, "yyyymm")

Est = False
fg_descarga
End Sub

Private Sub fpDateTime1_Change()
If IsDate(fpDateTime1.text) = False Then Exit Sub
vg_fecha = Format(fpDateTime1.text, "yyyymm")
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 0
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & " AND sub_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
    End If
'    RS.Open "SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: fpayuda(0).Caption = "": fpLongInteger1(1).Value = "": fpayuda(4).Caption = "": fpText(0).text = "": fpayuda(1).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing
    fpLongInteger1(1).Value = "": fpayuda(4).Caption = ""
    fpText(0).text = "": fpayuda(1).Caption = ""
Case 1
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    End If
'    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    If Est Then Exit Sub
    DataLoad
Case 2
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(6).Caption = "": Exit Sub
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & " and ser_indppr='" & vg_Indppr & "'")
    Else
      Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & "")
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "":  Exit Sub
    fpayuda(6).Caption = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    DataLoad
    
End Select
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText_Change(Index As Integer)
Dim RS As New ADODB.Recordset
Select Case Index
Case 0
    RS.Open "SELECT DISTINCT a.ing_codigo, a.ing_nombre FROM b_ingrediente a, b_receta b, b_recetadet c WHERE b.rec_codigo = c.red_codigo AND c.red_codpro = a.ing_codigo AND (b.rec_catdie = " & filcatdie & " OR " & filcatdie & " = 0) AND (b.rec_tippla = " & filtippla & " OR " & filtippla & " = 0) AND a.ing_codigo = '" & Trim(fpText(0).text) & "' AND (a.ing_Indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' <> '1')", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": vaSpread1.MaxRows = 0: Exit Sub
    fpayuda(1).Caption = Trim(RS!ing_nombre)
    RS.Close: Set RS = Nothing
   DataLoad

End Select
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    vg_left = fpayuda(0).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpText(0).text = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(1).SetFocus
Case 1
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = "": vg_filtippla = filtippla: vg_filcatdie = filcatdie
'   B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente", "Ingrec"
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "INgrediente", "AgregarIng"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    fpText(0).SetFocus
Case 2
    vg_left = fpayuda(2).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
    B_ArbEst.Show 1
    If vg_codigo = "" Then Exit Sub
    filcatdie = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre: vg_nombre = ""
    DataLoad
Case 3
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(3).Left + 3000
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
    B_ArbEst.Show 1
    If Trim(vg_codigo) = "" Then Exit Sub
    filtippla = Val(vg_codigo)
    fpayuda(3).Caption = vg_nombre: vg_nombre = ""
    DataLoad
Case 4
    vg_left = fpayuda(4).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    Est = False
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
    fpText(0).SetFocus
 Case 5
    vg_left = fpayuda(6).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser", filcatdie
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(6).Caption = vg_nombre
    fpDateTime1.SetFocus
End Select
End Sub

Private Sub MoverDetalle()
Dim RS As New ADODB.Recordset
vaSpread1.MaxRows = 0
If Val(fpLongInteger1(0).Value) = 0 Or LimpiaDato(Trim(fpText(0).text)) = "" Or Val(fpLongInteger1(1).Value) = 0 Then Exit Sub
fg_carga ""
Dim estgr As Boolean
Dim aAp As String
estgr = False
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
Set RS = vg_db.Execute("sgpadm_s_tablagramaje 2, " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", 0, '" & fpText(0).text & "', " & filcatdie & ", " & filtippla & "")
If RS.EOF Then vaSpread1.Visible = True: RS.Close: Set RS = Nothing: fg_descarga: Exit Sub 'MsgBox "Ingrediente no existe en recetarios", vbExclamation + vbOKOnly, "reemplazar ingrediente en receta": Exit Sub
Bar1.Visible = True: Bar1.Value = 0
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   Bar1.Value = Val((vaSpread1.MaxRows / RS!nReg) * 100)
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1: vaSpread1.CellType = CellTypeCheckBox:  vaSpread1.TypeCheckText = " ": vaSpread1.TypeHAlign = TypeHAlignCenter: vaSpread1.TypeCheckCenter = True: vaSpread1.text = "0" ' checked
   vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!rec_codigo
   vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!rec_nombre)
   vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!zon_codigo
   vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!Zon_nombre)
   If RS!opgramo <> "R" And Toolbar1.Buttons(4).Enabled = False Then estgr = True
   vaSpread1.Col = 8: vaSpread1.CellType = CellTypeEdit: vaSpread1.text = IIf(RS!opgramo = "R", "", Trim(RS!red_codpro))
   vaSpread1.Col = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = IIf(RS!opgramo = "R", "", Trim(RS!noming))
   vaSpread1.Col = 10:
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = vg_RDCa   'TypeFloatDecimalPlaces = vg_RDCa 'vg_dbndecimal
   vaSpread1.TypeFloatMin = "-99999999"
   vaSpread1.TypeFloatMax = "99999999"
   vaSpread1.TypeFloatMoney = False
   vaSpread1.TypeFloatSeparator = True
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.TypeFloatCurrencyChar = Asc("$")
   vaSpread1.TypeFloatDecimalChar = Asc(".")
   vaSpread1.TypeFloatSepChar = Asc(",")
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.text = Format(RS!red_canpro, fg_Pict(6, vg_RDCa)) '2))
   vaSpread1.ForeColor = IIf(RS!opgramo = "R", &HFF&, &HFF0000)
   RS.MoveNext
Loop
OpGrilla(1).Enabled = False
If estgr = True Then Toolbar1.Buttons(4).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Bar1.Visible = False
vaSpread1.Visible = True
'vaSpread1.SetFocus
RS.Close: Set RS = Nothing: fg_descarga
End Sub

Private Sub MoverDetalleOpcional()
Dim RS As New ADODB.Recordset
vaSpread1.MaxRows = 0
If Val(fpLongInteger1(0).Value) = 0 Or LimpiaDato(Trim(fpText(0).text)) = "" Or Val(fpLongInteger1(1).Value) = 0 Then Exit Sub
fg_carga ""
Dim estgr As Boolean
Dim aAp As String
estgr = False
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vg_fecha = ""
If Check2.Value = 1 Then
    vg_fecha = Mid(fpDateTime1.text, 4, 4) & Mid(fpDateTime1.text, 1, 2)
Else
    vg_fecha = ""
End If


Set RS = vg_db.Execute("sgpadm_s_tablagramajeminuta " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & fpText(0).text & "', " & filcatdie & ", " & filtippla & ", " & Val(vg_fecha) & ", " & Val(fpLongInteger1(2)) & ", " & ExraeCodCombo(Combo2(0)) & " ")
If RS.EOF Then vaSpread1.Visible = True: RS.Close: Set RS = Nothing: fg_descarga: Exit Sub 'MsgBox "Ingrediente no existe en recetarios", vbExclamation + vbOKOnly, "reemplazar ingrediente en receta": Exit Sub
Bar1.Visible = True: Bar1.Value = 0
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   Bar1.Value = Val((vaSpread1.MaxRows / RS!nReg) * 100)
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1: vaSpread1.CellType = CellTypeCheckBox:  vaSpread1.TypeCheckText = " ": vaSpread1.TypeHAlign = TypeHAlignCenter: vaSpread1.TypeCheckCenter = True: vaSpread1.text = "0" ' checked
   vaSpread1.Col = 4: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!rec_codigo
   vaSpread1.Col = 5: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!rec_nombre)
   vaSpread1.Col = 6: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = RS!zon_codigo
   vaSpread1.Col = 7: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = Trim(RS!Zon_nombre)
   If RS!opgramo <> "R" And Toolbar1.Buttons(4).Enabled = False Then estgr = True
   vaSpread1.Col = 8: vaSpread1.CellType = CellTypeEdit: vaSpread1.text = IIf(RS!opgramo = "R", "", Trim(RS!red_codpro))
   vaSpread1.Col = 9: vaSpread1.CellType = CellTypeStaticText: vaSpread1.text = IIf(RS!opgramo = "R", "", Trim(RS!noming))
   vaSpread1.Col = 10:
   vaSpread1.CellType = CellTypeCurrency
   vaSpread1.TypeCurrencyDecPlaces = vg_RDCa 'TypeFloatDecimalPlaces = vg_dbndecimal
   vaSpread1.TypeFloatMin = "-99999999"
   vaSpread1.TypeFloatMax = "99999999"
   vaSpread1.TypeFloatMoney = False
   vaSpread1.TypeFloatSeparator = True
   vaSpread1.TypeHAlign = TypeHAlignRight
   vaSpread1.TypeFloatCurrencyChar = Asc("$")
   vaSpread1.TypeFloatDecimalChar = Asc(".")
   vaSpread1.TypeFloatSepChar = Asc(",")
   vaSpread1.TypeCurrencyShowSymbol = False
   vaSpread1.text = Format(RS!red_canpro, fg_Pict(6, 2))
   vaSpread1.ForeColor = IIf(RS!opgramo = "R", &HFF&, &HFF0000)
   RS.MoveNext
Loop
OpGrilla(1).Enabled = False
If estgr = True Then Toolbar1.Buttons(4).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Bar1.Visible = False
vaSpread1.Visible = True
'vaSpread1.SetFocus
RS.Close: Set RS = Nothing: fg_descarga
End Sub

Private Sub Text1_Change(Index As Integer)
Dim Col As Integer
Col = 0
Select Case Index
    Case 1, 2
        Col = IIf(Index = 1, 4, 5)
        vaSpread1.Visible = False
        If Trim(Text1(Index).text) <> "" Then
           For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            vaSpread1.Col = Col
            indactivo = UCase(Trim(vaSpread1.Value)) Like "*" & UCase(Text1(Index).text) & "*"
            vaSpread1.Col = Col
            If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
               If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
            Else
               If vaSpread1.RowHidden = False Then vaSpread1.RowHidden = True
            End If
        Next i
            vaSpread1.SetActiveCell Col, 1
        End If
    '    vaSpread1_Click Index, 0
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(Text1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
        If Trim(Text1(Index).text) = "" Then
           For i = 1 To vaSpread1.MaxRows
               vaSpread1.Row = i
               If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           Next
           vaSpread1.SetActiveCell Col, vaSpread1.SearchCol(Col, 0, vaSpread1.MaxRows, Trim(Text1(Index).text), SearchFlagsGreaterOrEqual)
           vaSpread1.SetActiveCell Col, 1
        End If
        vaSpread1.Visible = True
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim RS As New ADODB.Recordset
Dim cantrg  As Double
Dim isel As Long, codrec As Long, codzon As Long
Dim coding As String
On Error GoTo Man_Error
Select Case Button.Index
Case 2, 4
    If vaSpread1.MaxRows < 1 Then Exit Sub
    RS.Open "SELECT ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: MsgBox "No existe ingrediente recetas", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    RS.Close: Set RS = Nothing
    isel = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = 1
        If vaSpread1.text = "1" Then isel = 1: Exit For
    Next i
    If isel = 0 Then MsgBox "Seleccione Uno o Más Recetas " & IIf(Button.Index = 2, "a Reemplazar", "a Borrar"), vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Esta Seguro ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    fg_carga ""
    Bar1.Visible = True: Bar1.Value = 0
    
    For i = 1 To vaSpread1.MaxRows
        Bar1.Value = Val((i / vaSpread1.MaxRows) * 100)
        vaSpread1.Row = i
        vaSpread1.Col = 1
        If vaSpread1.text = "1" Then
           Me.Refresh
           DoEvents
           vaSpread1.Col = 4: codrec = 0: codrec = Val(vaSpread1.text)
           vaSpread1.Col = 6: codzon = 0: codzon = vaSpread1.text
           vaSpread1.Col = 8: coding = "": coding = Trim(vaSpread1.text)
           vaSpread1.Col = 10: cantrg = 0: cantrg = vaSpread1.text
           If Trim(coding) = "" And vaSpread1.ForeColor = &HFF0000 Then
              RS.Open "SELECT DISTINCT a.rec_codigo FROM b_receta a, b_recetadet b WHERE a.rec_codigo = b.red_codigo AND a.rec_codigo = " & codrec & " AND b.red_codpro = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND b.red_canpro = " & cantrg & "", vg_db, adOpenStatic
              If RS.EOF Then coding = LimpiaDato(Trim(fpText(0).text))
              RS.Close: Set RS = Nothing
           End If
           If Trim(coding) = "" Then
'              vg_db.Execute "DELETE b_tablagramaje FROM b_tablagramaje WHERE tgr_subseg = " & Val(fpLongInteger1(0).Value) & " AND tgr_codreg = " & Val(fpLongInteger1(1).Value) & " AND tgr_codrec = " & codrec & " AND tgr_coding = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND tgr_codzon = " & codzon & ""
              vg_db.Execute "sgpadm_d_tablagramaje 'E', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & codrec & ", '" & LimpiaDato(Trim(fpText(0).text)) & "', " & codzon & ", '', 0"
           Else
              RS.Open "SELECT tgr_subseg FROM b_tablagramaje WHERE tgr_subseg = " & Val(fpLongInteger1(0).Value) & " AND tgr_codreg = " & Val(fpLongInteger1(1).Value) & " AND tgr_codrec = " & codrec & " AND tgr_coding = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND tgr_codzon = " & codzon & "", vg_db, adOpenStatic
              If Button.Index = 2 Then
                 If RS.EOF Then
'                    vg_db.Execute "INSERT INTO b_tablagramaje VALUES (" & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & codrec & ", '" & LimpiaDato(Trim(fpText(0).text)) & "', " & codzon & ", '" & coding & "', " & cantrg & ")"
                    vg_db.Execute "sgpadm_iu_tablagramaje 'A', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & codrec & ", '" & LimpiaDato(Trim(fpText(0).text)) & "', " & codzon & ", '" & coding & "', " & cantrg & ""
                 Else
'                    vg_db.Execute "UPDATE b_tablagramaje SET tgr_codins = '" & coding & "', tgr_cantgr = " & cantrg & " WHERE tgr_subseg = " & Val(fpLongInteger1(0).Value) & " AND tgr_codreg = " & Val(fpLongInteger1(1).Value) & " AND tgr_codrec = " & codrec & " AND tgr_coding = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND tgr_codzon = " & codzon & ""
                    vg_db.Execute "sgpadm_iu_tablagramaje 'M', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & codrec & ", '" & LimpiaDato(Trim(fpText(0).text)) & "', " & codzon & ", '" & coding & "', " & cantrg & ""
                 End If
              Else
'                 If Not RS.EOF Then vg_db.Execute "DELETE b_tablagramaje FROM b_tablagramaje WHERE tgr_subseg = " & Val(fpLongInteger1(0).Value) & " AND tgr_codreg = " & Val(fpLongInteger1(1).Value) & " AND tgr_codrec = " & codrec & " AND tgr_coding = '" & LimpiaDato(Trim(fpText(0).text)) & "' AND tgr_codzon = " & codzon & ""
                 If Not RS.EOF Then vg_db.Execute "sgpadm_d_tablagramaje 'E', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & codrec & ", '" & LimpiaDato(Trim(fpText(0).text)) & "', " & codzon & ", '', 0"
              End If
              RS.Close: Set RS = Nothing
           End If
        End If
    Next i
    
    Bar1.Visible = False
    fg_descarga
    MsgBox "Proceso finalizo sin problema", vbInformation + vbOKOnly, Msgtitulo
    DataLoad
    indsel = 0
'    vaSpread1.MaxRows = 0
Case 6
    filcatdie = 0: filtippla = 0
    fpayuda(2).Caption = "Todos": fpayuda(3).Caption = "Todos"
    RS.Open "SELECT par_valor FROM a_param WHERE par_codigo='catdefecto'", vg_db, adOpenStatic
    If Not RS.EOF Then filcatdie = RS!par_valor: fpayuda(2).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
    RS.Close: Set RS = Nothing
    DataLoad
Case 8
    'Historico tabla gramaje
    vg_codigo = "": vg_codregimen = 0
    B_HistPm.LlenarHistPlan "Histórico Tabla Gramaje", Val(fpLongInteger1(0).Value), 1, 2
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    Est = True
    fpLongInteger1(1).Value = ""
    Est = False
    fpLongInteger1(1).Value = vg_codregimen
    fpText(0).text = vg_codigo
Case 10
    I_TabGra.Show 1
Case 12
    M_CpTabGra.Show 1, Me
Case 14 'Salir
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
    DataLoad
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
Dim i As Long
Select Case BlockCol
Case 1
    vaSpread1.Col = 1
    For i = BlockRow To BlockRow2
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    Next
Case 8
    iblockcol = BlockCol
    iblockrow = BlockRow
    iblockcol2 = BlockCol2
    iblockrow2 = BlockRow2
End Select
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
iblockcol = Col
iblockrow = Row
iblockcol2 = Col
iblockrow2 = Row
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Or vaSpread1.MaxRows < 1 Then Exit Sub
Dim codrec As Long, coding As String, noming As String, i As Long, j As Long, codzon As Long
Select Case Col
Case 8
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = "": coding = "": noming = ""
'    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente", "Gen"
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
'    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    vaSpread1.Row = Row
    vaSpread1.Col = 8: vaSpread1.text = Trim(vg_codigo): coding = Trim(vg_codigo)
    vaSpread1.Col = 9: vaSpread1.text = Trim(vg_nombre): noming = Trim(vg_nombre)
    vaSpread1.Col = 1: vaSpread1.text = "1"
    vaSpread1.Col = 4: codrec = vaSpread1.text
    If Check1.Value = 0 Then Exit Sub
'    If MsgBox("Desea copiar ingrediente a las siguientes zona ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 4
        If vaSpread1.text = codrec And i <> Row Then
           vaSpread1.Col = 6: codzon = vaSpread1.text
           For j = 1 To TvwZon.Nodes.count
               If TvwZon.Nodes.Item(j).Checked = True And codzon = Mid(TvwZon.Nodes.Item(j).Key, 2, Len(TvwZon.Nodes.Item(j).Key)) Then
                  vaSpread1.Col = 1
                  vaSpread1.text = "1"
                  vaSpread1.Col = 8
                  vaSpread1.text = coding
                  vaSpread1.Col = 9
                  vaSpread1.text = noming
                  Exit For
               End If
           Next j
        End If
    Next i
End Select
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Row < 1 Then Exit Sub
Dim RS As New ADODB.Recordset
Dim codrec As Long, coding As String, noming As String, canbru As Double, i As Long, j As Long, codzon As Long
vaSpread1.Row = Row
'If ChangeMade = False And Col = 8 Then

If ChangeMade = True Then
   Select Case Col
   Case 8
       vaSpread1.Col = 8
       RS.Open "SELECT DISTINCT ing_codigo, ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "' AND (ing_Indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' <> '1')", vg_db, adOpenStatic
       vaSpread1.Col = 9
       If Not RS.EOF Then
          vaSpread1.text = Trim(RS!ing_nombre)
          coding = Trim(RS!ing_codigo)
          noming = Trim(RS!ing_nombre)
       Else
          vaSpread1.text = "": vaSpread1.Col = 8: vaSpread1.text = ""
       End If
       RS.Close: Set RS = Nothing
       vaSpread1.Col = 1: vaSpread1.text = "1"
       vaSpread1.Col = 4: codrec = vaSpread1.text
       If Check1.Value = 0 Then Exit Sub
'       If MsgBox("Desea copiar ingrediente a las siguientes zona ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = 4
           If vaSpread1.text = codrec And i <> Row Then
              vaSpread1.Col = 6: codzon = vaSpread1.text
              For j = 1 To TvwZon.Nodes.count
                  If TvwZon.Nodes.Item(j).Checked = True And codzon = Mid(TvwZon.Nodes.Item(j).Key, 2, Len(TvwZon.Nodes.Item(j).Key)) Then
                     vaSpread1.Col = 1
                     vaSpread1.text = "1"
                     vaSpread1.Col = 8
                     vaSpread1.text = coding
                     vaSpread1.Col = 9
                     vaSpread1.text = noming
                     Exit For
                  End If
              Next j
           End If
       Next i
   Case 10
'mod 20091120       vaSpread1.Col = 9
'mod 20091120       If Trim(vaSpread1.Text) = "" Then Exit Sub
       vaSpread1.Col = 10
       vaSpread1.ForeColor = &HFF0000
       canbru = 0: canbru = vaSpread1.text
       vaSpread1.Col = 1
       vaSpread1.text = "1"
       vaSpread1.Col = 4: codrec = vaSpread1.text
       If Check1.Value = 0 Then Exit Sub
'       If MsgBox("Desea copiar cantidad bruta a las siguientes zona ?", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
       For i = 1 To vaSpread1.MaxRows
           vaSpread1.Row = i
           vaSpread1.Col = 4
           If vaSpread1.text = codrec And i <> Row Then
              vaSpread1.Col = 6: codzon = vaSpread1.text
              For j = 1 To TvwZon.Nodes.count
                  If TvwZon.Nodes.Item(j).Checked = True And codzon = Mid(TvwZon.Nodes.Item(j).Key, 2, Len(TvwZon.Nodes.Item(j).Key)) Then
                     vaSpread1.Col = 1
                     vaSpread1.text = "1"
                     vaSpread1.Col = 10
                     vaSpread1.text = canbru
                     vaSpread1.ForeColor = &HFF0000
                     Exit For
                  End If
             Next j
           End If
       Next i
   End Select
End If
End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case Button
Case 2
    If vaSpread1.MaxRows < 1 Or vaSpread1.ActiveCol <> 8 And vaSpread1.ActiveCol <> 10 Then Exit Sub
    PopupMenu MenuDetalle
End Select
End Sub

Private Sub Opgrilla_Click(Index As Integer)
Select Case Index
Case 0
    If vaSpread1.ActiveCol = 10 Then
        OpGrilla(1).Enabled = True
        vaSpread1.Col = vaSpread1.ActiveCol
        vaSpread1.Row = vaSpread1.ActiveRow
        TmpCopiaGramaje = Trim(vaSpread1.text)

    Else
        OpGrilla(1).Enabled = True
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 8: coding = Trim(vaSpread1.text)
        vaSpread1.Col = 9: noming = Trim(vaSpread1.text)
    End If
    
    
Case 1
vaSpread1.GetSelection 1, Colini, FilIni, ColFin, FilFin

Dim i As Long
    If vaSpread1.ActiveCol = 10 Then
        For i = FilIni To FilFin
            vaSpread1.Row = i
            vaSpread1.Col = 10: vaSpread1.text = TmpCopiaGramaje
            vaSpread1.ForeColor = &HFF0000
            vaSpread1.Col = 1: vaSpread1.text = "1"
        Next i
    Else
        For i = FilIni To FilFin
            vaSpread1.Row = i
            vaSpread1.Col = 8: vaSpread1.text = coding
            vaSpread1.Col = 9: vaSpread1.text = noming
            vaSpread1.Col = 1: vaSpread1.text = "1"
        Next i
    End If
End Select
End Sub

Private Sub DataLoad()
    Dim RS As New ADODB.Recordset
    If Check2.Value = 0 Then MoverDetalle Else MoverDetalleOpcional
    If Check2.Value = 1 Then
        'Set RS = vg_db.Execute("sgpadm_s_traetipominuta " & fpLongInteger1(0).LongValue & ", " & fpLongInteger1(1).LongValue & ", '" & fpText(0).text & "' , " & filcatdie & ", " & filtippla & ", " & fpLongInteger1(2).LongValue & " , " & ExtraeFecha(fpDateTime1) & " ")
        Set RS = vg_db.Execute("sgpadm_s_traetipominuta " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", '" & fpText(0).text & "', " & filcatdie & ", " & filtippla & ", " & ExtraeFecha(fpDateTime1) & ", " & Val(fpLongInteger1(2)) & ", " & ExraeCodCombo(Combo2(0)) & " ")
        If (RS.EOF = True And RS.BOF = True) Then RS.Close: Set RS = Nothing: Exit Sub
        fpayuda(8).Caption = IIf(RS.Fields(0) = "1", "Real", "Propuesta")
        RS.Close: Set RS = Nothing
    End If
End Sub

