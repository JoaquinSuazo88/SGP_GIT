VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CpTabGra 
   Caption         =   "Copiar Tabla de Gramaje Destino"
   ClientHeight    =   5775
   ClientLeft      =   2745
   ClientTop       =   2865
   ClientWidth     =   9495
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Destino"
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
      Height          =   3045
      Index           =   1
      Left            =   120
      TabIndex        =   24
      Top             =   2640
      Width           =   8730
      Begin VB.Frame Frame3 
         Height          =   1335
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Width           =   8535
         Begin VB.OptionButton Option1 
            Caption         =   "Centro de Costo x Nivel"
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
            Index           =   5
            Left            =   3000
            TabIndex        =   43
            Top             =   120
            Width           =   2415
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Centro de Costo"
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
            Index           =   3
            Left            =   240
            TabIndex        =   10
            Top             =   120
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
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
            Height          =   255
            Index           =   2
            Left            =   6480
            TabIndex        =   11
            Top             =   120
            Width           =   1575
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   0
            Left            =   1785
            TabIndex        =   12
            Top             =   465
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
            Index           =   4
            Left            =   1785
            TabIndex        =   14
            Top             =   825
            Width           =   915
            _Version        =   196608
            _ExtentX        =   1614
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
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   3
            AutoAdvance     =   0   'False
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
            NegFormat       =   1
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
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   4
            Left            =   2655
            Picture         =   "M_CpTabGra.frx":0000
            Top             =   720
            Width           =   480
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
            Index           =   1
            Left            =   240
            TabIndex        =   40
            Top             =   870
            Width           =   1245
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   3105
            TabIndex        =   15
            Top             =   825
            Width           =   5175
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3105
            TabIndex        =   13
            Top             =   480
            Width           =   5175
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   2655
            Picture         =   "M_CpTabGra.frx":030A
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Costo"
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
            Left            =   240
            TabIndex        =   38
            Top             =   525
            Width           =   1380
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   3120
            TabIndex        =   39
            Top             =   510
            Width           =   5205
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   3120
            TabIndex        =   41
            Top             =   870
            Width           =   5205
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   5
         Left            =   1885
         TabIndex        =   16
         Top             =   1845
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   6
         Left            =   1885
         TabIndex        =   18
         Top             =   2175
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   2520
         TabIndex        =   31
         Top             =   2760
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   6
         Left            =   2765
         Picture         =   "M_CpTabGra.frx":0614
         Top             =   2100
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Zona"
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
         Left            =   440
         TabIndex        =   27
         Top             =   2235
         Width           =   900
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   3205
         TabIndex        =   19
         Top             =   2175
         Width           =   5175
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3205
         TabIndex        =   17
         Top             =   1845
         Width           =   5175
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
         Left            =   440
         TabIndex        =   25
         Top             =   1890
         Width           =   750
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   5
         Left            =   2765
         Picture         =   "M_CpTabGra.frx":091E
         Top             =   1740
         Width           =   480
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3240
         TabIndex        =   26
         Top             =   1890
         Width           =   5205
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   3240
         TabIndex        =   28
         Top             =   2220
         Width           =   5205
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos Origen"
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
      Height          =   2325
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   8730
      Begin VB.Frame Frame2 
         Height          =   1215
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   8535
         Begin VB.OptionButton Option1 
            Caption         =   "Centro de Costo x Nivel"
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
            Index           =   4
            Left            =   2880
            TabIndex        =   42
            Top             =   120
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Centro de Costo"
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
            Left            =   240
            TabIndex        =   0
            Top             =   120
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton Option1 
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
            Height          =   255
            Index           =   1
            Left            =   6240
            TabIndex        =   1
            Top             =   120
            Width           =   2175
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   1
            Left            =   1785
            TabIndex        =   2
            Top             =   465
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
            Left            =   1785
            TabIndex        =   4
            Top             =   825
            Width           =   915
            _Version        =   196608
            _ExtentX        =   1614
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
            AlignTextH      =   2
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   3
            AutoAdvance     =   0   'False
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
            NegFormat       =   1
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
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3105
            TabIndex        =   5
            Top             =   825
            Width           =   5175
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
            Index           =   3
            Left            =   240
            TabIndex        =   35
            Top             =   870
            Width           =   1245
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   2655
            Picture         =   "M_CpTabGra.frx":0C28
            Top             =   720
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   9
            Left            =   3105
            TabIndex        =   3
            Top             =   480
            Width           =   5175
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2655
            Picture         =   "M_CpTabGra.frx":0F32
            Top             =   360
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Centro de Costo"
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
            Left            =   240
            TabIndex        =   33
            Top             =   525
            Width           =   1380
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   10
            Left            =   3120
            TabIndex        =   34
            Top             =   510
            Width           =   5205
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3120
            TabIndex        =   36
            Top             =   870
            Width           =   5205
         End
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   2
         Left            =   1885
         TabIndex        =   6
         Top             =   1605
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   3
         Left            =   1885
         TabIndex        =   8
         Top             =   1935
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
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         NegFormat       =   1
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
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   2765
         Picture         =   "M_CpTabGra.frx":123C
         Top             =   1860
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Zona"
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
         Left            =   440
         TabIndex        =   29
         Top             =   1995
         Width           =   900
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3205
         TabIndex        =   9
         Top             =   1935
         Width           =   5175
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   2765
         Picture         =   "M_CpTabGra.frx":1546
         Top             =   1500
         Width           =   480
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
         Index           =   0
         Left            =   440
         TabIndex        =   22
         Top             =   1650
         Width           =   750
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   2
         Left            =   3205
         TabIndex        =   7
         Top             =   1605
         Width           =   5175
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3240
         TabIndex        =   23
         Top             =   1650
         Width           =   5205
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   5
         Left            =   3240
         TabIndex        =   30
         Top             =   1980
         Width           =   5205
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   5775
      Left            =   8955
      TabIndex        =   20
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   10186
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "M_CpTabGra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j As Integer
Dim MsgTitulo As String

Private Sub Form_Load()

On Error GoTo Man_Error

MsgTitulo = "Copiar Tabla de Gramaje"
fg_centra Me

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

'-------> Activar Ceco
fpText(1).Enabled = True
Image1(0).Enabled = True
'-------> desactivar sub-segmento
fpLongInteger1(1).Enabled = False
Image1(1).Enabled = False
Label2(5).Visible = False
fpLongInteger1(3).Visible = False
Image1(3).Visible = False
fpayuda(3).Visible = False
lblSOMBRA(5).Visible = False
Label2(4).Visible = False
fpLongInteger1(6).Visible = False
Image1(6).Visible = False
fpayuda(6).Visible = False
lblSOMBRA(4).Visible = False

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

Case 1
    
    RS.Open "SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing

Case 2
    
    If Val(fpLongInteger1(Index).Value) < 1 Then fpayuda(Index).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing

Case 3
    
    Set RS = vg_db.Execute("sgpadm_s_zona 9, " & Val(fpLongInteger1(Index).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!Zon_nombre)
    RS.Close: Set RS = Nothing

Case 4
    
    RS.Open "SELECT * FROM a_subsegmento WHERE sub_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing

Case 5
    
    If Val(fpLongInteger1(Index).Value) < 1 Then fpayuda(Index).Caption = "": Exit Sub
    RS.Open "SELECT * FROM a_regimen WHERE reg_codigo=" & Val(fpLongInteger1(Index).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing

Case 6
    
    Set RS = vg_db.Execute("sgpadm_s_zona 9, " & Val(fpLongInteger1(Index).Value) & ", ''")
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(Index).Caption = "": Exit Sub
    fpayuda(Index).Caption = Trim(RS!Zon_nombre)
    RS.Close: Set RS = Nothing
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
   
End Sub

Private Sub fpText_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index
    
    Case 0
       
       Sql = Trim(LimpiaDato(fpText(0).text))
       Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
       If RS.EOF Then
            fpayuda(0).Caption = ""
            RS.Close
            Set RS = Nothing
            Exit Sub
        End If
        fpayuda(0).Caption = Trim(RS!Cli_nombre)
    
    Case 1
       
       Sql = Trim(LimpiaDato(fpText(1).text))
       Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
       If RS.EOF Then
            fpayuda(9).Caption = ""
            RS.Close
            Set RS = Nothing
            Exit Sub
        End If
        fpayuda(9).Caption = Trim(RS!Cli_nombre)
    
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
   
End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
   
End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    
    Case 0
        
        vg_left = fpayuda(9).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
        B_TabEst.Show 1
        Me.Refresh
        Screen.MousePointer = 0
        If vg_codigo = "" Then Exit Sub
        fpText(1).text = vg_codigo: fpayuda(9).Caption = vg_nombre
        fpLongInteger1(2).SetFocus
    
    Case 7
        
        vg_left = fpayuda(9).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
        B_TabEst.Show 1
        Me.Refresh
        Screen.MousePointer = 0
        If vg_codigo = "" Then Exit Sub
        fpText(0).text = vg_codigo: fpayuda(0).Caption = vg_nombre
        fpLongInteger1(5).SetFocus
    
    Case 1
        
        vg_left = fpayuda(Index).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(Index).Value = Val(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        fpLongInteger1(2).SetFocus
    
    Case 2
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(Index).Value = Val(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        fpLongInteger1(Index).SetFocus
    
    Case 3
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Zon"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(Index).Value = Val(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
    
    Case 4
        
        vg_left = fpayuda(Index).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(Index).Value = Val(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        fpLongInteger1(5).SetFocus
    
    Case 5
        
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(Index).Value = Val(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
        fpLongInteger1(Index).SetFocus
    
    Case 6
    
        vg_left = fpayuda(1).Left + 2300
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_zona", "zon_", "Zona", "Zon"
        B_TabEst.Show 1
        Me.Refresh
        If vg_codigo = "" Then Exit Sub
        fpLongInteger1(Index).Value = Val(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
   
End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    
    Case 0
        
        '-------> Activar Ceco origen
        fpText(1).text = ""
        fpayuda(9).Caption = ""
        fpText(1).Enabled = True
        Image1(0).Enabled = True
        '-------> desactivar sub-segmento origen
        fpLongInteger1(1).text = ""
        fpayuda(1).Caption = ""
        fpLongInteger1(1).Enabled = False
        Image1(1).Enabled = False
        
        fpLongInteger1(2).text = ""
        fpayuda(2).Caption = ""
        
        Label2(5).Visible = False
        fpLongInteger1(3).text = ""
        fpLongInteger1(3).Visible = False
        Image1(3).Visible = False
        fpayuda(3).Visible = False
        fpayuda(3).Caption = ""
        lblSOMBRA(5).Visible = False
        Option1(3).Value = True
        
    Case 1
        
        '-------> Activar Ceco
        fpText(1).text = ""
        fpayuda(9).Caption = ""
        fpText(1).Enabled = False
        Image1(0).Enabled = False
        '-------> desactivar sub-segmento
        fpLongInteger1(1).Enabled = True
        Image1(1).Enabled = True
    
        fpLongInteger1(2).text = ""
        fpayuda(2).Caption = ""
    
        Label2(5).Visible = True
        fpLongInteger1(3).Visible = True
        fpLongInteger1(3).text = ""
        Image1(3).Visible = True
        fpayuda(3).Visible = True
        fpayuda(3).Caption = ""
        lblSOMBRA(5).Visible = True
        Option1(2).Value = True
        
    Case 3
        
        '-------> Activar Ceco origen
        fpText(0).text = ""
        fpText(0).Enabled = True
        fpayuda(0).Caption = ""
        Image1(7).Enabled = True
        '-------> desactivar sub-segmento origen
        fpLongInteger1(4).text = ""
        fpayuda(4).Caption = ""
        fpLongInteger1(4).Enabled = False
        Image1(4).Enabled = False
        
        fpLongInteger1(5).text = ""
        fpayuda(5).Caption = ""
        
        Label2(4).Visible = False
        fpLongInteger1(6).Visible = False
        fpLongInteger1(6).text = ""
        Image1(6).Visible = False
        fpayuda(6).Visible = False
        fpayuda(6).Caption = ""
        lblSOMBRA(4).Visible = False
        
    Case 2
    
        '-------> Activar Ceco
        fpText(0).text = ""
        fpayuda(0).Caption = ""
        fpText(0).Enabled = False
        Image1(7).Enabled = False
        '-------> desactivar sub-segmento
        fpLongInteger1(4).Enabled = True
        Image1(4).Enabled = True
    
        fpLongInteger1(5).text = ""
        fpayuda(5).Caption = ""
        
        Label2(4).Visible = True
        fpLongInteger1(6).Visible = True
        fpLongInteger1(6).text = ""
        Image1(6).Visible = True
        fpayuda(6).Visible = True
        fpayuda(6).Caption = ""
        lblSOMBRA(4).Visible = True
    
    Case 4
    
        '-------> Activar Ceco origen
        fpText(1).text = ""
        fpayuda(9).Caption = ""
        fpText(1).Enabled = True
        Image1(0).Enabled = True
        '-------> desactivar sub-segmento origen
        fpLongInteger1(1).text = ""
        fpayuda(1).Caption = ""
        fpLongInteger1(1).Enabled = False
        Image1(1).Enabled = False
        
        fpLongInteger1(2).text = ""
        fpayuda(2).Caption = ""
        
        Label2(5).Visible = False
        fpLongInteger1(3).text = ""
        fpLongInteger1(3).Visible = False
        Image1(3).Visible = False
        fpayuda(3).Visible = False
        fpayuda(3).Caption = ""
        lblSOMBRA(5).Visible = False
        Option1(5).Value = True
        
    Case 5
    
        '-------> Activar Ceco origen
        fpText(0).text = ""
        fpText(0).Enabled = True
        fpayuda(0).Caption = ""
        Image1(7).Enabled = True
        '-------> desactivar sub-segmento origen
        fpLongInteger1(4).text = ""
        fpayuda(4).Caption = ""
        fpLongInteger1(4).Enabled = False
        Image1(4).Enabled = False
        
        fpLongInteger1(5).text = ""
        fpayuda(5).Caption = ""
        
        Label2(4).Visible = False
        fpLongInteger1(6).Visible = False
        fpLongInteger1(6).text = ""
        Image1(6).Visible = False
        fpayuda(6).Visible = False
        fpayuda(6).Caption = ""
        lblSOMBRA(4).Visible = False

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
   
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim RS      As New ADODB.Recordset
Dim SubsegO As Long
Dim SubsegD As Long
Dim CodregO As Long
Dim CodregD As Long
Dim CodzonO As Long
Dim CodzonD As Long
Dim CecoO   As String
Dim CecoD   As String
Dim Subs, Reg, rec, ing, zon, ins, gr As Integer
Dim CantReg As Integer

Select Case Button.Index
    
    Case 1
        
        If Option1(0).Value = True And Option1(3).Value = False Then
        
           MsgBox "la option 1 debe ser igual origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        ElseIf Option1(1).Value = True And Option1(2).Value = False Then
        
           MsgBox "la option 3 debe ser igual origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        ElseIf Option1(4).Value = True And Option1(5).Value = False Then
        
            MsgBox "la option 2 debe ser igual origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
                   
        End If
        
        If Option1(0).Value = True And fpayuda(9).Caption = "" Then
           
           MsgBox "Debe ingresar ceco origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        ElseIf Option1(1).Value = True And fpayuda(1).Caption = "" Then
           
           MsgBox "Debe ingresar sub-segmento origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        ElseIf Option1(4).Value = True And fpayuda(9).Caption = "" Then
           
           MsgBox "Debe ingresar ceco origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        
        End If
        
        If Option1(3).Value = True And fpayuda(0).Caption = "" Then
           
           MsgBox "Debe ingresar ceco destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        ElseIf Option1(2).Value = True And fpayuda(4).Caption = "" Then
           
           MsgBox "Debe ingresar sub-segmento destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        ElseIf Option1(5).Value = True And fpayuda(0).Caption = "" Then
           
           MsgBox "Debe ingresar ceco destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        
        End If
        
        '-------> Validar regimen origen
        If fpayuda(2).Caption = "" Then
           
           MsgBox "Debe ingresar regimen origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        '-------> Validar regimen destino
        If fpayuda(5).Caption = "" Then
           
           MsgBox "Debe ingresar regimen destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        '-------> Validar zona origen
        If Option1(1).Value = True And fpayuda(3).Caption = "" Then
           
           MsgBox "Debe ingresar zona origen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        '-------> Validar zona destino
        If Option1(2).Value = True And fpayuda(6).Caption = "" Then
           
           MsgBox "Debe ingresar zona destino", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        
        End If
        
        CecoO = ""
        SubsegO = 0
        CodregO = 0
        CodzonO = 0
        
        CecoD = ""
        SubsegD = 0
        CodregD = 0
        CodzonD = 0
        
        CecoO = LimpiaDato(Trim(fpText(1).text))
        If Val(fpLongInteger1(1).text) > 0 Then SubsegO = fpLongInteger1(1).Value
        If Val(fpLongInteger1(2).Value) > 0 Then CodregO = fpLongInteger1(2).Value
        If Val(fpLongInteger1(3).Value) > 0 Then CodzonO = fpLongInteger1(3).Value
        
        CecoD = LimpiaDato(Trim(fpText(0).text))
        If Val(fpLongInteger1(4).Value) > 0 Then SubsegD = fpLongInteger1(4).Value
        If Val(fpLongInteger1(5).Value) > 0 Then CodregD = fpLongInteger1(5).Value
        If Val(fpLongInteger1(6).Value) > 0 Then CodzonD = fpLongInteger1(6).Value
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        '------->Validar Existencia Datos Origen
        If Option1(0).Value = True Then
           
           RS.Open "select tgc_ceco from b_tablagramajececo  WITH ( NOLOCK ) where tgc_ceco='" & CecoO & "' and tgc_codreg= " & CodregO & "", vg_db, adOpenStatic
           
           If RS.EOF Then
              
              MsgBox "No Existe Datos de Origen", vbExclamation + vbOKOnly, MsgTitulo
              RS.Close
              Set RS = Nothing
              Exit Sub
           
           End If
           RS.Close
           Set RS = Nothing
        
        ElseIf Option1(1).Value = True Then
            
            RS.Open "select tgr_subseg from b_tablagramaje  WITH ( NOLOCK ) where tgr_subseg=" & SubsegO & " and tgr_codreg= " & CodregO & " AND tgr_codzon=" & CodzonO & "", vg_db, adOpenStatic
            
            If RS.EOF Then
               
               MsgBox "No Existe Datos de Origen", vbExclamation + vbOKOnly, MsgTitulo
               RS.Close
               Set RS = Nothing
               Exit Sub
            
            End If
            
            RS.Close
            Set RS = Nothing
        
        ElseIf Option1(4).Value = True Then
        
        
            RS.Open "select top 1  IdCeco from b_tablagramajececo_nivel  WITH ( NOLOCK ) where IdCeco='" & CecoO & "'", vg_db, adOpenStatic
            
            If RS.EOF Then
               
               MsgBox "No Existe Datos de Origen", vbExclamation + vbOKOnly, MsgTitulo
               RS.Close
               Set RS = Nothing
               Exit Sub
            
            End If
            
            RS.Close
            Set RS = Nothing
        
        
        End If
        
        '-------> Validar Existencia de Datos de Destino
        If Option1(3).Value = True Then
           
           RS.Open "select distinct tgc_ceco from b_tablagramajececo WITH ( NOLOCK ) where tgc_ceco='" & CecoD & "' and tgc_codreg= " & CodregD & "", vg_db, adOpenStatic
           
           If Not RS.EOF Then
              
              If MsgBox("Existe información en Tabla Gramaje destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                 
                 RS.Close
                 Set RS = Nothing
                 Exit Sub
              
              End If
              
           End If
           RS.Close
           Set RS = Nothing
        
        ElseIf Option1(2).Value = True Then
           
           RS.Open "select distinct tgr_subseg from b_tablagramaje WITH ( NOLOCK ) where tgr_subseg=" & SubsegD & " and tgr_codreg= " & CodregD & " AND tgr_codzon=" & CodzonD & "", vg_db, adOpenStatic
           
           If Not RS.EOF Then
              
              If MsgBox("Existe información en Tabla Gramaje destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                 
                 RS.Close
                 Set RS = Nothing
                 Exit Sub
              
              End If
           
           End If
           
           RS.Close
           Set RS = Nothing
        
        ElseIf Option1(5).Value = True Then
        
        
            RS.Open "select top 1  IdCeco from " & _
                    "b_tablagramajececo_nivel  WITH ( NOLOCK ) " & _
                    "where IdCeco='" & CecoD & "'", vg_db, adOpenStatic
            
           If Not RS.EOF Then
              
              If MsgBox("Existe información en Tabla Gramaje Nivel destino. se borrara la información existente ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                 
                 RS.Close
                 Set RS = Nothing
                 Exit Sub
              
              End If
              
           End If
           RS.Close
           Set RS = Nothing
        
        End If
               
        Sql = ""
        Sql = Sql & "sgpadm_Ins_TablaGramajeRegZona "
        
        If Option1(0).Value = True Then
           
           Sql = Sql & CecoO
           CodzonO = 0
        
        ElseIf Option1(1).Value = True Then
           
           Sql = Sql & SubsegO
        
        ElseIf Option1(4).Value = True Then
        
            Sql = ""
            Sql = Sql & "sgpadm_DelIns_TablaGramajeCecoNivel "
        
        End If
        
    
        If Option1(3).Value = True Then
           
           Sql = Sql & ", " & CodregO
           Sql = Sql & ", " & CodzonO
           Sql = Sql & ", " & CecoD
        
        ElseIf Option1(2).Value = True Then
           
           Sql = Sql & ", " & CodregO
           Sql = Sql & ", " & CodzonO
           Sql = Sql & ", " & SubsegD
        
        ElseIf Option1(5).Value = True Then
           
           Sql = Sql & " " & CecoO
           Sql = Sql & ", " & CecoD
           Sql = Sql & ", " & CodregO
           Sql = Sql & ", " & CodregD
           Sql = Sql & ", " & vg_NUsr
           
        End If
        
        If Option1(0).Value = True Then
           
            Sql = Sql & ", " & CodregD
            Sql = Sql & ", " & CodzonD
            Sql = Sql & ", " & "'2'"
        
        ElseIf Option1(1).Value = True Then
           
            Sql = Sql & ", " & CodregD
            Sql = Sql & ", " & CodzonD
            Sql = Sql & ", " & "'1'"
        
        End If
        
        If Option1(2).Value = True Then
           
           Sql = Sql & ", " & "'1'"
        
        ElseIf Option1(3).Value = True Then
           
           Sql = Sql & ", " & "'2'"
        
        End If
        
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("" & Sql & "")
        If Not RS.EOF Then
           
           If RS(0) > 0 Then
              
              MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
           End If
        
        End If
        RS.Close
        Set RS = Nothing
        MsgBox "Proceso terminado exitosamente."
        
    Case 3
    
        Me.Hide
        Unload Me
        M_TabGra.WindowState = 0
        
End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume

End Sub
