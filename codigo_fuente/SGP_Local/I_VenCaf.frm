VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form I_VenCaf 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Venta Cafetería"
   ClientHeight    =   4170
   ClientLeft      =   2505
   ClientTop       =   2985
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3315
      Left            =   0
      TabIndex        =   13
      Top             =   240
      Width           =   7470
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   60
         TabIndex        =   22
         Top             =   2520
         Width           =   7335
         Begin VB.OptionButton OptTipCli 
            Caption         =   "Todos"
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
            Index           =   3
            Left            =   3510
            TabIndex        =   9
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton OptTipCli 
            Caption         =   "Uno"
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
            Index           =   2
            Left            =   2520
            TabIndex        =   8
            Top             =   0
            Width           =   690
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   2
            Left            =   630
            TabIndex        =   10
            Top             =   255
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2520
            TabIndex        =   23
            Top             =   255
            Width           =   4065
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2025
            Picture         =   "I_VenCaf.frx":0000
            Top             =   150
            Width           =   480
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2565
            TabIndex        =   24
            Top             =   300
            Width           =   4065
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   675
         Left            =   60
         TabIndex        =   19
         Top             =   1740
         Width           =   7335
         Begin VB.OptionButton OptTipCli 
            Caption         =   "Uno"
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
            Left            =   2520
            TabIndex        =   5
            Top             =   0
            Width           =   690
         End
         Begin VB.OptionButton OptTipCli 
            Caption         =   "Todos"
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
            Left            =   3510
            TabIndex        =   6
            Top             =   0
            Width           =   855
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   630
            TabIndex        =   7
            Top             =   255
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
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2025
            Picture         =   "I_VenCaf.frx":030A
            Top             =   150
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2520
            TabIndex        =   11
            Top             =   255
            Width           =   4065
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2565
            TabIndex        =   20
            Top             =   300
            Width           =   4065
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
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
         Height          =   1620
         Left            =   60
         TabIndex        =   14
         Top             =   75
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   1305
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   675
            Width           =   3885
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   1335
            TabIndex        =   0
            Top             =   210
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
               Weight          =   700
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
            Left            =   1320
            TabIndex        =   3
            Top             =   1170
            Width           =   1470
            _Version        =   196608
            _ExtentX        =   2593
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
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
            Text            =   "17/08/2004"
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
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   4680
            TabIndex        =   4
            Top             =   1185
            Width           =   1470
            _Version        =   196608
            _ExtentX        =   2593
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
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
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
            Text            =   "17/08/2004"
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
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   1350
            TabIndex        =   25
            Top             =   765
            Width           =   3885
         End
         Begin VB.Label Label1 
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
            Height          =   225
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   270
            Width           =   1170
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   2685
            Picture         =   "I_VenCaf.frx":0614
            Top             =   120
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3135
            TabIndex        =   1
            Top             =   225
            Width           =   4065
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
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
            Left            =   165
            TabIndex        =   17
            Top             =   1240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Fecha Termino"
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
            Left            =   3240
            TabIndex        =   16
            Top             =   1240
            Width           =   1275
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
            Index           =   1
            Left            =   180
            TabIndex        =   15
            Top             =   760
            Width           =   660
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   3180
            TabIndex        =   21
            Top             =   270
            Width           =   4065
         End
      End
   End
End
Attribute VB_Name = "I_VenCaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim Msgtitulo As String
Public lc_Aux As String

Private Sub Combo1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpDateTime1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_GotFocus(Index As Integer)
Select Case Index
Case 1
    Select Case lc_Aux
    Case "VenCaf2", "VenCaf3"
        If Trim(fpText1(1).text) = "" Then Exit Sub
        fpText1(1).text = fg_DespintaRut(Trim(fpText1(1).text))
        fpText1(1).text = Mid(fpText1(1).text, 1, Len(Trim(fpText1(1).text)) - 1)
    End Select
End Select
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
On Local Error GoTo Error_CargaVta
'-------> Ajusta el ancho y el largo del form.
Me.Height = 3990
Me.Width = 7560
EspFecha fpDateTime1(0)
EspFecha fpDateTime1(1)
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
'-------> Crea Botones ---
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'----------------------------> Valida permisos para impresión
Toolbar1.Buttons.Item(1).Enabled = IIf(Val(Mid(ValidarUsuario(Me), 4, 1)) = 1, True, False)
'-------> Centra el formulario
fg_centra Me
Frame2.Visible = True
Frame4.Visible = True

If lc_Aux = "VenCaf1" Then
    Msgtitulo = "Ventas por artículo de cafetería"
    Me.Caption = "Ventas por artículo de cafetería"
    Frame2.Visible = False
    Frame4.Caption = "Articulo de cafetería"
    Frame4.Top = 1740
    Frame1.Height = Frame1.Height - 795
    Me.Height = Me.Height - 795
ElseIf lc_Aux = "VenCaf2" Then
    Msgtitulo = "Ventas de cafetería por cliente y centro de costo"
    Me.Caption = "Ventas de cafetería por cliente y centro de costo"
    Frame4.Visible = False
    Frame2.Caption = "Cliente"
    Frame1.Height = Frame1.Height - 795
    Me.Height = Me.Height - 795
ElseIf lc_Aux = "VenCaf3" Then
    Msgtitulo = "Ventas de cafetería por cliente y centro de costo detallado"
    Me.Caption = "Ventas de cafetería por cliente y centro de costo detallado"
    Frame2.Caption = "Cliente"
    Frame4.Caption = "Articulo de cafetería"
ElseIf lc_Aux = "VenCaf4" Then
    Msgtitulo = "Salida de bodega por ventas de cafetería"
    Me.Caption = "Salida de bodega por ventas de cafetería"
    Frame2.Caption = "Producto"
    Frame4.Caption = "Familia"
End If
'-------> Cargar Combo Bodega
CargarDatoCombo Combo1, 0, "b_clientes", "cli_", "CliBod", "N"
If Combo1(0).listcount > -1 Then Combo1(0).ListIndex = 0
OptTipCli(1).Value = True: OptTipCli(3).Value = True
fpDateTime1(0).text = "01/" & Format(Date, "mm/yyyy"): fpDateTime1(1).text = Date
fpText1(0).Enabled = ModCasino
Image1(0).Enabled = ModCasino
fpText1(0).text = MuestraCasino(1)
fpText1_LostFocus 0
Exit Sub
Error_CargaVta:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
 If fpText1(Index).text = "" Then Exit Sub
Select Case Index
Case 0 '
    RS1.Open RutinaLectura.Cliente(1, LimpiaDato(Trim(fpText1(0).text)), ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
        Do While Not RS1.EOF
            fpayuda(Index).Caption = RS1!cli_nombre
            RS1.MoveNext
        Loop
    Else
        RS1.Close: Set RS1 = Nothing
        fpText1(0).text = ""
        fpayuda(Index).Caption = ""
        MsgBox "Contrato no existe...", vbExclamation + vbOKOnly, Msgtitulo
        If fpText1(0).Enabled = True Then fpText1(0).SetFocus
        Exit Sub
    End If
    RS1.Close: Set RS1 = Nothing
Case 1
    Select Case lc_Aux
    Case "VenCaf2", "VenCaf3"
        fpText1(1).text = fg_RutDig(Trim(fpText1(1).text))
        RS1.Open RutinaLectura.Cliente(2, fg_DespintaRut(LimpiaDato(Trim(fpText1(1).text))), ""), vg_db, adOpenStatic
        If Not RS1.EOF Then
            fpText1(1).text = fg_PintaRut(RS1!cli_codigo)
            fpayuda(Index).Caption = RS1!cli_nombre
        Else
            RS1.Close: Set RS1 = Nothing
            fpText1(1).text = ""
            fpayuda(Index).Caption = ""
            MsgBox "Cliente no existe...", vbExclamation + vbOKOnly, Msgtitulo
            Exit Sub
        End If
        RS1.Close: Set RS1 = Nothing
    Case "VenCaf4"
        RS1.Open "SELECT DISTINCT a.pro_nombre, a.pro_codigo FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_codigo = '" & LimpiaDato(Trim(fpText1(1).text)) & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then
            fpText1(1).text = RS1!pro_codigo
            fpayuda(Index).Caption = RS1!pro_nombre
        Else
            RS1.Close: Set RS1 = Nothing
            fpText1(1).text = ""
            fpayuda(Index).Caption = ""
            MsgBox "Producto no existe...", vbExclamation + vbOKOnly, Msgtitulo
            Exit Sub
        End If
        RS1.Close: Set RS1 = Nothing
    End Select
Case 2
    Select Case lc_Aux
    Case "VenCaf1", "VenCaf3"
        RS1.Open "SELECT tpc_nombre, tpc_codigo FROM b_totpreciocaf WHERE tpc_cencos = '" & MuestraCasino(1) & "' AND tpc_codigo = '" & LimpiaDato(Trim(fpText1(2).text)) & "'", vg_db, adOpenStatic
        If Not RS1.EOF Then
            fpText1(2).text = RS1!tpc_codigo
            fpayuda(Index).Caption = RS1!tpc_nombre
        Else
            RS1.Close: Set RS1 = Nothing
            fpText1(2).text = ""
            fpayuda(Index).Caption = ""
            MsgBox "Articulo no existe...", vbExclamation + vbOKOnly, Msgtitulo
            Exit Sub
        End If
        RS1.Close: Set RS1 = Nothing
    Case "VenCaf4"
        fpayuda(Index).Caption = ""
        fpayuda(Index).Caption = fg_BuscaenArbol(Val(fpText1(2).text), "a_tipopro", "tip_codigo")
        If Trim(fpayuda(Index).Caption) = "" Then
            fpayuda(Index).Caption = ""
            fpText1(2).text = ""
            MsgBox "No existe codigo en la tabla..."
        End If
    End Select
End Select
End Sub

Private Sub image1_Click(Index As Integer)
vg_codigo = 0
Select Case Index
Case 0
    vg_left = fpayuda(Index).Left + 1920
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
    fpText1(Index) = Trim(vg_codigo)
    fpayuda(Index).Caption = vg_nombre
    fpText1_LostFocus Index
Case 1
    Select Case lc_Aux
    Case "VenCaf2", "VenCaf3"
        vg_nombre = "": vg_codigo = ""
        vg_left = fpayuda(Index).Left + 1920
        B_TabEst.LlenaDatos "b_clientes", "cli_", "Clientes", "Cliente"
        B_TabEst.Show 1
        Me.Refresh
        If Val(vg_codigo) = 0 Then Exit Sub
        fpText1(Index) = Trim(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
    Case "VenCaf4"
        vg_left = fpayuda(Index).Left + 1920
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Pst"
        B_TabEst.Show 1
        If Val(vg_codigo) = 0 Then Exit Sub
        fpText1(Index) = Trim(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
    End Select
Case 2
    Select Case lc_Aux
    Case "VenCaf1", "VenCaf3"
        vg_left = fpayuda(Index).Left + 1920
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_totpreciocaf", "tpc_", "Artículos", "Tpc"
        B_TabEst.Show 1
        If Val(vg_codigo) = 0 Then Exit Sub
        fpText1(Index) = Trim(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
    Case "VenCaf4"
        vg_left = fpayuda(Index).Left + 1920
        vg_nombre = "": vg_codigo = ""
        B_ArbEst.MoverDatosTvwDir "a_tipopro", "tip_", "Familia del Producto"
        B_ArbEst.Show 1
        Me.Refresh
        If Val(vg_codigo) = 0 Then Exit Sub
        fpText1(Index) = Trim(vg_codigo)
        fpayuda(Index).Caption = vg_nombre
    End Select
End Select
End Sub

Private Sub OptTipCli_Click(Index As Integer)
Select Case Index
Case 0
     OptTipCli(0).Value = True
     fpText1(1).Enabled = True: fpayuda(1).Enabled = True: Image1(1).Enabled = True
Case 1
    OptTipCli(1).Value = True
    fpText1(1).Enabled = False: fpayuda(1).Enabled = False: Image1(1).Enabled = False
    fpText1(1).text = "": fpayuda(1).Caption = ""
Case 2
     OptTipCli(2).Value = True
     fpText1(2).Enabled = True: fpayuda(2).Enabled = True: Image1(2).Enabled = True
Case 3
    OptTipCli(3).Value = True
    fpText1(2).Enabled = False: fpayuda(2).Enabled = False: Image1(2).Enabled = False
    fpText1(2).text = "": fpayuda(2).Caption = ""
End Select
End Sub

Private Sub OptTipCli_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Local Error GoTo Error_SalirVent
Select Case Button.Index
Case 1
    If Combo1(0).ListIndex = -1 Then MsgBox "Seleccione Bodega...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If OptTipCli(0).Value And Len(fpText1(1).text) = 0 Then MsgBox "Seleccione " & Trim(Frame2.Caption) & "...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If OptTipCli(2).Value And Len(fpText1(2).text) = 0 Then MsgBox "Seleccione " & Trim(Frame4.Caption) & "...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    Select Case lc_Aux
    Case "VenCaf1"
        I_VenCafArt Me
    Case "VenCaf2"
        I_VenCafCli Me
    Case "VenCaf3"
        I_VenCafCliArt Me
    Case "VenCaf4"
        I_VenCafPro Me
    End Select
Case 3
    Me.Hide
    Unload Me
End Select
Exit Sub
Error_SalirVent:
    MsgBox Err.Number & " " & Err.Description, vbExclamation + vbOKOnly, Msgtitulo
    Exit Sub
End Sub
