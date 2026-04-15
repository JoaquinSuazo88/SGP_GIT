VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Parame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros Sistemas"
   ClientHeight    =   8925
   ClientLeft      =   3090
   ClientTop       =   1590
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   120
      TabIndex        =   42
      Top             =   480
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Parámetros Web Service - Ftp - Correo"
      TabPicture(0)   =   "M_Parame.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame9"
      Tab(0).Control(1)=   "Frame8(0)"
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(4)=   "Frame1"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Sistemas"
      TabPicture(1)   =   "M_Parame.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame4"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Recetas"
      TabPicture(2)   =   "M_Parame.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame5"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Parametro SGP Local"
      TabPicture(3)   =   "M_Parame.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame8(1)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame8(2)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame8(3)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame8(4)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame8(5)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame Frame8 
         Caption         =   "Contraseña Comensales Diarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   5
         Left            =   6240
         TabIndex        =   124
         Top             =   600
         Width           =   4935
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   1080
            TabIndex        =   130
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   44
            Left            =   120
            TabIndex        =   125
            Top             =   795
            Width           =   810
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Contraseña Reabrir Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   4
         Left            =   240
         TabIndex        =   122
         Top             =   6000
         Width           =   4935
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   3
            Left            =   1080
            TabIndex        =   129
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   43
            Left            =   120
            TabIndex        =   123
            Top             =   795
            Width           =   810
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Contraseña Caratula Inventario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   3
         Left            =   240
         TabIndex        =   120
         Top             =   4200
         Width           =   4935
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   2
            Left            =   1080
            TabIndex        =   128
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   42
            Left            =   120
            TabIndex        =   121
            Top             =   795
            Width           =   810
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Contraseña Anular Envio Inventario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   2
         Left            =   240
         TabIndex        =   118
         Top             =   2400
         Width           =   4935
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   1
            Left            =   1080
            TabIndex        =   127
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   41
            Left            =   120
            TabIndex        =   119
            Top             =   795
            Width           =   810
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Contraseña Ajuste Precio Inventario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   1
         Left            =   240
         TabIndex        =   116
         Top             =   600
         Width           =   4935
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   0
            Left            =   1080
            TabIndex        =   126
            Top             =   600
            Width           =   3615
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   40
            Left            =   120
            TabIndex        =   117
            Top             =   795
            Width           =   810
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "Push"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   -69840
         TabIndex        =   112
         Top             =   2520
         Width           =   6375
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   25
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ButtonStyle     =   0
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
            MaxLength       =   30
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   26
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ButtonStyle     =   0
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
            MaxLength       =   30
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   27
            Left            =   120
            TabIndex        =   17
            Top             =   2280
            Width           =   6195
            _Version        =   196608
            _ExtentX        =   10927
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
            ButtonStyle     =   0
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
         Begin VB.Label Label1 
            Caption         =   "Ruta Actualizador"
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
            Index           =   39
            Left            =   120
            TabIndex        =   115
            Top             =   1920
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Versión SGPSDX"
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
            Index           =   38
            Left            =   120
            TabIndex        =   114
            Top             =   1080
            Width           =   1530
         End
         Begin VB.Label Label1 
            Caption         =   "Versión SGP Local"
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
            Index           =   37
            Left            =   120
            TabIndex        =   113
            Top             =   360
            Width           =   1770
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Cambio Contaseña ADM SGP LOCAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Index           =   0
         Left            =   -69840
         TabIndex        =   110
         Top             =   600
         Width           =   6375
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   24
            Left            =   1080
            TabIndex        =   14
            Top             =   360
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ButtonStyle     =   0
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
            MaxLength       =   30
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
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   36
            Left            =   120
            TabIndex        =   111
            Top             =   440
            Width           =   810
         End
      End
      Begin VB.Frame Frame4 
         Height          =   7695
         Left            =   -74160
         TabIndex        =   61
         Top             =   480
         Width           =   10095
         Begin VB.Frame Frame7 
            Caption         =   "Códigos Exento Impuesto SAP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Left            =   120
            TabIndex        =   98
            Top             =   5400
            Width           =   9255
            Begin EditLib.fpText fpText1 
               Height          =   315
               Index           =   16
               Left            =   2880
               TabIndex        =   33
               Top             =   360
               Width           =   1035
               _Version        =   196608
               _ExtentX        =   1826
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
               ButtonStyle     =   0
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
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               AutoCase        =   1
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
               MaxLength       =   2
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
            Begin EditLib.fpText fpText1 
               Height          =   315
               Index           =   17
               Left            =   2880
               TabIndex        =   35
               Top             =   720
               Width           =   1035
               _Version        =   196608
               _ExtentX        =   1826
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
               ButtonStyle     =   0
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
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               AutoCase        =   1
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
               MaxLength       =   2
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
            Begin EditLib.fpText fpText1 
               Height          =   315
               Index           =   18
               Left            =   7320
               TabIndex        =   34
               Top             =   360
               Width           =   1035
               _Version        =   196608
               _ExtentX        =   1826
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
               ButtonStyle     =   0
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
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               AutoCase        =   1
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
               MaxLength       =   2
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
            Begin EditLib.fpText fpText1 
               Height          =   315
               Index           =   19
               Left            =   7320
               TabIndex        =   36
               Top             =   720
               Width           =   1035
               _Version        =   196608
               _ExtentX        =   1826
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
               ButtonStyle     =   0
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
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
               AutoBeep        =   0   'False
               AutoCase        =   1
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
               MaxLength       =   2
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
            Begin VB.Label Label1 
               Caption         =   "Servicios Casinos No Gravados"
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
               Index           =   30
               Left            =   4440
               TabIndex        =   102
               Top             =   780
               Width           =   2730
            End
            Begin VB.Label Label1 
               Caption         =   "Servicios Casinos Gravados"
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
               Index           =   29
               Left            =   4440
               TabIndex        =   101
               Top             =   420
               Width           =   2610
            End
            Begin VB.Label Label1 
               Caption         =   "Insumos Casinos No Gravados"
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
               Index           =   28
               Left            =   120
               TabIndex        =   100
               Top             =   780
               Width           =   2610
            End
            Begin VB.Label Label1 
               Caption         =   "Insumos Casinos Gravados"
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
               Index           =   27
               Left            =   120
               TabIndex        =   99
               Top             =   420
               Width           =   2370
            End
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Calculo Digito Verificador"
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
            Left            =   240
            TabIndex        =   37
            Top             =   6720
            Width           =   2535
         End
         Begin VB.Frame Frame6 
            Caption         =   "Impuesto IVA"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3255
            Left            =   480
            TabIndex        =   62
            Top             =   7320
            Visible         =   0   'False
            Width           =   8895
            Begin VB.CommandButton Command1 
               Caption         =   "Agregar IVA"
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
               Left            =   5280
               TabIndex        =   64
               Top             =   2640
               Visible         =   0   'False
               Width           =   1575
            End
            Begin VB.CommandButton Command1 
               Caption         =   "Eliminar IVA"
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
               Left            =   6960
               TabIndex        =   63
               Top             =   2640
               Visible         =   0   'False
               Width           =   1575
            End
            Begin FPSpread.vaSpread vaSpread1 
               Height          =   2295
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Visible         =   0   'False
               Width           =   8655
               _Version        =   393216
               _ExtentX        =   15266
               _ExtentY        =   4048
               _StockProps     =   64
               DisplayRowHeaders=   0   'False
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               MaxCols         =   2
               SpreadDesigner  =   "M_Parame.frx":0070
            End
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   3000
            TabIndex        =   18
            Top             =   360
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "10"
            MinValue        =   "1"
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   2
            Left            =   3000
            TabIndex        =   19
            Top             =   720
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "10"
            MinValue        =   "1"
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   4
            Left            =   3000
            TabIndex        =   24
            Top             =   2520
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "10"
            MinValue        =   "1"
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   10
            Left            =   3000
            TabIndex        =   25
            Top             =   2880
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   5
            Left            =   3000
            TabIndex        =   27
            Top             =   3600
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   6
            Left            =   3000
            TabIndex        =   28
            Top             =   3960
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   7
            Left            =   3000
            TabIndex        =   29
            Top             =   4320
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   9
            Left            =   3000
            TabIndex        =   26
            Top             =   3240
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   11
            Left            =   3000
            TabIndex        =   20
            Top             =   1080
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   12
            Left            =   3000
            TabIndex        =   21
            Top             =   1455
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   13
            Left            =   3000
            TabIndex        =   22
            Top             =   1815
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   14
            Left            =   2760
            TabIndex        =   38
            Top             =   6990
            Width           =   6435
            _Version        =   196608
            _ExtentX        =   11351
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
            ButtonStyle     =   0
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
            MaxLength       =   180
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
            Index           =   8
            Left            =   3000
            TabIndex        =   30
            Top             =   4680
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   15
            Left            =   3000
            TabIndex        =   31
            Top             =   5040
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
            MaxLength       =   3
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
            Index           =   10
            Left            =   7440
            TabIndex        =   32
            Top             =   5040
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            MaxValue        =   "10"
            MinValue        =   "0"
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   20
            Left            =   3000
            TabIndex        =   23
            Top             =   2180
            Width           =   1035
            _Version        =   196608
            _ExtentX        =   1826
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
            ButtonStyle     =   0
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
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   1
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   10
            Left            =   4515
            TabIndex        =   105
            Top             =   2180
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   9
            Left            =   4020
            Picture         =   "M_Parame.frx":192A
            Top             =   2090
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable Flete Insumo"
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
            Index           =   32
            Left            =   240
            TabIndex        =   104
            Top             =   2220
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Días Holguras"
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
            Index           =   31
            Left            =   5520
            TabIndex        =   103
            Top             =   5115
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Moneda SAP"
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
            Index           =   26
            Left            =   240
            TabIndex        =   97
            Top             =   5115
            Width           =   2370
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   8
            Left            =   4020
            Picture         =   "M_Parame.frx":1C34
            Top             =   1725
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   9
            Left            =   4515
            TabIndex        =   87
            Top             =   1800
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   4020
            Picture         =   "M_Parame.frx":1F3E
            Top             =   1365
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   8
            Left            =   4515
            TabIndex        =   86
            Top             =   1455
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   4020
            Picture         =   "M_Parame.frx":2248
            Top             =   960
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   4515
            TabIndex        =   85
            Top             =   1080
            Width           =   4830
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable Movilización"
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
            Index           =   24
            Left            =   240
            TabIndex        =   84
            Top             =   1905
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable Desechable"
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
            Index           =   23
            Left            =   240
            TabIndex        =   83
            Top             =   1530
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Cta. Contable Alimentación"
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
            Left            =   240
            TabIndex        =   82
            Top             =   1155
            Width           =   2370
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   4515
            TabIndex        =   81
            Top             =   3240
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   4020
            Picture         =   "M_Parame.frx":2552
            Top             =   3150
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Retención Iva Cigarrillo"
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
            Index           =   21
            Left            =   240
            TabIndex        =   80
            Top             =   3315
            Width           =   2370
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   4515
            TabIndex        =   79
            Top             =   4320
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   4020
            Picture         =   "M_Parame.frx":285C
            Top             =   4245
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Retención Hortofruticola"
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
            Index           =   19
            Left            =   240
            TabIndex        =   78
            Top             =   4440
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Retención Ica"
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
            Index           =   18
            Left            =   240
            TabIndex        =   77
            Top             =   4065
            Width           =   2370
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   4020
            Picture         =   "M_Parame.frx":2B66
            Top             =   3885
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   4515
            TabIndex        =   76
            Top             =   3960
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   4020
            Picture         =   "M_Parame.frx":2E70
            Top             =   3525
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   4515
            TabIndex        =   75
            Top             =   3600
            Width           =   4830
         End
         Begin VB.Label Label1 
            Caption         =   "Retención en la Fuente"
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
            Index           =   17
            Left            =   240
            TabIndex        =   74
            Top             =   3675
            Width           =   2370
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   4515
            TabIndex        =   73
            Top             =   2880
            Width           =   4830
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   4020
            Picture         =   "M_Parame.frx":317A
            Top             =   2805
            Width           =   480
         End
         Begin VB.Label Label1 
            Caption         =   "Pais"
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
            Index           =   16
            Left            =   240
            TabIndex        =   72
            Top             =   2955
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Decimales En Cantidades"
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
            Index           =   11
            Left            =   240
            TabIndex        =   71
            Top             =   435
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Decimales En Precios"
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
            Index           =   12
            Left            =   240
            TabIndex        =   70
            Top             =   795
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "% Cuota Hortofruticola"
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
            Index           =   15
            Left            =   240
            TabIndex        =   69
            Top             =   2595
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Empresa"
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
            Index           =   20
            Left            =   240
            TabIndex        =   68
            Top             =   7080
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Retención Iva"
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
            Index           =   25
            Left            =   240
            TabIndex        =   67
            Top             =   4785
            Width           =   2370
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   4020
            Picture         =   "M_Parame.frx":3484
            Top             =   4605
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   4515
            TabIndex        =   66
            Top             =   4680
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   5
            Left            =   4560
            TabIndex        =   88
            Top             =   1125
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   4560
            TabIndex        =   89
            Top             =   1500
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   7
            Left            =   4560
            TabIndex        =   90
            Top             =   1860
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   12
            Left            =   4560
            TabIndex        =   91
            Top             =   2925
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   4560
            TabIndex        =   92
            Top             =   3285
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   4560
            TabIndex        =   93
            Top             =   3645
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   4560
            TabIndex        =   94
            Top             =   4005
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   4560
            TabIndex        =   95
            Top             =   4365
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   4560
            TabIndex        =   96
            Top             =   4725
            Width           =   4830
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   8
            Left            =   4560
            TabIndex        =   106
            Top             =   2230
            Width           =   4830
         End
      End
      Begin VB.Frame Frame5 
         Height          =   6255
         Left            =   -74280
         TabIndex        =   57
         Top             =   540
         Width           =   10095
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "M_Parame.frx":378E
            Left            =   3120
            List            =   "M_Parame.frx":3790
            Style           =   2  'Dropdown List
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   420
            Width           =   3765
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   40
            Top             =   840
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
            MaxValue        =   "10"
            MinValue        =   "1"
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
            Caption         =   "Decimales En Cantidades"
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
            Index           =   14
            Left            =   600
            TabIndex        =   60
            Top             =   915
            Width           =   2370
         End
         Begin VB.Label Label1 
            Caption         =   "Lista Precios"
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
            Index           =   13
            Left            =   600
            TabIndex        =   59
            Top             =   560
            Width           =   1410
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   3190
            TabIndex        =   58
            Top             =   500
            Width           =   3735
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Correo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1905
         Left            =   -74640
         TabIndex        =   52
         Top             =   5940
         Width           =   4770
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   6
            Left            =   2055
            TabIndex        =   10
            Top             =   360
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   7
            Left            =   2055
            TabIndex        =   11
            Top             =   720
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   8
            Left            =   2055
            TabIndex        =   12
            Top             =   1080
            Width           =   1875
            _Version        =   196608
            _ExtentX        =   3307
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
            ButtonStyle     =   0
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
            MaxLength       =   30
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   9
            Left            =   2055
            TabIndex        =   13
            Top             =   1440
            Width           =   1875
            _Version        =   196608
            _ExtentX        =   3307
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
            ButtonStyle     =   0
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
            MaxLength       =   30
            MultiLine       =   0   'False
            PasswordChar    =   "*"
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
         Begin VB.Label Label1 
            Caption         =   "Usuario"
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
            Index           =   5
            Left            =   135
            TabIndex        =   56
            Top             =   1080
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Cuenta"
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
            Index           =   6
            Left            =   135
            TabIndex        =   55
            Top             =   720
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
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
            Index           =   7
            Left            =   135
            TabIndex        =   54
            Top             =   360
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   8
            Left            =   135
            TabIndex        =   53
            Top             =   1440
            Width           =   1290
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ftp Push"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3105
         Left            =   -74640
         TabIndex        =   46
         Top             =   2540
         Width           =   4770
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   2
            Left            =   2055
            TabIndex        =   2
            Top             =   240
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   4
            Left            =   2055
            TabIndex        =   7
            Top             =   2040
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ButtonStyle     =   0
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
            MaxLength       =   30
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   5
            Left            =   2055
            TabIndex        =   8
            Top             =   2400
            Width           =   1635
            _Version        =   196608
            _ExtentX        =   2884
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
            ButtonStyle     =   0
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
            MaxLength       =   30
            MultiLine       =   0   'False
            PasswordChar    =   "*"
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   3
            Top             =   600
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
            Left            =   2040
            TabIndex        =   9
            Top             =   2760
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
            MaxValue        =   "10000"
            MinValue        =   "1"
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   21
            Left            =   2040
            TabIndex        =   4
            Top             =   960
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   22
            Left            =   2040
            TabIndex        =   5
            Top             =   1320
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   23
            Left            =   2040
            TabIndex        =   6
            Top             =   1680
            Width           =   2595
            _Version        =   196608
            _ExtentX        =   4577
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
            ButtonStyle     =   0
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
            MaxLength       =   50
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
         Begin VB.Label Label1 
            Caption         =   "Nombre(Push SGP)"
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
            Index           =   35
            Left            =   120
            TabIndex        =   109
            Top             =   1750
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Directorio(Push SGP)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   34
            Left            =   120
            TabIndex        =   108
            Top             =   1400
            Width           =   1890
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre(Act. SGP)"
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
            Index           =   33
            Left            =   120
            TabIndex        =   107
            Top             =   1040
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Servidor"
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
            Left            =   135
            TabIndex        =   51
            Top             =   340
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
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
            Left            =   135
            TabIndex        =   50
            Top             =   2100
            Width           =   1650
         End
         Begin VB.Label Label1 
            Caption         =   "Password"
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
            Index           =   4
            Left            =   135
            TabIndex        =   49
            Top             =   2490
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Directorio(Act. SGP)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   10
            Left            =   135
            TabIndex        =   48
            Top             =   700
            Width           =   1770
         End
         Begin VB.Label Label1 
            Caption         =   "Puerto"
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
            Index           =   9
            Left            =   120
            TabIndex        =   47
            Top             =   2820
            Width           =   1290
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Web Service"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1785
         Left            =   -74640
         TabIndex        =   43
         Top             =   540
         Width           =   4770
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   2055
            TabIndex        =   0
            Top             =   255
            Width           =   1875
            _Version        =   196608
            _ExtentX        =   3307
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
            ButtonStyle     =   0
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
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   2055
            TabIndex        =   1
            Top             =   720
            Width           =   1875
            _Version        =   196608
            _ExtentX        =   3307
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
            ButtonStyle     =   0
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
            MaxLength       =   20
            MultiLine       =   0   'False
            PasswordChar    =   "*"
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
         Begin VB.Label Label1 
            Caption         =   "Sap Usuario"
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
            Left            =   135
            TabIndex        =   45
            Top             =   405
            Width           =   1290
         End
         Begin VB.Label Label1 
            Caption         =   "Sap Password"
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
            TabIndex        =   44
            Top             =   780
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "M_Parame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim Est As Boolean
Public lc_Aux As String
Dim vecftpcor() As Variant

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim i As Long

Select Case Index

    Case 0 '-------> Agregar Impuesto
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_impuesto", "imp_", "Impuesto", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        '-------> validar si existe impuesto
        If vaSpread1.SearchCol(1, 0, vaSpread1.MaxRows, Trim(vg_codigo), SearchFlagsNone) <> -1 Then
           MsgBox "Impuesto ya existe en la grilla...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        End If
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 1
        vaSpread1.text = vg_codigo
        vaSpread1.Col = 2
        vaSpread1.text = vg_nombre
    
    Case 1 '-------> Eliminar impuesto
        
        If vaSpread1.MaxRows < 1 Then Exit Sub
        If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un impuesto...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        If MsgBox("Elimina impuesto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1

End Select

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Me.Width = 12000
Me.Height = 9435
fg_centra Me
Me.HelpContextID = vg_OpcM
MsgTitulo = "Parametros Sistemas"
Est = True
Dim BtnX As Object

Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", True, False): BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

   
   If lc_Aux <> "Parsgplocal" Then
   
      SSTab1.Tab = 0
      SSTab1.TabVisible(0) = True
      SSTab1.TabVisible(1) = True
      SSTab1.TabVisible(2) = True
      SSTab1.TabVisible(3) = True
      
      Frame1.Caption = "Parametros Web Service"
   
   Else
   
      SSTab1.Tab = 3
      SSTab1.TabVisible(0) = False
      SSTab1.TabVisible(1) = False
      SSTab1.TabVisible(2) = False
      SSTab1.TabVisible(3) = True
   
   End If
   
   ReDim Preserve vecftpcor(15, 2)
   
   '-------> Usuario web service
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'sapusu'")
   If Not RS.EOF Then fpText1(0).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(1, 1) = "sapusu"
   vecftpcor(1, 2) = "Sap Usuario Web Service"
   
   '-------> Password web service
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'sappas'")
   If Not RS.EOF Then fpText1(1).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(2, 1) = "sappas"
   vecftpcor(2, 2) = "Sap Password Web Service"
   
   '-------> Password ftp servidor
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftpser'")
   If Not RS.EOF Then fpText1(2).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(3, 1) = "ftpser"
   vecftpcor(3, 2) = "Ftp Servidor"
   
   '-------> Password ftp directorio
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftpdir'")
   If Not RS.EOF Then fpText1(3).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(4, 1) = "ftpdir"
   vecftpcor(4, 2) = "Ftp Directorio"
   
   '-------> Password ftp usuario
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftpusu'")
   If Not RS.EOF Then fpText1(4).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(5, 1) = "ftpusu"
   vecftpcor(5, 2) = "Ftp Usuario"
   
   '-------> Password ftp password
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftppas'")
   If Not RS.EOF Then fpText1(5).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(6, 1) = "ftppas"
   vecftpcor(6, 2) = "Ftp Password"
   
   '-------> Password correo servidor
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'corser'")
   If Not RS.EOF Then fpText1(6).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(7, 1) = "corser"
   vecftpcor(7, 2) = "Correo Cuenta Mail"
   
   '-------> Password correo cuenta mail
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'corcum'")
   If Not RS.EOF Then fpText1(7).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(8, 1) = "corcum"
   vecftpcor(8, 2) = "Correo Cuenta Mail"
   
   '-------> Password correo usuario
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'corusu'")
   If Not RS.EOF Then fpText1(8).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(9, 1) = "corusu"
   vecftpcor(9, 2) = "Correo Usuario"
   
   '-------> Password correo password
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'corpas'")
   If Not RS.EOF Then fpText1(9).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(10, 1) = "corpas"
   vecftpcor(10, 2) = "Correo Password"
   
   '-------> Password ftp puerto
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftppue'")
   If Not RS.EOF Then fpLongInteger1(0).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(11, 1) = "ftppue"
   vecftpcor(11, 2) = "ftp Puerto"
   
   '-------> Nombre Actualizador SGP
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftpnarch'")
   If Not RS.EOF Then fpText1(21).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(12, 1) = "ftpnarch"
   vecftpcor(12, 2) = "Ftp Nombre archivo SGP"
   
   '-------> Directorio Push
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftpdirp'")
   If Not RS.EOF Then fpText1(22).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(13, 1) = "ftpdirp"
   vecftpcor(13, 2) = "Ftp Directorio Push"
   
   '-------> Nombre Actualizador Push
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ftpnarchp'")
   If Not RS.EOF Then fpText1(23).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(14, 1) = "ftpnarchp"
   vecftpcor(14, 2) = "Ftp Nombre archivo Push"
   
   '-------> Cambio de contraseña sgp local
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'csenaadm'")
   If Not RS.EOF Then fpText1(24).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   vecftpcor(15, 1) = "csenaadm"
   vecftpcor(15, 2) = "Cambio Contraseña desde Adm SGP"
      
   '-------> Push sgp local
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient

   Set RS = vg_db.Execute("select isnull(bvc.VersionSGP,'') as VersionSGP, " & _
                          "isnull(bvc.VersionSGPSDX,'') as VersionSGPSDX, " & _
                          "isnull(bvc.rutaArchivoActualizador,'') as rutaArchivoActualizador " & _
                          "from b_versionescasino as bvc with (nolock)")
   If Not RS.EOF Then
   
      fpText1(25).text = RS!VersionSGP
      fpText1(26).text = RS!VersionSGPSDX
      fpText1(27).text = RS!rutaArchivoActualizador
         
   End If
   RS.Close
   Set RS = Nothing
   
      
   '-------> Mover Parametros recetas
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Combo1(0).Clear
   Set RS = vg_db.Execute("sgpadm_s_listaprecio 4, 0, 0, '" & vg_NUsr & "'")
   If Not RS.EOF Then
      Do While Not RS.EOF
         Combo1(0).AddItem Trim(RS!lpr_nombre) & Space(150) & "(" & fg_pone_cero(RS!lpr_codigo, 10) & ")"
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
   
   '-------> cantidad decimales
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parrcandec'")
   If Not RS.EOF Then fpLongInteger1(3).text = RS!par_valor
   RS.Close: Set RS = Nothing
   
   '-------> Usuario web service
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo='parlprrec'")
   If Not RS.EOF Then Combo1(0).ListIndex = fg_buscacbo(Combo1, 0, 10, fg_pone_cero(Str(IIf(IsNull(RS!par_valor), "", RS!par_valor)), 10))
   RS.Close: Set RS = Nothing

   '-------> cantidad decimales
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parcandec'")
   If Not RS.EOF Then fpLongInteger1(1).text = RS!par_valor
   RS.Close: Set RS = Nothing
   
   '-------> precio decimales
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parpredec'")
   If Not RS.EOF Then fpLongInteger1(2).text = RS!par_valor
   RS.Close: Set RS = Nothing

   '-------> Cuenta Contable Alimentación
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.cta_nombre FROM a_param a, a_ctacontable b WHERE a.par_valor = b.cta_codigo AND a.par_codigo = 'ctainsumo'")
   If Not RS.EOF Then fpText1(11).text = Trim(RS!par_valor): fpayuda(7).Caption = Trim(RS!cta_nombre)
   RS.Close: Set RS = Nothing
   
   '-------> Cuenta Limpieza Desechable
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.cta_nombre FROM a_param a, a_ctacontable b WHERE a.par_valor = b.cta_codigo AND a.par_codigo = 'ctalimdes'")
   If Not RS.EOF Then fpText1(12).text = Trim(RS!par_valor): fpayuda(8).Caption = Trim(RS!cta_nombre)
   RS.Close: Set RS = Nothing
   
   '-------> Cuenta Movilización
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.cta_nombre FROM a_param a, a_ctacontable b WHERE a.par_valor = b.cta_codigo AND a.par_codigo = 'ctamovil'")
   If Not RS.EOF Then fpText1(13).text = Trim(RS!par_valor): fpayuda(9).Caption = Trim(RS!cta_nombre)
   RS.Close: Set RS = Nothing
   
   '-------> Cuenta Flete Insumo
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.cta_nombre FROM a_param a, a_ctacontable b WHERE a.par_valor = b.cta_codigo AND a.par_codigo = 'ctafleins'")
   If Not RS.EOF Then fpText1(20).text = Trim(RS!par_valor): fpayuda(10).Caption = Trim(RS!cta_nombre)
   RS.Close: Set RS = Nothing
   
   '-------> pais
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.pai_nombre FROM a_param a, a_pais b WHERE a.par_valor = b.pai_codigo AND a.par_codigo = 'parpais'")
   If Not RS.EOF Then fpText1(10).text = Trim(RS!par_valor): fpayuda(1).Caption = Trim(RS!pai_nombre)
   RS.Close: Set RS = Nothing

'   '-------> Parametro Iva
'   vaSpread1.MaxRows = 0
'   Set RS = vg_db.Execute("SELECT par_valor FROM a_param a WHERE par_codigo = 'pariva'")
'   If Not RS.EOF Then
'      If Trim(RS!par_valor) <> "" Then
'         RS.Close: Set RS = Nothing
'         Set RS = vg_db.Execute("SELECT b.imp_codigo, b.imp_nombre FROM a_impuesto b WHERE b.imp_codigo IN (" & fg_CambiaChar(GetParametro("pariva"), ";", ",") & ")")
'         Do While Not RS.EOF
'            vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'            vaSpread1.Row = vaSpread1.MaxRows
'            vaSpread1.Col = 1
'            vaSpread1.text = RS!imp_codigo
'            vaSpread1.Col = 2
'            vaSpread1.text = Trim(RS!imp_nombre)
'            RS.MoveNext
'         Loop
'         RS.Close: Set RS = Nothing
'      Else
'         RS.Close: Set RS = Nothing
'      End If
'   Else
'      RS.Close: Set RS = Nothing
'   End If

   
   '-------> Parametro Iva Cigarrillo
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.imp_nombre FROM a_param a, a_impuesto b WHERE a.par_valor = b.imp_codigo AND a.par_codigo = 'parivacig'")
   If Not RS.EOF Then fpLongInteger1(9).Value = Trim(RS!par_valor): fpayuda(6).Caption = Trim(RS!imp_nombre)
   RS.Close: Set RS = Nothing
   
   '-------> Retencion en la Fuente
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.imp_nombre FROM a_param a, a_impuesto b WHERE a.par_valor = b.imp_codigo AND a.par_codigo = 'parretfue'")
   If Not RS.EOF Then fpLongInteger1(5).Value = Trim(RS!par_valor): fpayuda(2).Caption = Trim(RS!imp_nombre)
   RS.Close: Set RS = Nothing

   '-------> Retención Ica
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.imp_nombre FROM a_param a, a_impuesto b WHERE a.par_valor = b.imp_codigo AND a.par_codigo = 'parretica'")
   If Not RS.EOF Then fpLongInteger1(6).text = Trim(RS!par_valor): fpayuda(3).Caption = Trim(RS!imp_nombre)
   RS.Close: Set RS = Nothing

   '-------> Retención Hortofruticola
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.imp_nombre FROM a_param a, a_impuesto b WHERE a.par_valor = b.imp_codigo AND a.par_codigo = 'parrethorf'")
   If Not RS.EOF Then fpLongInteger1(7).text = Trim(RS!par_valor): fpayuda(4).Caption = Trim(RS!imp_nombre)
   RS.Close: Set RS = Nothing

   '-------> Retención Iva
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT a.par_valor, b.imp_nombre FROM a_param a, a_impuesto b WHERE a.par_valor = b.imp_codigo AND a.par_codigo = 'retiva'")
   If Not RS.EOF Then fpLongInteger1(8).text = Trim(RS!par_valor): fpayuda(5).Caption = Trim(RS!imp_nombre)
   RS.Close: Set RS = Nothing

   '-------> % Cuota Hortifruticola
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parhorfru'")
   If Not RS.EOF Then fpLongInteger1(4).text = RS!par_valor Else fpLongInteger1(4).text = 0
   RS.Close: Set RS = Nothing

   '-------> tipo Moneda SAP
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'tipmonsap'")
   If Not RS.EOF Then fpText1(15).text = IIf(IsNull(RS!par_valor), "", (Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> Código sap Exento Nº1
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe1'")
   If Not RS.EOF Then fpText1(16).text = IIf(IsNull(RS!par_valor), "", (Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> Código sap Exento Nº2
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe2'")
   If Not RS.EOF Then fpText1(17).text = IIf(IsNull(RS!par_valor), "", (Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> Código sap Exento Nº3
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe3'")
   If Not RS.EOF Then fpText1(18).text = IIf(IsNull(RS!par_valor), "", (Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> Código sap Exento Nº4
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe4'")
   If Not RS.EOF Then fpText1(19).text = IIf(IsNull(RS!par_valor), "", (Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> calculo digito verificador
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Check1.Value = 1
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parcaldig'")
   If Not RS.EOF Then Check1.Value = IIf(RS!par_valor = "S", 1, 0)
   RS.Close: Set RS = Nothing

   '-------> Nombre Empresa
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'nomempresa'")
   If Not RS.EOF Then fpText1(14).text = IIf(IsNull(RS!par_valor), "Sodexo Chile S.A.", Trim(RS!par_valor))
   RS.Close: Set RS = Nothing

   '-------> Días Holguras
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'pardiaholg'")
   If Not RS.EOF Then fpLongInteger1(10).Value = IIf(IsNull(RS!par_valor), "", Trim(RS!par_valor))
   RS.Close: Set RS = Nothing

   '-------> contraseña ajuste precio inventario
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parconajpi'")
   If Not RS.EOF Then Text1(0).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
     
   '-------> contraseña Anular Envio Inventario
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parconaein'")
   If Not RS.EOF Then Text1(1).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> contraseña Caratula Inventario
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parconcain'")
   If Not RS.EOF Then Text1(2).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> contraseña Reabrir Mes
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parconreme'")
   If Not RS.EOF Then Text1(3).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
   '-------> contraseña comensales Diarios
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parcomdia'")
   If Not RS.EOF Then Text1(4).text = IIf(IsNull(RS!par_valor), "", fg_Desencripta(Trim(RS!par_valor)))
   RS.Close: Set RS = Nothing
   
BloquearOpSistema

Est = False

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub BloquearOpSistema()

On Error GoTo Man_Error

'-------> bloquear opciones del sistema si el pasi = chile
Label1(15).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(17).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(18).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(19).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(25).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)

'Check1.Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)

Image1(1).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Image1(2).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Image1(3).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Image1(4).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(4).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(5).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(6).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(7).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpLongInteger1(8).Enabled = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Frame7.Caption = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", "Exento Impuesto SAP", "Códigos Exento Impuesto SAP")
Label1(27).Caption = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", "Código Exento Impuesto", "Insumos Casinos Gravados")
Label1(28).Visible = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(29).Visible = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
Label1(30).Visible = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpText1(17).Visible = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpText1(18).Visible = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)
fpText1(19).Visible = IIf(Trim(vg_pais) = "CL" Or Trim(vg_pais) = "", False, True)

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

    Case 5
        
        Set RS = vg_db.Execute("sgpadm_s_impuesto 5, " & Val(fpLongInteger1(5).Value) & ", ''")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
        fpayuda(2).Caption = Trim(RS!imp_nombre)
        RS.Close: Set RS = Nothing
    
    Case 6
        
        Set RS = vg_db.Execute("sgpadm_s_impuesto 5, " & Val(fpLongInteger1(6).Value) & ", ''")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(3).Caption = "": Exit Sub
        fpayuda(3).Caption = Trim(RS!imp_nombre)
        RS.Close: Set RS = Nothing
    
    Case 7
        
        Set RS = vg_db.Execute("sgpadm_s_impuesto 5, " & Val(fpLongInteger1(7).Value) & ", ''")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(4).Caption = "": Exit Sub
        fpayuda(4).Caption = Trim(RS!imp_nombre)
        RS.Close: Set RS = Nothing
    
    Case 8
        
        Set RS = vg_db.Execute("sgpadm_s_impuesto 5, " & Val(fpLongInteger1(8).Value) & ", ''")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(5).Caption = "": Exit Sub
        fpayuda(5).Caption = Trim(RS!imp_nombre)
        RS.Close: Set RS = Nothing
    
    Case 9
        
        Set RS = vg_db.Execute("sgpadm_s_impuesto 5, " & Val(fpLongInteger1(9).Value) & ", ''")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "": Exit Sub
        fpayuda(6).Caption = Trim(RS!imp_nombre)
        RS.Close: Set RS = Nothing

End Select

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_Change(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

    Case 10
        
        Set RS = vg_db.Execute("SELECT * FROM a_pais WHERE pai_codigo = '" & Trim(LimpiaDato(fpText1(10).text)) & "'")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
        fpayuda(1).Caption = Trim(RS!pai_nombre)
        RS.Close: Set RS = Nothing
        BloquearOpSistema
    
    Case 11
        
        Set RS = vg_db.Execute("SELECT * FROM a_ctacontable WHERE cta_codigo = '" & Trim(LimpiaDato(fpText1(11).text)) & "'")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(7).Caption = "": Exit Sub
        fpayuda(7).Caption = Trim(RS!cta_nombre)
        RS.Close: Set RS = Nothing
    
    Case 12
        
        Set RS = vg_db.Execute("SELECT * FROM a_ctacontable WHERE cta_codigo = '" & Trim(LimpiaDato(fpText1(12).text)) & "'")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(8).Caption = "": Exit Sub
        fpayuda(8).Caption = Trim(RS!cta_nombre)
        RS.Close: Set RS = Nothing
    
    Case 13
        
        Set RS = vg_db.Execute("SELECT * FROM a_ctacontable WHERE cta_codigo = '" & Trim(LimpiaDato(fpText1(13).text)) & "'")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(9).Caption = "": Exit Sub
        fpayuda(9).Caption = Trim(RS!cta_nombre)
        RS.Close: Set RS = Nothing
    
    Case 20
        
        Set RS = vg_db.Execute("SELECT * FROM a_ctacontable WHERE cta_codigo = '" & Trim(LimpiaDato(fpText1(20).text)) & "'")
        If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(10).Caption = "": Exit Sub
        fpayuda(10).Caption = Trim(RS!cta_nombre)
        RS.Close: Set RS = Nothing

End Select

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

    Case 0
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_pais", "pai_", "Pais", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(Index).Caption = Trim(vg_nombre)
        fpText1(10) = Trim(vg_codigo)
        On Error Resume Next ': fpLongInteger1(8).SetFocus
    
    Case 1
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_impuesto", "imp_", "Impuesto", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(2).Caption = Trim(vg_nombre)
        fpLongInteger1(5).Value = Val(vg_codigo)
        On Error Resume Next: fpLongInteger1(6).SetFocus
    
    Case 2
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_impuesto", "imp_", "Impuesto", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(3).Caption = Trim(vg_nombre)
        fpLongInteger1(6).Value = Val(vg_codigo)
        On Error Resume Next: fpLongInteger1(7).SetFocus
    
    Case 3
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_impuesto", "imp_", "Impuesto", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(4).Caption = Trim(vg_nombre)
        fpLongInteger1(7).Value = Val(vg_codigo)
        On Error Resume Next: fpLongInteger1(8).SetFocus
    
    Case 4
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_impuesto", "imp_", "Impuesto", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(5).Caption = Trim(vg_nombre)
        fpLongInteger1(8).Value = Val(vg_codigo)
        On Error Resume Next: Check1.SetFocus
    
    Case 5
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_impuesto", "imp_", "Impuesto", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(6).Caption = Trim(vg_nombre)
        fpLongInteger1(9).Value = Val(vg_codigo)
        On Error Resume Next: fpLongInteger1(5).SetFocus
    
    Case 6
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cuenta Contable", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(7).Caption = Trim(vg_nombre)
        fpText1(11).text = Val(vg_codigo)
        On Error Resume Next: fpText1(12).SetFocus
    
    Case 7
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cuenta Contable", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(8).Caption = Trim(vg_nombre)
        fpText1(12).text = Val(vg_codigo)
        On Error Resume Next: fpText1(13).SetFocus
    
    Case 8
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cuenta Contable", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(9).Caption = Trim(vg_nombre)
        fpText1(13).text = Val(vg_codigo)
        On Error Resume Next: fpText1(20).SetFocus
    
    Case 9
        
        vg_left = fpayuda(1).Left + 1920
        vg_codigo = ""
        vg_nombre = ""
        B_TabEst.LlenaDatos "a_ctacontable", "cta_", "Cuenta Contable", "Gen"
        B_TabEst.Show 1
        Me.Refresh
        If Trim(vg_codigo) = "" Then Exit Sub
        fpayuda(10).Caption = Trim(vg_nombre)
        fpText1(20).text = Val(vg_codigo)
        On Error Resume Next: fpLongInteger1(4).SetFocus

End Select

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim codlpr As Long
Dim i      As Long
Dim j      As Long
Dim varrep As String
Dim RS     As New ADODB.Recordset

Select Case Button.Index

    Case 2 '-------> Actualizar parametros generales
           
           If lc_Aux = "Parsgplocal" Then Exit Sub
           
           If Trim(fpayuda(1).Caption) = "" Or Trim(fpText1(0).text) = "" Or Trim(fpText1(1).text) = "" Or Trim(fpText1(2).text) = "" Or Trim(fpText1(3).text) = "" & _
              Trim(fpText1(4).text) = "" Or Trim(fpText1(5).text) = "" Or Trim(fpText1(6).text) = "" Or Trim(fpText1(7).text) = "" & _
              Trim(fpText1(8).text) = "" Or Trim(fpText1(9).text) = "" Or Trim(fpText1(21).text) = "" Or Trim(fpText1(22).text) = "" Or Trim(fpText1(23).text) = "" Or fpLongInteger1(0).text = "" Or fpLongInteger1(1).text = "" Or fpLongInteger1(2).text = "" Or fpLongInteger1(3).text = "" _
              Or Trim(fpayuda(7).Caption) = "" Or Trim(fpayuda(8).Caption) = "" Or Trim(fpayuda(9).Caption) = "" Then MsgBox "Falta información en los parametros...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           If Trim(fpText1(25).text) = "" Then
           
              MsgBox "Falta información de la Versión SGP LOCAL...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           End If
           
           If Trim(fpText1(26).text) = "" Then
           
              MsgBox "Falta información de la Versión SGPSDX LOCAL...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           End If
           
           If Trim(fpText1(27).text) = "" Then
           
              MsgBox "Falta información de la ruta SGP LOCAL...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           
           End If
           
           '-------> Inicio Web Service - Ftp - Correo
           j = 0
           For i = 1 To 15
               
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
    
               Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = '" & vecftpcor(i, 1) & "'")
               
               varrep = ""
               
               If i < 11 Then
                  
                  varrep = fg_Encripta(LimpiaDato(Trim(fpText1(j).text)))
               
               ElseIf i = 12 Then
               
                  varrep = fg_Encripta(LimpiaDato(Trim(fpText1(21).text)))
               
               ElseIf i = 13 Then
               
                  varrep = fg_Encripta(LimpiaDato(Trim(fpText1(22).text)))
               
               ElseIf i = 14 Then
               
                  varrep = fg_Encripta(LimpiaDato(Trim(fpText1(23).text)))
               
               ElseIf i = 15 Then
               
                  varrep = fg_Encripta(LimpiaDato(Trim(fpText1(24).text)))
               
               Else
                  
                  varrep = fg_Encripta(LimpiaDato(Trim(fpLongInteger1(0).text)))
               
               End If
               
               If RS.EOF Then
                  
                  vg_db.Execute "sgpadm_iu_param 'A', '" & vecftpcor(i, 1) & "', '" & vecftpcor(i, 2) & "', 'C', '" & varrep & "'"
               
               Else
                  
                  vg_db.Execute "sgpadm_iu_param 'M', '" & vecftpcor(i, 1) & "', '', '', '" & varrep & "'"
               
               End If
               
               RS.Close
               Set RS = Nothing
               j = j + 1
           
           Next i
           
           '-------> Parametros Generales -------
           '-------> Push SGP LOCAL
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT top 1 isnull(versionSGP,'') as versionSGP FROM b_versionescasino")
           If RS.EOF Then
           
              vg_db.Execute "insert into b_versionescasino (VersionSGP, VersionSGPSDX, rutaArchivoActualizador) values ('" & Trim(LimpiaDato(fpText1(25).text)) & "', '" & Trim(LimpiaDato(fpText1(26).text)) & "', '" & Trim(LimpiaDato(fpText1(27).text)) & "')"
           
           Else
              
              vg_db.Execute "update b_versionescasino set VersionSGP = '" & Trim(LimpiaDato(fpText1(25).text)) & "', VersionSGPSDX = '" & Trim(LimpiaDato(fpText1(26).text)) & "', rutaArchivoActualizador = '" & Trim(LimpiaDato(fpText1(27).text)) & "'"
           
           End If
           
           RS.Close
           Set RS = Nothing
           
           '-------> cantidad decimales
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parcandec'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parcandec', 'Parametro Cantidad Decimales', 'C', '" & fpLongInteger1(1).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parcandec', '', '', '" & fpLongInteger1(1).text & "'"
           End If
           RS.Close: Set RS = Nothing
           vg_DCa = fpLongInteger1(1).text
           
           '-------> precio decimales
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parpredec'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parpredec', 'Parametro Precios Decimales', 'C', '" & fpLongInteger1(2).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parpredec', '', '', '" & fpLongInteger1(2).text & "'"
           End If
           RS.Close: Set RS = Nothing
           vg_DPr = fpLongInteger1(2).text
           vg_DPr = fpLongInteger1(2).text
           
           '-------> Cuenta Contable Alimentación
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ctainsumo'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'ctainsumo', 'Cuentas de Insumos', 'C', '" & fpText1(11).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'ctainsumo', '', '', '" & fpText1(11).text & "'"
           End If
           RS.Close: Set RS = Nothing
    
           '-------> Cuenta Contable Desechable
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ctalimdes'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'ctalimdes', 'Cuentas Limpieza y Desechables', 'C', '" & fpText1(12).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'ctalimdes', '', '', '" & fpText1(12).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Cuenta Contable Mivilización
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ctamovil'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'ctamovil', 'Movilizacion', 'C', '" & fpText1(13).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'ctamovil', '', '', '" & fpText1(13).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Cuenta Flete Insumo
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'ctafleins'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'ctafleins', 'Cuenta Flete Insumo', 'C', '" & fpText1(20).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'ctafleins', '', '', '" & fpText1(20).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Pais
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parpais'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parpais', 'Parametro Pais', 'C', '" & fpText1(10).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parpais', '', '', '" & fpText1(10).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
    '       '-------> parametro iva
    '       Dim parval As String
    '       parval = ""
    '       For i = 1 To vaSpread1.MaxRows
    '           vaSpread1.Row = i
    '           vaSpread1.Col = 1
    '           parval = parval + Trim(vaSpread1.text) + ";"
    '       Next i
    '       If Trim(parval) <> "" Then parval = Mid(parval, 1, Len(parval) - 1)
    '       Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'pariva'")
    '       If RS.EOF Then
    '          vg_db.Execute "sgpadm_iu_param 'A', 'pariva', 'Parametro Iva', 'C', '" & parval & "'"
    '       Else
    '          vg_db.Execute "sgpadm_iu_param 'M', 'pariva', '', '', '" & parval & "'"
    '       End If
    '       RS.Close: Set RS = Nothing
           
           '-------> parametro iva cigarrillo
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parivacig'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parivacig', 'Parametro Iva Cigarrillo', 'C', '" & fpLongInteger1(9).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parivacig', '', '', '" & fpLongInteger1(9).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> parametro retencion en la fuente
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parretfue'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parretfue', 'Parametro Retención en la Fuente', 'C', '" & fpLongInteger1(5).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parretfue', '', '', '" & fpLongInteger1(5).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> parametro retencion ica
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parretica'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parretica', 'Parametro Retención Ica', 'C', '" & fpLongInteger1(6).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parretica', '', '', '" & fpLongInteger1(6).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> parametro retención hortifruticola
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parrethorf'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parrethorf', 'Parametro Retención Hortofruticola', 'C', '" & fpLongInteger1(7).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parrethorf', '', '', '" & fpLongInteger1(7).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> % Cuota Hortofruticola
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parhorfru'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parhorfru', '% Cuota Hortofruticola', 'C', '" & fpLongInteger1(4).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parhorfru', '', '', '" & fpLongInteger1(4).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Retención Iva
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'retiva'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'retiva', 'Retención Iva', 'C', '" & fpLongInteger1(8).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'retiva', '', '', '" & fpLongInteger1(8).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Fin Web Service - Ftp - Correo
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           If Combo1(0).ListIndex = -1 Then MsgBox "Debe seleccionar lista precio... ", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
           codlpr = Trim(fg_codigocbo(Combo1, 0, 10, 0))
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parlprrec'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parlprrec', 'Parametro Lista Precio Recetas', 'C', '" & codlpr & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parlprrec', '', '', '" & codlpr & "'"
           End If
           RS.Close: Set RS = Nothing
        
           '-------> cantidad decimales
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parrcandec'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parrcandec', 'Parametro Cantidad Decimales Recetas', 'C', '" & fpLongInteger1(3).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parrcandec', '', '', '" & fpLongInteger1(3).text & "'"
           End If
           RS.Close: Set RS = Nothing
           vg_RDCa = fpLongInteger1(3).text
        
           '-------> Tipo Moneda SAP
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'tipmonsap'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'tipmonsap', 'Parametro Tipo Moneda SAP', 'C', '" & fpText1(15).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'tipmonsap', '', '', '" & fpText1(15).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Exento Impuesto Nº1
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe1'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'codsapexe1', 'Código sap Exento Nº1', 'C', '" & fpText1(16).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'codsapexe1', '', '', '" & fpText1(16).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Exento Impuesto Nº2
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe2'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'codsapexe2', 'Código sap Exento Nº2', 'C', '" & fpText1(17).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'codsapexe2', '', '', '" & fpText1(17).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Exento Impuesto Nº3
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe3'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'codsapexe3', 'Código sap Exento Nº3', 'C', '" & fpText1(18).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'codsapexe3', '', '', '" & fpText1(18).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> Exento Impuesto Nº4
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'codsapexe4'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'codsapexe4', 'Código sap Exento Nº4', 'C', '" & fpText1(19).text & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'codsapexe4', '', '', '" & fpText1(19).text & "'"
           End If
           RS.Close: Set RS = Nothing
           
           '-------> calculo digito verificador
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'parcaldig'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'parcaldig', 'Parametro Calculo Digito Verificador', 'C', '" & IIf(Check1.Value = 1, "S", "N") & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'parcaldig', '', '', '" & IIf(Check1.Value = 1, "S", "N") & "'"
           End If
           RS.Close: Set RS = Nothing
           vg_Dig = IIf(Check1.Value = 1, "S", "N")
        
           '-------> Nombre empresa
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'nomempresa'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'nomempresa', 'Parametro Nombre Empresa', 'C', '" & LimpiaDato(fpText1(14).text) & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'nomempresa', '', '', '" & LimpiaDato(fpText1(14).text) & "'"
           End If
           RS.Close: Set RS = Nothing
        
           '-------> Días Holguras
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           Set RS = vg_db.Execute("SELECT par_valor FROM a_param WHERE par_codigo = 'pardiaholg'")
           If RS.EOF Then
              vg_db.Execute "sgpadm_iu_param 'A', 'pardiaholg', 'Días Holguras', 'C', '" & LimpiaDato(fpLongInteger1(10).Value) & "'"
           Else
              vg_db.Execute "sgpadm_iu_param 'M', 'pardiaholg', '', '', '" & LimpiaDato(fpLongInteger1(10).Value) & "'"
           End If
           RS.Close: Set RS = Nothing
        
           vg_pais = GetParametro("parpais")
           BloquearOpSistema
        
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
            
           Partida.StatusBar1.Panels(7).text = "Pais : """
           Set RS = vg_db.Execute("sgpadm_s_pais 1, '" & vg_pais & "', ''")
           If Not RS.EOF Then
              Partida.StatusBar1.Panels(7).text = "Pais : " & Trim(RS!pai_nombre) & " "
           End If
           RS.Close: Set RS = Nothing
        
           MsgBox "Información fué grabada...", vbInformation, MsgTitulo
    
    
    Case 5
        
        Me.Hide
        Unload Me

End Select

Exit Sub
Man_Error:
'If Err = -2147217900 Or 3704 Or 3265 Then Resume Next
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
