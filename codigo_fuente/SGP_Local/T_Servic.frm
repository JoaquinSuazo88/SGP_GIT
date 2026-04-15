VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form T_Servic 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Servicio"
   ClientHeight    =   6750
   ClientLeft      =   4050
   ClientTop       =   2325
   ClientWidth     =   10215
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   10215
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6165
      Left            =   120
      TabIndex        =   2
      Top             =   450
      Width           =   9960
      _ExtentX        =   17568
      _ExtentY        =   10874
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
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
      Tab(1).Control(0)=   "sombra(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblNOMBRE(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fpayuda(3)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2(4)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Image1(3)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "fpText(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "vaSpread2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command1"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Comensales Estimados"
      TabPicture(2)   =   "T_Servic.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "vaSpread3"
      Tab(2).Control(1)=   "fpText(2)"
      Tab(2).Control(2)=   "fpayuda(4)"
      Tab(2).Control(3)=   "Label2(5)"
      Tab(2).Control(4)=   "Image1(4)"
      Tab(2).Control(5)=   "sombra(4)"
      Tab(2).Control(6)=   "Label1(5)"
      Tab(2).Control(7)=   "Label1(3)"
      Tab(2).Control(8)=   "Label1(2)"
      Tab(2).Control(9)=   "lblNOMBRE(1)"
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Costo Techo"
      TabPicture(3)   =   "T_Servic.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame2"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame2 
         Height          =   5535
         Left            =   -74400
         TabIndex        =   17
         Top             =   480
         Width           =   7935
         Begin VB.Frame Frame3 
            Height          =   1335
            Left            =   150
            TabIndex        =   18
            Top             =   240
            Width           =   7575
            Begin EditLib.fpLongInteger fpLongInteger1 
               Height          =   315
               Index           =   0
               Left            =   1515
               TabIndex        =   20
               Top             =   540
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
               Left            =   1515
               TabIndex        =   19
               Top             =   210
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
               Left            =   1515
               TabIndex        =   21
               Top             =   870
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
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   3255
               TabIndex        =   28
               Top             =   540
               Width           =   3975
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   3255
               TabIndex        =   27
               Top             =   210
               Width           =   3975
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
               Left            =   240
               TabIndex        =   26
               Top             =   645
               Width           =   750
            End
            Begin VB.Label Label2 
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
               Index           =   1
               Left            =   240
               TabIndex        =   25
               Top             =   315
               Width           =   735
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   2745
               Picture         =   "T_Servic.frx":0070
               Top             =   120
               Width           =   480
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   1
               Left            =   2745
               Picture         =   "T_Servic.frx":037A
               Top             =   465
               Width           =   480
            End
            Begin VB.Image Image1 
               Enabled         =   0   'False
               Height          =   480
               Index           =   2
               Left            =   2745
               Picture         =   "T_Servic.frx":0684
               Top             =   795
               Width           =   480
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
               Left            =   240
               TabIndex        =   24
               Top             =   960
               Width           =   705
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3255
               TabIndex        =   22
               Top             =   870
               Width           =   3975
            End
            Begin VB.Label sombra 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   3300
               TabIndex        =   29
               Top             =   255
               Width           =   3975
            End
            Begin VB.Label sombra 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   3300
               TabIndex        =   30
               Top             =   585
               Width           =   3975
            End
            Begin VB.Label sombra 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   3300
               TabIndex        =   31
               Top             =   915
               Width           =   3975
            End
         End
         Begin FPSpread.vaSpread vaSpread4 
            Height          =   3375
            Left            =   225
            TabIndex        =   23
            Top             =   1680
            Width           =   7455
            _Version        =   393216
            _ExtentX        =   13150
            _ExtentY        =   5953
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
            MaxCols         =   2
            MaxRows         =   20
            ScrollBars      =   2
            SpreadDesigner  =   "T_Servic.frx":098E
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   2085
            Top             =   5190
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Mes Habilitados"
            Height          =   195
            Index           =   0
            Left            =   2445
            TabIndex        =   33
            Top             =   5160
            Width           =   1125
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H008484FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   240
            Top             =   5190
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Mes Bloqueados"
            Height          =   195
            Index           =   1
            Left            =   600
            TabIndex        =   32
            Top             =   5160
            Width           =   1185
         End
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
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1935
         Visible         =   0   'False
         Width           =   315
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -73300
         TabIndex        =   3
         Top             =   510
         Width           =   6015
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "T_Servic.frx":0C9F
            Left            =   2175
            List            =   "T_Servic.frx":0CA9
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Left            =   2175
            TabIndex        =   0
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
            TabIndex        =   7
            Top             =   300
            Width           =   1380
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
            Index           =   0
            Left            =   4755
            TabIndex        =   5
            Top             =   645
            Width           =   585
         End
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   4710
         Left            =   165
         TabIndex        =   9
         Top             =   1290
         Width           =   9645
         _Version        =   393216
         _ExtentX        =   17013
         _ExtentY        =   8308
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
         MaxCols         =   7
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_Servic.frx":0CBD
         ClipboardOptions=   0
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   1245
         Left            =   -73530
         TabIndex        =   10
         Top             =   1680
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
         SpreadDesigner  =   "T_Servic.frx":27CD
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   4335
         Left            =   -74720
         TabIndex        =   8
         Top             =   1620
         Width           =   9165
         _Version        =   393216
         _ExtentX        =   16166
         _ExtentY        =   7646
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
         MaxCols         =   9
         ProcessTab      =   -1  'True
         ScrollBars      =   2
         SpreadDesigner  =   "T_Servic.frx":2DAE
         ScrollBarTrack  =   1
         ClipboardOptions=   0
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   2115
         TabIndex        =   34
         Top             =   570
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   2
         Left            =   -72645
         TabIndex        =   38
         Top             =   810
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   -70905
         TabIndex        =   41
         Top             =   810
         Width           =   3975
      End
      Begin VB.Label Label2 
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
         Index           =   5
         Left            =   -73920
         TabIndex        =   40
         Top             =   915
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   -71415
         Picture         =   "T_Servic.frx":49A8
         Top             =   720
         Width           =   480
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   4
         Left            =   -70860
         TabIndex        =   39
         Top             =   855
         Width           =   3975
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   3345
         Picture         =   "T_Servic.frx":4CB2
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label2 
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
         Index           =   4
         Left            =   840
         TabIndex        =   36
         Top             =   675
         Width           =   735
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3855
         TabIndex        =   35
         Top             =   570
         Width           =   3975
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
         Left            =   -74400
         TabIndex        =   15
         Top             =   2460
         Width           =   645
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
         Left            =   -74400
         TabIndex        =   14
         Top             =   2190
         Width           =   750
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
         Left            =   -74400
         TabIndex        =   13
         Top             =   1950
         Width           =   690
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
         TabIndex        =   12
         Top             =   960
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
         Index           =   1
         Left            =   -73710
         TabIndex        =   11
         Top             =   1320
         Width           =   5280
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   3900
         TabIndex        =   37
         Top             =   615
         Width           =   3975
      End
   End
End
Attribute VB_Name = "T_Servic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String
Dim OpGr As Boolean
Dim ibusca As Long, iRow As Long, itop As Long
Public lc_Aux As String

Private Sub GrabaRegistro(Fila)

Dim coddet As Long, nomdet As String, orddet As Long, codenc As Long, codser As Long, nomenc As String, ordenc As Long, i As Long, j As Long
Dim nrorac As Long
Dim codsec As Long
Dim fecpat As Long
Dim valor As Double
Dim racmin As Long
Dim horcob As String
Dim horent As String
Dim codsap As String
Dim indfac As String
Dim indact As String
Dim marcaplatos As String
On Error GoTo Man_Error
OpGr = True
If Command1.Visible = True Then Command1.Visible = False
vaSpread1.Row = Fila
vaSpread1.Col = 1: codenc = Val(vaSpread1.Value)
vaSpread1.Col = 2: nomenc = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 3: ordenc = Val(vaSpread1.Value)
vaSpread1.Col = 4: codsap = vaSpread1.text
vaSpread1.Col = 5: indfac = vaSpread1.text
vaSpread1.Col = 6: indact = vaSpread1.text
vaSpread1.Col = 7: horcod = "00:00": horcob = IIf(IsNull(vaSpread1.text) Or Trim(vaSpread1.text) = "", "00:00", vaSpread1.text)
vaSpread1.Col = 8: horent = "00:00": horent = IIf(IsNull(vaSpread1.text) Or Trim(vaSpread1.text) = "", "00:00", vaSpread1.text)
If Trim(nomenc) = "" Or ordenc = 0 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell vaSpread1.ActiveCol, vaSpread1.ActiveRow: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" And SSTab1.Tab = 0 Then
   '------- ENCABEZADO
    MoverDatosGrillas2
'    vg_db.BeginTrans
    RS1.Open RutinaLectura.Servicio(3, 0, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
        RS1.MoveFirst
        codenc = RS1!ser_codigo + 1
        If codenc > 9999 Then
           vg_db.RollbackTrans
           vaSpread1.Row = Fila
           vaSpread1.DeleteRows vaSpread1.Row, 1
           vaSpread1.MaxRows = vaSpread1.MaxRows - 1
           RS1.Close: Set RS1 = Nothing
           MsgBox "No puede crear más registro, comuniquese con informatica...", vbExclamation + vbOKOnly, MsgTitulo
           modo = "": Gl_Ac_Botones Me, 1, IIf(vg_modpac, 5, 1), modo
           Exit Sub
        End If
    Else
        codenc = 1
    End If
    RS1.Close: Set RS1 = Nothing
    vg_db.Execute "INSERT INTO a_servicio (ser_codigo, ser_nombre, ser_orden, ser_codsap, ser_facturable, ser_activo, ser_horcob, ser_horent, ser_horpda) VALUES (" & codenc & ", '" & Trim(nomenc) & "', " & ordenc & ", '" & codsap & "', '" & indfac & "', '" & indact & "', '" & horcob & "', '" & horent & "', Null)"
'    vg_db.CommitTrans
    vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.Value = codenc
ElseIf modo = "M" And SSTab1.Tab = 0 Then
    '------- ENCABEZADO
'    vg_db.BeginTrans
    vg_db.Execute "UPDATE a_servicio SET ser_nombre='" & Trim(nomenc) & "', ser_orden=" & ordenc & ", ser_codsap='" & codsap & "', ser_facturable='" & indfac & "', ser_activo='" & indact & "', ser_horcob='" & horcob & "', ser_horent='" & horent & "' WHERE ser_codigo=" & codenc & ""
'    vg_db.CommitTrans
End If

'------- DETALLE
If vaSpread2.MaxRows > 0 And SSTab1.Tab = 1 Then
    vaSpread2.Row = vaSpread2.ActiveRow
'    vg_db.BeginTrans
    If modo = "A" Then
        RS1.Open RutinaLectura.EstServicio(3, 0, 0), vg_db, adOpenStatic
        If Not RS1.EOF Then
            RS1.MoveFirst
            coddet = RS1!ess_codigo + 1
            If coddet > 9999 Then
               vg_db.RollbackTrans
               vaSpread2.Row = vaSpread2.ActiveRow
               vaSpread2.DeleteRows vaSpread2.Row, 1
               vaSpread2.MaxRows = vaSpread2.MaxRows - 1
               RS1.Close: Set RS1 = Nothing
               MsgBox "No puede crear más registro, comuniquese con informatica...", vbExclamation + vbOKOnly, MsgTitulo
               modo = "": Gl_Ac_Botones Me, 1, 1, modo
               SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(1) = True: SSTab1.TabEnabled(2) = True
               Exit Sub
            End If
        Else
            coddet = 1
        End If
        RS1.Close: Set RS1 = Nothing
    Else
        vaSpread2.Col = 1: vaSpread2.TypeHAlign = TypeHAlignRight: coddet = Val(vaSpread2.Value)
    End If
    vaSpread2.Col = 2: nomdet = Mid(Trim(LimpiaDato(vaSpread2.Value)), 1, 30)
    vaSpread2.Col = 3: orddet = Val(vaSpread2.Value)
    vaSpread2.Col = 4: codsec = Val(vaSpread2.Value)
    vaSpread2.Col = 6: racmin = Val(vaSpread2.Value)
    vaSpread2.Col = 7: marcaplatos = vaSpread2.Value
    If Trim(nomdet) = "" Or orddet = 0 Or codsec = 0 Then
        If codsec = 0 Then MsgBox "Favor ingresar sector, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo
        If nomdet = "" Then MsgBox "Favor ingresar descripción, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo
        If orddet = 0 Then MsgBox "Favor ingresar orden, dado a que el campo se encuentra en blanco.", vbExclamation + vbOKOnly, MsgTitulo
'        MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo
        vaSpread2.Col = 2: vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow: vaSpread2.SetFocus
        'vg_db.CommitTrans
        OpGr = False
        Exit Sub
    End If
    If modo = "A" Then
        vg_db.Execute "INSERT INTO a_estservicio (ess_codser, ess_codigo, ess_nombre, ess_orden, ess_codsec, ess_racmin, ess_cencos, ess_marcaplatos) VALUES (" & codenc & ", " & coddet & ", '" & Trim(nomdet) & "', " & orddet & ", " & codsec & ", " & racmin & ", '" & Trim(fpText(1).text) & "', '" & marcaplatos & "')"
    Else
        vg_db.Execute "UPDATE a_estservicio SET ess_nombre='" & Trim(nomdet) & "', ess_orden=" & orddet & ", ess_codsec=" & codsec & ", ess_racmin=" & racmin & " , ess_marcaplatos = '" & marcaplatos & "' WHERE ess_cencos='" & Trim(fpText(1).text) & "' AND ess_codser=" & codenc & " AND ess_codigo=" & coddet
    End If
    vaSpread2.Col = 1: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.Value = coddet
'    vg_db.CommitTrans
End If
'------- Costo Techo
If SSTab1.Tab = 3 Then
   vaSpread4.Row = vaSpread4.ActiveRow
   vaSpread4.Col = 1: fecpat = Mid(vaSpread4.text, 4, 4) & Mid(fg_pone_cero(vaSpread4.text, 2), 1, 2)
   vaSpread4.Col = 2: valor = IIf(Trim(vaSpread4.text) = "", 0, vaSpread4.text)
   If valor = 0 Then MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo: vaSpread4.Row = vaSpread4.ActiveRow: vaSpread4.Col = 2: vaSpread4.SetActiveCell vaSpread4.ActiveCol, vaSpread4.ActiveRow: vaSpread4.SetFocus: OpGr = False: Exit Sub
   RS1.Open RutinaLectura.CostoPatron(2, Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), fecpat), vg_db, adOpenStatic
   If RS1.EOF Then
      vg_db.Execute "INSERT INTO b_costopatron VALUES ('" & Trim(LimpiaDato(fpText(0).text)) & "', " & Val(fpLongInteger1(0).Value) & ", " & Val(fpLongInteger1(1).Value) & ", " & fecpat & ", 'TECHO', " & valor & ")"
   Else
      vg_db.Execute "UPDATE b_costopatron SET cpa_valor=" & valor & " WHERE cpa_cencos='" & Trim(LimpiaDato(fpText(0).text)) & "' AND cpa_codreg=" & Val(fpLongInteger1(0).Value) & " AND cpa_codser =" & Val(fpLongInteger1(1).Value) & " AND cpa_anomes=" & fecpat & " AND cpa_descripcion='TECHO'"
   End If
   RS1.Close: Set RS1 = Nothing
   Frame3.Enabled = True
End If
'------- Raciones estimadas
If SSTab1.Tab = 2 Then
   For i = 1 To (vaSpread3.MaxRows - 1)
       vaSpread3.Row = i
       For j = 1 To vaSpread3.MaxCols
           vaSpread3.Col = j: nrorac = Val(vaSpread3.Value)
'           If nrorac > 0 Then
              RS1.Open "SELECT * FROM a_serviciorac WHERE sra_cencos='" & Trim(fpText(2).text) & "' AND sra_codser=" & codenc & " AND sra_coditem=" & i & " AND sra_serdia=" & j & "", vg_db, adOpenStatic
              If Not RS1.EOF Then
                 vg_db.Execute "UPDATE a_serviciorac SET sra_raciones=" & nrorac & " WHERE sra_cencos='" & Trim(fpText(2).text) & "' AND sra_codser=" & codenc & " AND sra_coditem=" & i & " AND sra_serdia=" & j & ""
              Else
                 vg_db.Execute "INSERT INTO a_serviciorac (sra_codser, sra_coditem, sra_serdia, sra_raciones, sra_cencos) VALUES (" & codenc & ", " & i & ", " & j & ", " & nrorac & ", '" & Trim(fpText(2).text) & "')"
             End If
             RS1.Close: Set RS1 = Nothing
'           End If
       Next j
   Next i
   modo = "M"
   Gl_Ac_Botones Me, 1, 7, modo
Else
   Label2(0).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
   Combo1.Enabled = True: fpText1.Enabled = True
   modo = "": Gl_Ac_Botones Me, 1, IIf(vg_modpac, 5, 1), modo
   Bloquear
End If

If SSTab1.Tab = 3 Then Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = True
SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(1) = True: SSTab1.TabEnabled(2) = True: SSTab1.TabEnabled(3) = True
OpGr = False
Exit Sub
Man_Error:
If Err = -2147467259 Then
'   vg_db.RollbackTrans
   MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
   Exit Sub
End If
If Err = 3034 Then
'vg_db.RollbackTrans:
   
   Exit Sub

End If
'vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub Command1_Click()

vg_left = Command1.Left + 3801
vg_nombre = "": vg_codigo = ""
B_TabEst.LlenaDatos "a_sector", "sec_", "Sector", "Gen"
B_TabEst.Show 1
Me.Refresh
With vaSpread2
    If vg_codigo = "" Then .Col = 4: .Row = iRow: .SetActiveCell 4, iRow: .EditMode = True: .EditModeReplace = True: .SetFocus: Exit Sub
    .Row = iRow
    .Col = 4
    .Value = vg_codigo
    .Col = 5
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
Me.Height = 7230
Me.Width = 10050
MsgTitulo = "Servicio"
fg_centra Me
modo = ""
ibusca = 0
itop = 1

Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
Combo1.ListIndex = 1

If vg_modpac Then
   
   SSTab1.TabVisible(1) = False
   SSTab1.TabVisible(2) = False
   SSTab1.TabVisible(3) = False

End If
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

Private Sub fpLongInteger1_Change(Index As Integer)
Select Case Index
Case 0
    RS1.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(0).Value), ""), vg_db, adOpenStatic
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS1!reg_nombre)
    RS1.Close: Set RS1 = Nothing
    MoverVectorCtoTecho
End Select
End Sub

Private Sub fpText_Change(Index As Integer)
RS1.Open RutinaLectura.Cliente(1, Trim(LimpiaDato(fpText(Index).text)), ""), vg_db, adOpenStatic
If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(IIf(Index = 0, 0, IIf(Index = 1, 3, 4))).Caption = "": fpLongInteger1(0).Value = "": fpayuda(1).Caption = "": Exit Sub
fpayuda(IIf(Index = 0, 0, IIf(Index = 1, 3, 4))).Caption = Trim(RS1!cli_nombre)
RS1.Close: Set RS1 = Nothing
fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
MoverVectorCtoTecho
End Sub

Private Sub fpText1_Change()
Dim sql1 As String
If LimpiaDato(Trim(fpText1.text)) & Chr(KeyAscii) = "" Then Exit Sub
If Combo1.ItemData(Combo1.ListIndex) = 0 Then
    RS2.Open RutinaLectura.Servicio(4, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
    RS2.Open RutinaLectura.Servicio(5, 0, UCase(LimpiaDato(fpText1.text))), vg_db, adOpenStatic
End If
With vaSpread1
    .MaxRows = RS2.RecordCount
    i = 1
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          .Row = i
          i = i + 1
          .Col = 1: .TypeHAlign = TypeHAlignRight: .Value = RS2!ser_codigo
          
          .Col = 2
          .ColWidth(2) = IIf(vg_modpac, 19, 38.14)
    '      .ColWidth(2) = IIf(vg_modpac, 28.88, 38.14)
          .CellType = IIf(RS2!ser_codigo > 9999 Or vg_modpac, CellTypeStaticText, CellTypeEdit)
          .Value = IIf(IsNull(RS2!ser_nombre), "", Trim(RS2!ser_nombre))
          
          vaSpread1.Col = 3
          If RS2!ser_codigo > 9999 Or vg_modpac Then
             .CellType = CellTypeStaticText
             .TypeHAlign = TypeHAlignRight
          Else
             .CellType = CellTypeNumber
             .TypeNumberDecPlaces = 0
             .TypeIntegerMin = 1
             .TypeIntegerMax = 9999999
             .TypeHAlign = TypeHAlignRight
             .TypeSpin = False
             .TypeIntegerSpinInc = 1
             .TypeIntegerSpinWrap = False
         End If
         .Value = IIf(IsNull(RS2!ser_orden), "", Trim(RS2!ser_orden))
         
         .Col = 4: .Lock = IIf(RS2!ser_codigo > 9999 Or vg_modpac, True, False): .text = IIf(IsNull(RS2!ser_codsap), "", RS2!ser_codsap)
         .Col = 5: .Lock = IIf(RS2!ser_codigo > 9999 Or vg_modpac, True, False): .text = IIf(IsNull(RS2!ser_facturable), "0", RS2!ser_facturable)
    '     .Col = 6: .Lock = IIf(RS2!ser_codigo > 9999 Or vg_modpac, True, False): vaSpread1.text = IIf(IsNull(RS2!ser_activo), "0", RS2!ser_activo)
         .Col = 6: .Lock = False: .text = IIf(IsNull(RS2!ser_activo), "0", RS2!ser_activo)
         .Col = 7
         .ColHidden = IIf(vg_modpac, False, True)
         .TypeTime24Hour = TypeTime24Hour24HourClock
         .TypeTimeMin = "000000"
         .TypeTimeMax = "240000"
         .text = IIf(vg_modpac, Format(IIf(IsNull(RS2!ser_horcob), "0000", RS2!ser_horcob), "Hh:Nn"), "")
        
         .Col = 8
         .ColHidden = IIf(vg_modpac, False, True)
         .TypeTime24Hour = TypeTime24Hour24HourClock
         .TypeTimeMin = "000000"
         .TypeTimeMax = "240000"
         .text = IIf(vg_modpac, Format(IIf(IsNull(RS2!ser_horent), "0000", RS2!ser_horent), "Hh:Nn"), "")
        
         .Col = 9
         .ColHidden = IIf(vg_modpac, False, True)
         .Lock = True
         .TypeTime24Hour = TypeTime24Hour24HourClock
         .TypeTimeMin = "000000"
         .TypeTimeMax = "240000"
         .text = IIf(vg_modpac, Format(RS2!ser_horpda, "Hh:Nn"), "")
         
         RS2.MoveNext
       Loop
       SSTab1.TabEnabled(1) = True
       SSTab1.TabEnabled(2) = True
       Gl_Ac_Botones Me, 1, 1, modo
    Else
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(2) = False
       Gl_Ac_Botones Me, 1, 2, modo
    End If
    RS2.Close: Set RS2 = Nothing
    If fpText1.text = "" Then
       Label2(0).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    Else
       Label2(0).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
    End If
End With
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0, 3, 4
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Contratos", "Contrato"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(IIf(Index = 0, 0, IIf(Index = 3, 1, 2))).text = vg_codigo
    fpayuda(IIf(Index = 0, 0, IIf(Index = 3, 3, 4))).Caption = vg_nombre
    If Index <> 0 Then Exit Sub
    fpLongInteger1(0).Value = "": fpayuda(1).Caption = ""
Case 1
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim RS1 As New ADODB.Recordset
Select Case SSTab1.Tab
Case 0, 1
    itop = 1
    RS1.Open RutinaLectura.Servicio(6, 0, ""), vg_db, adOpenStatic
    If Not RS1.EOF Then
       Gl_Ac_Botones Me, 1, 1, modo
    Else
       Gl_Ac_Botones Me, 1, 2, modo
    End If
    RS1.Close: Set RS1 = Nothing
    Bloquear
    If SSTab1.Tab = 0 Then Exit Sub
    Me.Refresh
    fpText(1).Enabled = ModCasino
    Image1(3).Enabled = ModCasino
    fpText(1).text = MuestraCasino(1)
    fpayuda(3).Caption = MuestraCasino(2)
    vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(0).Caption = vaSpread1.Value
    MoverDatosGrillas2
    Command1.Visible = False: Command1.Top = 1935
Case 2
    Me.Refresh
    fpText(2).Enabled = ModCasino
    Image1(4).Enabled = ModCasino
    fpText(2).text = MuestraCasino(1)
    fpayuda(4).Caption = MuestraCasino(2)
    vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(1).Caption = vaSpread1.Value
    MoverDatosGrillas3
Case 3
    Gl_Ac_Botones Me, 1, 3, modo
    vaSpread4.Visible = False
    vaSpread4.MaxRows = 0
    fpText(0).Enabled = ModCasino
    Image1(0).Enabled = ModCasino
    fpText(0).text = MuestraCasino(1)
    fpayuda(0).Caption = MuestraCasino(2)
    fpLongInteger1(0).Value = ""
    If vaSpread1.MaxRows > 0 Then
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
       RS1.Open RutinaLectura.Servicio(6, codigo, ""), vg_db, adOpenStatic
       If Not RS1.EOF Then fpLongInteger1(1).Value = RS1!ser_codigo: fpayuda(2).Caption = Trim(RS1!ser_nombre)
       RS1.Close: Set RS1 = Nothing
    End If
    vaSpread4.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, Orden As String, codser As Long, fecpat As Long
On Error GoTo Man_Error
Select Case Button.Index
Case 1
    modo = "A"
    SSTab1.TabEnabled(IIf(SSTab1.Tab = 0, 1, 0)) = False
    If SSTab1.Tab = 0 Then
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        vaSpread2.MaxRows = 0
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = 3
        vaSpread1.TypeHAlign = TypeHAlignRight
        vaSpread1.Col = 2
        vaSpread1.text = ""
        vaSpread1.Col = 6
        vaSpread1.text = "1"
        vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus
    ElseIf SSTab1.Tab = 1 Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1
        If Val(vaSpread1.Value) > 9999 Then
           MsgBox "No puede crear estructura de servicio, para este servicio, comuniquese con informatica...", vbExclamation + vbOKOnly, MsgTitulo
           modo = "": Gl_Ac_Botones Me, 1, 1, modo
           SSTab1.TabEnabled(0) = True
           Exit Sub
        End If
        If vaSpread1.MaxRows < 1 Then Exit Sub
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(2) = False
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        iRow = vaSpread2.MaxRows: vaSpread2.Row = vaSpread2.MaxRows: vaSpread2.Col = 2: vaSpread2.SetActiveCell 2, vaSpread2.MaxRows: vaSpread2.SetFocus
        Command1.Visible = False
    ElseIf SSTab1.Tab = 3 Then
        '------- Validar contrato
        RS1.Open RutinaLectura.Cliente(1, Trim(LimpiaDato(fpText(0).text)), ""), vg_db, adOpenStatic
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(0).Caption = "": MsgBox "Debe seleccionar Contrato", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        RS1.Close: Set RS1 = Nothing
        '------- Validar Regimen
        RS1.Open RutinaLectura.Regimen(2, Val(fpLongInteger1(0).Value), ""), vg_db, adOpenStatic
        If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(1).Caption = "": MsgBox "Debe seleccionar Regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        RS1.Close: Set RS1 = Nothing
        SSTab1.TabEnabled(0) = False
        SSTab1.TabEnabled(1) = False
        SSTab1.TabEnabled(2) = False
        SSTab1.TabEnabled(3) = True
        Frame3.Enabled = False
        If vaSpread4.MaxRows > 0 Then
           vaSpread4.Row = vaSpread4.MaxRows: vaSpread4.Col = 1
           Fecha = BEoM("01/" & Mid(vaSpread4.text, 1, 2) & "/" & Mid(vaSpread4.text, 4, 4))
        Else
           Fecha = Format(Date, "dd/mm/yyyy")
'           Fecha = BEoM(Format(Date, "dd/mm/yyyy"))
        End If
        vaSpread4.MaxRows = vaSpread4.MaxRows + 1
        vaSpread4.Row = vaSpread4.MaxRows
        vaSpread4.Col = 1: vaSpread4.text = Mid(Fecha, 4, 7)
        vaSpread4.SetActiveCell 2, vaSpread4.Row: vaSpread4.SetFocus
    End If
    Gl_Ac_Botones Me, 1, 0, modo
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
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    If SSTab1.Tab = 0 Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
        vg_db.BeginTrans
        vg_db.Execute "DELETE a_servicio FROM a_servicio WHERE ser_codigo=" & codigo & ""
        vg_db.Execute "DELETE a_serviciorac FROM a_serviciorac WHERE sra_cencos='" & Trim(fpText(2).text) & "' AND sra_codser=" & codigo & ""
        vg_db.CommitTrans
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread3.MaxRows = 0: vaSpread3.MaxRows = 1
    ElseIf SSTab1.Tab = 1 Then
        vaSpread2.Row = vaSpread2.ActiveRow: vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread2.Col = 1: codigo = Val(vaSpread2.Value): vaSpread1.Col = 1: codser = Val(vaSpread1.Value)
        '------- Validar si existen datos en planificación
        RS1.Open "SELECT DISTINCT mid_estser FROM b_minutadet WHERE mid_estser=" & codigo & "", vg_db, adOpenStatic
        If Not RS1.EOF Then RS1.Close: Set RS1 = Nothing: MsgBox "El dato esta asociado planificación, no puede eliminar", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        RS1.Close: Set RS1 = Nothing
        '------- fin validar si existen datos en planificación
        vg_db.BeginTrans
        vg_db.Execute "DELETE a_estservicio FROM a_estservicio WHERE ess_cencos='" & Trim(fpText(1).text) & "' AND ess_codser=" & codser & " AND ess_codigo=" & codigo & ""
        vg_db.CommitTrans
        vaSpread2.DeleteRows vaSpread2.Row, 1
        vaSpread2.MaxRows = vaSpread2.MaxRows - 1
'    ElseIf SSTab1.Tab = 3 Then
'    '------- Costo Techo
'        vaSpread4.Row = vaSpread4.ActiveRow
'        vaSpread4.Col = 1: fecpat = Mid(vaSpread4.Text, 4, 4) & Mid(fg_pone_cero(vaSpread4.Text, 2), 1, 2)
'        RS1.Open "SELECT * FROM b_costopatron WHERE cpa_cencos='" & Trim(LimpiaDato(fpText.Text)) & "' AND cpa_codreg=" & Val(fpLongInteger1(0).Value) & " AND cpa_codser =" & Val(fpLongInteger1(1).Value) & " AND cpa_anomes=" & fecpat & " AND cpa_descripcion='TECHO'", vg_db, adOpenStatic
'        If Not RS1.EOF Then vg_db.Execute "UPDATE b_costopatron SET cpa_valor=0 WHERE cpa_cencos='" & Trim(LimpiaDato(fpText.Text)) & "' AND cpa_codreg=" & Val(fpLongInteger1(0).Value) & " AND cpa_codser =" & Val(fpLongInteger1(1).Value) & " AND cpa_anomes=" & fecpat & " AND cpa_descripcion='TECHO'"
'        RS1.Close: Set RS1 = Nothing
'        vaSpread4.Col = 2: vaSpread4.Text = 0
'        Frame3.Enabled = True
    End If
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 7
    fpText1.text = ""
    If SSTab1.Tab = 0 Then
       MoverDatosGrillas
    ElseIf SSTab1.Tab = 1 Then
       MoverDatosGrillas2
    ElseIf SSTab1.Tab = 2 Then
       MoverDatosGrillas3
    ElseIf SSTab1.Tab = 3 Then
       Frame3.Enabled = True
       MoverVectorCtoTecho
    End If
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    SSTab1.TabEnabled(0) = True: SSTab1.TabEnabled(1) = True: SSTab1.TabEnabled(2) = True: SSTab1.TabEnabled(3) = True
    If modo = "A" Then
       If SSTab1.Tab = 0 Then
          MoverDatosGrillas
          modo = "": Gl_Ac_Botones Me, 1, IIf(vg_modpac, 5, 1), modo
          Bloquear
       ElseIf SSTab1.Tab = 1 Then
          MoverDatosGrillas2
          modo = "": Gl_Ac_Botones Me, 1, 1, modo
       ElseIf SSTab1.Tab = 2 Then
          MoverDatosGrillas3
          modo = "": Gl_Ac_Botones Me, 1, 1, modo
       ElseIf SSTab1.Tab = 3 Then
          Frame3.Enabled = True
          MoverVectorCtoTecho
       End If
'       Combo1.Enabled = True: fpText1.Enabled = True
    Else
       Cancela
    End If
Case 12
    GrabaRegistro vaSpread1.ActiveRow
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If SSTab1.Tab = 0 Then
       I_Servic vg_modpac
    ElseIf SSTab1.Tab = 1 Then
       vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1
       I_EstructuraServicio Trim(LimpiaDato(fpText(1).text)), Val(vaSpread1.Value)
    ElseIf SSTab1.Tab = 2 Then
       vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1
       I_ComensalesEstimados Trim(LimpiaDato(fpText(2).text)), Val(vaSpread1.Value)
    ElseIf SSTab1.Tab = 3 Then
       I_CostoTecho Trim(LimpiaDato(fpText(0).text)), Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), fecpat
    End If
Case 18
    Me.Hide
    Unload Me
End Select
Exit Sub
Man_Error:
If Err = -2147467259 Then vg_db.RollbackTrans: MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If vaSpread1.MaxRows > 0 And modo <> "A" Then MoverDatosGrillas2
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)

If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(1) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False

End Sub

Private Sub MoverDatosGrillas()

With vaSpread1
    
    .MaxRows = 0
    RS1.Open RutinaLectura.Servicio(7, 0, ""), vg_db, adOpenStatic
    
    Do While Not RS1.EOF
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        If .Row = 1 Then MoverDatosGrillas2
        .Col = 1: .ColWidth(1) = IIf(vg_modpac, 5, 7.38): .TypeHAlign = TypeHAlignRight: .Value = RS1!ser_codigo
        
        .Col = 2
        .ColWidth(2) = IIf(vg_modpac, 19, 38.14)
        .CellType = IIf(RS1!ser_codigo > 9999 Or vg_modpac, CellTypeStaticText, CellTypeEdit)
        .Value = IIf(IsNull(RS1!ser_nombre), "", Trim(RS1!ser_nombre))
        
        .Col = 3
        .Lock = False
'        .ColWidth(3) = IIf(vg_modpac, 5, 7.88)
        
'        If RS1!ser_codigo > 9999 Or vg_modpac Then
           
'           .CellType = CellTypeStaticText
'           .TypeHAlign = TypeHAlignRight
        
'        Else
           
           .CellType = CellTypeNumber
           .TypeNumberDecPlaces = 0
           .TypeIntegerMin = 1
           .TypeIntegerMax = 9999999
           .TypeHAlign = TypeHAlignRight
           .TypeSpin = False
           .TypeIntegerSpinInc = 1
           .TypeIntegerSpinWrap = False
        
'        End If
        .Value = IIf(IsNull(RS1!ser_orden), "", Trim(RS1!ser_orden))
        
        .Col = 4
        .Lock = IIf(RS1!ser_codigo > 9999 Or vg_modpac, True, False): .text = IIf(IsNull(RS1!ser_codsap), "", RS1!ser_codsap)
        
        .Col = 5
        .Lock = IIf(RS1!ser_codigo > 9999 Or vg_modpac, True, False): .text = IIf(IsNull(RS1!ser_facturable), "0", RS1!ser_facturable)
        
        .Col = 6
        .Lock = False: .text = IIf(IsNull(RS1!ser_activo), "0", RS1!ser_activo)
    '    .Col = 6: .Lock = IIf(RS1!ser_codigo > 9999 Or vg_modpac, True, False): .text = IIf(IsNull(RS1!ser_activo), "0", RS1!ser_activo)
        
        .Col = 7
        .ColHidden = IIf(vg_modpac, False, True)
        .TypeTime24Hour = TypeTime24Hour24HourClock
        .TypeTimeMin = "000000"
        .TypeTimeMax = "240000"
        .text = IIf(vg_modpac, Format(IIf(IsNull(RS1!ser_horcob), "0000", RS1!ser_horcob), "Hh:Nn"), "")
        
        .Col = 8
        .ColHidden = IIf(vg_modpac, False, True)
        .TypeTime24Hour = TypeTime24Hour24HourClock
        .TypeTimeMin = "000000"
        .TypeTimeMax = "240000"
        .text = IIf(vg_modpac, Format(IIf(IsNull(RS1!ser_horent), "0000", RS1!ser_horent), "Hh:Nn"), "")
        
        .Col = 9
        .ColHidden = IIf(vg_modpac, False, True)
        .Lock = True
        .TypeTime24Hour = TypeTime24Hour24HourClock
        .TypeTimeMin = "000000"
        .TypeTimeMax = "240000"
        .text = IIf(vg_modpac, Format(RS1!ser_horpda, "Hh:Nn"), "")
        
        RS1.MoveNext
    
    Loop
    RS1.Close
    Set RS1 = Nothing
    Gl_Ac_Botones Me, 1, IIf(vg_modpac, 5, 1), modo
    Label2(0).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registro"
    Bloquear

End With

End Sub

Sub Bloquear()
Dim i As Long
Dim j As Long
'-------> bloquear servicio
Gl_Ac_Botones Me, 1, 7, modo
If SSTab1.Tab = 0 Then
   
   For i = 1 To vaSpread1.MaxRows
       
       vaSpread1.Row = i
       
       For j = 1 To 4
               
           If j <> 3 Then
              
              vaSpread1.Col = j
              vaSpread1.Lock = True
           
           End If
           
       Next j
   
   Next i

ElseIf SSTab1.Tab = 1 Then
    
    For i = 1 To vaSpread2.MaxRows
       
       vaSpread2.Row = i
       
       For j = 1 To 2
           vaSpread2.Col = j
           vaSpread2.Lock = True
       Next j
    Next i

End If
'-------> Fin bloqueo
End Sub

Private Sub MoverDatosGrillas2()

Dim codigo As Long
Dim RS2 As New ADODB.Recordset
OpGr = True

With vaSpread2
    
    Command1.Visible = False
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    .MaxRows = 0
'    RS2.Open "SELECT a.ess_codser, a.ess_codigo, a.ess_nombre, a.ess_orden, a.ess_codsec, a.ess_racmin, b.sec_nombre, ess_marcaplatos FROM a_estservicio a LEFT JOIN a_sector b ON a.ess_codsec = b.sec_codigo WHERE a.ess_cencos='" & Trim(fpText(1).text) & "' AND a.ess_codser=" & codigo & " ORDER BY a.ess_orden, a.ess_nombre", vg_db, adOpenStatic
    Set RS2 = vg_db.Execute("sgp_Sel_DetalleEstructuraServicio '" & Trim(fpText(1).text) & "', " & codigo & "")
    Do While Not RS2.EOF
        
        .MaxRows = .MaxRows + 1
        .Row = .MaxRows
        
        .Col = 1
        .TypeHAlign = TypeHAlignRight
        .Value = RS2!ess_codigo
        
        .Col = 2
        .CellType = IIf(codigo > 9999, CellTypeStaticText, CellTypeEdit)
        .Value = IIf(IsNull(RS2!ess_nombre), "", Trim(RS2!ess_nombre))
        
        .Col = 3
        .CellType = IIf(codigo > 9999, CellTypeStaticText, CellTypeEdit)
        .CellType = CellTypeEdit
        .TypeHAlign = TypeHAlignRight
        .Value = IIf(IsNull(RS2!ess_orden), "", RS2!ess_orden)
        
        .Col = 4
        .TypeHAlign = TypeHAlignLeft
        .Value = IIf(IsNull(RS2!ess_codsec) Or RS2!ess_codsec = 0, "", RS2!ess_codsec)
        
        .Col = 5
        .Value = IIf(IsNull(RS2!sec_nombre), "", Trim(RS2!sec_nombre))
        
        .Col = 6
        .Lock = IIf(codigo > 9999, True, False)
        .Value = IIf(IsNull(RS2!ess_racmin), "", Trim(RS2!ess_racmin))
        
        .Col = 7
        .Lock = False
        .text = IIf(IsNull(RS2!ess_marcaplatos), "0", RS2!ess_marcaplatos)
        
        RS2.MoveNext
    
    Loop
    RS2.Close
    Set RS2 = Nothing

End With
OpGr = False

End Sub

Private Sub MoverDatosGrillas3()
Dim codigo As Long
Dim RS2 As New ADODB.Recordset
With vaSpread3
    .Row = -1: .Col = -1:
    .BackColor = &H80000018
    vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
    .MaxRows = 0: .MaxRows = 2
    RS2.Open "SELECT * FROM a_serviciorac WHERE sra_cencos='" & Trim(fpText(2).text) & "' AND sra_codser=" & codigo & " ORDER BY sra_coditem, sra_serdia", vg_db, adOpenStatic
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          .Row = RS2!sra_coditem
          .Col = RS2!sra_serdia: .text = IIf(RS2!sra_raciones = 0, "", RS2!sra_raciones)
          RS2.MoveNext
       Loop
    End If
    .MaxRows = (.MaxRows + 1)
    .Row = .MaxRows
    .Col = 1
    .Col2 = .MaxCols
    .Row2 = .MaxRows
    .Lock = True
    .BlockMode = True
    ' Lock cells
    .Lock = True
    ' Protect the cells from being edited
    .Protect = True
    ' Turn block mode off
    .BlockMode = False
    .Col = -1: .BackColor = &HE0E0E0
End With
SumarTotales
RS2.Close: Set RS2 = Nothing
modo = "M"
Gl_Ac_Botones Me, 1, 7, modo
End Sub

Private Sub vaSpread1_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)

If (Col <> 3 And Col <> 5 And Col <> 6) Or Row = 0 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo

End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    
    GrabaRegistro Row

ElseIf Toolbar1.Buttons(12).Visible = False Then
    
    Cancela

End If

End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)

Select Case Col

Case Is <> 4
    
    Command1.Visible = False

Case 4
    
    Command1.Top = IIf(Row = 1, 1935, 1935 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    iRow = Row
    vaSpread2.Col = 4
    vaSpread2.TypeHAlign = TypeHAlignLeft

End Select

End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
Dim RS2 As New ADODB.Recordset
If SSTab1.Tab = 0 Then Exit Sub
iRow = Row
Command1.Top = IIf(Row = 1, 1935, 1935 + (240 * (Row - itop)))
Command1.Visible = True
If ChangeMade = False And Col <> 6 Then
   If Col <> 4 Then Command1.Visible = False
   Exit Sub
End If
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
Select Case Col
Case Is <> 4
    Command1.Visible = False
Case 4
    Command1.Top = IIf(Row = 1, 1935, 1935 + (240 * (Row - itop)))
    Command1.Visible = True
    vaSpread2.Row = Row
    vaSpread2.Col = Col
    RS2.Open "SELECT sec_nombre FROM a_sector WHERE sec_codigo=" & Val(vaSpread2.Value) & "", vg_db, adOpenStatic
    If RS2.EOF Then RS2.Close: Set RS2 = Nothing: vaSpread2.text = "": vaSpread2.Col = 5: vaSpread2.text = "": Exit Sub
    vaSpread2.Col = 5: vaSpread2.text = Trim(RS2!sec_nombre)
    RS2.Close: Set RS2 = Nothing
    Command1.Visible = False
End Select
End Sub

Private Sub vaSpread2_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
If (Col <> 7) Or Row = 0 Or OpGr Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False
SSTab1.TabEnabled(2) = False
SSTab1.TabEnabled(3) = False
End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 4 Then Command1.Visible = False
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro vaSpread1.ActiveRow
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub

Private Sub Cancela()
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim codigo As Long
If SSTab1.Tab = 0 Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
   RS1.Open RutinaLectura.Servicio(6, codigo, ""), vg_db, adOpenStatic
   If Not RS1.EOF Then
      vaSpread1.Col = 2: vaSpread1.Value = Trim(RS1!ser_nombre)
      vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignRight: vaSpread1.Value = Trim(RS1!ser_orden)
      vaSpread1.Col = 4: vaSpread1.text = IIf(IsNull(RS1!ser_codsap), "", Trim(RS1!ser_codsap))
      vaSpread1.Col = 5: vaSpread1.text = IIf(IsNull(RS1!ser_facturable), "0", Trim(RS1!ser_facturable))
      vaSpread1.Col = 6: vaSpread1.text = IIf(IsNull(RS1!ser_activo), "0", Trim(RS1!ser_activo))
   End If
   RS1.Close: Set RS1 = Nothing
   modo = "": Gl_Ac_Botones Me, 1, IIf(vg_modpac, 5, 1), modo
   Combo1.Enabled = True: fpText1.Enabled = True
   Bloquear
ElseIf SSTab1.Tab = 1 Then
   vaSpread2.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignRight: codser = Val(vaSpread1.Value)
   vaSpread2.Row = vaSpread2.ActiveRow: vaSpread2.Col = 1: vaSpread2.TypeHAlign = TypeHAlignRight: codigo = Val(vaSpread2.Value)
   RS1.Open "SELECT a.ess_codser, a.ess_codigo, a.ess_nombre, a.ess_orden, a.ess_codsec, b.sec_nombre, ess_marcaplatos FROM a_estservicio a LEFT JOIN a_sector b ON a.ess_codsec = b.sec_codigo WHERE a.ess_cencos='" & Trim(fpText(1).text) & "' AND a.ess_codser=" & codser & " AND a.ess_codigo=" & codigo, vg_db, adOpenStatic
   If Not RS1.EOF Then
      vaSpread2.Col = 2: vaSpread2.Value = Trim(RS1!ess_nombre)
      vaSpread2.Col = 3: vaSpread2.TypeHAlign = TypeHAlignRight: vaSpread2.Value = RS1!ess_orden
      vaSpread2.Col = 4: vaSpread2.TypeHAlign = TypeHAlignLeft: vaSpread2.Value = IIf(IsNull(RS1!ess_codsec) Or RS1!ess_codsec = 0, "", RS1!ess_codsec)
      vaSpread2.Col = 5: vaSpread2.Value = IIf(IsNull(RS1!sec_nombre), "", RS1!sec_nombre)
      vaSpread2.Col = 7: vaSpread2.Lock = False: vaSpread2.text = IIf(IsNull(RS1!ess_marcaplatos), "0", RS1!ess_marcaplatos)
   End If
   RS1.Close: Set RS1 = Nothing
   modo = "": Gl_Ac_Botones Me, 1, 1, modo
   Combo1.Enabled = True: fpText1.Enabled = True
   SSTab1.TabEnabled(0) = True
   SSTab1.TabEnabled(2) = True
   SSTab1.TabEnabled(3) = True
   Bloquear
ElseIf SSTab1.Tab = 2 Then
   Me.Refresh
   vaSpread1.Col = 2: vaSpread1.Row = vaSpread1.ActiveRow: lblNOMBRE(1).Caption = vaSpread1.Value
   MoverDatosGrillas3
ElseIf SSTab1.Tab = 3 Then
   If vaSpread4.MaxRows < 1 Then Exit Sub
   Me.Refresh
   vaSpread4.Row = vaSpread4.ActiveRow
   vaSpread4.Col = 1: fecpat = Mid(vaSpread4.text, 4, 4) & Mid(fg_pone_cero(vaSpread4.text, 2), 1, 2)
   RS1.Open "SELECT * FROM b_costopatron WHERE cpa_cencos='" & Trim(LimpiaDato(fpText(0).text)) & "' AND cpa_codreg=" & Val(fpLongInteger1(0).Value) & " AND cpa_codser =" & Val(fpLongInteger1(1).Value) & " AND cpa_anomes=" & fecpat & " AND cpa_descripcion='TECHO'", vg_db, adOpenStatic
   If Not RS1.EOF Then
      vaSpread4.Col = 2: vaSpread4.text = IIf(IsNull(RS1!cpa_valor), 0, RS1!cpa_valor)
   Else
   End If
   RS1.Close: Set RS1 = Nothing
   modo = "": Gl_Ac_Botones Me, 1, 1, modo
   Toolbar1.Buttons(5).Visible = False
   Toolbar1.Buttons(6).Visible = True
End If
End Sub

Private Sub vaSpread2_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)
itop = NewTop
Command1.Visible = False
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
With vaSpread3
    If .MaxRows < 1 Then Exit Sub
    .Row = Row
    .Col = Col
    If Val(.text) = 0 Then .text = ""
End With
End Sub

Sub SumarTotales()
Dim i As Long, j As Long, nrorac As Long
With vaSpread3
    For j = 1 To .MaxCols
        .Row = .MaxRows
        .Col = j: .text = ""
    Next j
    For i = 1 To (.MaxRows - 1)
        nrorac = 0
        For j = 1 To .MaxCols
            .Row = i
            .Col = j: nrorac = Val(.Value)
            .Row = .MaxRows
            If nrorac > 0 Then .Col = j: .Value = (Val(.Value) + nrorac)
        Next j
    Next i
End With
End Sub

Sub MoverVectorCtoTecho()
Dim codigo As Long
Dim diablq As Date
Dim est As Boolean
If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread4.Visible = False
vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = Val(vaSpread1.Value)
vaSpread4.MaxRows = 0
RS1.Open RutinaLectura.CostoPatron(3, Val(fpLongInteger1(0).Value), codigo, 0), vg_db, adOpenStatic
If Not RS1.EOF Then
   Do While Not RS1.EOF
      vaSpread4.MaxRows = vaSpread4.MaxRows + 1
      vaSpread4.Row = vaSpread4.MaxRows
      '------- Bloquea días de cierre en color rojo
      If TipoDato(GetParametro("diasbloq"), 0) <> 0 Then diablq = Format(Right("00" & Val(GetParametro("diasbloq")), 2) & Format(Date, "/mm/yyyy"), "dd/mm/yyyy") Else diablq = 0
      est = False
      If Format(Mid(RS1!cpa_anomes, 5, 2) & "/" & Mid(RS1!cpa_anomes, 1, 4), "yyyymm") < Format(Date, "yyyymm") Then
         est = True
      ElseIf Format(Mid(RS1!cpa_anomes, 5, 2) & "/" & Mid(RS1!cpa_anomes, 1, 4), "yyyymm") = Format(Date, "yyyymm") And Format(Date, "yyyymmdd") > Format(diablq, "yyyymmdd") Then
         est = True
      End If
      vaSpread4.Col = 1: vaSpread4.BackColor = IIf(est, Shape1(0).FillColor, Shape1(1).FillColor): vaSpread4.text = Mid(RS1!cpa_anomes, 5, 2) & "/" & Mid(RS1!cpa_anomes, 1, 4)
      vaSpread4.Col = 2: vaSpread4.text = IIf(IsNull(RS1!cpa_valor), 0, RS1!cpa_valor): vaSpread4.BackColor = IIf(est, Shape1(0).FillColor, Shape1(1).FillColor): vaSpread4.Lock = IIf(est, True, False)
      RS1.MoveNext
   Loop
   Gl_Ac_Botones Me, 1, 1, modo
Else
   Gl_Ac_Botones Me, 1, 2, modo
End If
RS1.Close: Set RS1 = Nothing
Toolbar1.Buttons(5).Visible = False
Toolbar1.Buttons(6).Visible = True
vaSpread4.Visible = True
End Sub

Private Sub vaSpread4_EditChange(ByVal Col As Long, ByVal Row As Long)
If vaSpread4.MaxRows < 1 Then Exit Sub
vaSpread4.Row = Row
vaSpread4.Col = Col
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(1) = False:: SSTab1.TabEnabled(2) = False:
End Sub

Private Sub vaSpread4_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
    GrabaRegistro Row
ElseIf Toolbar1.Buttons(12).Visible = False Then
    Cancela
End If
End Sub
