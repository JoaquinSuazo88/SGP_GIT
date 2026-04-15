VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_TabGra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tabla Gramaje"
   ClientHeight    =   10560
   ClientLeft      =   3570
   ClientTop       =   525
   ClientWidth     =   16260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10560
   ScaleWidth      =   16260
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   10455
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   18441
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Tabla gramaje x Receta Standar"
      TabPicture(0)   =   "M_TabGra.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1(1)"
      Tab(0).Control(1)=   "Frame1(0)"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Tabla gramaje Nivel"
      TabPicture(1)   =   "M_TabGra.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2(5)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Image1(7)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fpayuda(8)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fpayuda(11)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "fpText(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Frame8"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.Frame Frame8 
         Height          =   8535
         Left            =   120
         TabIndex        =   56
         Top             =   1440
         Width           =   15135
         Begin VB.CommandButton Command1 
            Caption         =   "Eliminar Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   13320
            TabIndex        =   63
            Top             =   7920
            Width           =   1335
         End
         Begin VB.CommandButton Command2 
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
            Index           =   4
            Left            =   10380
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   1000
            Width           =   315
         End
         Begin VB.CommandButton Command2 
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
            Index           =   3
            Left            =   7290
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   1000
            Width           =   315
         End
         Begin VB.CommandButton Command2 
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
            Index           =   2
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   1000
            Width           =   315
         End
         Begin VB.CommandButton Command2 
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
            Index           =   1
            Left            =   1215
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   1000
            Width           =   315
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Agregar Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   0
            Left            =   11760
            TabIndex        =   58
            Top             =   7920
            Width           =   1335
         End
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   7215
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   14895
            _Version        =   393216
            _ExtentX        =   26273
            _ExtentY        =   12726
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
            MaxCols         =   11
            MaxRows         =   30
            SpreadDesigner  =   "M_TabGra.frx":0038
         End
      End
      Begin VB.Frame Frame1 
         Height          =   4455
         Index           =   1
         Left            =   -74880
         TabIndex        =   43
         Top             =   5760
         Width           =   14715
         Begin VB.Frame Frame4 
            Height          =   435
            Left            =   285
            TabIndex        =   48
            Top             =   3720
            Width           =   795
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   1
               Left            =   45
               TabIndex        =   49
               Top             =   135
               Width           =   690
            End
         End
         Begin VB.Frame Frame5 
            Height          =   435
            Left            =   1080
            TabIndex        =   46
            Top             =   3720
            Width           =   4245
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   47
               Top             =   135
               Width           =   4150
            End
         End
         Begin VB.Frame Frame7 
            Height          =   435
            Left            =   11880
            TabIndex        =   44
            Top             =   3720
            Width           =   2685
            Begin VB.TextBox Text1 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   13
               Left            =   45
               TabIndex        =   45
               Top             =   135
               Width           =   2595
            End
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   3375
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   14430
            _Version        =   393216
            _ExtentX        =   25453
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
            MaxCols         =   13
            SpreadDesigner  =   "M_TabGra.frx":082D
         End
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   165
            Left            =   5400
            TabIndex        =   51
            Top             =   3900
            Visible         =   0   'False
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   291
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5250
         Index           =   0
         Left            =   -74760
         TabIndex        =   2
         Top             =   480
         Width           =   13755
         Begin VB.Frame Frame2 
            Height          =   2070
            Left            =   6810
            TabIndex        =   17
            Top             =   3060
            Width           =   4950
            Begin MSComctlLib.TreeView TvwZon 
               Height          =   1695
               Index           =   1
               Left            =   120
               TabIndex        =   18
               Top             =   225
               Width           =   4710
               _ExtentX        =   8308
               _ExtentY        =   2990
               _Version        =   393217
               Style           =   7
               Checkboxes      =   -1  'True
               Appearance      =   1
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1455
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   13335
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
               TabIndex        =   8
               Top             =   240
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
               Index           =   1
               Left            =   6480
               TabIndex        =   7
               Top             =   240
               Width           =   2175
            End
            Begin EditLib.fpText fpText 
               Height          =   315
               Index           =   1
               Left            =   1785
               TabIndex        =   9
               Top             =   585
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
               Index           =   0
               Left            =   1800
               TabIndex        =   10
               Top             =   1065
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
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   3105
               TabIndex        =   14
               Top             =   600
               Width           =   5175
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   6
               Left            =   2655
               Picture         =   "M_TabGra.frx":23EE
               Top             =   480
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
               Index           =   4
               Left            =   240
               TabIndex        =   13
               Top             =   640
               Width           =   1380
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   3105
               TabIndex        =   12
               Top             =   1065
               Width           =   5175
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   2670
               Picture         =   "M_TabGra.frx":26F8
               Top             =   960
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
               Index           =   0
               Left            =   240
               TabIndex        =   11
               Top             =   1120
               Width           =   1245
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3120
               TabIndex        =   15
               Top             =   630
               Width           =   5205
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   99
               Left            =   3120
               TabIndex        =   16
               Top             =   1095
               Width           =   5205
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Filtro Receta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   360
            TabIndex        =   3
            Top             =   4440
            Width           =   6135
            Begin VB.OptionButton Option2 
               Caption         =   "Planificación"
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
               Left            =   480
               TabIndex        =   5
               Top             =   360
               Value           =   -1  'True
               Width           =   1455
            End
            Begin VB.OptionButton Option2 
               Caption         =   "Receta"
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
               Left            =   4200
               TabIndex        =   4
               Top             =   360
               Width           =   1575
            End
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   0
            Left            =   1890
            TabIndex        =   19
            Top             =   2460
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
            Left            =   1890
            TabIndex        =   20
            Top             =   1710
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
         Begin MSComctlLib.Toolbar Toolbar2 
            Height          =   360
            Left            =   11820
            TabIndex        =   21
            Top             =   3720
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Style           =   1
            _Version        =   393216
            BorderStyle     =   1
         End
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   2
            Left            =   1890
            TabIndex        =   22
            Top             =   2085
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
            Index           =   0
            Left            =   1890
            TabIndex        =   23
            Top             =   2925
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   1
            Left            =   4770
            TabIndex        =   24
            Top             =   2925
            Width           =   1425
            _Version        =   196608
            _ExtentX        =   2514
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
            Left            =   345
            TabIndex        =   37
            Top             =   2475
            Width           =   1020
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2760
            Picture         =   "M_TabGra.frx":2A02
            Top             =   2355
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   3210
            TabIndex        =   36
            Top             =   2460
            Width           =   5175
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   3210
            TabIndex        =   35
            Top             =   1710
            Width           =   5175
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   2760
            Picture         =   "M_TabGra.frx":2D0C
            Top             =   1605
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
            Index           =   2
            Left            =   330
            TabIndex        =   34
            Top             =   1725
            Width           =   750
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
            Left            =   345
            TabIndex        =   33
            Top             =   2100
            Width           =   705
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   2760
            Picture         =   "M_TabGra.frx":3016
            Top             =   1980
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   3210
            TabIndex        =   32
            Top             =   2070
            Width           =   5175
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
            Left            =   6840
            TabIndex        =   31
            Top             =   2895
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha desde (dd/mm/aa)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   2
            Left            =   315
            TabIndex        =   30
            Top             =   2910
            Width           =   1200
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   1410
            Picture         =   "M_TabGra.frx":3320
            Top             =   3465
            Width           =   480
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
            Left            =   330
            TabIndex        =   29
            Top             =   3620
            Width           =   1020
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   1935
            TabIndex        =   28
            Top             =   3585
            Width           =   4695
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   1935
            TabIndex        =   27
            Top             =   4005
            Width           =   4695
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   1410
            Picture         =   "M_TabGra.frx":362A
            Top             =   3885
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
            Left            =   330
            TabIndex        =   26
            Top             =   4070
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Fecha hasta (dd/mm/aa)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Index           =   0
            Left            =   3435
            TabIndex        =   25
            Top             =   2910
            Width           =   1200
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   5
            Left            =   3225
            TabIndex        =   39
            Top             =   1725
            Width           =   5205
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   98
            Left            =   3225
            TabIndex        =   38
            Top             =   2085
            Width           =   5190
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   7
            Left            =   3225
            TabIndex        =   40
            Top             =   2475
            Width           =   5205
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   1980
            TabIndex        =   41
            Top             =   3615
            Width           =   4710
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   1980
            TabIndex        =   42
            Top             =   4035
            Width           =   4710
         End
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   2
         Left            =   1920
         TabIndex        =   53
         Top             =   720
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
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   11
         Left            =   3360
         TabIndex        =   55
         Top             =   720
         Width           =   5175
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   8
         Left            =   3360
         TabIndex        =   54
         Top             =   720
         Width           =   5205
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   7
         Left            =   2880
         Picture         =   "M_TabGra.frx":3934
         Top             =   600
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
         Index           =   5
         Left            =   360
         TabIndex        =   52
         Top             =   720
         Width           =   1380
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   10560
      Left            =   15720
      TabIndex        =   0
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   18627
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
Dim NomFor As String, MsgTitulo As String
Dim FilTipPla As Long
Dim iblockcol As Long, iblockrow As Long, iblockcol2 As Long, iblockrow2 As Long, IRow As Long
Dim CodIng As String, noming As String, Est As Boolean, TmpCopiaGramaje As Double
Dim rootNode As Node, nd As Node
Dim FilCatDie As Long

Dim FilIni As Variant, FilFin As Variant, Colini As Variant, ColFin As Variant
Dim itop As Long
Dim estelilinea As Boolean
Dim ValGrilla As Boolean

Private Sub Check1_Click()

On Error GoTo Man_Error

If Check1.Value = 0 Then
   
   TvwZon(1).Enabled = False

Else
   
   TvwZon(1).Enabled = True

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command1_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim Regimen     As String
Dim TipoPlato   As String
Dim Ingrediente As String
Dim var_row     As Long

MsgTitulo = "Tabla Gramaje x Nivel"
        
If ValidarGrillaDatosRepetidos Then
                                 
    Exit Sub
       
End If

If ValidarGrillaNivel(IIf(Index = 1, 1, 0)) Then
       
   Exit Sub
       
End If

If fpayuda(11).Caption = "" Then

   MsgBox "Debe ingresar centro de costo...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

SSTab1.TabEnabled(0) = False

Select Case Index

    Case 0
        
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 10
        vaSpread2.text = "1"

        vaSpread2.Col = 11
        vaSpread2.text = "1"

        vaSpread2.Col = 1
        vaSpread2.SetActiveCell 1, Row
        vaSpread2.SetFocus

    Case 1
        
        var_row = vaSpread2.ActiveRow
        vaSpread2.Row = var_row
        
        vaSpread2.Col = 1
        Regimen = vaSpread2.text
        
        vaSpread2.Col = 3
        TipoPlato = vaSpread2.text
        
        vaSpread2.Col = 5
        Ingrediente = vaSpread2.text
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        Set RS = vg_db.Execute("sgpadm_Sel_TablaGramajeNivelRegistro_V01 '" & Trim(LimpiaDato(fpText(2).text)) & "', '" & Regimen & "', '" & TipoPlato & "','" & Ingrediente & "'")
        If Not RS.EOF Then
   
           If MsgBox("Esta Seguro desactivar fila ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
        
              RS.Close
              Set RS = Nothing
                
              estelilinea = False
              Exit Sub
        
            End If
            
            vaSpread2.Row = var_row
            vaSpread2.Col = 10
            vaSpread2.text = "0"
            
            vaSpread2.Col = 11
            vaSpread2.text = "1"

        Else
        
           vaSpread2.Row = vaSpread2.MaxRows
           vaSpread2.DeleteRows vaSpread2.Row, 1
           vaSpread2.MaxRows = vaSpread2.MaxRows - 1
           Command2(1).Visible = False
           Command2(2).Visible = False
           Command2(3).Visible = False
           Command2(4).Visible = False
           
        End If
        RS.Close
        Set RS = Nothing
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Command2_Click(Index As Integer)

On Error GoTo Man_Error

estelilinea = True

Select Case Index

    Case 1
    
        vg_left = Command2(2).Left + 3801
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
        B_TabEst.Show 1
        Me.Refresh
        
        With vaSpread2
            
            If vg_codigo = "" Then
            
              .Col = 1
              .Row = IRow
              .SetActiveCell 1, IRow
              .EditMode = True
              .EditModeReplace = True
              .SetFocus
              estelilinea = False
              
              Exit Sub
            
            End If
            
            .Row = IRow
            .Col = 1
            .Value = vg_codigo
            .Col = 2
            .Value = vg_nombre
            .Col = 11
            .Value = "1"
        
        End With
        estelilinea = False

'        If modo <> "A" Then modo = "M"
'        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.TabEnabled(0) = False
    
    Case 2
    
        vg_left = Command2(2).Left + 3801
        vg_nombre = "": vg_codigo = ""
        B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
        B_ArbEst.Show 1
        Me.Refresh
        
        With vaSpread2
            
            If vg_codigo = "" Then
            
              .Col = 3
              .Row = IRow
              .SetActiveCell 3, IRow
              .EditMode = True
              .EditModeReplace = True
              .SetFocus
              estelilinea = False
              
              Exit Sub
            
            End If
            
            .Row = IRow
            .Col = 3
            .Value = vg_codigo
            .Col = 4
            .Value = vg_nombre
            .Col = 11
            .Value = "1"
        
        End With
        
'        If modo <> "A" Then modo = "M"
'        Gl_Ac_Botones Me, 1, 0, modo
        estelilinea = False

        SSTab1.TabEnabled(0) = False

    Case 3
    
        vg_left = Command2(3).Left + 3801
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente", "AgregarIngxReceta"
        B_TabEst.Show 1
        Me.Refresh
        
        With vaSpread2
            
            If vg_codigo = "" Then
            
              .Col = 5
              .Row = IRow
              .SetActiveCell 5, IRow
              .EditMode = True
              .EditModeReplace = True
              .SetFocus
              estelilinea = False

              Exit Sub
            
            End If
            
            .Row = IRow
            .Col = 5
            .Value = vg_codigo
            .Col = 6
            .Value = vg_nombre
            estelilinea = False
            .Col = 11
            .Value = "1"

        End With
        
'        If modo <> "A" Then modo = "M"
'        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.TabEnabled(0) = False

    Case 4
    
        vg_left = Command2(4).Left + 3801
        vg_nombre = "": vg_codigo = ""
        B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
        B_TabEst.Show 1
        Me.Refresh
        
        With vaSpread2
            
            If vg_codigo = "" Then
            
              .Col = 7
              .Row = IRow
              .SetActiveCell 7, IRow
              .EditMode = True
              .EditModeReplace = True
              .SetFocus
              estelilinea = False

              Exit Sub
            
            End If
            
            .Row = IRow
            .Col = 7
            .Value = vg_codigo
            .Col = 8
            .Value = vg_nombre
            estelilinea = False
            .Col = 11
            .Value = "1"

        End With
        
'        If modo <> "A" Then modo = "M"
'        Gl_Ac_Botones Me, 1, 0, modo
        SSTab1.TabEnabled(0) = False

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
fg_carga ""
Me.HelpContextID = vg_OpcM
Me.Height = 10995
Me.Width = 16350

MsgTitulo = "Tabla Gramaje"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Toolbar2.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar ": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Borrar ", , tbrDefault, "A_Borrar "): BtnX.Visible = True: BtnX.ToolTipText = "": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 3, 1) = "0", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): BtnX.Visible = True: BtnX.ToolTipText = "Deshacer"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Historico", , tbrDefault, "A_Historico"): BtnX.Visible = True: BtnX.ToolTipText = "Historico Tabla Gramaje"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Imprimir ", , tbrDefault, "A_Imprimir "): BtnX.Visible = True: BtnX.ToolTipText = "Imprimir": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "0", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Tabla Gramaje"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Excel log tabla gramaje nivel": BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
'Set btnX = Toolbar2.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set BtnX = Toolbar2.Buttons.Add(, "Proceso", , tbrDefault, "Proceso"): BtnX.Visible = True: BtnX.ToolTipText = "Proceso"

Toolbar1.Buttons(14).Enabled = False

vaSpread1.MaxRows = 0
vaSpread1.Row = -1: vaSpread1.Col = -1
vaSpread1.BackColor = &H80000018
vaSpread2.MaxRows = 0

Command2(1).Visible = False
Command2(2).Visible = False
Command2(3).Visible = False
Command2(4).Visible = False

itop = 1

fpayuda(2).Caption = "Todos": fpayuda(3).Caption = "Todos"
FilCatDie = 0
iayuda = 0
Est = True

'-------> Cargar zona
TvwZon(1).Nodes.Clear
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_zona 8, 0,''")
Do While Not RS.EOF
   
   Set rootNode = TvwZon(1).Nodes.Add(, , "H" & RS!zon_codigo, Trim(RS!Zon_nombre))
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing

'-------> activar centro de costo
fpText(1).text = ""
fpayuda(9).Caption = ""
fpText(1).Enabled = True
Image1(6).Enabled = True
'-------> desactivar sub-segmento
fpLongInteger1(0).Value = ""
fpayuda(0).Caption = ""
fpLongInteger1(0).Enabled = False
Image1(0).Enabled = False
'-------> desactiva concepto zona
Label3(0).Visible = False
Frame2.Enabled = False
Frame2.Visible = False
TvwZon(1).Visible = False
vaSpread1.Col = 7
vaSpread1.ColHidden = True

fpDateTime1(0).Value = " "
fpDateTime1(1).Value = " "
Est = False

SSTab1.Tab = 0
SSTab1.TabVisible(1) = False

Me.HelpContextID = 1121000
SSTab1.TabVisible(1) = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
Me.HelpContextID = vg_OpcM

fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

'If IsDate(fpDateTime1.text) = False Then Exit Sub
If Est Then Exit Sub
vg_fecha = Format(fpDateTime1(0).text, "yyyymmdd")
vg_fecha = Format(fpDateTime1(1).text, "yyyymmdd")

'DataLoad

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
If Est Then Exit Sub

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

Case 0
    
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & " AND sub_indppr = '" & vg_Indppr & "'")
    
    Else
      
      Set RS = vg_db.Execute("SELECT * FROM a_subsegmento WHERE sub_codigo = " & Val(fpLongInteger1(0).Value) & "")
    
    End If

    
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close
    Set RS = Nothing

Case 1
    
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & " AND reg_indppr = '" & vg_Indppr & "'")
    
    Else
      
      Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(fpLongInteger1(1).Value) & "")
    
    End If
    
    If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: fpayuda(4).Caption = "": Exit Sub
    fpayuda(4).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    If Est Then Exit Sub
'    DataLoad

Case 2
    
    If Val(fpLongInteger1(2).Value) < 1 Then fpayuda(6).Caption = "": Exit Sub
    
    If vg_Indppr = 1 Or vg_Indppr = 2 Then
      
      Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & " and ser_indppr='" & vg_Indppr & "'")
    
    Else
      
      Set RS = vg_db.Execute("SELECT * FROM a_servicio WHERE ser_codigo=" & Val(fpLongInteger1(2).Value) & "")
    
    End If
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(6).Caption = "":  Exit Sub
    fpayuda(6).Caption = Trim(RS!ser_nombre)
    RS.Close
    Set RS = Nothing
'    DataLoad
   
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
    
    RS.Open "SELECT DISTINCT a.ing_codigo, a.ing_nombre FROM b_ingrediente a WITH (NOLOCK), b_receta b WITH (NOLOCK), b_recetadet c WITH (NOLOCK) WHERE b.rec_codigo = c.red_codigo AND c.red_codpro = a.ing_codigo AND (b.rec_catdie = " & FilCatDie & " OR " & FilCatDie & " = 0) AND (b.rec_tippla = " & FilTipPla & " OR " & FilTipPla & " = 0) AND a.ing_codigo = '" & Trim(fpText(0).text) & "' AND (a.ing_Indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' <> '1')", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": vaSpread1.MaxRows = 0: Exit Sub
    fpayuda(1).Caption = Trim(RS!ing_nombre)
    RS.Close
    Set RS = Nothing
    If Me.fpText(0).text <> 0 Then
    '    Me.Toolbar2.Enabled = True
    Else
    '    Me.Toolbar2.Enabled = False
    End If

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

Case 2
   
   Sql = Trim(LimpiaDato(fpText(2).text))
   Set RS = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
   If RS.EOF Then
        
        fpayuda(11).Caption = ""
        RS.Close
        Set RS = Nothing
        Exit Sub
    
    End If
    fpayuda(11).Caption = Trim(RS!Cli_nombre)
    
    MoverDetalleTablaNivel
    
'    SSTab1.TabEnabled(0) = False

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MoverDetalleTablaNivel()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim i As Long

vaSpread2.MaxRows = 0

Command2(1).Visible = False
Command2(2).Visible = False
Command2(3).Visible = False
Command2(4).Visible = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
    
fg_carga ""
    
Set RS = vg_db.Execute("sgpadm_Sel_DetalleTablaGramajeNivel_V01 '" & Trim(LimpiaDato(fpText(2).text)) & "'")

If Not RS.EOF Then

    Do While Not RS.EOF
        
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 1
        vaSpread2.text = IIf(RS!Reg_Codigo = 0, "", RS!Reg_Codigo)
        
        vaSpread2.Col = 2
        vaSpread2.text = IIf(RS!reg_nombre = "", "", RS!reg_nombre)
        
        vaSpread2.Col = 3
        vaSpread2.text = IIf(RS!tip_codigo = 0, "", RS!tip_codigo)
        
        vaSpread2.Col = 4
        vaSpread2.text = IIf(RS!TipoPlato = "", "", RS!TipoPlato)
        
        vaSpread2.Col = 5
        vaSpread2.text = IIf(RS!ing_codigo = "", "", RS!ing_codigo)
        
        vaSpread2.Col = 6
        vaSpread2.text = IIf(RS!ing_nombre = "", "", RS!ing_nombre)
        
        vaSpread2.Col = 7
        vaSpread2.text = IIf(RS!ing_cambio = "", "", RS!ing_cambio)
        
        vaSpread2.Col = 8
        vaSpread2.text = IIf(RS!ing_nomcambio = "", "", RS!ing_nomcambio)
        
        vaSpread2.Col = 9
        vaSpread2.text = IIf(IsNull(RS!CantidadBruta), "", RS!CantidadBruta)
        
        vaSpread2.Col = 10
        vaSpread2.text = IIf(IsNull(RS!Activo), "", RS!Activo)
        
        vaSpread2.Col = 11
        vaSpread2.text = "0"
        
        RS.MoveNext
        
    Loop

    Me.HelpContextID = 1121000
    Toolbar1.Buttons(4).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
    Me.HelpContextID = vg_OpcM

Else

    MsgBox "No existe información para este centro de costo", vbCritical + vbOKOnly, MsgTitulo

    Toolbar1.Buttons(4).Enabled = False
    
End If

fg_descarga

RS.Close
Set RS = Nothing

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
    
    vg_left = fpayuda(0).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_subsegmento", "sub_", "Subsegmento", "Sub"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpText(0).text = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).Value = "": fpayuda(4).Caption = ""
    fpLongInteger1(1).SetFocus

Case 1
    
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = "": vg_filtippla = FilTipPla: vg_filcatdie = FilCatDie
'   B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente", "Ingrec"
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "INgrediente", "AgregarIng"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    fpDateTime1(0).SetFocus

Case 2
    
    vg_left = fpayuda(2).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica", 1
    B_ArbEst.Show 1
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    FilCatDie = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre: vg_nombre = ""
    'DataLoad

Case 3
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(3).Left + 3000
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato", 1
    B_ArbEst.Show 1
    Screen.MousePointer = 0
    If Trim(vg_codigo) = "" Then Exit Sub
    FilTipPla = Val(vg_codigo)
    fpayuda(3).Caption = vg_nombre: vg_nombre = ""
    'DataLoad

Case 4
    
    vg_left = fpayuda(4).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    Est = False
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre
    fpLongInteger1(2).SetFocus
 
 Case 5
    
    vg_left = fpayuda(6).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_servicio", "ser_", "Servicio", "Ser", FilCatDie
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(6).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
 
 Case 6
    
    vg_left = fpayuda(9).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo: fpayuda(9).Caption = vg_nombre
    fpLongInteger1(1).SetFocus
 
 Case 7
    
    vg_left = fpayuda(11).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText(2).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
    vaSpread2.SetFocus

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
    
    '-------> activar centro de costo
    fpText(1).text = ""
    fpayuda(9).Caption = ""
    fpText(1).Enabled = True
    Image1(6).Enabled = True
    '-------> desactivar sub-segmento
    fpLongInteger1(0).Value = ""
    fpayuda(0).Caption = ""
    fpLongInteger1(0).Enabled = False
    Image1(0).Enabled = False
    '-------> desactiva concepto zona
    Label3(0).Visible = False
    Frame2.Enabled = False
    Frame2.Visible = False
    TvwZon(1).Visible = False
    vaSpread1.Col = 7
    vaSpread1.ColHidden = True

Case 1
    
    '-------> activar sub-segmento
    fpLongInteger1(0).Value = ""
    fpayuda(0).Caption = ""
    fpLongInteger1(0).Enabled = True
    Image1(0).Enabled = True
    '-------> activa zona
    Frame2.Enabled = True
    '-------> desactivar centro de costo
    fpText(1).text = ""
    fpayuda(9).Caption = ""
    fpText(1).Enabled = False
    Image1(6).Enabled = False
    '-------> activa concepto zona
    Label3(0).Visible = True
    Frame2.Enabled = True
    Frame2.Visible = True
    TvwZon(1).Visible = True
    vaSpread1.Col = 7
    vaSpread1.ColHidden = False

End Select

vaSpread1.MaxRows = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error GoTo Man_Error

itop = 1

Select Case SSTab1.Tab

Case 0

    
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(14).Enabled = False
    
    MsgTitulo = "Tabla Gramaje Ceco"

Case 1
    
    MsgTitulo = "Tabla Gramaje x Nivel"
    
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(14).Enabled = True

    If fpText(1).text <> "" Then
    
       vaSpread2.MaxRows = 0
       
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub SSTab1_DblClick()

On Error GoTo Man_Error

itop = 1

Select Case SSTab1.Tab

Case 0

    
    Toolbar1.Buttons(8).Enabled = True
    Toolbar1.Buttons(14).Enabled = False
    
    MsgTitulo = "Tabla Gramaje Ceco"

Case 1
    
    MsgTitulo = "Tabla Gramaje x Nivel"
    
    Toolbar1.Buttons(8).Enabled = False
    Toolbar1.Buttons(14).Enabled = True

    If fpText(1).text <> "" Then
    
       vaSpread2.MaxRows = 0
       
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Text1_Change(Index As Integer)

On Error GoTo Man_Error

Dim Col As Integer
Col = 0

Select Case Index
    
    Case 1, 2, 13
        
        Col = IIf(Index = 1, 4, IIf(Index = 13, 13, 5))
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
'        vaSpread1.SortKey(1) = IIf(Trim(Text1(Index).text) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
'        vaSpread1.Sort -1, -1, vaSpread1.MaxCols, vaSpread1.MaxRows, SortByRow
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
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim RS As New ADODB.Recordset
Dim cantrg  As Double
Dim isel As Long
Dim CodRec As Long
Dim codzon As Long
Dim CodIng As String
Dim MyBuffer As String
Dim Sql As String

Dim nivel_regimen           As String
Dim nivel_tipoplato         As String
Dim nivel_ingredienteorigen As String
Dim nivel_ingredientecambio As String
Dim nivel_cantidadcambio    As String
Dim nivel_activo            As String

On Error GoTo Man_Error

Select Case Button.Index

Case 2, 4
    
    If SSTab1.Tab = 0 Then
        
        If vaSpread1.MaxRows < 1 Then Exit Sub
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS.Open "SELECT ing_nombre FROM b_ingrediente WITH ( NOLOCK ) WHERE ing_codigo = '" & LimpiaDato(Trim(fpText(0).text)) & "'", vg_db, adOpenStatic
        If RS.EOF Then RS.Close: Set RS = Nothing: vaSpread1.MaxRows = 0: MsgBox "No existe ingrediente recetas", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        RS.Close
        Set RS = Nothing
        isel = 0
        
        For i = 1 To vaSpread1.MaxRows
            
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" Then
            
               isel = 1
               Exit For
        
            End If
            
        Next i
        
        If isel = 0 Then MsgBox "Seleccione Uno o Más Recetas " & IIf(Button.Index = 2, "a Reemplazar", "a Borrar"), vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        If MsgBox("Esta Seguro " & IIf(Button.Index = 2, "a Reemplazar", "a Borrar") & " ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        fg_carga ""
        Bar1.Visible = True
        Bar1.Value = 0
    
        Let MyBuffer = ""
        Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
        
        If Button.Index = 2 Then
           
           Let MyBuffer = MyBuffer & "<GrabaTablaGramajes>"
        
        ElseIf Button.Index = 4 Then
           
           Let MyBuffer = MyBuffer & "<BorraTablaGramajes>"
     
        End If
        
        For i = 1 To vaSpread1.MaxRows
            
            Bar1.Value = Val((i / vaSpread1.MaxRows) * 100)
            vaSpread1.Row = i
            vaSpread1.Col = 1
            
            If vaSpread1.text = "1" Then
               
               Me.Refresh
               DoEvents
               
               vaSpread1.Col = 4
               CodRec = 0
               CodRec = Val(vaSpread1.text)
               
               vaSpread1.Col = 6
               codzon = 0
               codzon = vaSpread1.text
               
               vaSpread1.Col = 8
               CodIng = ""
               CodIng = IIf(Trim(vaSpread1.text) = "", LimpiaDato(Trim(fpText(0).text)), LimpiaDato(Trim(vaSpread1.text)))
               
               vaSpread1.Col = 12
               codReg = 0
               codReg = vaSpread1.text
               
               vaSpread1.Col = 10
               cantrg = 0
               cantrg = vaSpread1.text
    
               MyBuffer = MyBuffer & " <TablaGramaje"
               MyBuffer = MyBuffer & " CodReceta = " & Chr(34) & CodRec & Chr(34)
               MyBuffer = MyBuffer & " CodZona = " & Chr(34) & codzon & Chr(34)
               MyBuffer = MyBuffer & " CodIngrediente = " & Chr(34) & CodIng & Chr(34)
               MyBuffer = MyBuffer & " CodRegimen = " & Chr(34) & codReg & Chr(34)
               MyBuffer = MyBuffer & " CantGramaje = " & Chr(34) & cantrg & Chr(34)
               Let MyBuffer = MyBuffer & "/>"
            
            End If
        
        Next i
        
        If Button.Index = 2 Then
           
           Let MyBuffer = MyBuffer & "</GrabaTablaGramajes>"
        
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Grabar"), Me.HelpContextID, "", "", "")
        
        ElseIf Button.Index = 4 Then
           
           Let MyBuffer = MyBuffer & "</BorraTablaGramajes>"
        
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), Me.HelpContextID, "", "", "")
        
        End If
        
        Sql = ""
        If Button.Index = 2 Then
           
           Sql = Sql & "sgpadm_Ins_XmlTablaGramaje "
        
        ElseIf Button.Index = 4 Then
           
           Sql = Sql & "sgpadm_Del_XmlTablaGramaje "
        
        End If
        
        If Option1(0).Value = True Then
           
           Sql = Sql & "'" & LimpiaDato(Trim(fpText(1).text)) & "',"
        
        ElseIf Option1(1).Value = True Then
           
           Sql = Sql & "'" & fpLongInteger1(0).Value & "',"
        
        End If
        
        Sql = Sql & " " & "'" & LimpiaDato(Trim(fpText(0).text)) & "',"
        
        If Option1(0).Value = True Then
           
           Sql = Sql & " " & "'2'"
        
        ElseIf Option1(1).Value = True Then
           
           Sql = Sql & "'1'"
        
        End If
        Sql = Sql
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("" & Sql & ", '" & MyBuffer & "'")
        If Not RS.EOF Then
           
           If RS(0) > 0 Then
              
              If Button.Index = 2 Then
           
                 Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), Me.HelpContextID, "", "", "")
        
              ElseIf Button.Index = 4 Then
           
                 Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
        
              End If
              
              MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
           End If
        
        End If
        RS.Close
        Set RS = Nothing
        
        Bar1.Visible = False
        fg_descarga
        
        MsgBox "Proceso finalizo sin problema", vbInformation + vbOKOnly, MsgTitulo
        DataLoad
        indsel = 0
        SSTab1.TabEnabled(1) = True
        
    ElseIf SSTab1.Tab = 1 Then
    
      MsgTitulo = "Tabla Gramaje x Nivel"
        
        Command2(1).Visible = False
        Command2(2).Visible = False
        Command2(3).Visible = False
        Command2(4).Visible = False
       
       If ValidarGrillaDatosRepetidos Then
          
          Exit Sub
       
       End If
       
       If ValidarGrillaNivel(0) Then
       
          Exit Sub
       
       End If
       
       If Button.Index = 2 Then
        
            'validar centro costo
            If fpayuda(11).Caption = "" Then
        
                MsgBox "No esta definido centro de costo...", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub
            
            End If
        
            'validar que tenga datos detalle grilla
            If vaSpread2.MaxRows < 1 Then
        
                MsgBox "No existe datos ingresado en el detalle de la grilla...", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub
            
            End If
        
            isel = 0
            For i = 1 To vaSpread2.MaxRows
            
                vaSpread2.Row = i
                vaSpread2.Col = 11
            
                If vaSpread2.text = "1" Then
            
                    isel = 1
                    Exit For
        
                End If
            
            Next i
        
            If isel = 0 Then
        
                MsgBox "No existen datos que Agregar o bien Actualizar...", vbCritical + vbOKOnly, MsgTitulo
                Exit Sub
        
            End If
        
            If MsgBox("Esta Seguro de Insertar o bien Actualizar ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
        
                Exit Sub
        
            End If
        
            'validar que existan datos que insertar y modificar
            Let MyBuffer = ""
            Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
            Let MyBuffer = MyBuffer & "<GrabaTablaGramajes>"
            Bar1.Visible = True
            Bar1.Value = 0
        
            For i = 1 To vaSpread2.MaxRows
        
                vaSpread2.Row = i
                vaSpread2.Col = 11
            
                Bar1.Value = Val((i / vaSpread2.MaxRows) * 100)
            
                If vaSpread2.text = "1" Then
            
                    Me.Refresh
                    DoEvents
               
                    vaSpread2.Col = 1
                    nivel_regimen = ""
                    nivel_regimen = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
               
                    vaSpread2.Col = 3
                    nivel_tipoplato = ""
                    nivel_tipoplato = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)

                    vaSpread2.Col = 5
                    nivel_ingredienteorigen = ""
                    nivel_ingredienteorigen = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
               
                    vaSpread2.Col = 7
                    nivel_ingredientecambio = ""
                    nivel_ingredientecambio = IIf(Trim(vaSpread2.text) = "", nivel_ingredienteorigen, vaSpread2.text)
    
                    vaSpread2.Col = 9
                    nivel_cantidadcambio = ""
                    nivel_cantidadcambio = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
               
                    vaSpread2.Col = 10
                    nivel_activo = ""
                    nivel_activo = IIf(Trim(vaSpread2.text) = "", "0", vaSpread2.text)
               
                    MyBuffer = MyBuffer & " <TablaGramaje"
                    MyBuffer = MyBuffer & " IR = " & Chr(34) & nivel_regimen & Chr(34)
                    MyBuffer = MyBuffer & " ITP = " & Chr(34) & nivel_tipoplato & Chr(34)
                    MyBuffer = MyBuffer & " CIO = " & Chr(34) & nivel_ingredienteorigen & Chr(34)
                    MyBuffer = MyBuffer & " CIC = " & Chr(34) & nivel_ingredientecambio & Chr(34)
                    MyBuffer = MyBuffer & " CT = " & Chr(34) & nivel_cantidadcambio & Chr(34)
                    MyBuffer = MyBuffer & " A = " & Chr(34) & nivel_activo & Chr(34)
                    Let MyBuffer = MyBuffer & "/>"
            
                End If
            
        
            Next i
               
            Let MyBuffer = MyBuffer & "</GrabaTablaGramajes>"
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Grabar"), Me.HelpContextID, "", "", "")
        
            Sql = ""
            Sql = Sql & "sgpadm_InsOrUpd_XmlTablaGramajeNivel "
            Sql = Sql & "'" & LimpiaDato(Trim(fpText(2).text)) & "',"
            Sql = Sql & "'" & vg_NUsr & "' "
        
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
        
            Set RS = vg_db.Execute("" & Sql & ", '" & MyBuffer & "'")
            If Not RS.EOF Then
           
                If RS(0) > 0 Then
                        
                    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Agregado"), Me.HelpContextID, "", "", "")
                      
                    MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
                End If
        
            End If
            RS.Close
            Set RS = Nothing
        
            Bar1.Visible = False
            fg_descarga
        
            MsgBox "Proceso finalizo sin problema", vbInformation + vbOKOnly, MsgTitulo
            MoverDetalleTablaNivel

            indsel = 0
            SSTab1.TabEnabled(0) = True
        
       ElseIf Button.Index = 4 Then
       
       
            Command2(1).Visible = False
            Command2(2).Visible = False
            Command2(3).Visible = False
            Command2(4).Visible = False
            
            'validar centro costo
            If fpayuda(11).Caption = "" Then
        
                MsgBox "No esta definido centro de costo...", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub
            
            End If
        
            'validar que tenga datos detalle grilla
            If vaSpread2.MaxRows < 1 Then
        
                MsgBox "No existe datos ingresado en el detalle de la grilla...", vbExclamation + vbOKOnly, MsgTitulo
                Exit Sub
            
            End If
        
            
            If MsgBox("Esta Seguro de eliminar registro completo ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
        
                Exit Sub
                SSTab1.TabEnabled(0) = True
        
            End If
           
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Eliminar"), Me.HelpContextID, "", "", "")
            
            'validar que existan datos que insertar y modificar
            Let MyBuffer = ""
            Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
            Let MyBuffer = MyBuffer & "<GrabaTablaGramajes>"
            Bar1.Visible = True
            Bar1.Value = 0
        
            For i = 1 To vaSpread2.MaxRows
        
                vaSpread2.Row = i
                vaSpread2.Col = 11
            
                Bar1.Value = Val((i / vaSpread2.MaxRows) * 100)
            
'                If vaSpread2.text = "1" Then
            
                    Me.Refresh
                    DoEvents
               
                    vaSpread2.Col = 1
                    nivel_regimen = ""
                    nivel_regimen = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
               
                    vaSpread2.Col = 3
                    nivel_tipoplato = ""
                    nivel_tipoplato = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)

                    vaSpread2.Col = 5
                    nivel_ingredienteorigen = ""
                    nivel_ingredienteorigen = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
               
                    vaSpread2.Col = 7
                    nivel_ingredientecambio = ""
                    nivel_ingredientecambio = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
    
                    vaSpread2.Col = 9
                    nivel_cantidadcambio = ""
                    nivel_cantidadcambio = IIf(Trim(vaSpread2.text) = "", "", vaSpread2.text)
               
                    vaSpread2.Col = 10
                    nivel_activo = ""
                    nivel_activo = "0"
               
                    MyBuffer = MyBuffer & " <TablaGramaje"
                    MyBuffer = MyBuffer & " IR = " & Chr(34) & nivel_regimen & Chr(34)
                    MyBuffer = MyBuffer & " ITP = " & Chr(34) & nivel_tipoplato & Chr(34)
                    MyBuffer = MyBuffer & " CIO = " & Chr(34) & nivel_ingredienteorigen & Chr(34)
                    MyBuffer = MyBuffer & " CIC = " & Chr(34) & nivel_ingredientecambio & Chr(34)
                    MyBuffer = MyBuffer & " CT = " & Chr(34) & nivel_cantidadcambio & Chr(34)
                    MyBuffer = MyBuffer & " A = " & Chr(34) & nivel_activo & Chr(34)
                    Let MyBuffer = MyBuffer & "/>"
            
'                End If
            
        
            Next i
               
            Let MyBuffer = MyBuffer & "</GrabaTablaGramajes>"
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Grabar"), Me.HelpContextID, "", "", "")
        
            Sql = ""
            Sql = Sql & "sgpadm_InsOrUpd_XmlTablaGramajeNivel "
            Sql = Sql & "'" & LimpiaDato(Trim(fpText(2).text)) & "',"
            Sql = Sql & "'" & vg_NUsr & "' "
        
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
        
            Set RS = vg_db.Execute("" & Sql & ", '" & MyBuffer & "'")
            If Not RS.EOF Then
           
                If RS(0) > 0 Then
                        
                   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Eliminacion"), Me.HelpContextID, "", "", "")
                      
                    MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
           
                End If
        
            End If
            RS.Close
            Set RS = Nothing
        
            Bar1.Visible = False
            fg_descarga
        
            MsgBox "Proceso finalizo sin problema", vbInformation + vbOKOnly, MsgTitulo
            MoverDetalleTablaNivel
            indsel = 0
            SSTab1.TabEnabled(0) = True
            
       End If
       
    End If
    
Case 6 'deshacer
    
    If SSTab1.Tab = 0 Then
    
        SSTab1.TabEnabled(1) = True
        
        FilCatDie = 0
        FilTipPla = 0
        fpayuda(2).Caption = "Todos"
        fpayuda(3).Caption = "Todos"
        fpText(1).text = ""
        fpayuda(9).Caption = ""
        fpLongInteger1(0).text = ""
        fpayuda(0).Caption = ""
        fpLongInteger1(1).text = ""
        fpayuda(4).Caption = ""
        fpLongInteger1(2).text = ""
        fpayuda(6).Caption = ""
        fpText(0).text = ""
        fpayuda(1).Caption = ""
        fpDateTime1(0).text = ""
        fpDateTime1(1).text = ""
    
        For j = 1 To TvwZon(1).Nodes.count
            
            TvwZon(1).Nodes.item(j).Checked = False
    
        Next j
        vaSpread1.MaxRows = 0
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Deshacer"), Me.HelpContextID, "", "", "")

    ElseIf SSTab1.Tab = 1 Then
    
        Command2(1).Visible = False
        Command2(2).Visible = False
        Command2(3).Visible = False
        Command2(4).Visible = False

        'validar que tenga datos detalle grilla
        If vaSpread2.MaxRows < 1 Then
        
           MsgBox "No existe datos ingresado en el detalle de la grilla...", vbExclamation + vbOKOnly, MsgTitulo
           Exit Sub
            
        End If
        
        If MsgBox("Esta Seguro desahacer los cambios ?", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
        
           Exit Sub
        
        End If
       
       SSTab1.TabEnabled(0) = True
       MoverDetalleTablaNivel

    End If
    
Case 8 '-------> Historico tabla gramaje
    
    vg_codigo = ""
    vg_codregimen = 0
    If Option1(0).Value = True Then
        
        B_HistPm.LlenarHistPlan "Histórico Tabla Gramaje Ceco", fpText(1).text, 1, 4
    
    ElseIf Option1(1).Value = True Then
        
        B_HistPm.LlenarHistPlan "Histórico Tabla Gramaje Sub-Segmento", Val(fpLongInteger1(0).Value), 1, 2
    
    End If
    
    B_HistPm.Show 1
    If vg_codigo = "" Then Exit Sub
    Est = True
    fpLongInteger1(1).Value = ""
    Est = False
    fpLongInteger1(1).Value = vg_codregimen
    fpText(0).text = vg_codigo
    vaSpread1.MaxRows = 0

Case 10 'Imprimir
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), Me.HelpContextID, "", "", "")
    I_TabGra.Show 1

Case 12
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso_Copiar"), Me.HelpContextID, "", "", "")
    M_CpTabGra.Show 1, Me

Case 14 'emitir log excel
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Imprimir"), Me.HelpContextID, "", "", "")
    E_LogTablaNivel.Show 1, Me
    
Case 16 'Salir
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")
  
    Me.Hide
    Unload Me

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

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim tiene As Integer

   tiene = 0
    
   For j = 1 To TvwZon(1).Nodes.count
        
        If TvwZon(1).Nodes.item(j).Checked = True Then
            
            tiene = 1
        
        End If
   
   Next j
   
   If tiene = 0 And Option1(1).Value = True Then
      
      MsgBox "debe seleccionar Zona que son obligatorios", vbCritical + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   
   If (Trim(fpText(1).text) = "" Or Trim(fpayuda(9).Caption) = "") And Option1(0).Value = True Then
      
      MsgBox "debe ingresar a lo menos Centro costo", vbCritical + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   
   If (Val(fpLongInteger1(0).text) < 1 Or Trim(fpayuda(0).Caption) = "") And Option1(1).Value = True Then
      
      MsgBox "debe ingresar a lo menos Sub-segmento", vbCritical + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   
   If Val(fpLongInteger1(1).text) < 1 Then
      
      MsgBox "debe ingresar a lo menos Regimen", vbCritical + vbOKOnly, MsgTitulo
      Exit Sub
   
   End If
   
   If Trim(fpText(0).text) = "" Then
        
        MsgBox "Debe ingresar Ingrediente es obligatorio", vbCritical + vbOKOnly, MsgTitulo
        Exit Sub
   
   End If
   
   DataLoad

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

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

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

iblockcol = Col
iblockrow = Row
iblockcol2 = Col
iblockrow2 = Row

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)


On Error GoTo Man_Error

If Row < 1 Or vaSpread1.MaxRows < 1 Then Exit Sub
Dim CodRec As Long, CodIng As String, noming As String, i As Long, j As Long, codzon As Long
Select Case Col

Case 8
    
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = "": CodIng = "": noming = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    vaSpread1.Row = Row
    
    vaSpread1.Col = 8
    vaSpread1.text = Trim(vg_codigo)
    CodIng = Trim(vg_codigo)
    
    vaSpread1.Col = 9
    vaSpread1.text = Trim(vg_nombre)
    noming = Trim(vg_nombre)
    vaSpread1.Col = 1
    vaSpread1.text = "1"
    
    vaSpread1.Col = 4
    CodRec = vaSpread1.text
'    If Check1.Value = 0 Then Exit Sub
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 4
        If vaSpread1.text = CodRec And i <> Row Then
           
           vaSpread1.Col = 6
           codzon = vaSpread1.text
           
           For j = 1 To TvwZon(1).Nodes.count
               
               If TvwZon(1).Nodes.item(j).Checked = True And codzon = Mid(TvwZon(1).Nodes.item(j).key, 2, Len(TvwZon(1).Nodes.item(j).key)) Then
                  
                  vaSpread1.Col = 1
                  vaSpread1.text = "1"
                  vaSpread1.Col = 8
                  vaSpread1.text = CodIng
                  vaSpread1.Col = 9
                  vaSpread1.text = noming
                  Exit For
               
               End If
           
           Next j
        
        End If
    
    Next i
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

If Row < 1 Then Exit Sub
Dim RS As New ADODB.Recordset
Dim CodRec As Long, CodIng As String, noming As String, canbru As Double, i As Long, j As Long, codzon As Long
vaSpread1.Row = Row

If ChangeMade = True Then
   
   SSTab1.TabEnabled(1) = False

   Select Case Col
   
   Case 8
       
       vaSpread1.Col = 8
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open "SELECT DISTINCT ing_codigo, ing_nombre FROM b_ingrediente WHERE ing_codigo = '" & LimpiaDato(Trim(vaSpread1.text)) & "' AND (ing_Indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' <> '1')", vg_db, adOpenStatic
       vaSpread1.Col = 9
       If Not RS.EOF Then
          
          vaSpread1.text = Trim(RS!ing_nombre)
          CodIng = Trim(RS!ing_codigo)
          noming = Trim(RS!ing_nombre)
       
       Else
          
          vaSpread1.text = ""
          vaSpread1.Col = 8
          vaSpread1.text = ""
       
       End If
       RS.Close
       Set RS = Nothing
       
       vaSpread1.Col = 1
       vaSpread1.text = "1"
       vaSpread1.Col = 4
       CodRec = vaSpread1.text

'       If Check1.Value = 0 Then Exit Sub
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = 4
           
           If vaSpread1.text = CodRec And i <> Row Then
              
              vaSpread1.Col = 6
              codzon = vaSpread1.text
              
              For j = 1 To TvwZon(1).Nodes.count
                  
                  If TvwZon(1).Nodes.item(j).Checked = True And codzon = Mid(TvwZon(1).Nodes.item(j).key, 2, Len(TvwZon(1).Nodes.item(j).key)) Then
                     
                     vaSpread1.Col = 1
                     vaSpread1.text = "1"
                     
                     vaSpread1.Col = 8
                     vaSpread1.text = CodIng
                     
                     vaSpread1.Col = 9
                     vaSpread1.text = noming
                     
                     Exit For
                  
                  End If
              
              Next j
           
           End If
       
       Next i
   
   Case 10
       
       vaSpread1.Col = 10
       vaSpread1.ForeColor = &HFF0000
       canbru = 0
       canbru = vaSpread1.text
       vaSpread1.Col = 1
       vaSpread1.text = "1"
       
       vaSpread1.Col = 4
       CodRec = vaSpread1.text
'       If Check1.Value = 0 Then Exit Sub
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = 4
           
           If vaSpread1.text = CodRec And i <> Row Then
              
              vaSpread1.Col = 6
              codzon = vaSpread1.text
              
              For j = 1 To TvwZon(1).Nodes.count
                  
                  If TvwZon(1).Nodes.item(j).Checked = True And codzon = Mid(TvwZon(1).Nodes.item(j).key, 2, Len(TvwZon(1).Nodes.item(j).key)) Then
                     
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

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

On Error GoTo Man_Error

Select Case Button

Case 2
    
    If vaSpread1.MaxRows < 1 Or vaSpread1.ActiveCol <> 8 And vaSpread1.ActiveCol <> 10 Then Exit Sub
    PopupMenu MenuDetalle

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Opgrilla_Click(Index As Integer)

On Error GoTo Man_Error

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
        
        vaSpread1.Col = 8
        CodIng = Trim(vaSpread1.text)
        
        vaSpread1.Col = 9
        noming = Trim(vaSpread1.text)
    
    End If
    
    
Case 1

    vaSpread1.GetSelection 1, Colini, FilIni, ColFin, FilFin

Dim i As Long
    
    If vaSpread1.ActiveCol = 10 Then
        
        For i = FilIni To FilFin
            
            vaSpread1.Row = i
            
            If vaSpread1.RowHidden = False Then
            
               vaSpread1.Col = 10
               vaSpread1.text = TmpCopiaGramaje
               vaSpread1.ForeColor = &HFF0000
            
               vaSpread1.Col = 1
               vaSpread1.text = "1"
       
            End If
            
        Next i
    
    Else
        
        For i = FilIni To FilFin
            
            vaSpread1.Row = i
            
                        
            If vaSpread1.RowHidden = False Then
            
               vaSpread1.Col = 8
               vaSpread1.text = CodIng
            
               vaSpread1.Col = 9
               vaSpread1.text = noming
            
               vaSpread1.Col = 1
               vaSpread1.text = "1"
            
            End If
            
        Next i
    
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub DataLoad()
    
On Error GoTo Man_Error

    Dim RS As New ADODB.Recordset
    Dim periodo As Long
    Dim Sql As String
    Dim MyBuffer As String
    Dim tiene As Integer
    
    Let MyBuffer = ""
    Let Sql = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<Zonas>"
    tiene = 0

'''***************zonas
    For j = 1 To TvwZon(1).Nodes.count
        
        If TvwZon(1).Nodes.item(j).Checked = True Then
            
            tiene = 1
            MyBuffer = MyBuffer & " <Zona"
            MyBuffer = MyBuffer & " zona = " & Chr(34) & Mid(TvwZon(1).Nodes.item(j).key, 2, Len(TvwZon(1).Nodes.item(j).key)) & Chr(34)
            Let MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next j
    Let MyBuffer = MyBuffer & "</Zonas>"
    Screen.MousePointer = 11
'''********************
    Sql = ""
    periodo = 0
    If Not IsNull(fpDateTime1(0).Value) And fpDateTime1(0).text <> "" Then
       
       periodo = Val(Mid(fpDateTime1(0).text, 4, 2) & Mid(fpDateTime1(0).text, 7, 4))
    
    End If
    
    If Me.fpText(0).text = "" Or tiene = 0 And Option1(1).Value = True Then
        
        Screen.MousePointer = 0
        MsgBox "Debe ingresar a lo menos Ing. Receta y Seleccionar Zona.....", vbInformation + vbOKOnly, MsgTitulo
        Exit Sub
    
    End If
    
    Sql = ""
       
    If Option1(0).Value = True Then
       
       Sql = " sgpadm_Sel_TablaGramajeCeco "
       If fpText(1).text <> "" Then
          
          Sql = Sql & "'" & fpText(1).text & "',"
       
       Else
          
          Sql = Sql & "'',"
       
       End If
    
    Else
       
       Sql = " sgpadm_Sel_TablaGramajeSubSegmento "
       
       If Val(fpLongInteger1(0).text) <> 0 Then
          
          Sql = Sql & Val(fpLongInteger1(0).text) & ","
       
       Else
          
          Sql = Sql & "0,"
       
       End If
    
    End If
    
    If Val(fpLongInteger1(1).text) <> 0 Then
        
        Sql = Sql & Val(fpLongInteger1(1).text) & ","
    
    Else
        
        Sql = Sql & "0,"
    
    End If
    
    If fpText(0).text <> "" Then
        
        Sql = Sql & "'" & fpText(0).text & "',"
    
    End If
    
    If Val(fpLongInteger1(2).text) > 0 Then
        
        Sql = Sql & Val(fpLongInteger1(2).text) & ","
    
    Else
        
        Sql = Sql & "0,"
    
    End If
    
    If Val(fpDateTime1(0).Value) > 0 Then
       
       Sql = Sql & Format(fpDateTime1(0).Value, "yyyymmdd") & ","
    
    Else
       
       Sql = Sql & "0,"
    
    End If
    
    If Val(fpDateTime1(1).Value) > 0 Then
        
        Sql = Sql & Format(fpDateTime1(1).Value, "yyyymmdd") & ","
    
    Else
       
       Sql = Sql & "0,"
    
    End If
        
    
    If FilCatDie <> 0 Then
        
        Sql = Sql & FilCatDie & ","
    
    Else
        
        Sql = Sql & "0,"
    
    End If
    
    If FilTipPla <> 0 Then
        
        Sql = Sql & FilTipPla & ","
    
    Else
        
        Sql = Sql & "0,"
    
    End If
    
    If Option2(0).Value = True Then
       
       Sql = Sql & "1"
    
    ElseIf Option2(1).Value = True Then
       
       Sql = Sql & "2"
    
    End If
    
    If Option1(1).Value = True Then
       
       Sql = Sql & ", " & "'" & MyBuffer & "'"
    
    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute(Sql)
    vaSpread1.MaxRows = 0
    If (RS.EOF = True And RS.BOF = True) Then MsgBox "No existe información", vbInformation, "Tabla Gramaje": Screen.MousePointer = 0: RS.Close: Set RS = Nothing: Exit Sub
'    fpayuda(8).Caption = IIf(RS.Fields(0) = "1", "Real", "Propuesta")
    Bar1.Visible = True: Bar1.Value = 0
    vaSpread1.MaxRows = 0
    Do While Not RS.EOF
       
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1
       If Val((vaSpread1.MaxRows / RS!cuantos) * 100) < 100 Then
        
        Bar1.Value = Val((vaSpread1.MaxRows / RS!cuantos) * 100)
       
       End If
       vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.Col = 1
       vaSpread1.CellType = CellTypeCheckBox
       vaSpread1.TypeCheckText = " "
       vaSpread1.TypeHAlign = TypeHAlignCenter
       vaSpread1.TypeCheckCenter = True
       vaSpread1.text = "0" ' checked
       
       vaSpread1.Col = 4
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = RS!rec_codigo
       
       vaSpread1.Col = 5
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = Trim(RS!rec_nombre)
       
       vaSpread1.Col = 6
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = RS!zon_codigo
       
       vaSpread1.Col = 7
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = Trim(RS!Zon_nombre)
       
       If RS!opgramo <> "R" And Toolbar1.Buttons(4).Enabled = False Then estgr = True
       
       vaSpread1.Col = 8
       vaSpread1.CellType = CellTypeEdit
       vaSpread1.text = IIf(RS!opgramo = "R", "", Trim(RS!red_codpro))
       
       vaSpread1.Col = 9
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = IIf(RS!opgramo = "R", "", Trim(RS!noming))
       
       vaSpread1.Col = 10
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
       'vaSpread1.ForeColor = IIf(RS!id = "1", &HFF&, &HFF0000)
       vaSpread1.ForeColor = IIf(RS!opgramo = "R", &HFF&, &HFF0000)
       
       vaSpread1.Col = 11
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = RS!codReg & " - " & Trim(RS!nomreg)
       
       vaSpread1.Col = 12
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = RS!codReg

       vaSpread1.Col = 13
       vaSpread1.CellType = CellTypeStaticText
       vaSpread1.text = RS!TipoPlato
       
       RS.MoveNext
       
    Loop
    
    OpGrilla(1).Enabled = False
    If estgr = True Then Toolbar1.Buttons(4).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 2) = "00", False, True)
    Bar1.Visible = False
    vaSpread1.Visible = True
    RS.Close
    Set RS = Nothing
    fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)

Select Case Col

Case Is <> 1
    
    Command2(1).Visible = False

Case Is <> 3

    Command2(2).Visible = False

Case Is <> 5

    Command2(3).Visible = False

Case Is <> 7

    Command2(4).Visible = False

End Select

Select Case Col

Case 1
    
    Command2(1).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
    Command2(1).Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 1
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 3
    
    Command2(2).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
    Command2(2).Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 3
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 5
    
    Command2(3).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
    Command2(3).Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 5
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 7
    
    Command2(4).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
    Command2(4).Visible = True
    vaSpread2.EditMode = True
    vaSpread2.EditModeReplace = True
    vaSpread2.Row = Row
    IRow = Row
    vaSpread2.Col = 7
    vaSpread2.TypeHAlign = TypeHAlignLeft

Case 10

    vaSpread2.Row = Row
    vaSpread2.Col = 11
    vaSpread2.text = "1"

End Select

End Sub

Private Sub vaSpread2_EditChange(ByVal Col As Long, ByVal Row As Long)

'If modo = "" Then modo = "M"
'Gl_Ac_Botones Me, 1, 0, modo
SSTab1.TabEnabled(0) = False

End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

Dim RS As New ADODB.Recordset

If SSTab1.Tab = 0 Then Exit Sub

IRow = Row

Select Case Col

    Case 1
   
     Command2(1).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
     Command2(1).Visible = True

    Case 3
   
     Command2(2).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
     Command2(2).Visible = True
    
    Case 5
   
     Command2(3).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
     Command2(3).Visible = True
    
    Case 7
   
     Command2(4).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
     Command2(4).Visible = True

End Select

If ChangeMade = False And Col <> 1 Then

   Command2(1).Visible = False
   Exit Sub

End If

If ChangeMade = False And Col <> 3 Then

   Command2(2).Visible = False
   Exit Sub

End If

If ChangeMade = False And Col <> 5 Then

   Command2(3).Visible = False
   Exit Sub

End If

If ChangeMade = False And Col <> 7 Then

   Command2(4).Visible = False
   Exit Sub

End If

Select Case Col
    
    Case 1
        
        Command2(1).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
        Command2(1).Visible = False
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("SELECT * FROM a_regimen WHERE reg_codigo = " & Val(vaSpread2.Value) & "")
        
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           vaSpread2.text = ""
           vaSpread2.Col = 2
           vaSpread2.text = ""
           Exit Sub
        
        End If
        
        vaSpread2.Col = 2
        vaSpread2.text = Trim(IIf(IsNull(RS!reg_nombre), "", RS!reg_nombre))
        RS.Close: Set RS = Nothing
        Command2(1).Visible = False
        
        vaSpread2.Col = 11
        vaSpread2.text = "1"
       
    Case 3
        
        Command2(2).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
        Command2(2).Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("SELECT * FROM a_recetatippla WHERE tip_codigo = " & Val(vaSpread2.Value) & "")
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           vaSpread2.text = ""
           
           vaSpread2.Col = 4
           vaSpread2.text = ""
           
           Exit Sub
               
        End If
        
        vaSpread2.Col = 4
        vaSpread2.text = Trim(IIf(IsNull(RS!tip_nombre), "", RS!tip_nombre))
        RS.Close
        Set RS = Nothing
        Command2(2).Visible = False
        
        vaSpread2.Col = 11
        vaSpread2.text = "1"
    
    Case 5
        
        Command2(3).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
        Command2(3).Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS.Open "SELECT top 1 a.ing_codigo, a.ing_nombre FROM b_receta b WITH (NOLOCK) " & _
                "inner join b_recetadet c WITH (NOLOCK) on b.rec_codigo = c.red_codigo " & _
                "inner join b_ingrediente a WITH (NOLOCK) on c.red_codpro = a.ing_codigo " & _
                "WHERE a.ing_codigo = '" & (vaSpread2.text) & "' ", vg_db, adOpenStatic
        
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           
           MsgBox "Ingrediente ingresado no esta asociado a ninguna receta o bien no existe ingrediente...", vbExclamation + vbOKOnly, MsgTitulo
           
           vaSpread2.text = ""
           vaSpread2.Col = 6
           vaSpread2.text = ""
           Exit Sub
        
        End If
        
        vaSpread2.Col = 6
        vaSpread2.text = Trim(RS!ing_nombre)
        RS.Close
        Set RS = Nothing
        Command2(3).Visible = False
        
        vaSpread2.Col = 11
        vaSpread2.text = "1"
    
    Case 7
        
        Command2(4).Top = IIf(Row = 1, 1000, 1000 + (240 * (Row - itop)))
        Command2(4).Visible = True
        vaSpread2.Row = Row
        vaSpread2.Col = Col
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS.Open "SELECT a.ing_codigo, a.ing_nombre FROM b_ingrediente a WITH (NOLOCK) WHERE a.ing_codigo = '" & (vaSpread2.text) & "' ", vg_db, adOpenStatic
        
        If RS.EOF Then
           
           RS.Close
           Set RS = Nothing
           
           MsgBox "No existe ingrediente...", vbExclamation + vbOKOnly, MsgTitulo
           
           vaSpread2.text = ""
           vaSpread2.Col = 8
           vaSpread2.text = ""
           Exit Sub
               
        End If
        
        vaSpread2.Col = 8
        vaSpread2.text = Trim(RS!ing_nombre)
        RS.Close
        Set RS = Nothing
        Command2(4).Visible = False
    
        vaSpread2.Col = 11
        vaSpread2.text = "1"
        
    Case 9
    
        vaSpread2.Col = 11
        vaSpread2.text = "1"

    Case 10
    
        vaSpread2.Col = 11
        vaSpread2.text = "1"

End Select

End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

If Row = 0 Or estelilinea Or NewRow = vaSpread2.MaxRows Or NewRow < 1 Then Exit Sub

' Validar solo si cambia de fila
If vaSpread2.MaxRows <> NewRow Then
   
   If ValidarGrillaDatosRepetidos Then
   
      Exit Sub
      
   End If

   If ValidarGrillaNivel(0) Then
       
      Exit Sub
       
   End If

End If


If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 1 Then Command2(1).Visible = False
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 3 Then Command2(2).Visible = False
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 5 Then Command2(3).Visible = False
If (Row <> NewRow And NewRow > 0) Or (Col <> NewCol And NewCol > 0) Or Col <> 7 Then Command2(4).Visible = False

End Sub

Function ValidarGrillaDatosRepetidos() As Boolean

Dim i             As Long
Dim Regimen       As String
Dim regimenII     As String
Dim Ingrediente   As String
Dim ingredienteII As String
Dim TipoPlato     As String
Dim TipoPlatoII   As String
   
   ValidarGrillaDatosRepetidos = False
   vaSpread2.Row = vaSpread2.MaxRows
   
   vaSpread2.Col = 1
   regimenII = vaSpread2.text
   
   vaSpread2.Col = 3
   TipoPlatoII = vaSpread2.text
   
   vaSpread2.Col = 5
   ingredienteII = vaSpread2.text
   
   For i = 1 To vaSpread2.MaxRows
                
       vaSpread2.Row = i
       vaSpread2.Col = 1
                

       Regimen = vaSpread2.text
       
       vaSpread2.Col = 3
       TipoPlato = vaSpread2.text
       
       vaSpread2.Col = 5
       Ingrediente = vaSpread2.text
       
       'validar primer nivel
       If Regimen = regimenII And TipoPlato = TipoPlatoII And Ingrediente = ingredienteII And vaSpread2.MaxRows <> i Then
                       
          ValidarGrillaDatosRepetidos = True
          
          MsgBox "ya existen dato en grilla, con los datos existen...", vbExclamation + vbOKOnly, MsgTitulo

          vaSpread2.Row = vaSpread2.MaxRows
          vaSpread2.Col = 1
          vaSpread2.SetActiveCell 2, vaSpread2.MaxRows
          vaSpread2.SetFocus
                   
          Exit Function
          
       End If
                
   Next i

Exit Function:

End Function

Function ValidarGrillaNivel(Est As String) As Boolean

Dim RS            As New ADODB.Recordset
Dim regimenII     As String
Dim ingredienteII As String
Dim TipoPlatoII   As String
   
   ValidarGrillaNivel = False
   
   vaSpread2.Row = vaSpread2.MaxRows
   
   vaSpread2.Col = 1
   regimenII = vaSpread2.text
   
   vaSpread2.Col = 3
   TipoPlatoII = vaSpread2.text
   
   vaSpread2.Col = 5
   ingredienteII = vaSpread2.text
   
   If Est = "1" Then
      
      If RS.State = 1 Then RS.Close
      RS.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      Set RS = vg_db.Execute("sgpadm_Sel_TablaGramajeNivelRegistro_V01 '" & Trim(LimpiaDato(fpText(2).text)) & "', '" & regimenII & "', '" & TipoPlatoII & "','" & ingredienteII & "'")
   
      If RS.EOF Then
   
        RS.Close
        Set RS = Nothing
      
        ValidarGrillaNivel = False
        Exit Function
        
      End If
      RS.Close
      Set RS = Nothing
   
   End If
   
   'validar primer nivel
   If ingredienteII = "" Then
                       
       ValidarGrillaNivel = True
          
       MsgBox "tiene que ingresar ingrediente origen...", vbExclamation + vbOKOnly, MsgTitulo

       vaSpread2.Row = vaSpread2.MaxRows
       vaSpread2.Col = 1
       vaSpread2.SetActiveCell 2, vaSpread2.MaxRows
       vaSpread2.SetFocus
                   
       Exit Function
          
   End If
   
   'validar tercer nivel
   If ingredienteII <> "" And regimenII = "" And TipoPlatoII <> "" Then
                       
       ValidarGrillaNivel = True
          
       MsgBox "si esta ingresando tipo de plato y ing. origen, debe ingresar regimen...", vbExclamation + vbOKOnly, MsgTitulo

       vaSpread2.Row = vaSpread2.MaxRows
       vaSpread2.Col = 1
       vaSpread2.SetActiveCell 2, vaSpread2.MaxRows
       vaSpread2.SetFocus
                   
       Exit Function
          
   End If
   
Exit Function:

End Function

Private Sub vaSpread2_TopLeftChange(ByVal OldLeft As Long, ByVal OldTop As Long, ByVal NewLeft As Long, ByVal NewTop As Long)

itop = NewTop
Command2(1).Visible = False
Command2(2).Visible = False
Command2(3).Visible = False
Command2(4).Visible = False

End Sub

