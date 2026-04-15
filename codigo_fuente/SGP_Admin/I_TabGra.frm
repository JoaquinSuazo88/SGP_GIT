VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form I_TabGra 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Tabla Gramaje"
   ClientHeight    =   5340
   ClientLeft      =   4635
   ClientTop       =   2295
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   4695
      Index           =   0
      Left            =   80
      TabIndex        =   14
      Top             =   480
      Width           =   8775
      Begin MSComDlg.CommonDialog CD 
         Left            =   1200
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Caption         =   "Filtro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   250
         TabIndex        =   21
         Top             =   3240
         Width           =   7965
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "(Opcional)"
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
            Left            =   6930
            TabIndex        =   25
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "(Opcional)"
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
            Left            =   6930
            TabIndex        =   24
            Top             =   390
            Width           =   885
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   1590
            TabIndex        =   13
            Top             =   630
            Width           =   5175
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   1140
            Picture         =   "I_TabGra.frx":0000
            Top             =   525
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   1140
            Picture         =   "I_TabGra.frx":030A
            Top             =   180
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
            Left            =   90
            TabIndex        =   23
            Top             =   690
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
            TabIndex        =   22
            Top             =   345
            Width           =   1020
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   1590
            TabIndex        =   12
            Top             =   285
            Width           =   5175
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   1605
            TabIndex        =   26
            Top             =   330
            Width           =   5205
         End
         Begin VB.Label lblSOMBRA 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   1605
            TabIndex        =   27
            Top             =   675
            Width           =   5205
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2655
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   480
         Width           =   8535
         Begin VB.Frame Frame3 
            Height          =   1575
            Left            =   120
            TabIndex        =   30
            Top             =   120
            Width           =   8295
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
               Left            =   120
               TabIndex        =   2
               Top             =   240
               Value           =   -1  'True
               Width           =   1815
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
               Left            =   4740
               TabIndex        =   3
               Top             =   240
               Width           =   1815
            End
            Begin EditLib.fpLongInteger fpLongInteger1 
               Height          =   315
               Index           =   0
               Left            =   1650
               TabIndex        =   6
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
            Begin EditLib.fpText fpText 
               Height          =   315
               Index           =   1
               Left            =   1650
               TabIndex        =   4
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
               Left            =   120
               TabIndex        =   33
               Top             =   645
               Width           =   1380
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   5
               Left            =   2535
               Picture         =   "I_TabGra.frx":0614
               Top             =   480
               Width           =   480
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   9
               Left            =   2985
               TabIndex        =   5
               Top             =   600
               Width           =   5175
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   0
               Left            =   3000
               TabIndex        =   7
               Top             =   1065
               Width           =   5175
            End
            Begin VB.Image Image1 
               Height          =   480
               Index           =   0
               Left            =   2520
               Picture         =   "I_TabGra.frx":091E
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
               Left            =   120
               TabIndex        =   31
               Top             =   1140
               Width           =   1245
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   99
               Left            =   3015
               TabIndex        =   32
               Top             =   1095
               Width           =   5205
            End
            Begin VB.Label fpayuda 
               Appearance      =   0  'Flat
               BackColor       =   &H80000010&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   10
               Left            =   3000
               TabIndex        =   34
               Top             =   630
               Width           =   5205
            End
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   0
            Left            =   1470
            TabIndex        =   10
            Top             =   2250
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
            Left            =   1470
            TabIndex        =   8
            Top             =   1890
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
            Left            =   120
            TabIndex        =   18
            Top             =   2310
            Width           =   1020
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   2400
            Picture         =   "I_TabGra.frx":0C28
            Top             =   2160
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   2880
            TabIndex        =   11
            Top             =   2250
            Width           =   5175
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   2880
            TabIndex        =   9
            Top             =   1890
            Width           =   5175
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   2400
            Picture         =   "I_TabGra.frx":0F32
            Top             =   1800
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
            Left            =   120
            TabIndex        =   17
            Top             =   1950
            Width           =   750
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   5
            Left            =   2895
            TabIndex        =   20
            Top             =   1905
            Width           =   5205
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   300
            Index           =   98
            Left            =   2895
            TabIndex        =   19
            Top             =   2265
            Width           =   5205
         End
      End
      Begin VB.OptionButton Option1 
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
         Height          =   255
         Index           =   1
         Left            =   7080
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
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
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   2040
         TabIndex        =   28
         Top             =   4440
         Visible         =   0   'False
         Width           =   4470
         _ExtentX        =   7885
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Generando Informe"
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
         TabIndex        =   29
         Top             =   4440
         Visible         =   0   'False
         Width           =   1635
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_TabGra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim FilTipPla As Long, FilCatDie As Long
Dim MsgTitulo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

fg_centra Me
fg_carga ""
MsgTitulo = "Imprimir Tabla Gramaje"
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

fpayuda(3).Caption = "Todos": fpayuda(4).Caption = "Todos"

fpText(1).text = ""
fpText(1).Enabled = True
Image1(5).Enabled = True
fpayuda(9).Caption = ""

fpLongInteger1(0).Value = ""
fpLongInteger1(0).Enabled = False
Image1(0).Enabled = False
fpayuda(0).Caption = ""

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Select Case Index

Case 0
    
    RS.Open "SELECT * FROM a_subsegmento with (nolock) WHERE sub_codigo=" & Val(fpLongInteger1(0).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!sub_nombre)
    RS.Close: Set RS = Nothing

Case 1
    
    RS.Open "SELECT * FROM a_regimen with (nolock) WHERE reg_codigo=" & Val(fpLongInteger1(1).Value) & "", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

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
    
    RS.Open "SELECT DISTINCT a.ing_codigo, a.ing_nombre FROM b_ingrediente a WITH ( NOLOCK ), b_receta b WITH ( NOLOCK ), b_recetadet c WITH ( NOLOCK ) WHERE b.rec_codigo=c.red_codigo AND c.red_codpro=a.ing_codigo AND a.ing_codigo='" & Trim(fpText(0).text) & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "":  Exit Sub
    fpayuda(2).Caption = Trim(RS!ing_nombre)
    RS.Close: Set RS = Nothing

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
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

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
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
    fpText(0).text = "": fpayuda(1).Caption = ""
    fpLongInteger1(1).SetFocus

Case 1
    
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_regimen", "reg_", "Regimen", "Reg"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    Est = False
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre
    fpText(0).SetFocus

Case 2
    
    vg_left = fpayuda(2).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingrediente", "Ingrec"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo
    fpayuda(2).Caption = vg_nombre

Case 3
    
    vg_left = fpayuda(2).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica", "1"
    B_ArbEst.Show 1
    If vg_codigo = "" Then Exit Sub
    FilCatDie = Val(vg_codigo)
    fpayuda(3).Caption = vg_nombre: vg_nombre = ""

Case 4
    
    vg_codigo = "": vg_nombre = ""
    vg_left = fpayuda(3).Left + 3000
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato", "1"
    B_ArbEst.Show 1
    If Trim(vg_codigo) = "" Then Exit Sub
    FilTipPla = Val(vg_codigo)
    fpayuda(4).Caption = vg_nombre: vg_nombre = ""

Case 5
    
    vg_left = fpayuda(9).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "CliAct"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText(1).text = vg_codigo: fpayuda(9).Caption = vg_nombre
    fpLongInteger1(1).SetFocus

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    Frame1(1).Enabled = True
'    fpLongInteger11(0).Value = ""
'    fpLongInteger11(1).Value = ""
'    fpText1(0).Value = ""
'    fpayuda1(0).Caption = ""
'    fpayuda1(1).Caption = ""
'    fpayuda1(2).Caption = ""

Case 1

'    Frame1(1).Enabled = False
    fpLongInteger1(0).Value = ""
    fpLongInteger1(1).Value = ""
    fpText(0).text = ""
    fpayuda(0).Caption = ""
    fpayuda(1).Caption = ""
    fpayuda(2).Caption = ""

Case 3
    
    fpText(1).text = ""
    fpText(1).Enabled = True
    Image1(5).Enabled = True
    fpayuda(9).Caption = ""
    
    fpLongInteger1(1).Value = ""
    fpLongInteger1(0).Value = ""
    fpLongInteger1(0).Enabled = False
    Image1(0).Enabled = False
    fpayuda(0).Caption = ""

Case 2
    
    fpText(1).text = ""
    fpText(1).Enabled = False
    Image1(5).Enabled = False
    fpayuda(9).Caption = ""
    
    fpLongInteger1(1).Value = ""
    fpLongInteger1(0).Value = ""
    fpLongInteger1(0).Enabled = True
    Image1(0).Enabled = True
    fpayuda(0).Caption = ""

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 1
       'Validar opción de ingreso
    If Option1(3).Value = True Then
       If fpayuda(9).Caption = "" Then MsgBox "No existe centro de costo", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    ElseIf Option1(2).Value = True Then
       If fpayuda(0).Caption = "" Then MsgBox "No existe sub-segmento", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    End If
    If fpayuda(1).Caption = "" Then MsgBox "No existe regimen", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    
'    Toolbar1.Enabled = False
'    Frame1(0).Enabled = False
    If Option1(3).Value = True Then
       
       'I_TablaGramajeCeco LimpiaDato(Trim(fpText(1).text)), Val(fpLongInteger1(1).Value), LimpiaDato(Trim(fpText(0).text)), FilCatDie, FilTipPla
       ImprimirTablaGrameCecoExcel
       
    ElseIf Option1(2).Value = True Then
       
       I_TablaGramaje Val(fpLongInteger1(0).Value), Val(fpLongInteger1(1).Value), LimpiaDato(Trim(fpText(0).text)), FilCatDie, FilTipPla
    
    End If
'    Toolbar1.Enabled = True
'    Frame1(0).Enabled = True

Case 3
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Informe Recetas"

End Sub

Sub ImprimirTablaGrameCecoExcel()

Dim i               As Long
Dim RS              As New ADODB.Recordset
Dim NomArchivoExcel As String
Dim Extension       As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object
Dim XL              As New excel.Application 'Crea el objeto excel

  '-------> Rescata Ceco Seleccionado
  fg_carga ""
  
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient

  Set RS = vg_db.Execute("sgpadm_Sel_ImpTablaGramajeCeco_V03 '" & LimpiaDato(Trim(fpText(1).text)) & "', " & Val(fpLongInteger1(1).Value) & ", '" & LimpiaDato(Trim(fpText(0).text)) & "', " & FilCatDie & ", " & FilTipPla & "")

  If Not RS.EOF Then
             
     If RS.RecordCount > 1020000 Then
      
        RS.Close
        Set RS = Nothing
      
        MsgBox "El resultado sobrepasa maximo de fila en excel, Debera seleccionar menos Recetas", vbCritical
        Exit Sub
   
     End If
             

    '-------> Guardar nombre archivo excel
    NomArchivoExcel = ""
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.Filter = "Todos los archivos *.xls,*.xlsx"
    On Error Resume Next
    CD.ShowSave
               
    '-------> JPAZ Permite controlar Boton Cancelar
    If Err.Number = 32755 Then
       
       MsgBox "Proceso cancelado"
       Exit Sub
    
    End If
                
    If CD.FileName = "" Then
       
       MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
       Exit Sub
    
    Else
       
       Extension = ""
       Extension = Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, ".")))
       
       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
          Exit Sub
       End If
       
       NomArchivoExcel = CD.FileName
    
    End If
       
       '-------> Create an instance of Excel and add a workbook
       Set xlApp = CreateObject("Excel.Application")
       Set xlWb = xlApp.Workbooks.Add
       Set xlWs = xlWb.Worksheets("Hoja1")
  
       '-------> Display Excel and give user control of Excel's lifetime
       xlApp.UserControl = True
    
       '-------> Check version of Excel
       Call encabezado(RS, xlWs)
        
       xlWs.Cells(2, 1).CopyFromRecordset RS

       '-------> Auto-fit the column widths and row heights
'       xlApp.Selection.CurrentRegion.Columns.AutoFit
'       xlApp.Selection.CurrentRegion.Rows.AutoFit
    
'       xlApp.Columns("A:B").Select
'       xlApp.Selection.Delete Shift:=xlToLeft
  
'       NomArchivoExcel = fg_ArchivoXls("ExportarExcel_EncabezadoReceta")
                    
       xlWb.Close True, NomArchivoExcel

''       Dim XL As New excel.Application 'Crea el objeto excel
       XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
       XL.Visible = True
       XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada

       '-- Cerrar Excel
       xlApp.Quit
      
       '-------> Release Excel references
       Set xlWs = Nothing
       Set xlWb = Nothing
       Set xlApp = Nothing
          
       MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
                                                    
  End If
  
  RS.Close
  Set RS = Nothing

  fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
