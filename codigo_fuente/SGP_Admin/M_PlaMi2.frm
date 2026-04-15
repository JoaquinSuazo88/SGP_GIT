VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_Plami2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Teórica"
   ClientHeight    =   7815
   ClientLeft      =   2235
   ClientTop       =   1665
   ClientWidth     =   11640
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7815
   ScaleWidth      =   11640
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   11400
      Top             =   4680
   End
   Begin VB.Frame Frame2 
      Height          =   2625
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   5190
      Visible         =   0   'False
      Width           =   15195
      Begin VB.Frame Frame2 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   56
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   57
            Top             =   150
            Visible         =   0   'False
            Width           =   2610
            _ExtentX        =   4604
            _ExtentY        =   318
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   120
            Picture         =   "M_PlaMi2.frx":0000
            Top             =   150
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   480
         Width           =   3795
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rac."
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
            Index           =   49
            Left            =   120
            TabIndex        =   67
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   48
            Left            =   2370
            TabIndex        =   66
            Top             =   1380
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   41
            Left            =   1020
            TabIndex        =   59
            Top             =   810
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   40
            Left            =   1020
            TabIndex        =   58
            Top             =   525
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo Medio"
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
            Index           =   39
            Left            =   1260
            TabIndex        =   55
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total del Mes"
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
            Left            =   2520
            TabIndex        =   54
            Top             =   240
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mat.Prima"
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
            TabIndex        =   53
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Est.Fija"
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
            TabIndex        =   52
            Top             =   810
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo"
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
            Left            =   120
            TabIndex        =   51
            Top             =   1095
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Patrón"
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
            TabIndex        =   50
            Top             =   1770
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
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
            TabIndex        =   49
            Top             =   1830
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   7
            Left            =   2370
            TabIndex        =   48
            Top             =   525
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   8
            Left            =   1020
            TabIndex        =   47
            Top             =   1095
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   9
            Left            =   1020
            TabIndex        =   46
            Top             =   1770
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   10
            Left            =   1020
            TabIndex        =   45
            Top             =   1830
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   11
            Left            =   2370
            TabIndex        =   44
            Top             =   810
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   12
            Left            =   2370
            TabIndex        =   43
            Top             =   1095
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   13
            Left            =   2370
            TabIndex        =   42
            Top             =   1770
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   14
            Left            =   2370
            TabIndex        =   41
            Top             =   1830
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Día 01/08/2004"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   2
         Left            =   3960
         TabIndex        =   30
         Top             =   480
         Width           =   3795
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   47
            Left            =   2370
            TabIndex        =   65
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   46
            Left            =   2370
            TabIndex        =   64
            Top             =   1380
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   45
            Left            =   960
            TabIndex        =   63
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   44
            Left            =   960
            TabIndex        =   62
            Top             =   1380
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Medio"
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
            Index           =   43
            Left            =   90
            TabIndex        =   61
            Top             =   1680
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rac."
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
            Index           =   42
            Left            =   90
            TabIndex        =   60
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mat.Prima"
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
            Left            =   90
            TabIndex        =   39
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Est.Fija"
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
            Left            =   90
            TabIndex        =   38
            Top             =   810
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Total"
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
            Index           =   17
            Left            =   90
            TabIndex        =   37
            Top             =   1095
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Realizado"
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
            Index           =   18
            Left            =   1410
            TabIndex        =   36
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Food Cost"
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
            Index           =   19
            Left            =   2790
            TabIndex        =   35
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   20
            Left            =   960
            TabIndex        =   34
            Top             =   525
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   21
            Left            =   960
            TabIndex        =   33
            Top             =   810
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   22
            Left            =   960
            TabIndex        =   32
            Top             =   1095
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   23
            Left            =   2370
            TabIndex        =   31
            Top             =   1095
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Acumulado hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Index           =   3
         Left            =   7800
         TabIndex        =   14
         Top             =   480
         Width           =   3795
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Mat.Prima"
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
            Index           =   24
            Left            =   90
            TabIndex        =   29
            Top             =   465
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Est.Fija"
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
            Index           =   25
            Left            =   90
            TabIndex        =   28
            Top             =   750
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Total"
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
            Index           =   26
            Left            =   60
            TabIndex        =   27
            Top             =   1050
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rac."
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
            Index           =   27
            Left            =   90
            TabIndex        =   26
            Top             =   1350
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Medio"
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
            Index           =   28
            Left            =   90
            TabIndex        =   25
            Top             =   1650
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Realizado"
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
            Index           =   29
            Left            =   1425
            TabIndex        =   24
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Food Cost"
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
            Index           =   30
            Left            =   2790
            TabIndex        =   23
            Top             =   210
            Width           =   870
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   31
            Left            =   960
            TabIndex        =   22
            Top             =   465
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   32
            Left            =   960
            TabIndex        =   21
            Top             =   750
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   33
            Left            =   960
            TabIndex        =   20
            Top             =   1050
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   34
            Left            =   960
            TabIndex        =   19
            Top             =   1350
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   35
            Left            =   960
            TabIndex        =   18
            Top             =   1650
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   36
            Left            =   2370
            TabIndex        =   17
            Top             =   1050
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   37
            Left            =   2370
            TabIndex        =   16
            Top             =   1350
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Index           =   38
            Left            =   2370
            TabIndex        =   15
            Top             =   1650
            Width           =   1320
         End
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      ScaleHeight     =   1035
      ScaleWidth      =   6675
      TabIndex        =   3
      Top             =   2880
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   210
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin MSComctlLib.ProgressBar gauge1 
         Height          =   210
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Día"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   195
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   10905
      Begin VB.CheckBox Check1 
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   9180
         TabIndex        =   71
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Costo"
         Height          =   255
         Index           =   1
         Left            =   7365
         TabIndex        =   70
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "N.Rac."
         Height          =   195
         Index           =   0
         Left            =   5160
         TabIndex        =   69
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Semana Nº"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   150
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H80000018&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   9180
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         Height          =   195
         Index           =   0
         Left            =   9540
         TabIndex        =   11
         Top             =   135
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00D9D9FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   7365
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Bloqueada"
         Height          =   195
         Index           =   1
         Left            =   7725
         TabIndex        =   10
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00DEFEDE&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   5160
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estructura de Servicio"
         Height          =   195
         Index           =   2
         Left            =   5520
         TabIndex        =   9
         Top             =   135
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   30
         Left            =   4200
         TabIndex        =   68
         Top             =   240
         Width           =   135
         _ExtentX        =   238
         _ExtentY        =   53
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      DragIcon        =   "M_PlaMi2.frx":030A
      Height          =   3600
      Left            =   0
      TabIndex        =   8
      Top             =   1560
      Width           =   11055
      _Version        =   393216
      _ExtentX        =   19500
      _ExtentY        =   6350
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
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
      MaxCols         =   250
      MaxRows         =   100
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RestrictRows    =   -1  'True
      SpreadDesigner  =   "M_PlaMi2.frx":074C
      UserResize      =   1
      VisibleCols     =   1
      VisibleRows     =   100
      TextTip         =   2
      TextTipDelay    =   0
      ScrollBarTrack  =   3
   End
   Begin VB.Label fpayuda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   2
      Left            =   45
      TabIndex        =   72
      Top             =   125
      Width           =   4575
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planificación Minutas Teórica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   10905
   End
   Begin VB.Menu Main 
      Caption         =   "Menú"
      Index           =   0
      NegotiatePosition=   2  'Middle
      Begin VB.Menu Plantilla 
         Caption         =   "&Grabar Semana"
         Enabled         =   0   'False
         Index           =   0
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Ver &Receta"
         Index           =   5
      End
      Begin VB.Menu Plantilla 
         Caption         =   "C&opiar Minutas"
         Index           =   8
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Aporte &Nutricionales x Días"
         Index           =   10
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Costo Receta"
         Index           =   11
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Frecuencia Recetas"
         Index           =   12
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Ac&tualizar Costo Planificación"
         Index           =   13
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Exportar Recetas"
         Index           =   14
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Parámetro de Grabado"
         Index           =   16
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   17
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Cerrar"
         Index           =   22
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Plato Menú"
      Index           =   1
      NegotiatePosition=   2  'Middle
      Begin VB.Menu Plato 
         Caption         =   "&Deshacer"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Plato 
         Caption         =   "Cambiar Plato &Menú"
         Index           =   2
      End
      Begin VB.Menu Plato 
         Caption         =   "Come&ntario"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Plato 
         Caption         =   "&Insertar"
         Index           =   5
      End
      Begin VB.Menu Plato 
         Caption         =   "&Eliminar"
         Index           =   6
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Plato 
         Caption         =   "&Subir"
         Index           =   8
      End
      Begin VB.Menu Plato 
         Caption         =   "&Bajar"
         Index           =   9
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu Plato 
         Caption         =   "Cor&tar"
         Index           =   11
         Shortcut        =   ^X
      End
      Begin VB.Menu Plato 
         Caption         =   "C&opiar"
         Index           =   12
         Shortcut        =   ^C
      End
      Begin VB.Menu Plato 
         Caption         =   "&Pegar"
         Enabled         =   0   'False
         Index           =   13
         Shortcut        =   ^V
      End
      Begin VB.Menu Plato 
         Caption         =   "Pegado &Especial"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu Plato 
         Caption         =   "&Buscar Recetas o Ingredientes"
         Index           =   15
         Shortcut        =   ^B
      End
      Begin VB.Menu Plato 
         Caption         =   "Crear Estr&uctura"
         Index           =   17
      End
      Begin VB.Menu Plato 
         Caption         =   "&Agrega Estructura"
         Index           =   18
         Begin VB.Menu Estructura1 
            Caption         =   ""
            Index           =   0
         End
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Ver"
      Index           =   2
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu Ver 
         Caption         =   "Días &Pantalla"
         Index           =   0
         Visible         =   0   'False
         Begin VB.Menu Dias 
            Caption         =   "&1"
            Index           =   0
         End
         Begin VB.Menu Dias 
            Caption         =   "&2"
            Index           =   1
         End
         Begin VB.Menu Dias 
            Caption         =   "&3"
            Index           =   2
         End
         Begin VB.Menu Dias 
            Caption         =   "&4"
            Index           =   3
         End
         Begin VB.Menu Dias 
            Caption         =   "&5"
            Index           =   4
         End
         Begin VB.Menu Dias 
            Caption         =   "&6"
            Index           =   5
         End
         Begin VB.Menu Dias 
            Caption         =   "&7"
            Index           =   6
         End
      End
      Begin VB.Menu Ver 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "&Semana Siguiente"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "Semana &Anterior"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "-"
         Index           =   4
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "Costo Minutas"
         Index           =   5
         Visible         =   0   'False
      End
      Begin VB.Menu Ver 
         Caption         =   "Aporte &Nutricional x Día"
         Index           =   6
      End
      Begin VB.Menu Ver 
         Caption         =   "&Gramos Productos Mensual"
         Index           =   7
      End
      Begin VB.Menu Ver 
         Caption         =   "&Frecuencia De Recetas"
         Index           =   8
      End
      Begin VB.Menu Ver 
         Caption         =   "&Costo Minuta Resumido"
         Index           =   9
      End
   End
   Begin VB.Menu MenuDetalle 
      Caption         =   ""
      NegotiatePosition=   2  'Middle
      Visible         =   0   'False
      Begin VB.Menu OpGrilla 
         Caption         =   "Deshacer"
         Enabled         =   0   'False
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Cambiar Plato &Menú"
         Index           =   2
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Come&ntario"
         Index           =   3
         Visible         =   0   'False
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Insertar"
         Index           =   5
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Eliminar"
         Index           =   6
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Subir"
         Index           =   8
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Bajar"
         Index           =   9
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Cor&tar"
         Index           =   11
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "C&opiar"
         Index           =   12
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Pegar"
         Enabled         =   0   'False
         Index           =   13
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Pegado Especial"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Buscar Receta"
         Index           =   15
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Ag&rega Estructura Personalizada"
         Index           =   16
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Crear Estr&uctura"
         Index           =   17
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Agrega Estructura"
         Index           =   18
         Begin VB.Menu Estructura2 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "M_Plami2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim i As Long, j As Long, indcortarpegar As Long, Fecha As Long, MaxColumna As Long, maxfila As Long
Dim iblockrow As Integer, iblockrow2 As Integer, iblockcol As Integer, iblockcol2 As Integer, SwSalir As Integer
Dim aiblockrow As Integer, aiblockrow2 As Integer, aiblockcol As Integer, aiblockcol2 As Integer, indactivo As Integer
Dim indcos As Boolean, estgra As Boolean, estapo As Boolean
Dim veccos() As Variant
Dim vectorcol() As Long
Dim MsgTitulo As String
Dim TipoCopia As String, NameTemp As String
Dim SpresdText As String
Dim CellTex As String
Dim SpreadClon As New M_Plami2
Dim xColIni As Variant, xRowIni As Variant, xcolfin As Variant, xRowFin As Variant
Dim CorDes As Long
'**** Samuel melendez ------------------------------------
Private Declare Function sendmessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wmsg As Long, _
    ByVal wparam As Long, lparam As Any) As Long
    Private Const EM_CANUNDO = &HC6
    Private Const EM_UNDO = &HC7
'***-------------------------------------------------------


'******* Id de Proceso SQL ***********************************************
'** Las siguientes variables RSSpid y Spid, sirven a
'** algunos procesos los cuelaes necesitan identificar el
'** turno de usuario que nos asigna SQL Server de manera unica.
'** estas variables solo deben ocuparse
'** para consultar el numero de proceso, el cual sera siempre el mismo
'** mientras no se cierre este formulario, asi mismo estas se destruyen
'** cuando es cerrado el formulario.
Dim RSSpid As New ADODB.Recordset
Dim spid As Long
'**----------------------------------------------------------------------
'************************************************************************
Enum DeshacerType
    AddFile = 1
    DelFile = 2
End Enum

Private Sub Check1_Click(Index As Integer)
HabilitaCol Index
End Sub

Private Sub Estructura1_Click(Index As Integer)
    LlenaSubMenu Estructura1, Index
End Sub

Sub LlenaSubMenu(SubMenu As Object, Index As Integer)
Dim auxest As Long
Dim xrow As Long, i As Long

xrow = vaSpread1.ActiveRow
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.maxcols
If Trim(vaSpread1.text) <> "" Then auxest = Val(vaSpread1.text)
vaSpread1.Col = 1

DesqloqSubMenu vaSpread1.text
vaSpread1.text = SubMenu(Index).Caption

ActualizaEstructuraInferior vaSpread1, SubMenu(Index).Caption
vaSpread1.Col = vaSpread1.maxcols: vaSpread1.text = SubMenu(Index).HelpContextID
'------->
For i = xrow + 1 To vaSpread1.MaxRows - 1
    vaSpread1.Row = i
    vaSpread1.Col = vaSpread1.maxcols
    If Val(vaSpread1.text) = auxest Then
       vaSpread1.text = SubMenu(Index).HelpContextID
    Else
       Exit For
    End If
Next i

Estructura1.item(Index).Enabled = False: Estructura2.item(Index).Enabled = False

Plantilla(0).Enabled = True
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
End Sub

Private Sub Estructura2_Click(Index As Integer)
LlenaSubMenu Estructura2, Index
End Sub

Private Sub Form_Load()
' **La variable "SpreadClon" contiene una copia de la grilla "vaSpread1"
' **tal como se inicio cuando se cargó el formulario
'Set SpreadClon = New vaSpread1

'*********-----> Identificacion que asigna el servidor SQL
        '** se mantiene mientras este abierto formulario
Dim RS As New ADODB.Recordset
Set RSSpid = vg_db.Execute("Select @@Spid")
If Not (RSSpid.EOF And RSSpid.BOF) Then spid = RSSpid.Fields(0)
RSSpid.Close: Set RSSpid = Nothing
'********---->Validar minuta en uso <-

Me.HelpContextID = vg_OpcM
Me.Height = 6765
Me.Width = 11055
fg_centra Me
MsgTitulo = "Planificación Teórica"
fg_carga ""

' Ejecuta el timer cada 1 segundo
Timer1.Interval = 1000
vg_TemSeg = 0
CorDes = 0
Label4.Caption = M_Plami1.fpayuda(0).Caption & "(" & M_Plami1.fpLongInteger1(0).Value & ")" & " - " & M_Plami1.fpayuda(1).Caption & " - " & M_Plami1.fpayuda(2).Caption & " - " & " Tipo: " & IIf(vg_IndpprSelec = "1", "Real", "Propuesta") & " - Zona : " & Trim(Mid(M_Plami1.Combo2(0).text, 1, 150))
Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "
indcos = False
estapo = False
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = " "
Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = "Grabar Datos": BtnX.Enabled = IIf(Mid(ValidarUsuario(M_Plami1), 2, 2) = "0", False, True)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Cortar", , tbrDefault, "A_Cortar"): BtnX.Visible = True: BtnX.ToolTipText = "Cortar"
Set BtnX = Toolbar1.Buttons.Add(, "A_Copiar", , tbrDefault, "A_Copiar"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar"
Set BtnX = Toolbar1.Buttons.Add(, "I_Pegar", , tbrDefault, "I_Pegar"): BtnX.Visible = True: BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, "A_Pegar", , tbrDefault, "A_Pegar"): BtnX.Visible = False: BtnX.ToolTipText = "Pegar"
'Set btnX = Toolbar1.Buttons.Add(, "I_PegadoEspecial", , tbrDefault, "I_PegadoEspecial"): btnX.Visible = True: btnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, "I_PegadoEspecial", , tbrDefault, "I_PegadoEspecial"): BtnX.Visible = True: BtnX.ToolTipText = ""  ' Activé visiblemente esta opcion (True)  02/09/09 Samuel Melendez
Set BtnX = Toolbar1.Buttons.Add(, "A_PegadoEspecial", , tbrDefault, "A_PegadoEspecial"): BtnX.Visible = False: BtnX.ToolTipText = "Pegado Especial"
Set BtnX = Toolbar1.Buttons.Add(, "A_Buscar", , tbrDefault, "A_Buscar"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar Recetas o Ingredientes"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): BtnX.Visible = True: BtnX.ToolTipText = "Insertar"
Set BtnX = Toolbar1.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): BtnX.Visible = True: BtnX.ToolTipText = "Eliminar"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_SubirF", , tbrDefault, "A_SubirF"): BtnX.Visible = True: BtnX.ToolTipText = "Subir"
Set BtnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): BtnX.Visible = True: BtnX.ToolTipText = "Bajar"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_VerReceta", , tbrDefault, "A_VerReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Ver Recetas"
Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Planificación Teórica"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Aportes", , tbrDefault, "A_Aportes"): BtnX.Visible = True: BtnX.ToolTipText = "Aportes Nutricionales x Días"
Set BtnX = Toolbar1.Buttons.Add(, "A_Costo", , tbrDefault, "A_Costo"): BtnX.Visible = True: BtnX.ToolTipText = "Visualizar Costo"
Set BtnX = Toolbar1.Buttons.Add(, "A_Frecuencia", , tbrDefault, "A_Frecuencia"): BtnX.Visible = True: BtnX.ToolTipText = "Frecuencia Recetas"
Set BtnX = Toolbar1.Buttons.Add(, "A_ExporReceta", , tbrDefault, "A_ExporReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Recetas"
Set BtnX = Toolbar1.Buttons.Add(, "A_ActCostoReceta", , tbrDefault, "A_ActCostoReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Planificación"
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Planificación Minuta a Excel "
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "Calorias", , tbrDefault, "Calorias"): BtnX.Visible = True: BtnX.ToolTipText = "Minuta con Calorias"
Set BtnX = Toolbar1.Buttons.Add(, "Ingrediente", , tbrDefault, "Ingrediente"): BtnX.Visible = True: BtnX.ToolTipText = "Frecuencia de Ingrediente"
Set BtnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): BtnX.Visible = True: BtnX.ToolTipText = "Deshacer"
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
vg_ActCalorias = False
DetallePlantillaMinuta

'-------> Llena Sub Menu estructura
Toolbar1.Buttons(31).Enabled = False
CargarListaMenu
 
'Estructura1(0).Visible = False
'Estructura2(0).Visible = False
If Mid(ValidaPerfil(M_Plami2), 1, 4) = "1000" = True Then BlocSoloAcceso
End Sub

Sub CargarListaMenu()
'-------> Llena Sub Menu estructura
Dim X As Long
For X = 1 To Estructura2.count - 1
    Unload Estructura2(X)
Next X
For X = 1 To Estructura1.count - 1
    Unload Estructura1(X)
Next X

Set RS = vg_db.Execute("sgpadm_s_estservicio 2, " & vg_codservicio & ",''")
If Not RS.EOF Then
    X = 1
    Do While Not RS.EOF
        Load Estructura1(X): Load Estructura2(X)
        Estructura1(X).Caption = Trim(RS!ess_nombre): Estructura2(X).Caption = Trim(RS!ess_nombre)
        Estructura1(X).HelpContextID = RS!ess_codigo: Estructura2(X).HelpContextID = RS!ess_codigo
        Estructura1(X).Enabled = True: Estructura2(X).Enabled = True
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Col = vaSpread1.maxcols: vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then
                If Val(vaSpread1.text) = RS!ess_codigo Then Estructura1(X).Enabled = False: Estructura2(X).Enabled = False
            End If
        Next
        X = X + 1
        RS.MoveNext
    Loop
End If
RS.Close: Set RS = Nothing
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 675
If Me.WindowState <> 1 Then vaSpread1.Move 0, 1560, ScaleWidth, ScaleHeight - 1560
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Man_Error
M_Plami1.DropTebleTmp (NameTemp)
If SwSalir <> 0 Then Exit Sub
If Toolbar1.Buttons(2).Visible = False Then Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Cancel = -1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
If Toolbar1.Buttons(2).Visible = True And Cancel <> -1 Then GrabarPlantillaMinuta
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
SwSalir = 1
vg_PartePlani = False
Me.Hide
Unload Me
Set SpreadClon = Nothing
M_Plami1.WindowState = 0
Man_Error:
End Sub

Private Sub Plantilla_Click(Index As Integer)
Dim RS As New ADODB.Recordset
Dim StrRec As String, StrRecb As String
Dim j As Long, i As Long, CodRec As Long, tiprec As Long
Dim cosali As Double, CosDes As Double
Dim desc As String
vg_RecetaReal = 0
estgra = False
Select Case Index
Case 0 '-------> Actualizar planificación
    If Toolbar1.Buttons(2).Enabled = False Then estgra = False: Exit Sub
    If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Cancel = -1: estgra = False: Exit Sub
    If ValidaEstructuras = False Then MsgBox "No puede grabar, si exiten recetas sin ser asignadas a una estructura": Exit Sub
    If Toolbar1.Buttons(2).Visible = True Then
       Toolbar1.Enabled = False
       Toolbar1.Buttons(31).Enabled = False
       CorDes = 0
       GrabarPlantillaMinuta
       CorDes = 0
       If Dir(LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6") <> "" Then Kill LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6"
       Toolbar1.Enabled = True
    End If
    vg_TemSeg = 0
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
Case 3 '-------> Visualizar costo
    If Frame2(0).Visible = True Then Frame2(0).Visible = False: vaSpread1.Move 0, 1480, ScaleWidth, ScaleHeight - 1480: estgra = False: Exit Sub
    vaSpread1.Move 0, 1480, ScaleWidth, ScaleHeight - 4080 '4000
    Frame2(0).Move 0, ScaleHeight - 2600, ScaleWidth, ScaleHeight - 1200
    Frame2(0).Visible = True
    CargarCosto
Case 5 '-------> Visualizar receta
    Dim xcol As Integer, auxtiprec  As Long
    vaSpread1.Row = vaSpread1.ActiveRow ': cand = vaSpread1.text
    vaSpread1.Col = vaSpread1.ActiveCol: desc = vaSpread1.text
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    vg_newestrec = True
    vg_modreceta = True
    xcol = 0
    For i = 1 To MaxColumna
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 4)) And Trim(vaSpread1.text) <> "" Then xcol = vectorcol(i): Exit For
    Next i
    If xcol = 0 Then MsgBox "No existe receta ha vizualizar", vbCritical + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    If vg_newestrec = True Then
       vg_fecval = 0: vg_fecval = Val(vg_fecha) & Right("0" & (Int(xcol / 6) + 1), 2)
       Set RS = vg_db.Execute("sgpadm_s_planifminuta 3, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & vg_fecval & ", 0, 0," & vg_IndpprSelec & "")
       If Not RS.EOF Then vg_fecval = RS!mid_fecval: vg_opcion = 2
       RS.Close: Set RS = Nothing
    End If
    vaSpread1.Col = xcol
    vaSpread1.Row = 0
    If vaSpread1.text = "R" Then
      vaSpread1.Col = vaSpread1.ActiveCol + 1: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 4
      StrRec = vaSpread1.text
    ElseIf vaSpread1.text = "N.Rac." Then
      vaSpread1.Col = vaSpread1.ActiveCol - 1: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 2
      StrRec = vaSpread1.text
    ElseIf vaSpread1.text = "Costo" Then
      vaSpread1.Col = vaSpread1.ActiveCol - 2: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 1
      StrRec = vaSpread1.text
    ElseIf vaSpread1.text = "Calorias" Then
      vaSpread1.Col = vaSpread1.ActiveCol - 1: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol - 1
      StrRec = vaSpread1.text
    Else
      vaSpread1.Row = vaSpread1.ActiveRow: desc = vaSpread1.text
      vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = xcol + 3
      StrRec = vaSpread1.text
    End If

    If Len(StrRec) <> 0 Then
       Do While InStr(StrRec, ";") <> 0
          StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
          StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
          vg_newcodrec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
          vg_newcodrec = IIf(vg_newcodrec = 0, BuscarCodReceta(desc), vg_newcodrec)
          vg_tiprec = Val(Mid(StrRecb, 1))
          vg_PartePlani = True
       Loop
    End If
    auxtiprec = vg_tiprec
    Vg_FechaDesde = vg_fecha
    Dim Receta As New M_Receta
    vg_RecetaReal = 1
    Receta.Show 1, Me
    Set Receta = Nothing

    vg_newestrec = False
    If vg_newcodrec <> 0 And Trim(vg_newnomrec) <> "" And vaSpread1.BackColor <> Shape1(1).FillColor And auxtiprec = vg_tiprec Then
        vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = vaSpread1.ActiveCol
        vaSpread1.Col = xcol + 3
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = xcol
        '-------> Limpiar Datos y Formato Celda
        vaSpread1.Action = 3
        '-------> Retorna Modo de la columna
        vaSpread1.BlockMode = False
        vaSpread1.Font.Bold = False
        vaSpread1.Font.Size = 8
        vaSpread1.text = vg_newnomrec
        
        vaSpread1.Col = xcol + 2
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 1
        '-------> Calcular costo alimentación y deshechable
        cosali = Format(fg_CalCtoRecListaPrecio(Val(vg_newcodrec), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
        CosDes = Format(fg_CalCtoRecListaPrecio(Val(vg_newcodrec), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
        vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
        
        vaSpread1.Col = xcol + 3
        vaSpread1.text = vg_newcodrec & "&" & vg_tiprec & "&;"
        
        '-------> Revizar si existe receta iguales en el mes y actualizar
        For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 5
            vaSpread1.Col = i + 4
            For j = 1 To (vaSpread1.MaxRows - 1)
                vaSpread1.Row = j: CodRec = 0
                vaSpread1.Col = i + 1
                If vaSpread1.BackColor = Shape1(1).FillColor Then Exit For
                vaSpread1.Col = i + 4
                If Trim(vaSpread1.text) <> "" Then
                   StrRec = vaSpread1.text
                   If Len(StrRec) <> 0 Then
                      Do While InStr(StrRec, ";") <> 0
                         StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                         StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                         CodRec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                         tiprec = Val(Mid(StrRecb, 1))
                      Loop
                   End If
                   If CodRec = vg_newcodrec Then
                      vaSpread1.Col = i + 4
                      vaSpread1.text = vg_newcodrec & "&" & vg_tiprec & "&;"
                      
                      vaSpread1.Col = i + 3
                      vaSpread1.CellType = 5
                      vaSpread1.TypeHAlign = 1
                      vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
                   End If
                End If
            Next j
        Next i
        If indcos = True Then
           For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        '-------> Actualizar lista receta
        If B_Receta.vaSpread1.MaxRows > 0 Then
            B_Receta.vaSpread1.Row = B_Receta.vaSpread1.SearchCol(1, -1, B_Receta.vaSpread1.MaxRows, Val(vg_newcodrec), SearchFlagsEqual)
            B_Receta.vaSpread1.Col = 3: B_Receta.vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
        End If
        vg_newcodrec = 0: vg_newnomrec = "": vg_tiprec = -1
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    End If
    If indcos = True Then Me.Refresh: Toolbar1.Refresh: Frame2(0).Refresh: Frame2(1).Refresh: Frame2(2).Refresh: Frame2(3).Refresh: Frame2(4).Refresh
    vg_newcodrec = 0
Case 8 '-------> Copiar planificación
    M_CPlaTe.Show 1, Me
Case 10 '-------> Visualizar aportes x día
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    j = 0
    For i = 1 To MaxColumna
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then j = vectorcol(i): Exit For
    Next i
    vaSpread1.Col = j: vaSpread1.Row = 0
    C_ApoPla.LlenarApoPlan Me, "Aporte Planificación Real " & vaSpread1.text, vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha), 2, j
    C_ApoPla.Show 1, Me
Case 11 '-------> Visualizar frecuencia de recetas
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    Call C_FrePla.LlenarFrecPlan("Frecuencia Planificación " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha), 1, 0)
    C_FrePla.Show 1, Me
Case 13 '-------> Actualizar costo recetas y planificación
    If IndGrabado = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    fg_carga ""
    '-------> Rutina actualizar precio planificación
    vg_db.Execute "sgpadm_p_actuaplanif " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_codlpr & ", " & Val(vg_fecha) & ""
    Dim vecactrec As Variant
    '-------> Traer total de receta desde planificación de minutas y luego calcular costo
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 11, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & "," & Val(vg_fecha) & ", 0,0," & vg_IndpprSelec & "")
    If RS.EOF Or RS!nReg < 1 Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    ReDim vecactrec(RS!nReg, 4)
    RS.Close: Set RS = Nothing
    For i = 1 To UBound(vecactrec)
        DoEvents
        vecactrec(i, 1) = 0 '-------> codigo receta
        vecactrec(i, 2) = 0 '-------> tipo receta
        vecactrec(i, 3) = 0 '-------> costo receta alimentación
        vecactrec(i, 4) = 0 '-------> costo receta desechable
    Next i
    i = 1
    Dim IndDia As Long
    gauge1.Value = 0: gauge.Value = 0: Fecha = 0: IndDia = 1: Fecha = 0: cosali = 0: CosDes = 0
    Picture1.Visible = True: Label2.Visible = False: Label3.Visible = True: Label3.Caption = "Recopilando información, un momento....": gauge.Visible = True: gauge.Visible = False
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 12, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & " , " & Val(vg_fecha) & ", 0,0," & vg_IndpprSelec & " ")
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    Do While Not RS.EOF
       DoEvents
       vecactrec(i, 1) = RS!mid_codrec
       vecactrec(i, 2) = RS!mid_tiprec
       vecactrec(i, 3) = Format(fg_CalCtoRecListaPrecio(Val(RS!mid_codrec), RS!mid_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))       'Format(IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec), fg_Pict(6, 2))
       vecactrec(i, 4) = Format(IIf(IsNull(RS!mid_cosdes), 0, RS!mid_cosdes), fg_Pict(6, 2))
       
       RS.MoveNext: i = i + 1
    Loop
    RS.Close: Set RS = Nothing
    
    gauge1.Value = 0: gauge.Value = 0: Fecha = 0: IndDia = 1: Fecha = 0: cosali = 0: CosDes = 0
    Picture1.Visible = True: Label2.Visible = False: Label3.Visible = True: Label3.Caption = "Actualizando costo receta, en planificación": gauge.Visible = True: gauge.Visible = False
    For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
        DoEvents
        gauge1.Value = Val((i / vaSpread1.maxcols) * 100)
        existedat = 0
        vaSpread1.Row = 1: vaSpread1.Col = i
        Fecha = Val(vg_fecha) & fg_pone_cero(IndDia, 2)
        If vaSpread1.BackColor <> Shape1(1).FillColor Then
           For j = 1 To (vaSpread1.MaxRows - 1)
               vaSpread1.Row = j
               vaSpread1.Col = i + 1
               If Trim(vaSpread1.text) <> "" Then existedat = 1: Exit For
           Next j
           If existedat > 0 Then
              For j = 1 To (vaSpread1.MaxRows - 1)
                  vaSpread1.Row = j: vaSpread1.Col = i + 1: CodRec = 0
                  If Trim(vaSpread1.text) <> "" Then
                    vaSpread1.Col = i + 4: StrRec = vaSpread1.text
                    If Len(StrRec) <> 0 Then
                       Do While InStr(StrRec, ";") <> 0
                          StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                          StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                          CodRec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                          tiprec = Val(Mid(StrRecb, 1))
                       Loop
                    End If
                    vaSpread1.Col = i + 3
                    '-------> Traer costo alimentación y desechables
                    For X = 1 To UBound(vecactrec)
                        If CodRec = vecactrec(X, 1) And tiprec = vecactrec(X, 2) Then
                           cosali = vecactrec(X, 3)
                           CosDes = vecactrec(X, 4)
                           Exit For
                        End If
                    Next
                    vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
                    IndGrabado = 1
                  End If
              Next j
           End If
        End If
        IndDia = IndDia + 1
    Next i
    Label2.Visible = True: Picture1.Visible = False: gauge.Visible = False
    vaSpread1.Refresh
    If IndGrabado = 1 Then fg_descarga: MsgBox "Actualización costo receta finalizado sin problema, luego grabe información", vbInformation + vbOKOnly, MsgTitulo: Plantilla(0).Enabled = True: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True: estgra = False: Exit Sub
    fg_descarga
Case 14 '-------> Visualizar detalle de recetas
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    C_ExpRec.LlenarExporReceta "Exportar Recetas " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha)
    C_ExpRec.Show 1, Me
Case 16 '-------> Parámetro de grabado
    M_ParGra.Show 1, Me
Case 20
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    C_IngPla.LlenarFrecIng "Frecuencia Planificación Ingrediente " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codsubseg, vg_codregimen, vg_codservicio, Val(vg_fecha), 1
    C_IngPla.Show 1, Me
Case 21
    Deshacer "Spread" & vg_NUsr & CorDes & ".ss6"
    If CorDes < 1 Then: Toolbar1.Buttons(31).Visible = True: Toolbar1.Buttons(31).Enabled = False
Case 22 '-------> Salir
    vg_PartePlani = False
    SwSalir = 0
    If Toolbar1.Buttons(2).Visible = False Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0: estgra = False: Exit Sub
    If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
    If Toolbar1.Buttons(2).Visible = True Then GrabarPlantillaMinuta
    SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
End Select
estgra = False
End Sub

Private Sub Plato_Click(Index As Integer)
On Error GoTo Man_Error
Dim RS As New ADODB.Recordset
If Toolbar1.Buttons(2).Enabled = False Then estgra = False: Exit Sub
Dim Del_Row As Integer, IndCol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer, indrow3 As Long, xx As Long
Dim Col As Long, fil As Long, codest As Long, cosali As Double, CosDes As Double
Dim VecSelGrid As Variant: Dim VecRacPegar As Variant
Dim contador, contador_b, cantCol As Integer, LargoVec As Integer
Dim accion As String
Dim ColumnaActiva, FilaActiva, ColumnaAntActiva, n, n1, NFilas As Integer
contador = 0: contador_b = 0: cantCol = 0: LargoVec = 0:  accion = "": n1 = 0: n = 0: NFilas = 0
estgra = True
Select Case Index
Case 2 '-------> Ingresa recetas
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Then Exit Sub
    iblockcol = vaSpread1.ActiveCol: aiblockcol = vaSpread1.ActiveCol
    iblockcol2 = vaSpread1.ActiveCol: aiblockcol2 = vaSpread1.ActiveCol
    iblockrow = vaSpread1.ActiveRow: aiblockrow = vaSpread1.ActiveRow
    iblockrow2 = vaSpread1.ActiveRow: aiblockrow2 = vaSpread1.ActiveRow
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Col = vaSpread1.ActiveCol + 4: vaSpread1.Row = 0:
    If vaSpread1.text = "Calorias" Then
       If vaSpread1.ColHidden = False Then vg_ActCalorias = True Else vg_ActCalorias = False
    End If
    
    vg_RecetaReal = 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
    
    j = 0
    For i = 1 To MaxColumna
        If vaSpread1.Col = vectorcol(i) Then j = vectorcol(i): Exit For
    Next i
    If j = 0 Then estgra = False: Exit Sub
    vg_codigo = "": vg_nombre = "": vg_tiprec = -1
    vaSpread1.Row = vaSpread1.ActiveRow
    B_Receta.vaSpread1.Col = 6
    If vg_ActCalorias = True Then
       B_Receta.vaSpread1.ColHidden = False
    Else
       B_Receta.vaSpread1.ColHidden = True
    End If
    B_Receta.Show 1, Me
    
    vg_RecetaReal = 0
    B_Receta.vaSpread1.Col = 6
    If vg_ActCalorias = True Then
       B_Receta.vaSpread1.ColHidden = False
    Else
       B_Receta.vaSpread1.ColHidden = True
    End If

    If Trim(vg_codigo) = "" Or Trim(vg_nombre) = "" Then estgra = False: Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = j - 1
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 2
    vaSpread1.Value = "R"
    vaSpread1.ForeColor = &HFF&
    vaSpread1.BackColor = &H80FF80
    
    GrabarCambios 1, 1, ""
    
    vaSpread1.Col = j
    '-------> Limpiar Datos y Formato Celda
    vaSpread1.Action = 3
    '-------> Retorna Modo de la columna
    vaSpread1.BlockMode = False
    vaSpread1.Font.Bold = False
    vaSpread1.Font.Size = 8

    vaSpread1.text = vg_nombre
    
    vaSpread1.Col = j + 1
    If Trim(vaSpread1.text) = "" Then
       '-------> Asignar raciones estimadas
       codest = 0
       vaSpread1.Row = vaSpread1.ActiveRow
       For i = (IIf(vaSpread1.Row = 1, 1, vaSpread1.Row + 1 - 1)) To 1 Step -1
           vaSpread1.Row = i
           vaSpread1.Col = 1
           If Trim(vaSpread1.text) <> "" Then vaSpread1.Col = vaSpread1.maxcols: codest = Val(vaSpread1.text): Exit For
       Next i
       Set RS = vg_db.Execute("SELECT * FROM a_estservicio With(NoLock) WHERE ess_codser=" & vg_codservicio & " AND ess_codigo=" & codest & "")
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = j + 1
       vaSpread1.CellType = 3
       vaSpread1.TypeIntegerMin = 1
       vaSpread1.TypeIntegerMax = 9999999
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.text = IIf(RS.EOF, 0, RS!ess_racmin)
       vaSpread1.ForeColor = &HFF0000
       RS.Close: Set RS = Nothing
    End If
    
    vaSpread1.Col = j + 2
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 1
    '------> Calcular costo planificación alimento y desechable
    cosali = 0: CosDes = 0
    cosali = Format(fg_CalCtoRecListaPrecio(Val(vg_codigo), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
    CosDes = Format(fg_CalCtoRecListaPrecio(Val(vg_codigo), vg_tiprec, vg_codlpr, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")), Val(vg_fecha)), fg_Pict(6, 2))
    vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
    
    vaSpread1.Col = j + 3
    vaSpread1.Col = vaSpread1.maxcols - 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.text = EstructuraSuperior(vaSpread1, vaSpread1.Row)

    vaSpread1.Col = j + 3
    vaSpread1.text = Val(vg_codigo) & "&" & vg_tiprec & "&;"

    If indcos = True Then Calctodia vaSpread1.Row, j
    
    RS1.Open "sgpadm_s_AporteNutricionales 2," & IIf(CodRec = 0, BuscarCodReceta(vg_nombre), CodRec) & "," & vg_codsubseg & "," & vg_codregimen & ", " & vg_Zona & "", vg_db, adOpenForwardOnly ', adOpenStatic
    If Not RS1.EOF Then
      vg_Calorias = RS1!candiet
    End If
    RS1.Close: Set RS1 = Nothing
    
    vaSpread1.Col = j + 4
    vaSpread1.Row = 0
    If vaSpread1.text = "Calorias" Then
      vaSpread1.Col = j + 4
      vaSpread1.Row = vaSpread1.ActiveRow
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = Format(vg_Calorias, fg_Pict(9, 2))

    End If
    
    If Mid(ValidaPerfil(M_Plami2), 1, 4) = "1000" = True Then
        BlocSoloAcceso
    Else
        vaSpread1.Row = iblockrow
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    End If
Case 5 '-------> Insertar linea
    ' se agregó esta asignacion a estas variables, para indicarle la seleccion de las celdas
    vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
    GrabarCambios Val(xRowIni), Val(xRowFin), "Insertar"
    vaSpread1.Enabled = False
    IndCol = iblockcol
    iblockcol = 1: iblockcol2 = vaSpread1.maxcols
    vaSpread1.MaxRows = vaSpread1.MaxRows + ((xRowFin - xRowIni) + 1) '1
    vaSpread1.InsertRows xRowIni, ((xRowFin - xRowIni) + 1)
    
    If vg_IndpprSelec <> "2" Then
       For i = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
           vaSpread1.Row = 0: vaSpread1.Col = i
           If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
              Dim f As Long, c As Long
              For c = i - 1 To i + 2
                  vaSpread1.Row = xRowIni: vaSpread1.Col = c
                  vaSpread1.BackColor = Shape1(1).FillColor
              Next c
           End If
       Next i
    End If
    iblockcol = IndCol
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    fg_descarga
    vaSpread1.Enabled = True
Case 6 '-------> Eliminar línea
    vaSpread1.Enabled = False
    Dim x_iblockrow As Variant, x_iblockrow2 As Variant, x_iblockcol As Variant, x_iblockcol2 As Variant
    
    vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
    
    '-- INICIO -- ULTIMA FILA
        'Saca de la seleccion la ultima fila cuando se encuantra seleccionada por el usuario para borrar
      If xRowFin = vaSpread1.MaxRows Then xRowFin = xRowFin - 1
    '-- FIN --  ULTIMA FILA
  
    
    ' se agregó esta asignacion a estas variablea, las cuales corresponden
    ' al rango de celda seleccionado, ya que como se estaban asignando
    ' a veces se producia inconsistencias en la asignacion devolviendo rangos malos
    iblockrow = xRowIni
    iblockrow2 = xRowFin
    iblockcol = xColIni
    iblockcol2 = xcolfin
    '*********************------------------
    
    fg_carga ""
    IndCol = iblockcol
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.maxcols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    
    NFilas = (aiblockrow2 - aiblockrow) + 1
    
    If NFilas > 1 Then
        'debido a que las siguientes variable cambian de valor, al salir el mensaje
        ' al usuario, debido a que se ejecuta un evento que las cambia al
        ' perder el foco, aqui se intenta rescatar su valor en variables de paso
        ' para una vez mostrado el mensaje, se les devuelva su valor anterior

    
        If MsgBox("Cuando se intenta eliminar mas de una fila, no es posible recuperar la informacion contenida en ella mediante la opcion deshacer  ¿Desea Continuar? ", vbInformation + vbYesNo) = vbNo Then
                fg_descarga
                vaSpread1.Enabled = True
                Exit Sub
        End If
        
        For i = xRowIni To xRowFin
            vaSpread1.Col = 1
            vaSpread1.Row = i
            DesqloqSubMenu (vaSpread1.text)
        Next i
        'aqui se recupera su valor anterior
        iblockrow = xRowIni
        iblockrow2 = xRowFin
        iblockcol = xColIni
        iblockcol2 = xcolfin
    End If
    
    If vaSpread1.BackColor = Shape1(1).FillColor And Trim(vaSpread1.text) <> "" Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Or Trim(vaSpread1.text) = "" Then GoTo Paso
    'Validación Dias Bloqueados
    If vg_IndpprSelec <> "2" Then
       If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
          For i = 1 To MaxColumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Existen Días Bloqueado, utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For j = iblockrow To iblockrow2
                 vaSpread1.Row = j
                 If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Bloque seleccionado existen días bloqueado, utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
              Next j
          Next i
       End If
    End If
    For i = 1 To MaxColumna
        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
    Next i
    For i = 1 To MaxColumna
        If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
        If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
    Next i
    IndCol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    If indcos = True Then
       For i = iblockcol To iblockcol2 Step 6
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
      
    End If
    '-------> Fin validar días modificados
    iblockcol = AuxCol
    vaSpread1.BlockMode = False
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indactivo = 0
Paso:
    vaSpread1.Row = vaSpread1.ActiveRow
    'Validación Dias Bloqueados
    If vg_IndpprSelec <> "2" Then
       If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
          For i = 1 To MaxColumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Existen Días Bloqueado, utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For j = iblockrow To iblockrow2
                  vaSpread1.Row = j
                  If vaSpread1.BackColor = Shape1(1).FillColor And Shape1(1).FillColor = &H8080FF Then MsgBox "Bloque seleccionado existen días bloqueado,utilizar opción Suprimir para eliminar recetas.", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
              Next j
          Next i
       End If
    End If
    vaSpread1.Row = iblockrow2
    SpreadClon.vaSpread1.Row = iblockrow2
    vaSpread1.Col = iblockcol
    SpreadClon.vaSpread1.Col = iblockcol
    
    GrabarCambios Val(iblockrow), Val(NFilas), "Eliminar"
   
'Esta función pega datos de una grila a otra
'    'Single Block Selected
'    Dim array1, array2 As Long
'    'Get the size of the block
'    array1 = vaSpread1.SelBlockRow2 - vaSpread1.SelBlockRow
'    array2 = vaSpread1.SelBlockCol2 - vaSpread1.SelBlockCol
'    'Init array size
'    ReDim fparray(array1, array2) As Variant
'    'Get data: ColLeft, RowTop
'    vaSpread1.GetArray vaSpread1.SelBlockCol, vaSpread1.SelBlockRow, fparray
'    'Display the selected data
'    vaSpread2.SetArray 1, 1, fparray
    
    
    vaSpread1.DeleteRows iblockrow, NFilas
    SpreadClon.vaSpread1.DeleteRows iblockrow, NFilas
    vaSpread1.MaxRows = vaSpread1.MaxRows - NFilas
    SpreadClon.vaSpread1.MaxRows = vaSpread1.MaxRows - NFilas
    vaSpread1.Col = vaSpread1.Row
    vaSpread1.Visible = True
    SpreadClon.vaSpread1.Visible = True
    
    iblockcol = IndCol
    If vg_IndpprSelec <> "2" Then
       For i = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
           vaSpread1.Row = 0: vaSpread1.Col = i
           If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
               For Col = 0 To i - 4
                   vaSpread1.Row = (vaSpread1.MaxRows - 1): vaSpread1.Col = Col + 2
                   vaSpread1.BackColor = Shape1(1).FillColor
               Next Col
           End If
       Next i
    End If
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    fg_descarga
    vaSpread1.Enabled = True
Case 8 '-------> Subir línea
    vaSpread1.Enabled = False
    
    vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
    iblockrow = xRowIni
    iblockrow2 = xRowFin
    iblockcol = xColIni
    iblockcol2 = xcolfin
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = 1 Or vaSpread1.Row = vaSpread1.MaxRows Then estgra = False: vaSpread1.Enabled = True: Exit Sub
    If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
       For i = 1 To MaxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For j = iblockrow To iblockrow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
           Next j
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col > 1 Then
        IndCol = iblockcol
        vaSpread1.Col = 1
        If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        If (iblockrow - ((iblockrow2 - iblockrow) + 1)) < 1 Then
           MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        End If
        If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.maxcols
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        GrabarCambios vaSpread1.Row, 1, "Subir Linea"
        '-------> Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = True
        vaSpread1.MoveRange iblockcol, (iblockrow - 1), iblockcol2, (iblockrow - 1), iblockcol, vaSpread1.MaxRows
        
        '-------> Copiar datos fila seleccionada
        vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow - 1), False
        vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow - 1)
        
        '---> SpreadClon
        SpreadClon.vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow - 1), False
        SpreadClon.vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow - 1)
        '---
        
        '-------> Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
        
        '-------> Devolver datos fila y restar ultima fila
        SpreadClon.vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        SpreadClon.vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
       
        vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
        vaSpread1.DeleteRows vaSpread1.MaxRows, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.text) = "" Then estgra = False: vaSpread1.Enabled = True: Exit Sub
        For i = iblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next i
        For z = iblockrow + 1 To (vaSpread1.MaxRows - 1) 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then
            For fil = (vaSpread1.MaxRows - 1) To 1 Step -1
                For Colu = 1 To vaSpread1.maxcols
                    vaSpread1.Col = Colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next Colu
                If z <= (vaSpread1.MaxRows) Then Exit For
            Next fil
        End If
        FilaAct = iblockrow         'Fila actual
        FilaAnt = IIf(i < 1, 1, i)  'Fila anterior
        FilaPos = z                 'Fila posterior
        
        '------- Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (FilaAct - FilaAnt)
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows + (FilaAct - FilaAnt)
            vaSpread1.Row = i
            vaSpread1.RowHidden = True
        Next i
        vaSpread1.MoveRange 1, FilaAnt, vaSpread1.maxcols, (FilaAct - 1), 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1
        
        '-------> Mover estructura
        vaSpread1.MoveRange 1, FilaAct, vaSpread1.maxcols, (FilaPos - 1), 1, FilaAnt
        
        '-------> Devolver respaldo
        vaSpread1.MoveRange 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1, vaSpread1.maxcols, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 + (FilaAct - FilaAnt - 1), 1, FilaAnt + (FilaPos - FilaAct)
        
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows
            vaSpread1.DeleteRows i, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Next i
        vaSpread1.SetActiveCell 1, FilaAnt
    End If
    vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    Plato(14).Enabled = False
    OpGrilla(14).Enabled = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    vaSpread1.Col = 1
    For i = 1 To (vaSpread1.MaxRows - 1)
        vaSpread1.Row = i
        vaSpread1.BackColor = Shape1(2).FillColor
    Next i
    vaSpread1.Enabled = True
Case 9 '-------> Bajar línea
    vaSpread1.Enabled = False
    vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
    
    iblockrow = xRowIni
    iblockrow2 = xRowFin
    iblockcol = xColIni
    iblockcol2 = xcolfin
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = vaSpread1.MaxRows Then estgra = False: vaSpread1.Enabled = True: Exit Sub
    If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
       For i = 1 To MaxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For j = iblockrow To iblockrow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
           Next j
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    '-------> Grabar Evento
    GrabarCambios vaSpread1.Row, j, "Bajar Linea"
    If vaSpread1.Col > 1 Then
        vaSpread1.Col = 1
        vaSpread1.Row = vaSpread1.ActiveRow + 1
        If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow - 1
        If (iblockrow2 + ((iblockrow2 - iblockrow) + 1)) > (vaSpread1.MaxRows - 1) Then
           MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, MsgTitulo: estgra = False: vaSpread1.Enabled = True: Exit Sub
        End If
        IndCol = iblockcol
        If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.maxcols
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        '-------> Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = True
        vaSpread1.MoveRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), iblockcol, vaSpread1.MaxRows
    
        '-------> Copiar datos fila Seleccionada
        vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
        vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
                
        '------->  SpreadClon
        SpreadClon.vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
        SpreadClon.vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
       
        '-------> Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
        
        '------->  SpreadClon
        SpreadClon.vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        SpreadClon.vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow

        vaSpread1.DeleteRows vaSpread1.MaxRows, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.Row = iblockrow + 1: vaSpread1.Col = iblockcol
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.text) = "" Then estgra = False: vaSpread1.Enabled = True: Exit Sub
        For z = iblockrow + 1 To (vaSpread1.MaxRows - 1) 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then estgra = False: vaSpread1.Enabled = True: Exit Sub
        vaSpread1.Col = vaSpread1.ActiveCol
        AuxIblockrow = z
        For i = AuxIblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next i
        For z = AuxIblockrow + 1 To (vaSpread1.MaxRows - 1) 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then
            For fil = (vaSpread1.MaxRows - 1) To 1 Step -1
                For Colu = 1 To vaSpread1.maxcols
                    vaSpread1.Col = Colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next Colu
                If z <= (vaSpread1.MaxRows - 1) Then Exit For
            Next fil
        End If
        FilaAct = AuxIblockrow         'Fila actual
        FilaAnt = IIf(i < 1, 1, i)  'Fila anterior
        FilaPos = z                 'Fila posterior
        
        '-------> Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (FilaAct - FilaAnt)
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows + (FilaAct - FilaAnt)
            vaSpread1.Row = i
            vaSpread1.RowHidden = True
        Next i
        vaSpread1.MoveRange 1, FilaAnt, vaSpread1.maxcols, (FilaAct - 1), 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1
        
        '-------> Mover estructura
        vaSpread1.MoveRange 1, FilaAct, vaSpread1.maxcols, (FilaPos - 1), 1, FilaAnt
        
        
        '-------> Devolver respaldo
        vaSpread1.MoveRange 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1, vaSpread1.maxcols, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 + (FilaAct - FilaAnt - 1), 1, FilaAnt + (FilaPos - FilaAct)
        

        
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows
            vaSpread1.DeleteRows i, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Next i
        vaSpread1.SetActiveCell 1, FilaAnt + (FilaPos - FilaAct)
    End If
    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    Plato(14).Enabled = False
    OpGrilla(14).Enabled = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    vaSpread1.Col = 1
    For i = 1 To (vaSpread1.MaxRows - 1)
        vaSpread1.Row = i
        vaSpread1.BackColor = Shape1(2).FillColor
    Next i
    vaSpread1.Enabled = True
Case 11, 12 '-------> Copiar y pegar linea
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Then estgra = False: Exit Sub
    If Index = 11 Then
       If iblockcol < 1 Then
          For i = 1 To MaxColumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For j = iblockrow To iblockrow2
                 vaSpread1.Row = j
                 If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
              Next j
          Next i
       End If
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    '------> Verificar si copiar receta o raciones solamente
    vaSpread1.Row = 0
    If vaSpread1.text = "N.Rac." Then
      TipoCopia = "Copiar Raciones"
    Else
      TipoCopia = "Copiar Receta"
    End If
       
    aiblockrow = iblockrow: aiblockrow2 = iblockrow2
    aiblockcol = iblockcol: aiblockcol2 = iblockcol2
    
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(7).Visible = True
       
    Plato(13).Enabled = True: OpGrilla(13).Enabled = True
    Plato(14).Enabled = True: OpGrilla(14).Enabled = True
    If iblockcol < 1 Then aiblockcol = 2: aiblockcol2 = vaSpread1.maxcols
    indcortarpegar = 1
    If Index = 11 Then
       indcortarpegar = 0
       Toolbar1.Buttons(8).Visible = True
       Toolbar1.Buttons(9).Visible = False
       Plato(14).Enabled = False
       OpGrilla(14).Enabled = False
    Else
       Toolbar1.Buttons(8).Visible = False
       Toolbar1.Buttons(9).Visible = True ' Cambié opcion a "True"  02/09/09 Samuel Melendez
       Plato(14).Enabled = True
       OpGrilla(14).Enabled = True
    End If
Case 13, 14 '-------> Copiar y pegar
    If indcortarpegar = 0 Then
       If (iblockcol2 - iblockcol) > (aiblockcol2 - aiblockcol) Or (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then MsgBox "Imposible Pegar la infomación ya que el área de Cortar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
       indcortarpegar = 0
    Else
       If (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
       If iblockcol2 + (aiblockcol2 - aiblockcol) > (vaSpread1.maxcols - MaxColumna) Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub   'aiblockcol <> iblockcol2 Or aiblockcol = 1 Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    End If
    If iblockcol < 1 Then
       For i = 1 To MaxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For j = iblockrow To iblockrow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor And Index <> 14 Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
           Next j
       Next i
    End If
    
    vaSpread1.Col = 1
    If vaSpread1.text = "Comensales" Then estgra = False: Exit Sub ' Valida que no se peguen recetas en la Línea de Comensales.
    
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then estgra = False: Exit Sub
    If indcortarpegar = 0 Then Toolbar1.Buttons(6).Visible = True: Toolbar1.Buttons(7).Visible = False
    '-------> Destinacion de copiar y pegar datos
    If iblockcol < 1 Then
       iblockcol = 2: iblockcol2 = vaSpread1.maxcols
    End If
    
    If aiblockcol2 = vaSpread1.maxcols Then aiblockcol2 = vaSpread1.maxcols - 1
    If aiblockcol2 = (vaSpread1.maxcols - MaxColumna - 1) Then aiblockcol2 = (vaSpread1.maxcols - MaxColumna - 1)
    vaSpread1.Row = 0: vaSpread1.Col = iblockcol

    vaSpread1.Row = 0
    If vaSpread1.text = "N.Rac." And TipoCopia = "Copiar Raciones" Then
       If (aiblockrow2 - aiblockrow) + iblockrow2 >= vaSpread1.MaxRows Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo: estgra = False: Exit Sub
        cantCol = aiblockcol2 - aiblockcol
        CantCol1 = iblockcol2 - iblockcol
    Else
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For  'Inicio pegar
        Next i
        For i = 1 To MaxColumna
             If (vectorcol(i) - 1) = iblockcol2 Or vectorcol(i) = iblockcol2 Or (vectorcol(i) + 1) = iblockcol2 Or (vectorcol(i) + 2) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 3)): Exit For ' Fin pegar
        Next i
    
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = aiblockcol Or vectorcol(i) = aiblockcol Or (vectorcol(i) + 1) = aiblockcol Or (vectorcol(i) + 2) = aiblockcol Then aiblockcol = (vectorcol(i) - 1): Exit For ' Columna de inicio copia
        Next i
        For i = 1 To MaxColumna
'            If (vectorcol(i) - 1) = aiblockcol2 Or vectorcol(i) = aiblockcol2 Or (vectorcol(i) + 1) = aiblockcol2 Or (vectorcol(i) + 2) = aiblockcol2 = (vectorcol(i) + 1) Or aiblockcol2 = (vectorcol(i) + 3) Then aiblockcol2 = (vectorcol(i) + 3): Exit For 'Fin copia   aiblockcol2 = (vectorcol(i) + 3) Copiaba hasta el cod receta ahora hasta Calorias
            If (vectorcol(i) - 1) = aiblockcol2 Or vectorcol(i) = aiblockcol2 Or (vectorcol(i) + 1) = aiblockcol2 Or (vectorcol(i) + 2) = aiblockcol2 Or aiblockcol2 = (vectorcol(i) + 3) Then aiblockcol2 = (vectorcol(i) + 4): Exit For  'Fin copia   aiblockcol2 = (vectorcol(i) + 3) Copiaba hasta el cod receta ahora hasta Calorias
        Next i
        
        cantCol = aiblockcol2 - aiblockcol
        CantCol1 = iblockcol2 - iblockcol
    End If
    '-----> Llena vectores con las raciones
    LargoVec = aiblockrow2 - aiblockrow + 1
    If aiblockcol > 1 And aiblockrow > 0 Then
       ReDim VecSelGrid(0)
       ReDim VecSelGrid(20000)
       For i = aiblockcol To aiblockcol2
           vaSpread1.Col = i
           vaSpread1.Row = 0
           d = vaSpread1.text
           If vaSpread1.text = "N.Rac." Then
              For j = aiblockrow To aiblockrow + LargoVec - 1
                  vaSpread1.Col = i: vaSpread1.Row = j: d = vaSpread1.text
                  contador = contador + 1
                  If Trim(vaSpread1.text) <> "" Then VecSelGrid(contador) = vaSpread1.text    ' Almacena las raciones a copiar
              Next j
           End If
       Next i
    End If
    
    If vaSpread1.ActiveCol > 1 And vaSpread1.ActiveRow > 0 Then
       ReDim VecRacPegar(0)
       ReDim VecRacPegar(20000)
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           vaSpread1.Row = 0
           If vaSpread1.text = "N.Rac." Then
              For j = vaSpread1.ActiveRow To vaSpread1.ActiveRow + contador - 1 'vaSpread1.MaxRows - 1
                  vaSpread1.Col = i: vaSpread1.Row = j
                  contador_b = contador_b + 1
                  If Trim(vaSpread1.text) <> "" Then VecRacPegar(contador_b) = vaSpread1.text ' Almacena las raciones a reemplazar
              Next j
           End If
       Next i
    End If
    
    IndCol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    If Index = 14 And indcortarpegar = 1 Then
'       If (aiblockrow2 - aiblockrow) <> 0 Or (aiblockcol2 - aiblockcol) <> 4 Then MsgBox "Por esta opción solamente puede copiar una receta", vbInformation + vbOKOnly, Msgtitulo: iblockcol = vaSpread1.ActiveCol: iblockcol2 = indcol2: estgra = False: Exit Sub
       If (aiblockrow2 - aiblockrow) <> 0 Or (aiblockcol2 - aiblockcol) <> 5 Then MsgBox "Por esta opción solamente puede copiar una receta", vbInformation + vbOKOnly, MsgTitulo: iblockcol = vaSpread1.ActiveCol: iblockcol2 = indcol2: estgra = False: Exit Sub
       '-------> Rutina pegado especial
       Dim nrodia As String
       vaSpread1.Row = 0: nrodia = ""
       For i = aiblockcol To aiblockcol2 Step 6
           vaSpread1.Col = i + 1
           nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
       Next i
       For i = 1 To MaxColumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then vaSpread1.Row = 0: nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
       Next i
       
       vg_codigo = ""
       M_CpRPla.Inicio "Copia Especial Recetas en Planificación Real", "PLAREA", vg_fecha, nrodia
       M_CpRPla.Show 1
       If Trim(vg_codigo) = "" Then
          iblockcol = vaSpread1.ActiveCol: iblockcol2 = indcol2
          estgra = False: Exit Sub
       End If
       '-------> Grabar Evento Pegado Especial
       GrabarCambios 1, 1, "Pegado Especial"
       Dim VecDia() As String
       Dim xser As Long, iser As Long
       '-------> Mover días no permitidos
       ReDim Preserve VecDia(0)
       ValLcntH = "": i = 0
       For j = 1 To Len(vg_codigo)
           If Asc(Mid(vg_codigo, j, 1)) <> 59 Then
              ValLcntH = ValLcntH + Mid(vg_codigo, j, 1)
           Else
              ReDim Preserve VecDia(i): VecDia(i) = ValLcntH: ValLcntH = "": i = i + 1
           End If
       Next j
       If Trim(ValLcntH) <> "" Then ReDim Preserve VecDia(i): VecDia(i) = ValLcntH
       vaSpread1.Enabled = False
       fg_carga ""
       For i = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
           vaSpread1.Row = aiblockrow
           vaSpread1.Col = vaSpread1.maxcols
           iser = Val(vaSpread1.text)
           vaSpread1.Row = 0
           vaSpread1.Col = i
           L = 0
           nrodia = Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2))
           For j = 0 To UBound(VecDia)
               If nrodia = VecDia(j) Then
                  vaSpread1.Row = aiblockrow: vaSpread1.Col = i - 1
                  If Trim(vaSpread1.text) <> "" Then
                     For X = aiblockrow + 1 To vaSpread1.MaxRows
                         vaSpread1.Row = X: vaSpread1.Col = vaSpread1.maxcols: xser = Val(vaSpread1.text)
                         vaSpread1.Col = i + 1
                         If vaSpread1.Row = vaSpread1.MaxRows Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows X, 1: L = X: Exit For
                         If xser <> iser And xser > 0 Then
                            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows X, 1: L = X: Exit For
                         ElseIf Trim(vaSpread1.text) <> "" And xser > 0 Then
                            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows X + 1, 1: X = X + 1: L = X: Exit For
                         ElseIf Trim(vaSpread1.text) = "" Then
                            Exit For
                         End If
                     Next X
                     vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, X
                     vaSpread1.Row = X: accion = "Copiar"
                  Else
                  '-----> Copia los elemenos seleccionados
                     vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, aiblockrow
                       
                     '--***--> este procedimiento guarda los cambios de manera que despues se puedan deshacer
                     vaSpread1.Row = aiblockrow: accion = "Copiar"

                  End If
                  '-------> Asignar colores
                  For X = (i - 1) To (i - 1) + 4
                      vaSpread1.Col = X
                      vaSpread1.BackColor = Shape1(0).FillColor
                      For xx = 1 To MaxColumna
                          If (vectorcol(xx) - 1) = vaSpread1.Col Then
                              vaSpread1.Col = X + 2
                              vaSpread1.CellType = CellTypeNumber
                              vaSpread1.TypeNumberDecPlaces = 0
                              vaSpread1.TypeIntegerMin = 1
                              vaSpread1.TypeIntegerMax = 9999999
                              vaSpread1.TypeHAlign = TypeHAlignRight
                              vaSpread1.TypeSpin = False
                              vaSpread1.TypeIntegerSpinInc = 1
                              vaSpread1.TypeIntegerSpinWrap = False
                              Exit For
                          End If
                      Next xx
                      vaSpread1.Col = X
                      If X = (i - 1) Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                  Next X
                  If L > 0 Then
                     z = L
                     For L = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
                         vaSpread1.Row = 1: vaSpread1.Col = L
                         If vaSpread1.BackColor = Shape1(1).FillColor Then
                            vaSpread1.Row = z
                            For X = (L - 1) To (L - 1) + 4
                                vaSpread1.Col = X
                                vaSpread1.BackColor = Shape1(1).FillColor
                            Next X
                         End If
                     Next L
                  End If
                  '-------> Fin asignar colores
                  Exit For
               End If
           Next j
       Next i
    Else
       '-------> Grabar Evento Copiado y Pegado
       GrabarCambios vaSpread1.Row, j, "Copiado y Pegado"
       indrow3 = vaSpread1.MaxRows
       For i = iblockcol To iblockcol2 Step 6
           If indcortarpegar = 1 Then
              vaSpread1.Row = aiblockrow: vaSpread1.Col = aiblockcol
              If vaSpread1.BackColor = Shape1(1).FillColor Then
                 vaSpread1.MaxRows = vaSpread1.MaxRows + (aiblockrow2 - aiblockrow) + 1
                vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)

                 accion = "Copiar"
                 '-------> Asignar colores
                 For j = vaSpread1.MaxRows - (aiblockrow2 - aiblockrow) To vaSpread1.MaxRows
                     vaSpread1.Row = j
                     For X = (i) To (i) + 4
                         vaSpread1.Col = X
                         vaSpread1.BackColor = Shape1(0).FillColor
                         For xx = 1 To MaxColumna
                             If (vectorcol(xx) - 1) = vaSpread1.Col Then
                                vaSpread1.Col = X + 2
                                vaSpread1.CellType = CellTypeNumber
                                vaSpread1.TypeNumberDecPlaces = 0
                                vaSpread1.TypeIntegerMin = 1
                                vaSpread1.TypeIntegerMax = 9999999
                                vaSpread1.TypeHAlign = TypeHAlignRight
                                vaSpread1.TypeSpin = False
                                vaSpread1.TypeIntegerSpinInc = 1
                                vaSpread1.TypeIntegerSpinWrap = False
                                Exit For
                             End If
                         Next xx
                         vaSpread1.Col = X
                         If X = (i) And Trim(vaSpread1.text) <> "" Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                     Next X
                 Next j
                 '-------> Fin asignar colores
                 vaSpread1.CopyRange iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 vaSpread1.MaxRows = indrow3: accion = "Copiar"
              Else
                 vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
                 accion = "Copiar"
            End If
           ElseIf indcortarpegar = 0 Then
              vaSpread1.Row = aiblockrow: vaSpread1.Col = aiblockcol
              If vaSpread1.BackColor = Shape1(1).FillColor Then
                 vaSpread1.MaxRows = vaSpread1.MaxRows + (aiblockrow2 - aiblockrow) + 1
                 vaSpread1.MoveRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)
                 '-------> Asignar colores
                 For j = vaSpread1.MaxRows - (aiblockrow2 - aiblockrow) To vaSpread1.MaxRows
                     vaSpread1.Row = j
                     For X = (i) To (i) + 4
                         vaSpread1.Col = X
                         vaSpread1.BackColor = Shape1(0).FillColor
                         For xx = 1 To MaxColumna
                             If (vectorcol(xx) - 1) = vaSpread1.Col Then
                                vaSpread1.Col = X + 2
                                vaSpread1.CellType = CellTypeNumber
                                vaSpread1.TypeNumberDecPlaces = 0
                                vaSpread1.TypeIntegerMin = 1
                                vaSpread1.TypeIntegerMax = 9999999
                                vaSpread1.TypeHAlign = TypeHAlignRight
                                vaSpread1.TypeSpin = False
                                vaSpread1.TypeIntegerSpinInc = 1
                                vaSpread1.TypeIntegerSpinWrap = False
                                Exit For
                             End If
                         Next xx
                         vaSpread1.Col = X
                         If X = (i) And Trim(vaSpread1.text) <> "" Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                     Next X
                 Next j
                 '-------> Fin asignar colores
                 vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 vaSpread1.MaxRows = indrow3: accion = "Cortar"
              Else
                 '------- Funcion CORTAR Y PEGAR
                 vaSpread1.MoveRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
                 accion = "Cortar"
                 OpGrilla(13).Enabled = False
                 Toolbar1.Buttons(6).Visible = True
                 Toolbar1.Buttons(7).Visible = False
              End If
           End If
           For j = vaSpread1.ActiveRow To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
               vaSpread1.Row = j
               If indcortarpegar = 0 Then
               Else
                  If vaSpread1.ActiveRow = vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) And ((iblockcol + (aiblockcol2 - aiblockcol)) - 6) > 6 Then
                  Else
                  End If
               End If
           Next j
           '-------> Fin validar días modificados
       Next i
    End If
    If indcos = True Then
       For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
       
    End If
    '------> Se trabaja como excel las raciones
    ColumnaActiva = vaSpread1.ActiveCol: FilaActiva = vaSpread1.ActiveRow: ColumnaAntActiva = ColumnaActiva - 1
    vaSpread1.Col = ColumnaActiva: vaSpread1.Row = 0
    If ColumnaActiva > 1 And accion = "Copiar" Then
      vaSpread1.Row = 0
      '-------->  Copia en posición Ración
      If vaSpread1.text = "N.Rac." Then
            If contador = 1 Then
              n = 1: n1 = 1: Max = contador: max1 = contador_b
              For ff = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Row = FilaActiva
                If Trim(VecRacPegar(n1)) = "" Then
                  vaSpread1.Col = f - 1: desc = vaSpread1.text
                  vaSpread1.Col = f: vaSpread1.Row = FilaActiva
                  vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                ElseIf Trim(VecRacPegar(n1)) = "0" Then
                  vaSpread1.text = IIf(Trim(VecRacPegar(n1)) = "0", VecSelGrid(n), VecRacPegar(n1))
                Else
                  If TipoCopia = "Copiar Raciones" Then
                    vaSpread1.text = Trim(VecSelGrid(n))
                  End If
                End If
                If n <= Max Then n = n + 1
                If n1 <= max1 Then n1 = n1 + 1
                If n > Max Then n = 1
                If n1 > max1 Then Exit For
              Next ff
            ElseIf contador > 1 Then
              n = 1: n1 = 1: Max = contador: max1 = contador_b
              For g = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow) 'vaSpread1.MaxRows
                vaSpread1.Col = ColumnaActiva: vaSpread1.Row = g
                    If Trim(VecRacPegar(n1)) = "" Then
                      vaSpread1.Col = ColumnaAntActiva: vaSpread1.Row = g: desc = vaSpread1.text
                      vaSpread1.Col = ColumnaActiva: vaSpread1.Row = g
                      vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                    ElseIf Trim(VecRacPegar(n1)) = "0" Then
                      vaSpread1.text = IIf(Trim(VecRacPegar(n1)) = "0", VecSelGrid(n), VecRacPegar(n1))
                    Else
                      If TipoCopia = "Copiar Raciones" Then
                        vaSpread1.text = Trim(VecSelGrid(n))
                      Else
                        vaSpread1.text = Trim(VecRacPegar(n1))
                      End If
                    End If
                    If n <= Max Then n = n + 1
                    If n1 <= max1 Then n1 = n1 + 1
                    If n > Max Then n = 1
                    If n1 > max1 Then Exit For
              Next g
            End If
      '-------->  Copia en posición Costo
      ElseIf vaSpread1.text = "Costo" Then
        tope = ColumnaActiva - CantCol1
        For f = ColumnaActiva To tope Step -1
          vaSpread1.Col = f: vaSpread1.Row = 0
          If vaSpread1.text = "N.Rac." Then
              n = 1: n1 = 1: Max = contador: max1 = contador_b
              For g = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
                vaSpread1.Col = f: vaSpread1.Row = g
                  If Trim(VecRacPegar(n1)) = "" Then
                    vaSpread1.Col = f - 1: desc = vaSpread1.text
                    vaSpread1.Col = f
                    vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                  ElseIf Trim(VecRacPegar(n1)) > 0 Then
                    vaSpread1.text = Trim(VecRacPegar(n1))
                  Else
                    vaSpread1.text = VecSelGrid(n)
                  End If
                  If n <= Max Then n = n + 1
                  If n1 <= max1 Then n1 = n1 + 1
                  If n > Max Then n = 1
                  If n1 > max1 Then Exit For
              Next g
          End If
        Next f
      Else
        '-------->  Distinta posición a la anterior
        For f = ColumnaActiva To vaSpread1.ActiveCol + CantCol1
          vaSpread1.Col = f: vaSpread1.Row = 0
          If vaSpread1.text = "N.Rac." Then
            If contador = 1 Then
              n = 1: n1 = 1: Max = contador: max1 = contador_b
              For ff = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
                vaSpread1.Row = FilaActiva
                If Trim(VecRacPegar(n1)) = "" Then
                  vaSpread1.Col = f - 1: desc = vaSpread1.text
                  vaSpread1.Col = f: vaSpread1.Row = FilaActiva
                  vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                Else
                  vaSpread1.text = IIf(Trim(VecRacPegar(n1)) = 0, VecSelGrid(n), VecRacPegar(n1))
                End If
                If n <= Max Then n = n + 1
                If n1 <= max1 Then n1 = n1 + 1
                If n > Max Then n = 1
                If n1 > max1 Then Exit For
              Next ff
            Else
              n = 1: n1 = 1: Max = contador: max1 = contador_b
              For g = FilaActiva To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
                vaSpread1.Col = f: vaSpread1.Row = g
                  If Trim(VecRacPegar(n1)) = "" Then
                    vaSpread1.Col = f - 1: desc = vaSpread1.text
                    vaSpread1.Col = f
                    vaSpread1.text = IIf(Trim(desc) = "", "", 0)
                  ElseIf Trim(VecRacPegar(n1)) > 0 Then
                    vaSpread1.text = Trim(VecRacPegar(n1))
                  Else
                    vaSpread1.text = VecSelGrid(n)
                  End If
                  If n <= Max Then n = n + 1
                  If n1 <= max1 Then n1 = n1 + 1
                  If n > Max Then n = 1
                  If n1 > max1 Then Exit For
              Next g
            End If
          End If
        Next f
      End If
    End If
    '------>
    aiblockcol = IndCol: iblockcol2 = indcol2
    aiblockrow = indrow: aiblockrow2 = indrow2
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    vaSpread1.Enabled = True
    fg_descarga
Case 15
    B_BusVas.Partidas Me
    B_BusVas.Show 1
Case 16
    If vaSpread1.ActiveCol = 1 And vaSpread1.ActiveRow <> vaSpread1.MaxRows And Trim(vaSpread1.text) <> "" And vg_IndpprSelec = 2 Then
       G_Proc.CellEdite B_CelEdi, "Editar Estructura", "Nombre Estructura", vaSpread1, "1"
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = True
    End If
Case 17
    DoEvents
    T_Servic.SSTab1.TabEnabled(0) = False
    DoEvents
    T_Servic.SSTab1.TabEnabled(2) = False
    DoEvents
    T_Servic.CallForm = Me.Name
    DoEvents
    T_Servic.SSTab1.Tab = 1
    DoEvents
    T_Servic.MoverDatosGrillas2 vg_codservicio
    DoEvents
    T_Servic.Show 1
    DoEvents
End Select
CargarAporteCalorico
estgra = False

Exit Sub
Man_Error:
vaSpread1.Enabled = True
fg_descarga
End Sub

Private Sub Timer1_Timer()
On Error GoTo Man_Error
If Toolbar1.Buttons(2).Visible = False Then Exit Sub
' variable estática para acumular la cantidad de segundos
'Static Temp_Seg As Long
' incrementa
vg_TemSeg = vg_TemSeg + 1
' comprueba que los segundos no sea igual a la cantidad de minutos _
  que queremos , en este caso 5 minutos
If (vg_TemSeg * 60) >= (vg_IntMin * 60) * 60 And Not estgra Then
   ' reestablece
   estgra = True
   If MsgBox(" Actualiza planificación ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Cancel = -1: vg_TemSeg = 0: estgra = False: Exit Sub
   If Toolbar1.Buttons(2).Visible = True Then
      Toolbar1.Enabled = False
      Toolbar1.Buttons(31).Enabled = False
      CorDes = 0
      GrabarPlantillaMinuta
      CorDes = 0
      If Dir(LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6") <> "" Then Kill LCase(App.Path) & "\" & "Spread" & vg_NUsr & "*" & ".ss6"
      Toolbar1.Enabled = True
   End If
   Toolbar1.Buttons(1).Visible = True
   Toolbar1.Buttons(2).Visible = False
   estgra = False
   vg_TemSeg = 0
End If
Exit Sub
Man_Error:
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 2
    Plantilla_Click (0)
Case 4
    Plato_Click (11)
Case 5
    Plato_Click (12)
Case 7
    Plato_Click (13)
Case 9
    Plato_Click (14)
Case 10
   Plato_Click (15)
Case 12
    Plato_Click (5)
Case 13
    Plato_Click (6)
Case 15
    Plato_Click (8)
Case 16
    Plato_Click (9)
Case 18
    Plantilla_Click (5)
Case 19
    Plantilla_Click (8)
Case 21
    Plantilla_Click (10)
Case 22
    Plantilla_Click (3)
Case 23
    Plantilla_Click (11)
Case 24
    Plantilla_Click (14)
Case 25
    Plantilla_Click (13)
Case 27
    ExportarExcel
Case 29
    HabilitaCeldaCalorias
Case 30
    Plantilla_Click (20)
Case 31
    Plantilla_Click (21)
Case 32
    Plantilla_Click (22)
End Select
End Sub

Sub ExportarExcel()
On Error GoTo Ex_Error
CargarAporteCalorico
Dim NashXl As excel.Application
Dim IRow As Long, irow2 As Long
Dim NColumnas As Integer
fg_carga ""
Set NashXl = CreateObject("excel.application")
Set NashXl = New excel.Application
NashXl.SheetsInNewWorkbook = 1
NashXl.Workbooks.Add
NashXl.Range("A1").Select
NashXl.ActiveCell.FormulaR1C1 = "Sub-Segmento : " & vg_codsubseg & "-" & vg_nomsubseg
NashXl.Range("A2").Select
NashXl.ActiveCell.FormulaR1C1 = "Regimen      : " & vg_codregimen & "-" & vg_nomreg
NashXl.Range("A3").Select
NashXl.ActiveCell.FormulaR1C1 = "Servicio     : " & vg_codservicio & "-" & vg_nomser
NashXl.Range("A4").Select
NashXl.ActiveCell.FormulaR1C1 = "Fecha        : " & vg_fecha

MaxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
NColumnas = (MaxColumna * 6) + 1
vaSpread1.AllowMultiBlocks = True
vaSpread1.SetSelection 1, -1, NColumnas, vaSpread1.MaxRows + 3
vaSpread1.ClipboardCopy

IRow = vaSpread1.MaxRows + 5
'------- Pegar vaspread1(0) - Planilla Excel
NashXl.Range("A5").Select
NashXl.ActiveSheet.Paste
'------- Asignar color
'------- Colorear titulo
NashXl.Range("A5:GE5").Select
With NashXl.Selection.Interior
     .ColorIndex = 15
     .Pattern = xlSolid
End With
'------- Dibujar marco
NashXl.Range("A5:GE" & IRow).Select
NashXl.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
NashXl.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
With NashXl.Selection.Borders(xlEdgeLeft)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeTop)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeBottom)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlEdgeRight)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideVertical)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
With NashXl.Selection.Borders(xlInsideHorizontal)
     .LineStyle = xlContinuous
     .Weight = xlThin
     .ColorIndex = xlAutomatic
End With
NashXl.Range("A2" & ":" & "A" & IRow).Select
NashXl.Selection.NumberFormat = "#,##0.00"

'------- Asigna Colores a Estructura de Servicio
NashXl.Range("A6:" & "A" & IRow).Select
With NashXl.Selection.Interior
     .ColorIndex = 10
     .Pattern = xlSolid
End With
'------- Aplicar totales

NashXl.Selection.Font.Bold = True

NashXl.Range("B" & IRow & ":" & "B" & 2).Select
NashXl.Selection.NumberFormat = "#,##0.00"
'------- Ajustar columna
NashXl.Cells.Select
NashXl.Cells.EntireColumn.AutoFit
vaSpread1.AllowMultiBlocks = False: vaSpread1.SetSelection 1, 0, vaSpread1.maxcols, vaSpread1.MaxRows
Dim aa As Variant
NashXl.Cells.Replace What:="&0&;", Replacement:="", LookAt:=xlPart, SearchOrder _
      :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
If Not IsEmpty(aa) Then
   NashXl.Cells.Replace What:="&0&;", Replacement:="", LookAt:=xlPart, SearchOrder _
         :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End If
If Not IsEmpty(aa) Then
   NashXl.Cells.Replace What:="&-1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
         :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End If

aa = NashXl.Cells.Find(What:="&" & vg_codregimen & "&;", LookAt:=xlPart, SearchOrder _
      :=xlByRows, MatchCase:=False, SearchFormat:=False)
      
If Not IsEmpty(aa) Then
   NashXl.Cells.Replace What:="&" & vg_codregimen & "&;", Replacement:="", LookAt:=xlPart, SearchOrder _
         :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
End If

'NashXl.Cells.Replace What:="&-1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
'      :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
'
'NashXl.Cells.Replace What:="&1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
'      :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        
fg_descarga
NashXl.Visible = True
Ex_Error:
    Resume Next
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
indactivo = 1
iblockrow = BlockRow
iblockrow2 = BlockRow2
iblockcol = BlockCol
iblockcol2 = IIf(estapo = False, BlockCol2 + 1, BlockCol2)
If BlockRow < 0 Then iblockrow = 1
If BlockRow2 < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
If BlockRow2 >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Then Exit Sub
OpGrilla(15).Enabled = IIf(Col = 1, True, False)
Plato(15).Enabled = IIf(Col = 1, True, False)
indactivo = 1
iblockrow = vaSpread1.ActiveRow
iblockrow2 = vaSpread1.ActiveRow
iblockcol = vaSpread1.ActiveCol
iblockcol2 = vaSpread1.ActiveCol
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Col = 1 Then Plato_Click (16): Exit Sub
If Row < 1 Or Col = 1 Then Exit Sub
Plato_Click (2)
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
    If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
    vaSpread1.Row = Row: vaSpread1.Col = Col
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If vaSpread1.ChangeMade = False Or Col = 1 Or Mode = 1 Then i = IIf(vaSpread1.text = "", "0", vaSpread1.text): Exit Sub
    '-------> Grabar Evento Modificar
    GrabarCambios 1, j, "Modificar Estructura"
    If vaSpread1.ChangeMade = True Then vaSpread1.Col = (MaxColumna * 6 + 1) + (vaSpread1.Col / 6): vaSpread1.text = 1: If indcos = True Then vaSpread1.Col = Col: j = Col - 1:  Calctodia vaSpread1.Row, j 'veccos((Int(J / 5) + 1), 4) = Round(veccos((Int(J / 5) + 1), 4) - (i), vg_DPr): veccos((Int(J / 5) + 1), 4) = Round(veccos((Int(J / 5) + 1), 4) + (vaSpread1.Text), vg_DPr)
    vaSpread1.Row = Row
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    CargarAporteCalorico
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
Dim delrow As Integer, IndCol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer
Select Case KeyCode
Case 65 To 90
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    ws_respuesta = ""
    ws_respuesta = Chr(KeyCode)
    Plato_Click (2)
Case 86
    Exit Sub
Case 46
    Select Case vaSpread1.ActiveCol
    Case 1
        Dim xrow As Integer
        Dim codest As Long, auxest As Long
        Dim AvisoEst As Boolean
        vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
        
        iblockcol = xColIni
        iblockrow = xRowIni
        iblockcol2 = xColIni
        iblockrow2 = xRowFin
        
        If vaSpread1.MaxRows = vaSpread1.ActiveRow Or vaSpread1.MaxRows = iblockrow Or vaSpread1.MaxRows = iblockrow2 Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = vaSpread1.ActiveCol
        If vaSpread1.Col = 1 And vaSpread1.Row <> vaSpread1.MaxRows Then
           '-------> Grabar Evento Modificar Estructura
           GrabarCambios 1, 1, "Modificar Estructura"
            
           DesqloqSubMenu (vaSpread1.text)
           vaSpread1.Row = vaSpread1.ActiveRow
           vaSpread1.Col = vaSpread1.ActiveCol
           vaSpread1.text = ""
           vaSpread1.Col = vaSpread1.maxcols
           If Trim(vaSpread1.text) <> "" Then auxest = Val(vaSpread1.text)
           vaSpread1.Col = 1
           xrow = vaSpread1.ActiveRow
           AvisoEst = False
           For i = vaSpread1.ActiveRow To 1 Step -1
               vaSpread1.Row = i
               If Trim(vaSpread1.text) <> "" Then
                   vaSpread1.Col = vaSpread1.maxcols
                   codest = vaSpread1.text
                   vaSpread1.Row = xrow
                   vaSpread1.text = codest
                   vaSpread1.Col = 1
                   AvisoEst = True
                   Exit For
               End If
           Next i
           '-------> eliminar siguiente estructura
           For i = xrow + 1 To vaSpread1.MaxRows - 1
               vaSpread1.Row = i
               vaSpread1.Col = vaSpread1.maxcols
               If Val(vaSpread1.text) = auxest Then
                  vaSpread1.text = codest
               Else
                  Exit For
               End If
           Next i
           If AvisoEst = False And Trim(StrRec) <> "" Then
              If Dir(LCase(App.Path) & "\" & StrRec) <> "" Then Kill LCase(App.Path) & "\" & StrRec
              CorDes = CorDes - 1
              MsgBox "Al borrar esta estructura, dejará recetas sin asignar", vbCritical
           End If
           Toolbar1.Buttons(1).Visible = False
           Toolbar1.Buttons(2).Visible = True
        End If
        If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        
        j = 0
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then j = (vectorcol(i) - 1): Exit For
        Next i
        If j = 0 Then Exit Sub
        Plato(0).Enabled = True
        OpGrilla(0).Enabled = True
        Plato(13).Enabled = False
        OpGrilla(13).Enabled = False
        If indactivo = 0 Or iblockcol < 1 Or iblockrow < 1 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
        
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.maxcols
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        IndCol = aiblockcol: indcol2 = iblockcol2
        indrow = aiblockrow: indrow2 = IIf(aiblockrow2 = vaSpread1.MaxRows, (aiblockrow2 - 1), aiblockrow2)
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False
        
        If indcos = True Then
           For i = iblockcol To iblockcol2 Step 6
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        iblockcol = AuxCol
        vaSpread1.BlockMode = False
        Plantilla(0).Enabled = True
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        indactivo = 0
    Case Is > 1
        vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
        
        iblockcol = xColIni
        iblockrow = xRowIni
        iblockcol2 = xcolfin
        iblockrow2 = xRowFin
        
        If vaSpread1.MaxRows = vaSpread1.ActiveRow Or vaSpread1.MaxRows = iblockrow Or vaSpread1.MaxRows = iblockrow2 Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = vaSpread1.ActiveCol
        If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        j = 0
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then j = (vectorcol(i) - 1): Exit For
        Next i
        If j = 0 Then Exit Sub
        Plato(0).Enabled = True
        OpGrilla(0).Enabled = True
        Plato(13).Enabled = False
        OpGrilla(13).Enabled = False
        If indactivo = 0 Or iblockcol < 1 Or iblockrow < 1 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
         
        '-------> Grabar Evento Eliminación recetas
        GrabarCambios 1, 1, "Eliminación Recetas"
         
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.maxcols
        aiblockcol = iblockcol
        aiblockrow = iblockrow
        aiblockcol2 = iblockcol2
        aiblockrow2 = iblockrow2
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 4): Exit For
        Next i
        IndCol = aiblockcol: indcol2 = iblockcol2
        indrow = aiblockrow: indrow2 = IIf(aiblockrow2 = vaSpread1.MaxRows, (aiblockrow2 - 1), aiblockrow2)
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False
        If indcos = True Then
           For i = iblockcol To iblockcol2 Step 6
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        iblockcol = AuxCol
        vaSpread1.BlockMode = False
        Plantilla(0).Enabled = True
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
        indactivo = 0
   
    End Select
End Select
CargarAporteCalorico
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If vaSpread1.MaxRows < 1 Then Exit Sub
iblockrow = NewRow
iblockrow2 = NewRow
iblockcol = NewCol
iblockcol2 = NewCol
If NewRow < 0 Then iblockrow = 1
If NewRow < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
If NewRow >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)
If indcos = False Or NewCol < 1 Then Exit Sub
MostrarCosto NewCol
End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
    If vaSpread1.Visible <> True Then Exit Sub
    Indvaspread1 = 0
    If Mid(ValidaPerfil(M_Plami2), 1, 4) <> "1000" = True Then
        PopupMenu MenuDetalle
        
    End If
End Select
End Sub

Private Sub Opgrilla_Click(Index As Integer)
Select Case Index
Case 0
    Plato_Click (0)
Case 2
    Plato_Click (2)
Case 3
    Plato_Click (3)
Case 5
    Plato_Click (5)
Case 6
    Plato_Click (6)
Case 8
    Plato_Click (8)
Case 9
    Plato_Click (9)
Case 11
    Plato_Click (11)
Case 12
    Plato_Click (12)
Case 13
    Plato_Click (13)
Case 14
    Plato_Click (14)
Case 15
    Plato_Click (15)
Case 16
    Plato_Click (16)
Case 17
    Plato_Click (17)
End Select
End Sub

Private Sub GrabarPlantillaMinuta()
Dim RS2 As New ADODB.Recordset
Dim desc As String, StrRec As String, StrRecb As String, NameEstManual As String, NameEst As String
Dim CodRec As Long, numrac As Long, estser As Long, Fecha As Long, ConRegDet As Long, indice As Long, existedat As Long, IndDia As Long, tiprec As Long
Dim fechasis As Long, FecIni As Long, FecFin As Long, totrac As Long
Dim cosali As Double, cospro As Double, CosDes As Double
Dim i As Long, j As Long
Dim MyBuffer     As String
Dim estgrapla    As Boolean

On Error GoTo Man_Error
Main(0).Enabled = False
Main(1).Enabled = False
vaSpread1.Enabled = False
estgrapla = False
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
NameEstManual = ""
IndDia = 1: ConRegDet = 0: gauge1.Value = 0: gauge.Value = 0: Fecha = 0: FecIni = 0: FecFin = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh
fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
fg_carga ""
'-------> Grabar planificación minutas
For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
    DoEvents
    estgrapla = False
    If IndDia > MaxColumna Then Exit Sub
    gauge1.Value = Val((IndDia / MaxColumna) * 100)
    Label3.Caption = "": Label3.Caption = "Día : " & IndDia
    existedat = 0: vaSpread1.Row = 1: vaSpread1.Col = i
    If (vaSpread1.MaxRows - 1) = 0 Then
       existedat = 0
       Fecha = Val(vg_fecha) & fg_pone_cero(IndDia, 2)
    Else
       For j = 1 To (vaSpread1.MaxRows - 1)
           vaSpread1.Row = j
           Fecha = Val(vg_fecha) & fg_pone_cero(IndDia, 2)
           vaSpread1.Col = i + 1
           If Trim(vaSpread1.text) <> "" Then existedat = 1: Exit For
       Next j
    End If
    indice = 0
    vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = i + 2: totrac = Val(vaSpread1.text)
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 4, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & ", " & Val(Fecha) & ", 0, 0, " & vg_IndpprSelec & "")
    If Not RS.EOF Then
       indice = RS!min_codigo: RS.Close: Set RS = Nothing
       If indice > 0 And existedat = 0 Then
          vg_db.Execute "sgpadm_d_minutadet 'E2', " & indice & ", '1', 0, 0, 0, 0, '', 0, 0, 0, 0"
          vg_db.Execute "sgpadm_d_minuta 'E', " & indice & ", 0, 0, 0, 0, 0, 0, 0, 0, '', '" & vg_IndpprSelec & "'"
       Else
          vaSpread1.Row = 1
          If vaSpread1.BackColor <> Shape1(1).FillColor Then
             vg_db.Execute "sgpadm_iu_minuta 'M1', " & indice & ", " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Val(Fecha) & ", 0, " & totrac & ", " & totrac & ", 0, '', '" & vg_IndpprSelec & "'"
          End If
       End If
    Else
       RS.Close: Set RS = Nothing
       If existedat > 0 Then
          Set RS = vg_db.Execute("sgpadm_iu_minuta 'A', 0, '" & vg_codsubseg & "', " & vg_codregimen & ", " & vg_codservicio & ", " & Val(Fecha) & ", 0, " & totrac & ", " & totrac & ", 0, '', '" & vg_IndpprSelec & "'")
          If Not RS.EOF Then
             indice = RS!indice
          End If
          RS.Close: Set RS = Nothing
       End If
    End If
    gauge.Value = 0: ConRegDet = 0: estser = 0
    If existedat > 0 Then
'       If maxfila > vaSpread1.MaxRows Then
'          '-------> Si maximo de fila es mayor que grilla borra detalle
'          For j = vaSpread1.MaxRows To maxfila
'              vg_db.Execute "sgpadm_d_minutadet 'E1', " & indice & ", '1', " & j & ", 0, 0, 0, '', 0, 0, 0, 0"
'          Next j
'       End If
       
       Let MyBuffer = ""
       Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
       Let MyBuffer = MyBuffer & "<GrabaMinuta>"
       
       '-------> Actualizar detalle minutas
       For j = 1 To (vaSpread1.MaxRows - 1)
           ConRegDet = ConRegDet + 1
           gauge.Value = Val((ConRegDet / (vaSpread1.MaxRows - 1)) * 100)
           desc = "": CodRec = 0: NumRec = 0: cosali = 0: CosDes = 0: tiprec = 0
           vaSpread1.Row = j
          
           '---------- Samuel Melendez 28/09/09
           '** Si el nombre de la estructura fue ingresado manualmente
           '** por el usuario se llena la variable "NameEstManual", sino queda vacia
           'NameEstManual = ValidaNombreEstructura(j, vaSpread1)
'jpaz           NameEstManual = EstructuraSuperior(vaSpread1, j)
           '-----------------------------------
                      
           vaSpread1.Col = vaSpread1.maxcols
           If Trim(vaSpread1.text) <> "" Then
              estser = vaSpread1.text
              vaSpread1.Col = 1
              If vg_IndpprSelec = "2" And Trim(vaSpread1.text) <> "" Then
                 NameEst = vaSpread1.text
                 vaSpread1.Col = vaSpread1.maxcols - 1
                 NameEstManual = vaSpread1.text
                 
                 If NameEstManual <> NameEst Then
                    NameEstManual = NameEst
                 Else
                    NameEstManual = IIf(Trim(NameEstManual) = "", "", NameEstManual)
                 End If
              ElseIf vg_IndpprSelec = "1" Then
                 NameEstManual = ""
              End If
           End If
           vaSpread1.Col = i + 1: desc = Trim(vaSpread1.text)
           
'           If desc <> "" And estser > 0 Then
            If desc = "" Then 'And estser < 1 Then
               a = estser
            End If
              vaSpread1.Col = i + 2: numrac = IIf(Trim(vaSpread1.text) = "", 0, Val(vaSpread1.text))
              vaSpread1.Col = i + 3: cosali = IIf(Trim(vaSpread1.text) = "", 0, Val(vaSpread1.text))
              vaSpread1.Col = i + 4: d = vaSpread1.text
              
              StrRec = Trim(vaSpread1.text)
              If Len(StrRec) <> 0 Then
                 Do While InStr(StrRec, ";") <> 0
                    StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                    StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                    CodRec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)):
                    StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                    tiprec = Val(Mid(StrRecb, 1))
                 Loop
              End If
              
              '----> Traer costo receta alimentación y desechable
              cosali = 0
              CosDes = 0
'              If codrec = 0 Then
'                 codrec = BuscarCodReceta(desc)
'              End If
              
'              Set RS = vg_db.Execute("sgpadm_s_minutadet 1, " & indice & ", '1', " & j & ", 0, 0, 0, '', 0, 0, 0, 0")
              Let MyBuffer = MyBuffer & " <Minuta"
'              If Not RS.EOF Then
'                 'Actualiza
'                 MyBuffer = MyBuffer & " Op = " & Chr(34) & 1 & Chr(34)
'              Else
'                 'Graba
'                 MyBuffer = MyBuffer & " Op = " & Chr(34) & 2 & Chr(34)
'              End If
'              RS.Close: Set RS = Nothing
              MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)
              estgrapla = True
                  
              desc = Replace(Trim(desc), Chr(34), "&quot;")
              desc = Replace(Trim(desc), Chr(38), "&amp;")
              desc = Replace(Trim(desc), Chr(39), "&apos;")
              desc = Replace(Trim(desc), Chr(60), "&lt;")
              desc = Replace(Trim(desc), Chr(62), "&gt;")
              
              NameEstManual = Replace(Trim(NameEstManual), Chr(34), "&quot;")
              NameEstManual = Replace(Trim(NameEstManual), Chr(38), "&amp;")
              NameEstManual = Replace(Trim(NameEstManual), Chr(39), "&apos;")
              NameEstManual = Replace(Trim(NameEstManual), Chr(60), "&lt;")
              NameEstManual = Replace(Trim(NameEstManual), Chr(62), "&gt;")
                            
              MyBuffer = MyBuffer & " Codigo = " & Chr(34) & indice & Chr(34)
              MyBuffer = MyBuffer & " TipMin = " & Chr(34) & 1 & Chr(34)
              MyBuffer = MyBuffer & " NumLin = " & Chr(34) & j & Chr(34)
              MyBuffer = MyBuffer & " EstSer = " & Chr(34) & estser & Chr(34)
              MyBuffer = MyBuffer & " CodRec = " & Chr(34) & CodRec & Chr(34)
              MyBuffer = MyBuffer & " NumRac = " & Chr(34) & numrac & Chr(34)
              MyBuffer = MyBuffer & " DesCri = " & Chr(34) & Mid(desc, 1, 50) & Chr(34)
              MyBuffer = MyBuffer & " CosRec = " & Chr(34) & cosali & Chr(34)
              MyBuffer = MyBuffer & " FecVal = " & Chr(34) & 0 & Chr(34)
              MyBuffer = MyBuffer & " TipRec = " & Chr(34) & tiprec & Chr(34)
              MyBuffer = MyBuffer & " CosDes = " & Chr(34) & CosDes & Chr(34)
              MyBuffer = MyBuffer & " DesEst = " & Chr(34) & NameEstManual & Chr(34)
              Let MyBuffer = MyBuffer & "/>"
'              If Not RS.EOF Then
'                 RS.Close: Set RS = Nothing
''                 vaSpread1.Col = (maxColumna * 6 + 1) + ((i + 3) / 6) ' este estaba
'                 vg_db.Execute "sgpadm_iu_minutadet 'M', " & indice & ", '1', " & j & ", " & estser & ", " & CodRec & ", " & numrac & ", '" & Mid(desc, 1, 50) & "', " & cosali & ", 0, " & tiprec & ", " & cosdes & ", '" & NameEstManual & "'"
'              Else
'                 RS.Close: Set RS = Nothing
'                 vg_db.Execute "sgpadm_iu_minutadet 'A', " & indice & ", '1', " & j & ", " & estser & ", " & CodRec & ", " & numrac & ", '" & Trim(Mid(desc, 1, 50)) & "', " & cosali & ", 0, " & 1 & ", " & cosdes & ", '" & NameEstManual & "'"
'              End If
'           Else
'               vg_db.Execute "sgpadm_d_minutadet 'E1', " & indice & ", '1', " & j & ", 0, 0, 0, '', 0, 0, 0, 0"
'           End If
       Next j
       Let MyBuffer = MyBuffer & "</GrabaMinuta>"
'       MyBuffer = Replace(Trim(MyBuffer), "&", "&amp;")
'  "  &#34; &quot;   quotation mark
'  &  &#38; &amp; ampersand
'  '  &#39; &apos; (does not work in IE)  apostrophe
'  <  &#60; &lt;  less-than
'  >  &#62; &gt;  greater-than
       If estgrapla Then Set RS2 = vg_db.Execute("sgpadm_iu_minutadetadm '" & MyBuffer & "'")
       Set RS2 = Nothing
       estgrapla = False
    
    End If
    IndDia = IndDia + 1
Next i
FecFin = Fecha
Picture1.Visible = False: gauge.Visible = False
vaSpread1.Enabled = True
Main(0).Enabled = True
Main(1).Enabled = True
vaSpread1.Refresh
fg_descarga
CargarAporteCalorico
Toolbar1.Buttons(31).Enabled = False
Exit Sub
Man_Error:
vaSpread1.Enabled = True
Main(0).Enabled = True
Main(1).Enabled = True
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
Resume
End Sub

Sub HabilitaCeldaCalorias()
CargarAporteCalorico
vaSpread1.Visible = False
For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
  vaSpread1.Row = 0
  vaSpread1.Col = i + 5
  If vaSpread1.ColHidden = True Then
     vaSpread1.ColHidden = False
     estapo = False
  Else
     vaSpread1.ColHidden = True
     estapo = True
  End If
Next i
vaSpread1.Visible = True
End Sub

Sub DetallePlantillaMinuta()
fg_carga ""
Dim indrow3 As Long, IndDia As Long, Fecha As String, spid As Long
Dim precio As Double
Dim sw As Boolean: sw = False

SwSalir = 0: MaxColumna = 0: indactivo = 0
iblockrow = 0: iblockrow2 = 0: iblockcol = 0: iblockcol2 = 0: SwSalir = 0
aiblockrow = 0: aiblockrow2 = 0: aiblockcol = 0: aiblockcol2 = 0

vg_db.Execute "DELETE paso_servicio WHERE ser_spid = @@spid and ser_usr = '" & vg_NUsr & "'"
'--isel = 0
'-------> Buscar spid
Set RS = vg_db.Execute("SELECT @@spid spid")
If Not RS.EOF Then spid = RS!spid: vg_db.Execute "INSERT INTO paso_servicio VALUES (" & spid & ", '" & vg_NUsr & "', " & Val(vg_codservicio) & ")"
RS.Close: Set RS = Nothing

'-------> Formatear columna
MaxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
vaSpread1.MaxRows = 1000
vaSpread1.maxcols = 0: vaSpread1.maxcols = 6 * MaxColumna + 1: vaSpread1.Row = 0
vaSpread1.Col = 1
vaSpread1.ColsFrozen = 1
vaSpread1.VisibleCols = 1
vaSpread1.ColWidth(1) = 15
vaSpread1.text = "Estructura Servicio"
ReDim Preserve vectorcol(0)
For i = 2 To vaSpread1.maxcols Step 6
    vaSpread1.Col = i
    vaSpread1.ColWidth(i) = 1.5
    vaSpread1.text = " "
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 1
    vaSpread1.ColWidth(i + 1) = 21
    If i = 2 Then
       ReDim Preserve vectorcol(1)
       vectorcol(1) = 3
       vaSpread1.text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & (i - 1), 2), 1), 1, 3) & " " & (i - 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
    Else
       vaSpread1.text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & CLng((i / 6) + 1), 2), 1), 1, 3) & " " & CLng((i / 6) + 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
       ReDim Preserve vectorcol(CLng((i / 6) + 1))
       vectorcol(CLng((i / 6) + 1)) = i + 1
    End If
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 2
    vaSpread1.ColWidth(i + 2) = 6
    vaSpread1.text = "N.Rac."
    vaSpread1.ColHidden = False
   
    vaSpread1.Col = i + 3
    vaSpread1.ColWidth(i + 3) = 9
    vaSpread1.text = "Costo"
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 4
    vaSpread1.text = "Cod. Receta"
    vaSpread1.ColHidden = True
    
    vaSpread1.Col = i + 5
    vaSpread1.ColWidth(i + 3) = 9
    vaSpread1.text = "Calorias"
    vaSpread1.ColHidden = True
    
    For j = 1 To vaSpread1.MaxRows
        vaSpread1.Row = j

        vaSpread1.Col = i
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = ""

        vaSpread1.Col = i + 1
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 2
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 3
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 4
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignLeft
        vaSpread1.text = " "

        vaSpread1.Col = i + 5
        vaSpread1.CellType = CellTypeStaticText
        vaSpread1.TypeHAlign = TypeHAlignRight
        vaSpread1.text = " " 'aca debe venir el codigo receta

    Next j
    vaSpread1.Row = 0
Next i

vaSpread1.Row = 0
For i = 1 To MaxColumna
   vaSpread1.maxcols = vaSpread1.maxcols + 1
   vaSpread1.Col = vaSpread1.maxcols
   vaSpread1.text = "Estado"
   vaSpread1.ColHidden = True
Next i
vaSpread1.maxcols = vaSpread1.maxcols + 1
vaSpread1.Col = vaSpread1.maxcols
vaSpread1.ColWidth(vaSpread1.maxcols) = 5
vaSpread1.text = "Còd. Est."
vaSpread1.ColHidden = True

vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
vaSpread1.Row = -1: vaSpread1.Col = 1
vaSpread1.Font.Bold = True
vaSpread1.Font.Size = 9
vaSpread1.BackColor = Shape1(2).FillColor 'Verde
If vg_Zona = "" Then vg_Zona = 0
j = 0: i = 0: indrow3 = 0 'sgpadm_s_PlanMinutaDetreal 50, 10013,10001,1, 200811,2, 'adm', 66
Set RS = vg_db.Execute("sgpadm_s_PlanMinutaDetreal " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(vg_fecha) & ", " & vg_codlpr & ",'" & vg_NUsr & "'," & spid & "," & vg_IndpprSelec & "")
DoEvents
If Not RS.EOF Then
  sw = True   '-------> Calcula el costo plato según su gramaje
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 6) - 6) + 1) + 1
      vaSpread1.Row = RS!mid_numlin
      If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
      If RS!ess_codigo <> i Then
         vaSpread1.Col = 1
         If IIf(IsNull(RS!mid_desest), "", RS!mid_desest) <> "" And vg_IndpprSelec = 2 Then
            vaSpread1.text = RS!mid_desest
            
            vaSpread1.Col = vaSpread1.maxcols - 1
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignCenter
            vaSpread1.text = IIf(IsNull(RS!mid_desest), "", RS!mid_desest)
         
         Else
            vaSpread1.text = Trim(RS!ess_nombre)
            
            vaSpread1.Col = vaSpread1.maxcols - 1
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignCenter
            vaSpread1.text = Trim(RS!ess_nombre)
         
         End If
         
         vaSpread1.Col = vaSpread1.maxcols
         vaSpread1.CellType = CellTypeStaticText
         vaSpread1.TypeHAlign = TypeHAlignCenter
         vaSpread1.text = RS!ess_codigo
         i = RS!ess_codigo
        
      End If
      
      vaSpread1.Col = j
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.Value = "R"
      vaSpread1.ForeColor = &HFF&
      vaSpread1.BackColor = &H80FF80
           
      vaSpread1.Col = j + 1
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignLeft
      vaSpread1.text = Trim(RS!pas_nombre)
                         
      vaSpread1.Col = j + 2
      vaSpread1.CellType = CellTypeNumber
      vaSpread1.TypeNumberDecPlaces = 0
      vaSpread1.TypeIntegerMin = 1
      vaSpread1.TypeIntegerMax = 9999999
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.TypeSpin = False
      vaSpread1.TypeIntegerSpinInc = 1
      vaSpread1.TypeIntegerSpinWrap = False
      vaSpread1.Value = RS!mid_numrac
      vaSpread1.ForeColor = &HFF0000
                       
      vaSpread1.Col = j + 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignRight
      precio = Format(IIf(IsNull(RS!pas_prerec) Or Trim(RS!pas_prerec) = 0, 0, RS!pas_prerec), fg_Pict(6, vg_DPr))
      vaSpread1.text = Format(precio, fg_Pict(6, vg_DPr))
      
      vaSpread1.Col = j + 4: vaSpread1.text = RS!pas_codrec & "&" & RS!mid_tiprec & "&;"
          
      vaSpread1.Col = j + 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.text = IIf(IsNull(RS!candiet) Or RS!candiet = 0, "", Format(Trim(RS!candiet), fg_Pict(6, vg_DPr)))
      
      vaSpread1.Col = vaSpread1.maxcols
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.TypeHAlign = TypeHAlignCenter
      vaSpread1.text = RS!ess_codigo
      
      
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing: fg_descarga
Else
    '-------> Retorna minuta sin precio
    Set RS = vg_db.Execute("sgpadm_s_planifminuta 1, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & "," & vg_Zona & ", " & Val(vg_fecha) & ", 0,0," & vg_IndpprSelec & "")
    DoEvents
    If Not RS.EOF Then '-------> Consulta trae productos sin costo
      sw = True
        Do While Not RS.EOF
              DoEvents
              j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 6) - 6) + 1) + 1
              vaSpread1.Row = RS!mid_numlin
              If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
              If RS!mid_estser <> i Then
                 vaSpread1.Col = 1
                 vaSpread1.text = RS!ess_nombre
                 vaSpread1.Col = vaSpread1.maxcols
                 vaSpread1.CellType = CellTypeStaticText
                 vaSpread1.TypeHAlign = TypeHAlignCenter
                 vaSpread1.text = RS!mid_estser
                 i = RS!mid_estser
              End If
              vaSpread1.Col = j
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignCenter
              vaSpread1.Value = "R"
              vaSpread1.ForeColor = &HFF&
              vaSpread1.BackColor = &H80FF80
                   
              vaSpread1.Col = j + 1
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignLeft
              vaSpread1.text = Trim(RS!mid_descri)
                                 
              vaSpread1.Col = j + 2
              vaSpread1.CellType = CellTypeNumber
              vaSpread1.TypeNumberDecPlaces = 0
              vaSpread1.TypeIntegerMin = 1
              vaSpread1.TypeIntegerMax = 9999999
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.TypeSpin = False
              vaSpread1.TypeIntegerSpinInc = 1
              vaSpread1.TypeIntegerSpinWrap = False
              vaSpread1.Value = RS!mid_numrac
              vaSpread1.ForeColor = &HFF0000
                               
              vaSpread1.Col = j + 3
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.text = Format((IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec) + IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec)), fg_Pict(6, 2))
              
              vaSpread1.Col = j + 4: vaSpread1.text = Val(RS!mid_codrec) & "&" & vg_tiprec & "&;"
              
              vaSpread1.Col = j + 5
              vaSpread1.CellType = CellTypeStaticText
              vaSpread1.TypeHAlign = TypeHAlignRight
              vaSpread1.text = Format(Trim(RS!candiet), fg_Pict(6, 2))
                          
             RS.MoveNext
           Loop
        End If
   RS.Close: Set RS = Nothing: fg_descarga
End If

If Not sw And vg_IndpprSelec = 1 Then    '--->Trae estructura completa si no hay registros de minuta.
   Set RS = vg_db.Execute("sgpadm_s_estservicio 1, " & vg_codservicio & ",''")
   If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
   Do While Not RS.EOF
      vaSpread1.Row = RS!ess_orden
      If indrow3 < vaSpread1.Row Then indrow3 = vaSpread1.Row
      vaSpread1.Col = 1
      vaSpread1.text = RS!ess_nombre
      For i = 2 To vaSpread1.maxcols Step 6
          vaSpread1.Col = vaSpread1.maxcols
          vaSpread1.text = RS!ess_codigo
      Next i
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
ElseIf Not sw And vg_IndpprSelec = 2 Then
   indrow3 = 20
End If

If vg_IndpprSelec <> 2 Then
    For i = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
        vaSpread1.Row = 0: vaSpread1.Col = i
        If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
            Dim fil As Long, Col As Long
            For fil = 1 To (vaSpread1.MaxRows - 1)
                For Col = i - 1 To i + 2
                    vaSpread1.Row = fil: vaSpread1.Col = Col
                    If vaSpread1.CellType = CellTypeNumber Then
                       vaSpread1.CellType = CellTypeStaticText
                       vaSpread1.TypeHAlign = TypeHAlignRight
                    End If
                    vaSpread1.BackColor = Shape1(1).FillColor
                Next Col
            Next fil
        End If
    Next i

   For i = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
       vaSpread1.Row = 0: vaSpread1.Col = i
       If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
          For fil = 1 To (vaSpread1.MaxRows - 1)
              For Col = i - 1 To i + 4
                  vaSpread1.Row = fil: vaSpread1.Col = Col
                  If vaSpread1.CellType = CellTypeNumber Then
                     vaSpread1.CellType = CellTypeStaticText
                     vaSpread1.TypeHAlign = TypeHAlignRight
                  End If
                  vaSpread1.BackColor = Shape1(1).FillColor
              Next Col
          Next fil
       End If
   Next i
End If

vaSpread1.MaxRows = indrow3 + 1
vaSpread1.Row = vaSpread1.MaxRows
maxfila = vaSpread1.MaxRows
vaSpread1.Col = 1
vaSpread1.text = "Comensales"
vaSpread1.Col = -1: vaSpread1.BackColor = &HE0E0E0
'-------> formatear ultima columna
For i = 2 To (vaSpread1.maxcols - MaxColumna) Step 6
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = i + 2
    vaSpread1.CellType = CellTypeNumber
    vaSpread1.TypeNumberDecPlaces = 0
    vaSpread1.TypeIntegerMin = 1
    vaSpread1.TypeIntegerMax = 9999999
    vaSpread1.TypeHAlign = TypeHAlignRight
    vaSpread1.TypeSpin = False
    vaSpread1.TypeIntegerSpinInc = 1
    vaSpread1.TypeIntegerSpinWrap = False
    vaSpread1.Value = Format(0, fg_Pict(6, 0))
    vaSpread1.ForeColor = &HFF0000
Next i

Set RS = vg_db.Execute("sgpadm_s_planifminuta 2, " & vg_codsubseg & ", " & vg_codregimen & ", " & vg_codservicio & ", " & vg_Zona & "," & Val(vg_fecha) & ", 0, 0," & vg_IndpprSelec & "")
DoEvents
If Not RS.EOF Then
   Do While Not RS.EOF
      DoEvents
      j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 6) - 6) + 1) + 1
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = j + 2
      vaSpread1.CellType = CellTypeNumber
      vaSpread1.TypeNumberDecPlaces = 0
      vaSpread1.TypeIntegerMin = 1
      vaSpread1.TypeIntegerMax = 9999999
      vaSpread1.TypeHAlign = TypeHAlignRight
      vaSpread1.TypeSpin = False
      vaSpread1.TypeIntegerSpinInc = 1
      vaSpread1.TypeIntegerSpinWrap = False
      vaSpread1.Value = IIf(IsNull(RS!min_racteo), 0, RS!min_racteo)
      vaSpread1.ForeColor = &HFF0000
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
Else
   RS.Close: Set RS = Nothing
   Set RS = vg_db.Execute("sgpadm_s_servraciones " & vg_codservicio & "")
   DoEvents
   If Not RS.EOF Then
      Do While Not RS.EOF
         IndDia = 1
         For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
             If RS!sra_serdia = IIf(fg_Dia(vg_fecha & fg_pone_cero(IndDia, 2)) = 1, 7, Val(fg_Dia(vg_fecha & fg_pone_cero(IndDia, 2)) - 1)) Then
                vaSpread1.Col = i + 2
                vaSpread1.CellType = CellTypeNumber
                vaSpread1.TypeNumberDecPlaces = 0
                vaSpread1.TypeIntegerMin = 1
                vaSpread1.TypeIntegerMax = 9999999
                vaSpread1.TypeHAlign = TypeHAlignRight
                vaSpread1.TypeSpin = False
                vaSpread1.TypeIntegerSpinInc = 1
                vaSpread1.TypeIntegerSpinWrap = False
                vaSpread1.Value = RS!Raciones
                vaSpread1.ForeColor = &HFF0000
             End If
             IndDia = IndDia + 1
         Next i
         RS.MoveNext
      Loop
   End If
   RS.Close: Set RS = Nothing
End If

If vg_IndpprSelec <> 2 Then
   For i = 3 To (vaSpread1.maxcols - MaxColumna) Step 6
       vaSpread1.Row = 0: vaSpread1.Col = i
       If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
          For fil = 1 To (vaSpread1.MaxRows - 1)
              For Col = i - 1 To i + 2
                  vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = Col
                  If vaSpread1.CellType = CellTypeNumber Then
                     vaSpread1.CellType = CellTypeStaticText
                     vaSpread1.TypeHAlign = TypeHAlignRight
                  End If
              Next Col
          Next fil
       End If
   Next i
End If
vaSpread1.Row = 1: vaSpread1.Col = 1
iblockrow = vaSpread1.Row: aiblockrow = vaSpread1.Row
iblockrow2 = vaSpread1.Row: aiblockrow2 = vaSpread1.Row
iblockcol = vaSpread1.Col: aiblockcol = vaSpread1.Col
iblockcol2 = vaSpread1.Col: aiblockcol2 = vaSpread1.Col
End Sub

Sub Calctodia(Row As Long, Col As Long)
Dim X As Long, numrac As Long
Dim cosdia As Double
veccos((Int(Col / 6) + 1), 1) = 0: veccos((Int(Col / 6) + 1), 4) = 0
For X = 1 To (vaSpread1.MaxRows - 1)
    vaSpread1.Row = X
    vaSpread1.Col = Col + 1: numrac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
    vaSpread1.Col = Col + 2: cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
    vaSpread1.Col = Col + 3
    If Trim(vaSpread1.text) <> "" And numrac > 0 Then
       vaSpread1.Col = Col + 2: veccos((Int(Col / 6) + 1), 1) = Round(veccos((Int(Col / 6) + 1), 1) + (cosdia * numrac), vg_DCa)
'       vaSpread1.Col = Col + 1: veccos((Int(Col / 5) + 1), 4) = Round(veccos((Int(Col / 5) + 1), 4) + numrac, vg_DCa)
    End If
Next X
vaSpread1.Row = vaSpread1.MaxRows
vaSpread1.Col = Col + 1: veccos((Int(Col / 6) + 1), 4) = Round(veccos((Int(Col / 6) + 1), 4) + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DCa)
End Sub

Sub MostrarCosto(Col As Long)
Dim xcol As Long
Dim toapla As Double, toaesf As Double, toafoo As Double, totdia As Double, totesf As Double, nracre As Double, nracfo As Double, totrac As Double
vaSpread1.Col = Col
xcol = 0
For i = 1 To MaxColumna
    If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then xcol = vectorcol(i): Exit For
Next i
vaSpread1.Row = 0: vaSpread1.Col = xcol: Frame2(2).Caption = vaSpread1.text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
toapla = 0: toaesf = 0: toafoo = 0: totdia = 0: totesf = 0: nracre = 0: nracfo = 0: totrac = 0
For i = 1 To UBound(veccos)
    If i <= (Int(xcol / 5) + 1) Then
       toapla = CCur(toapla + veccos(i, 1))
       toaesf = CCur(toaesf + veccos(i, 2))
       toafoo = CCur(toafoo + veccos(i, 3))
       nracre = CCur(nracre + veccos(i, 4))
       nracfo = CCur(nracfo + veccos(i, 5))
    End If
    totrac = CCur(totrac + veccos(i, 4))
    totdia = CCur(totdia + veccos(i, 1))
    totesf = CCur(totesf + veccos(i, 2))
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
If totrac > 0 Then Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2)) Else Label1(40).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(41).Caption = Format(CCur(totesf / totrac), fg_Pict(6, 2)) Else Label1(41).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(8).Caption = Format(CCur((totdia + totesf) / totrac), fg_Pict(6, 2)) Else Label1(8).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(48).Caption = Format(totrac, fg_Pict(6, 2)) Else Label1(48).Caption = Format(0, fg_Pict(6, 2))
Label1(20).Caption = Format(veccos((Int(xcol / 6) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format(veccos((Int(xcol / 6) + 1), 2), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))), fg_Pict(6, 2))
Label1(23).Caption = Format(veccos((Int(xcol / 6) + 1), 3), fg_Pict(6, 2))
Label1(44).Caption = Format(veccos((Int(xcol / 6) + 1), 4), fg_Pict(6, 2))
If veccos((Int(xcol / 6) + 1), 4) > 0 Then Label1(45).Caption = Format(CCur((veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))) / veccos((Int(xcol / 6) + 1), 4)), fg_Pict(6, 2)) Else Label1(45).Caption = Format(0, fg_Pict(6, 2))
Label1(46).Caption = Format(veccos((Int(xcol / 6) + 1), 5), fg_Pict(6, 2))
If veccos((Int(xcol / 6) + 1), 5) > 0 Then Label1(47).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 3) / veccos((Int(xcol / 6) + 1), 5)), fg_Pict(6, 2)) Else Label1(47).Caption = Format(0, fg_Pict(6, 2))
Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
Label1(32).Caption = Format((toaesf), fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(toapla + (toaesf)), fg_Pict(6, 2))
Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((toapla + toaesf) / nracre), fg_Pict(6, 2)) Else Label1(35).Caption = Format(0, fg_Pict(6, 2))
Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2)) Else Label1(38).Caption = Format(0, fg_Pict(6, 2))
End Sub

Sub CargarAporteCalorico()
Dim i, j, racion As Long
Dim caloria, RacxCal, SumCal, TotalAporte As Double
racion = 0: caloria = 0: RacxCal = 0: SumCal = 0: TotalAporte = 0
For i = 2 To vaSpread1.maxcols Step 6
  SumCal = 0: TotalAporte = 0
  For j = 1 To vaSpread1.MaxRows
    vaSpread1.Col = i: vaSpread1.Row = j
    If vaSpread1.text = "R" Then
      RacxCal = 0: caloria = 0
      vaSpread1.Col = i + 2 'Racion
      If Trim(vaSpread1.text) <> "" Then
         racion = Val(vaSpread1.text)
      Else
         racion = 0
      End If
      
      vaSpread1.Col = i + 5 'Caloria
      
      caloria = Trim(vaSpread1.text)
      If caloria <> "" Then
      RacxCal = (racion * caloria) ' Ración por Caloria
      End If
      SumCal = SumCal + RacxCal
    End If
  Next j
  vaSpread1.Col = i + 2 ' Posición Ración
  vaSpread1.Row = vaSpread1.MaxRows
  If IsNull(vaSpread1.text) = False And vaSpread1.text > "0" Then
     TotalAporte = SumCal / vaSpread1.text
  
     vaSpread1.Col = i + 5: vaSpread1.Row = vaSpread1.MaxRows
     vaSpread1.text = Format(TotalAporte, fg_Pict(6, 2))
     vaSpread1.ForeColor = &HFF0000
'     vaSpread1.Lock = True
  Else
     vaSpread1.Col = i + 5: vaSpread1.Row = vaSpread1.MaxRows
     vaSpread1.text = ""
     vaSpread1.ForeColor = &HFF0000
'     vaSpread1.Lock = True
  End If
Next i
End Sub

Sub CargarCosto()
fg_carga ""
vaSpread1.Col = vaSpread1.ActiveCol
If vaSpread1.Col = 1 Then vaSpread1.Col = 3
Dim cosdia As Double, totdia As Double, totesf As Double, totrac As Double
Dim Fecha As Long, xcol As Long, IndDia As Long, fecesf As Double, nracre As Long, nracfo As Long
Dim aAp As String
j = 0: fecval = 0: cosdia = 0: totdia = 0: totesf = 0: fecesf = 0: IndDia = 1: numrac = 0: totrac = 0
For i = 1 To MaxColumna
    If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then xcol = vectorcol(i): Exit For
Next i
vaSpread1.Row = 0: vaSpread1.Col = xcol: Frame2(2).Caption = vaSpread1.text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
ReDim veccos(MaxColumna, 5)
'------------ Calcular costo día planificado & estructura fija & salida
Bar1(0).Min = 0: Bar1(0).Value = 0: Bar1(0).Max = MaxColumna: Frame2(4).Visible = True: Bar1(0).Visible = True
For j = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
    Bar1(0).Value = Bar1(0).Value + 1
    If IndDia > MaxColumna Then Exit Sub
    Fecha = Val(vg_fecha) & Right("0" & IndDia, 2)
    veccos(IndDia, 1) = 0: veccos(IndDia, 2) = 0: veccos(IndDia, 3) = 0: veccos(IndDia, 4) = 0: veccos(IndDia, 5) = 0
    For i = 1 To (vaSpread1.MaxRows - 1)
        vaSpread1.Row = i
        vaSpread1.Col = j + 2: numrac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = j + 3: cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = j + 4
        If Trim(vaSpread1.text) <> "" And numrac > 0 Then
           totdia = Round(totdia + (cosdia * numrac), vg_DCa)
           veccos(IndDia, 1) = Round(veccos(IndDia, 1) + (cosdia * numrac), vg_DCa)
        End If
    Next i
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = j + 2
    veccos(IndDia, 4) = Round(veccos(IndDia, 4) + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
    totrac = Round(totrac + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
    If fecesf > 0 Then
    End If
    IndDia = IndDia + 1
Next j
Frame2(4).Visible = False
Bar1(0).Visible = False
'------------ Fin Calcular costo día
toapla = 0: toaesf = 0: toafoo = 0: numrac = 0: nracfo = 0
For i = 1 To (Int(xcol / 6) + 1)
    toapla = Round(toapla + veccos(i, 1), vg_DPr)
    toaesf = Round(toaesf + veccos(i, 2), vg_DPr)
    toafoo = Round(toafoo + veccos(i, 3), vg_DPr)
    nracre = Round(nracre + veccos(i, 4), vg_DPr)
    nracfo = Round(nracfo + veccos(i, 5), vg_DPr)
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, vg_DPr))
Label1(11).Caption = Format(totesf, fg_Pict(6, vg_DPr))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, vg_DPr))
If totrac > 0 Then Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, vg_DPr)) Else Label1(40).Caption = Format(0, fg_Pict(6, vg_DPr))
If totrac > 0 Then Label1(41).Caption = Format(CCur(totesf / totrac), fg_Pict(6, vg_DPr)) Else Label1(41).Caption = Format(0, fg_Pict(6, vg_DPr))
If totrac > 0 Then Label1(8).Caption = Format(CCur((totdia + totesf) / totrac), fg_Pict(6, vg_DPr)) Else Label1(8).Caption = Format(0, fg_Pict(6, vg_DPr))
If totrac > 0 Then Label1(48).Caption = Format(totrac, fg_Pict(6, vg_DPr)) Else Label1(48).Caption = Format(0, fg_Pict(6, vg_DPr))
Label1(20).Caption = Format(veccos((Int(xcol / 6) + 1), 1), fg_Pict(6, vg_DPr))
Label1(21).Caption = Format((veccos((Int(xcol / 6) + 1), 2)), fg_Pict(6, vg_DPr))
Label1(22).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))), fg_Pict(6, vg_DPr))
Label1(23).Caption = Format(veccos((Int(xcol / 6) + 1), 3), fg_Pict(6, vg_DPr))
Label1(44).Caption = Format(nracre, fg_Pict(6, vg_DPr))
If nracre > 0 Then Label1(45).Caption = Format(CCur((veccos((Int(xcol / 6) + 1), 1) + (veccos((Int(xcol / 6) + 1), 2))) / nracre), fg_Pict(6, vg_DPr)) Else Label1(45).Caption = Format(0, fg_Pict(6, vg_DPr))
Label1(46).Caption = Format(nracfo, fg_Pict(6, vg_DPr))
If nracfo > 0 Then Label1(47).Caption = Format(CCur(veccos((Int(xcol / 6) + 1), 3) / nracfo), fg_Pict(6, vg_DPr)) Else Label1(47).Caption = Format(0, fg_Pict(6, vg_DPr))
Label1(31).Caption = Format(toapla, fg_Pict(6, vg_DPr))
Label1(32).Caption = Format((toaesf), fg_Pict(6, vg_DPr))
Label1(33).Caption = Format(CCur(toapla + (toaesf)), fg_Pict(6, vg_DPr))
Label1(34).Caption = Format(nracre, fg_Pict(6, vg_DPr))
If nracre > 0 Then Label1(35).Caption = Format(CCur((toapla + toaesf) / nracre), fg_Pict(6, vg_DPr)) Else Label1(35).Caption = Format(0, fg_Pict(6, vg_DPr))
Label1(36).Caption = Format(toafoo, fg_Pict(6, vg_DPr))
Label1(37).Caption = Format(nracfo, fg_Pict(6, vg_DPr))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, vg_DPr)) Else Label1(38).Caption = Format(0, fg_Pict(6, vg_DPr))
indcos = True
fg_descarga
End Sub

'Function ValidaMinuta(subseg As String, reg As String, Serv As String, TipPlan As String, Zona As String, Fec As String) As Boolean
''*****************---->Validar minuta en uso <---------------------------
''------ Esta funcion crea una tabla temporal concatenando los parametros ingresaods
''------ para la minuta, de esta manera permanece una tabla temporal identificando
''------ que alguien se encuentra conectado a esa minuta, si alguien
''------ mas quiere acceder, se dara un aviso que esta en uso
''------ esta tabla temporal se destruye cuando se cierra este formulario (evento Unload)
''------ y tambien si el usuario cierra la sesion SQL Server la destruye automaticamente.
''----------------------------------------------------------------------
'
'    Dim RSTempCheck As New ADODB.Recordset
'    Dim RS As New ADODB.Recordset
'    Dim RSTem As New ADODB.Recordset
'    NameTemp = subseg & reg & Serv & TipPlan & Zona & Fec
'
'    Set RSTempCheck = vg_db.Execute("select * from tempdb.dbo.sysobjects where xtype = 'U' and name = '##ValidaMinuta_" & NameTemp & "'")
'
'    If RSTempCheck.EOF And RSTempCheck.BOF Then
'        Set RSTem = vg_db.Execute("CREATE TABLE ##ValidaMinuta_" & NameTemp & " (usu_codigo VarChar(20))")
'        Set RS = vg_db.Execute("INSERT INTO ##ValidaMinuta_" & NameTemp & " (usu_codigo) values ('" & vg_NUsr & "')")
'        Set RS = Nothing
'        Set RSTem = Nothing
'        ValidaMinuta = True
'    Else
'        ValidaMinuta = False
'        Set RS = vg_db.Execute("SELECT usu_codigo from ##ValidaMinuta_" & NameTemp & " ")
'        If Not (RS.EOF = True And RS.BOF = True) Then
'            RS.MoveFirst
'            MsgBox "La minuta con los parametros ingresados, actualmente esta siendo usada por el usuario: '" & UCase(RS!usu_codigo) & "', podra ingresar cuando el usuario termine de trabajar en ella"
'        End If
'        RS.Close: Set RS = Nothing
'    End If
'
'RSTempCheck.Close
'Set RSTempCheck = Nothing
'End Function

Sub BlocSoloAcceso()
        ' en caso si tiene solo autorizacion para ver sin modificar ni grabar
        Toolbar1.Buttons(25).Enabled = False ' Actualizar
        Toolbar1.Buttons(4).Enabled = False ' Cortar
        Toolbar1.Buttons(5).Enabled = False ' copiar
        Toolbar1.Buttons(2).Enabled = False
        Toolbar1.Buttons(6).Enabled = False
        'Toolbar1.Buttons(10).Enabled = False ' buscar
        Toolbar1.Buttons(12).Enabled = False ' insertar fila
        Toolbar1.Buttons(13).Enabled = False ' eliminar fila
        Toolbar1.Buttons(15).Enabled = False ' subir fila
        Toolbar1.Buttons(16).Enabled = False ' bajar fila
        Toolbar1.Buttons(19).Enabled = False ' Copiar Minuta
End Sub

Sub CellEditEstruct()
' este procedimiento no se esta usando
' deja editables solo las celdas de la columna uno cuando su contenido
' es distinto de vacio
Dim i As Integer
    If ExraeCodCombo(M_Plami1.Combo2(1)) = 2 Then
        vaSpread1.Col = 1
       
        For i = 1 To vaSpread1.MaxRows
        DoEvents
            vaSpread1.Row = i
            If vaSpread1.text <> "" Then
                vaSpread1.TypeEditCharSet = TypeEditCharSetAlphanumeric
                vaSpread1.CellType = CellTypeEdit
            End If
            DoEvents
        Next i
    End If
End Sub

Sub HabilitaCol(Index As Integer)
Select Case Index
Case 0
    For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
    DoEvents
      vaSpread1.Row = 0
      vaSpread1.Col = i + 2
      If vaSpread1.ColHidden = True Then
         vaSpread1.ColHidden = False
         estapo = False
         DoEvents
      Else
         vaSpread1.ColHidden = True
         estapo = True
         DoEvents
      End If
    Next i
Case 1
    For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
    DoEvents
      vaSpread1.Row = 0
      vaSpread1.Col = i + 3
      DoEvents
      If vaSpread1.ColHidden = True Then
         vaSpread1.ColHidden = False
         estapo = False
      Else
         vaSpread1.ColHidden = True
         estapo = True
      End If
      DoEvents
    Next i
Case 2

    'Samuel Quie me quede no vuelve a mostrar columna
    vaSpread1.Row = -1
    vaSpread1.Col = 2
    If vaSpread1.ColHidden = True Then
       vaSpread1.ColHidden = False
       estapo = False
    Else
       vaSpread1.ColHidden = True
       estapo = True
    End If
    
    For i = 2 To (vaSpread1.maxcols - MaxColumna - 1) Step 6
        DoEvents
        vaSpread1.Row = 0
        vaSpread1.Col = i + 6
        If Trim(vaSpread1.text) <> "Estado" Then
        If vaSpread1.ColHidden = True Then
           vaSpread1.ColHidden = False
           estapo = False
        Else
           vaSpread1.ColHidden = True
           estapo = True
        End If
        End If
        DoEvents
    Next i
End Select
vaSpread1.Visible = True
End Sub

Private Function ValidaEstructuras() As Boolean
'----- el objetivo de esta funcion es encontrar recetas que no esten asignadas
'----- a una estructura, para lo cual comienza recorriendo desde la columna uno
'----- hacia abajo, preguntando hasta que encuentre recetas sin estructura
'----- por ejemplo si la celda (columna 1, fila 1) esta en blanco y la celda (columna 1, fila 2)
'----- es distinta de vacia, devolvera FALSE
Dim i As Integer, j As Integer
        ValidaEstructuras = True
        xrow = vaSpread1.ActiveRow
        vaSpread1.Row = 1
        vaSpread1.Col = 1
        
        If Trim(vaSpread1.text) <> "" Then ValidaEstructuras = True: Exit Function
        
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Row = i
            
            If Trim(vaSpread1.text) = "" Then
            
                For j = 2 To vaSpread1.maxcols
                DoEvents
                    vaSpread1.Col = j
                    If Trim(vaSpread1.text) <> "" Then
                        ValidaEstructuras = False
                        DoEvents
                        Exit Function
                    End If
                Next j

            Else
                ValidaEstructuras = True: Exit Function
            End If
            
        Next i
End Function

Function ValidaNombreEstructura(ByVal xrow As Integer, ByVal xSpread As vaSpread) As String
Dim Estruc1 As String, Estruc2 As String
Dim RSEst As New ADODB.Recordset
ValidaNombreEstructura = ""
Dim xResCol As Long, xResRow2 As Long


    xResCol = xSpread.Col
    xResRow2 = xSpread.Row
    
    xSpread.Row = xrow
    xSpread.Col = 1
    Estruc2 = xSpread.text
    
    If Estruc2 = "" Then ValidaNombreEstructura = "": Exit Function
    
    xSpread.Col = xSpread.maxcols
    Estruc1 = G_Proc.fg_ExtraeServicio(vg_codservicio, xSpread.text)
    
    If Trim(Estruc1) = Trim(Estruc2) Then
        ValidaNombreEstructura = ""
    Else
        ValidaNombreEstructura = Estruc2
    End If
    
    xSpread.Col = xResCol
    xSpread.Row = xResRow2
    
End Function

Sub AddEstructuraMenu(ByVal CodigoEst As Long, ByVal NombreEst As String)
On Error GoTo errSub
Dim i As Long
Load Estructura1.item(Estructura1.count)
Load Estructura2.item(Estructura2.count)

Estructura1.item(Estructura1.count - 1).Caption = NombreEst
Estructura1.item(Estructura1.count - 1).HelpContextID = CodigoEst
Estructura1.item(Estructura1.count - 1).Enabled = True
Estructura1.item(Estructura1.count - 1).Visible = True

Estructura2.item(Estructura2.count - 1).Caption = NombreEst
Estructura2.item(Estructura2.count - 1).HelpContextID = CodigoEst
Estructura2.item(Estructura2.count - 1).Enabled = True
Estructura2.item(Estructura2.count - 1).Visible = True
'Estructura2.Item(Estructura2.count - 1) = True

Exit Sub
errSub:
    On Local Error Resume Next
    MsgBox Err.Description, vbCritical
End Sub

Function EstructuraSuperior(ByVal Spread As vaSpread, ByVal Fila As Long) As String
Dim xRespRow As Long, xRespCol As Long
Dim x1 As Long

EstructuraSuperior = ""
xRespRow = Spread.Row
xRespCol = Spread.Col

'AvisoEst = False
Spread.Row = Fila
Spread.Col = Spread.maxcols - 1
For x1 = Spread.Row To 1 Step -1
    Spread.Row = x1
    If Trim(Spread.text) <> "" Then
       EstructuraSuperior = Spread.text
       Exit For
    End If
Next x1
        
Spread.Row = xRespRow
Spread.Col = xRespCol
End Function

Sub ActualizaEstructuraInferior(ByVal Spread As vaSpread, ByVal NameEstruct As String, Optional UltimoCambio As Long)
Dim xRespRow As Long, xRespCol As Long, EstructuraAnterior As String
Dim x1 As Long
xRespRow = Spread.Row
xRespCol = Spread.Col
Spread.Col = Spread.maxcols - 1
EstructuraAnterior = Spread.text
For x1 = Spread.Row To Spread.MaxRows - 1
    Spread.Row = x1
    If Trim(Spread.text) = "" Or Spread.Row = Spread.MaxRows Or Trim(EstructuraAnterior) <> Trim(Spread.text) Then Exit For
    Spread.text = NameEstruct
Next x1
Spread.Row = xRespRow
Spread.Col = xRespCol
End Sub

Sub DesqloqSubMenu(OpcioneMenu As String)
Dim iA As Integer
For iA = 1 To Estructura2.count - 1
    If Trim(Estructura2.item(iA).Caption) = Trim(OpcioneMenu) Then
       Estructura1.item(iA).Enabled = True
       Estructura2.item(iA).Enabled = True
    End If
Next iA
End Sub

Sub Deshacer(StrRec As Variant)
'load in file
Dim ret As Integer
Screen.MousePointer = 11
ret = vaSpread1.LoadFromFile(LCase(App.Path) & "\" & StrRec)
If Dir(LCase(App.Path) & "\" & StrRec) <> "" Then Kill LCase(App.Path) & "\" & StrRec
CorDes = CorDes - 1
Screen.MousePointer = 0
End Sub

Sub GrabarCambios(ifil As Long, icol As Long, estado As String)
Dim ret
CorDes = CorDes + 1
ret = vaSpread1.SaveToFile(LCase(App.Path) & "\" & "spread" & vg_NUsr & CorDes & ".ss6", False)
Toolbar1.Buttons(31).Visible = True
Toolbar1.Buttons(31).Enabled = True
End Sub
