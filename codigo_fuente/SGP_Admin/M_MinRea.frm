VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_MinRea 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Real"
   ClientHeight    =   6930
   ClientLeft      =   405
   ClientTop       =   2310
   ClientWidth     =   11685
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6930
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   3765
      Index           =   0
      Left            =   330
      TabIndex        =   13
      Top             =   2850
      Visible         =   0   'False
      Width           =   9705
      Begin VB.Frame Frame2 
         Height          =   1365
         Index           =   4
         Left            =   60
         TabIndex        =   56
         Top             =   180
         Width           =   4755
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   240
            Picture         =   "M_MinRea.frx":0000
            Top             =   240
            Width           =   480
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Index           =   1
         Left            =   60
         TabIndex        =   40
         Top             =   1560
         Width           =   4755
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo día"
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
            Left            =   2190
            TabIndex        =   55
            Top             =   300
            Width           =   840
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
            Left            =   3480
            TabIndex        =   54
            Top             =   300
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Materia Prima"
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
            Left            =   150
            TabIndex        =   53
            Top             =   555
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estructura Fija"
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
            Top             =   840
            Width           =   1245
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
            Top             =   1140
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
            Top             =   1440
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
            Top             =   1740
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
            Left            =   3090
            TabIndex        =   48
            Top             =   555
            Width           =   1560
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
            Left            =   1500
            TabIndex        =   47
            Top             =   1140
            Width           =   1530
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
            Left            =   1500
            TabIndex        =   46
            Top             =   1440
            Width           =   1530
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
            Left            =   1500
            TabIndex        =   45
            Top             =   1740
            Width           =   1530
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
            Left            =   3090
            TabIndex        =   44
            Top             =   840
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   43
            Top             =   1140
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   42
            Top             =   1440
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   41
            Top             =   1740
            Width           =   1560
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
         Height          =   1365
         Index           =   2
         Left            =   4890
         TabIndex        =   30
         Top             =   180
         Width           =   4755
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Materia Prima"
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
            Top             =   435
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estructura Fija"
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
            Top             =   720
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo Total"
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
            Top             =   1005
            Width           =   990
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
            Left            =   2130
            TabIndex        =   36
            Top             =   150
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
            Left            =   3750
            TabIndex        =   35
            Top             =   150
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
            Left            =   1440
            TabIndex        =   34
            Top             =   435
            Width           =   1560
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
            Left            =   1440
            TabIndex        =   33
            Top             =   720
            Width           =   1560
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
            Left            =   1440
            TabIndex        =   32
            Top             =   1005
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   31
            Top             =   1005
            Width           =   1560
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
         Height          =   2115
         Index           =   3
         Left            =   4890
         TabIndex        =   14
         Top             =   1560
         Width           =   4755
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Materia Prima"
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
            Top             =   555
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estructura Fija"
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
            Top             =   840
            Width           =   1245
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo Total"
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
            Top             =   1140
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comensales"
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
            Top             =   1440
            Width           =   1020
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
            Index           =   28
            Left            =   90
            TabIndex        =   25
            Top             =   1740
            Width           =   1065
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
            Left            =   2145
            TabIndex        =   24
            Top             =   300
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
            Left            =   3750
            TabIndex        =   23
            Top             =   300
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
            Left            =   1440
            TabIndex        =   22
            Top             =   555
            Width           =   1560
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
            Left            =   1440
            TabIndex        =   21
            Top             =   840
            Width           =   1560
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
            Left            =   1440
            TabIndex        =   20
            Top             =   1140
            Width           =   1560
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
            Left            =   1440
            TabIndex        =   19
            Top             =   1440
            Width           =   1560
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
            Left            =   1440
            TabIndex        =   18
            Top             =   1740
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   17
            Top             =   1140
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   16
            Top             =   1440
            Width           =   1560
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
            Left            =   3090
            TabIndex        =   15
            Top             =   1740
            Width           =   1560
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
      Height          =   435
      Left            =   30
      TabIndex        =   0
      Top             =   840
      Width           =   10905
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estructura de Servicio"
         Height          =   195
         Index           =   2
         Left            =   5505
         TabIndex        =   12
         Top             =   135
         Width           =   1560
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   5160
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Bloqueada"
         Height          =   195
         Index           =   1
         Left            =   7650
         TabIndex        =   11
         Top             =   135
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008484FF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   1
         Left            =   7305
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         Height          =   195
         Index           =   0
         Left            =   9405
         TabIndex        =   10
         Top             =   135
         Width           =   1155
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   9060
         Top             =   165
         Width           =   300
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
         TabIndex        =   1
         Top             =   150
         Width           =   1215
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      DragIcon        =   "M_MinRea.frx":030A
      Height          =   4140
      Left            =   -15
      TabIndex        =   9
      Top             =   1305
      Width           =   11055
      _Version        =   393216
      _ExtentX        =   19500
      _ExtentY        =   7303
      _StockProps     =   64
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
      MaxCols         =   249
      MaxRows         =   100
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RestrictRows    =   -1  'True
      SpreadDesigner  =   "M_MinRea.frx":074C
      UserResize      =   1
      VisibleCols     =   1
      VisibleRows     =   100
      TextTip         =   2
      TextTipDelay    =   0
      ScrollBarTrack  =   3
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planificación Minutas Real"
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
      Top             =   345
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
         Visible         =   0   'False
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
         Caption         =   "Aportes &Nutricionales x Días"
         Index           =   10
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Frecuencia Recetas"
         Index           =   11
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Cerrar"
         Index           =   20
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
         Index           =   13
         Shortcut        =   ^V
      End
      Begin VB.Menu Plato 
         Caption         =   "&Agregar Estructura"
         Index           =   14
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
         Index           =   13
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Agrega Estructura"
         Index           =   14
         Begin VB.Menu Estructura2 
            Caption         =   ""
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "M_MinRea"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset, RS1 As New ADODB.Recordset
Dim NewRow As Long, i As Long, J As Long, indcortarpegar As Long, wsmaxfilas As Long, fecha As Long, maxcolumna As Long
Dim iblockrow As Integer, iblockrow2 As Integer, iblockcol As Integer, iblockcol2 As Integer, SwSalir As Integer
Dim aiblockrow As Integer, aiblockrow2 As Integer, aiblockcol As Integer, aiblockcol2 As Integer, indactivo As Integer, indgrabado As Integer
Dim indcos As Boolean
Dim veccos() As Variant
Dim vectorcol() As Long
Dim MsgTitulo As String

Private Sub Estructura1_Click(Index As Integer)
LlenaSubMenu Estructura1, Index
End Sub

Sub LlenaSubMenu(SubMenu As Object, Index As Integer)
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
If Trim(vaSpread1.Text) <> "" Then MsgBox "Seleccione una celda de estructura que este vacía...", vbInformation, MsgTitulo: Exit Sub
vaSpread1.Text = SubMenu(Index).Caption
vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.Text = SubMenu(Index).HelpContextID
Estructura1(Index).Enabled = False: Estructura2(Index).Enabled = False
indgrabado = 1
Plantilla(0).Enabled = True
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
End Sub

Private Sub Estructura2_Click(Index As Integer)
LlenaSubMenu Estructura2, Index
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 6765
Me.Width = 11055
fg_centra Me
MsgTitulo = "Planificación Real"
fg_carga (ss)
Label4.Caption = M_Plami1.fpayuda(1).Text & "(" & M_Plami1.fpText.Text & ")" & " - " & M_Plami1.fpayuda(2).Text & " - " & M_Plami1.fpayuda(3).Text
Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "
indcos = False
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): btnX.Visible = True: btnX.ToolTipText = " "
Set btnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): btnX.Visible = False: btnX.ToolTipText = "Grabar Datos": btnX.Enabled = IIf(Mid(ValidarUsuario(M_Plami1), 2, 2) = "0", False, True)
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Cortar", , tbrDefault, "A_Cortar"): btnX.Visible = True: btnX.ToolTipText = "Cortar"
Set btnX = Toolbar1.Buttons.Add(, "A_Copiar", , tbrDefault, "A_Copiar"): btnX.Visible = True: btnX.ToolTipText = "Copiar"
Set btnX = Toolbar1.Buttons.Add(, "I_Pegar", , tbrDefault, "I_Pegar"): btnX.Visible = True: btnX.ToolTipText = ""
Set btnX = Toolbar1.Buttons.Add(, "A_Pegar", , tbrDefault, "A_Pegar"): btnX.Visible = False: btnX.ToolTipText = "Pegar"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): btnX.Visible = True: btnX.ToolTipText = "Insertar"
Set btnX = Toolbar1.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): btnX.Visible = True: btnX.ToolTipText = "Eliminar"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_SubirF", , tbrDefault, "A_SubirF"): btnX.Visible = True: btnX.ToolTipText = "Subir"
Set btnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): btnX.Visible = True: btnX.ToolTipText = "Bajar"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_VerReceta", , tbrDefault, "A_VerReceta"): btnX.Visible = True: btnX.ToolTipText = "Ver Recetas"
Set btnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): btnX.Visible = True: btnX.ToolTipText = "Copiar Planificación Teórica"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Aportes", , tbrDefault, "A_Aportes"): btnX.Visible = True: btnX.ToolTipText = "Aportes Nutricionales x Días"
Set btnX = Toolbar1.Buttons.Add(, "A_Costo", , tbrDefault, "A_Costo"): btnX.Visible = False: btnX.ToolTipText = "Costo"
Set btnX = Toolbar1.Buttons.Add(, "A_Frecuencia", , tbrDefault, "A_Frecuencia"): btnX.Visible = True: btnX.ToolTipText = "Frecuencia Recetas"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): btnX.Visible = False: btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"

DetallePlantillaMinuta
' Llena Sub Menu estructura
Dim X As Long
RS.Open "select ess_nombre, ess_codigo from a_estservicio where ess_codser=" & vg_codservicio & " order by ess_orden", vg_db, adOpenStatic

If Not RS.EOF Then
    X = 1
    Do While Not RS.EOF
        Load Estructura1(X): Load Estructura2(X)
        Estructura1(X).Caption = Trim(RS!ess_nombre): Estructura2(X).Caption = Trim(RS!ess_nombre)
        Estructura1(X).HelpContextID = RS!ess_codigo: Estructura2(X).HelpContextID = RS!ess_codigo
        Estructura1(X).Enabled = True: Estructura2(X).Enabled = True
        For i = 1 To vaSpread1.MaxRows
            vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.Row = i
            If Trim(vaSpread1.Text) <> "" Then
                If Val(vaSpread1.Text) = RS!ess_codigo Then Estructura1(X).Enabled = False: Estructura2(X).Enabled = False
            End If
        Next
        X = X + 1
        RS.MoveNext
    Loop
End If
RS.Close: Set RS = Nothing
Estructura1(0).Visible = False: Estructura2(0).Visible = False
End Sub

Private Sub Form_Resize()
If Frame2(0).Visible = True Then
   If Me.WindowState = 0 Then
      vaSpread1.Height = 840
      Frame2(0).Top = 2230
      Frame2(0).Visible = True
      Exit Sub
   ElseIf Me.WindowState = 2 Then
'      vaSpread1.Height = 2540
      vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - (ScaleHeight / 2)
      Frame2(0).Move 0, ScaleHeight - 4000, (ScaleWidth - (ScaleWidth / 20)), (ScaleHeight - (ScaleHeight / 20)) '- 1380 'ScaleHeight - 3800 'Frame2(0).Top = 4000
      Frame2(0).Visible = True
   End If
Else
   If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
   If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 445
   If Me.WindowState <> 1 Then vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
If SwSalir <> 0 Then Exit Sub
If indgrabado <> 1 Then Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then indgrabado = 0: Cancel = -1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
If indgrabado = 1 Then GrabarPlantillaMinuta
indgrabado = 0
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
SwSalir = 1
Me.Hide
Unload Me
M_Plami1.WindowState = 0
End Sub

Private Sub image1_Click(Index As Integer)
fg_carga ""
vaSpread1.Col = vaSpread1.ActiveCol
If vaSpread1.Col = 1 Then vaSpread1.Col = 3
Dim cosdia As Double, totdia As Double, totesf As Double
Dim fecha As Long, xcol As Long, inddia As Long, fecesf As Double, nracre As Long, nracfo As Long
J = 0: fecval = 0: cosdia = 0: totdia = 0: totesf = 0: fecesf = 0: inddia = 1: numrac = 0
For i = 1 To maxcolumna
    If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then xcol = vectorcol(i): Exit For
Next i
vaSpread1.Row = 0: vaSpread1.Col = xcol: Frame2(1).Caption = vaSpread1.Text: Frame2(2).Caption = vaSpread1.Text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.Text
ReDim Preserve veccos(maxcolumna, 5)
'------------ Buscar fecha estructura fija
RS.Open "select max(mif_fecval) as fecval from b_minutafija " & _
        "where mif_cencos='" & vg_codcasino & "' " & _
        "and   mif_codreg=" & vg_codregimen & " " & _
        "and   mif_codser=" & vg_codservicio & "", vg_db, adOpenStatic
If Not RS.EOF And IsNull(RS!fecval) = False Then fecesf = RS!fecval
RS.Close: Set RS = Nothing
'------------

'------------ Calcular costo día planificado & estructura fija & salida
'    fecval = Val(vg_fecha) & Right("0" & (Int(j / 5) + 1), 2)
For J = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
    fecha = Val(vg_fecha) & Right("0" & inddia, 2)
    veccos(inddia, 1) = 0: veccos(inddia, 2) = 0: veccos(inddia, 3) = 0: veccos(inddia, 4) = 0: veccos(inddia, 5) = 0
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = J + 2: numrac = Val(vaSpread1.Text): vaSpread1.Col = J + 4
        If Trim(vaSpread1.Text) <> "" And numrac > 0 Then
           RS.Open "select distinct b_minutadet.mid_fecval from b_minuta, b_minutadet " & _
                   "where  b_minuta.min_codigo=b_minutadet.mid_codigo " & _
                   "and    b_minuta.min_cencos='" & vg_codcasino & "' " & _
                   "and    b_minuta.min_codreg=" & vg_codregimen & " " & _
                   "and    b_minuta.min_codser=" & vg_codservicio & " " & _
                   "and    b_minuta.min_fecmin=" & fecha & " " & _
                   "and    b_minutadet.mid_tipmin='2' " & _
                   "and    b_minutadet.mid_numrac>0 " & _
                   "and    b_minutadet.mid_fecval>0", vg_db, adOpenStatic
           If Not RS.EOF Then totdia = Round(totdia + (fg_CalCtoRecPlan(RS!mid_fecval, 2, vaSpread1.Text)), vg_DCa): veccos(inddia, 1) = Round(veccos(inddia, 1) + (fg_CalCtoRecPlan(RS!mid_fecval, 2, vaSpread1.Text)), vg_DCa): veccos(inddia, 4) = Round(veccos(inddia, 4) + numrac, vg_DPr)
           RS.Close: Set RS = Nothing
        End If
    Next i
    If fecesf > 0 Then
       RS.Open "select b_minutafija.mif_dianro, sum(b_productos.pro_propon*b_minutafija.mif_canpro) as cosesf " & _
               "from   b_productos, b_minutafija " & _
               "where  b_minutafija.mif_codpro=b_productos.pro_codigo " & _
               "and    b_productos.pro_ctacon in ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
               "and    b_minutafija.mif_cencos='" & vg_codcasino & "' " & _
               "and    b_minutafija.mif_codreg=" & vg_codregimen & " " & _
               "and    b_minutafija.mif_codser=" & vg_codservicio & " " & _
               "and    b_minutafija.mif_fecval=" & fecesf & " " & _
               "and    b_minutafija.mif_dianro=" & fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))) & " " & _
               "group by b_minutafija.mif_dianro", vg_db, adOpenStatic
       If Not RS.EOF Then totesf = Round(totesf + RS!cosesf, vg_DCa): veccos(inddia, 2) = Round(veccos(inddia, 2) + RS!cosesf, vg_DCa)
       RS.Close: Set RS = Nothing
    End If
    RS.Open "select b_totventas.tov_codreg, b_totventas.tov_codser, " & _
            "sum(b_totventas.tov_totdoc) as totdoc " & _
            "from  b_totventas " & _
            "where b_totventas.tov_codreg=" & vg_codregimen & " " & _
            "and   b_totventas.tov_codser=" & vg_codservicio & " " & _
            "and   b_totventas.tov_tipdoc='SP' " & _
            "and   b_totventas.tov_estdoc<>'A' " & _
            "and   b_totventas.tov_fecpro=cdate('" & fg_Ctod1(Val(vg_fecha) & Right("0" & inddia, 2)) & "') " & _
            "group by b_totventas.tov_codreg, b_totventas.tov_codser", vg_db, adOpenStatic
    If Not RS.EOF Then veccos(inddia, 3) = Round(veccos(inddia, 3) + RS!totdoc, vg_DCa)
    RS.Close: Set RS = Nothing
                    
    RS.Open "select b_totventas.tov_codreg, b_totventas.tov_codser, " & _
            "sum(b_totventas.tov_totdoc) as totdoc " & _
            "from  b_totventas " & _
            "where b_totventas.tov_codreg=" & vg_codregimen & " " & _
            "and   b_totventas.tov_codser=" & vg_codservicio & " " & _
            "and   b_totventas.tov_tipdoc='DP' " & _
            "and   b_totventas.tov_estdoc<>'A' " & _
            "and   b_totventas.tov_fecpro=cdate('" & fg_Ctod1(Val(vg_fecha) & Right("0" & inddia, 2)) & "') " & _
            "group by b_totventas.tov_codreg, b_totventas.tov_codser", vg_db, adOpenStatic
    If Not RS.EOF Then: veccos(inddia, 3) = Round(veccos(inddia, 3) - RS!totdoc, vg_DCa)
    RS.Close: Set RS = Nothing
    
    RS.Open "select sum(mir_nrorac) as mir_nrorac from b_minutaraciones " & _
            "where  mir_cencos='" & vg_codcasino & "' " & _
            "and    mir_codreg=" & vg_codregimen & " " & _
            "and    mir_codser=" & vg_codservicio & " " & _
            "and    mir_fecmin=" & Val(vg_fecha) & Right("0" & inddia, 2) & "", vg_db, adOpenStatic
    If Not RS.EOF Then veccos(inddia, 5) = Round(veccos(inddia, 5) + RS!mir_nrorac, vg_DPr)
    RS.Close: Set RS = Nothing
    inddia = inddia + 1
Next J
'------------ Fin Calcular costo día
toapla = 0: toaesf = 0: toafoo = 0: numrac = 0: nracfo = 0
For i = 1 To (Int(xcol / 5) + 1)
    toapla = Round(toapla + veccos(i, 1), vg_DCa)
    toaesf = Round(toaesf + veccos(i, 2), vg_DCa)
    toafoo = Round(toafoo + veccos(i, 3), vg_DCa)
    nracre = Round(nracre + veccos(i, 4), vg_DPr)
    nracfo = Round(nracfo + veccos(i, 5), vg_DPr)
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
Label1(8).Caption = Format(veccos((Int(xcol / 5) + 1), 1), fg_Pict(6, 2))
Label1(20).Caption = Format(veccos((Int(xcol / 5) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format(veccos((Int(xcol / 5) + 1), 2), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(veccos((Int(xcol / 5) + 1), 1) + veccos((Int(xcol / 5) + 1), 2)), fg_Pict(6, 2))
Label1(23).Caption = Format(veccos((Int(xcol / 5) + 1), 3), fg_Pict(6, 2))
Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
Label1(32).Caption = Format(toaesf, fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(toapla + toaesf), fg_Pict(6, 2))
Label1(34).Caption = nracre 'Format(veccos((Int(xcol / 5) + 1), 4), fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((toapla + toaesf) / nracre), fg_Pict(6, 2))
Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2))
indcos = True
fg_descarga
End Sub

Private Sub Plantilla_Click(Index As Integer)
Select Case Index
Case 0
    If Toolbar1.Buttons(2).Enabled = False Then indgrabado = 0: Exit Sub
    If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Cancel = -1: Exit Sub
    If indgrabado = 1 Then GrabarPlantillaMinuta
    indgrabado = 0
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
Case 3
    If Frame2(0).Visible = True Then Frame2(0).Visible = False: vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380: Exit Sub
    vaSpread1.Height = 2540
    Frame2(0).Top = 4000
    Frame2(0).Visible = True
Case 5
    Dim xcol As Integer
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Then vg_newestrec = True Else vg_newestrec = False
    xcol = 0
    For i = 1 To maxcolumna
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) And Trim(vaSpread1.Text) <> "" Then xcol = vectorcol(i): Exit For
    Next i
    If xcol = 0 Then MsgBox "No existe receta ha vizualizar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If vg_newestrec = True Then
       vg_fecval = 0: vg_fecval = Val(vg_fecha) & Right("0" & (Int(xcol / 5) + 1), 2)
       RS.Open "select b_minutadet.mid_fecval from  b_minuta, b_minutadet " & _
               "where b_minuta.min_codigo=b_minutadet.mid_codigo " & _
               "and b_minuta.min_cencos='" & vg_codcasino & "' " & _
               "and b_minuta.min_codreg=" & vg_codregimen & " " & _
               "and b_minuta.min_codser=" & vg_codservicio & " " & _
               "and b_minuta.min_fecmin=" & vg_fecval & " " & _
               "and b_minutadet.mid_tipmin='2' " & _
               "and b_minutadet.mid_fecval>0", vg_db, adOpenStatic
       If Not RS.EOF Then vg_fecval = RS!mid_fecval: vg_opcion = 2
       RS.Close: Set RS = Nothing
    End If
    vaSpread1.Col = xcol + 3
    vg_newnomrec = "": vg_newcodrec = Val(vaSpread1.Text)
    Dim Receta As New M_Receta
    Receta.Show 1, Me
    Me.Refresh
    Toolbar1.Refresh
    vg_newestrec = False
    If vg_newcodrec <> 0 And Trim(vg_newnomrec) <> "" And vaSpread1.BackColor <> Shape1(1).FillColor Then
        vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = vaSpread1.ActiveCol
        vaSpread1.Col = xcol + 3
        If vg_newcodrec = Val(vaSpread1.Text) Then Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = xcol
        ' Limpiar Datos y Formato Celda
        vaSpread1.Action = 3
        ' Retorna Modo de la columna
        vaSpread1.BlockMode = False
        vaSpread1.Font.Bold = False
        vaSpread1.Font.Size = 8
        vaSpread1.Text = vg_newnomrec
        
        vaSpread1.Col = xcol + 2
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 1
        vaSpread1.Text = Format(fg_CalCtoRecInv(Val(vg_newcodrec)), fg_Pict(6, 0))
        
        vaSpread1.Col = xcol + 3
        vaSpread1.Text = vg_newcodrec
        indgrabado = 1
        vg_newcodrec = 0: vg_newnomrec = ""
        
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    End If
Case 8
    M_CPlaTe.Show 1, Me
Case 10
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    J = 0
    For i = 1 To maxcolumna
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then J = vectorcol(i): Exit For
    Next i
    vaSpread1.Col = J: vaSpread1.Row = 0
    C_ApoPla.LlenarApoPlan M_MinRea, "Aporte Planificación Real " & vaSpread1.Text, vg_codcasino, vg_codregimen, vg_codservicio, Val(vg_fecha), 2, J
    C_ApoPla.Show 1, Me
Case 11
    If indgrabado = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    C_FrePla.LlenarFrecPlan "Frecuencia Planificación Real " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codcasino, vg_codregimen, vg_codservicio, Val(vg_fecha), 2
    C_FrePla.Show 1, Me
Case 20
    SwSalir = 0
    If Toolbar1.Buttons(2).Enabled = False Then indgrabado = 0
    If indgrabado <> 1 Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
    If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then indgrabado = 0
    If indgrabado = 1 Then GrabarPlantillaMinuta
    indgrabado = 0
    SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
End Select
End Sub

Private Sub Plato_Click(Index As Integer)
If Toolbar1.Buttons(2).Enabled = False Then indgrabado = 0: Exit Sub
Dim Del_Row As Integer, indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer
Dim Col As Long, fil As Long
Select Case Index
Case 2
    '------- Ingresar Recetas
    iblockcol = vaSpread1.ActiveCol: aiblockcol = vaSpread1.ActiveCol
    iblockcol2 = vaSpread1.ActiveCol: aiblockcol2 = vaSpread1.ActiveCol
    iblockrow = vaSpread1.ActiveRow: aiblockrow = vaSpread1.ActiveRow
    iblockrow2 = vaSpread1.ActiveRow: aiblockrow2 = vaSpread1.ActiveRow
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    J = 0
    For i = 1 To maxcolumna
        If vaSpread1.Col = vectorcol(i) Then J = vectorcol(i): Exit For
    Next i
    If J = 0 Then Exit Sub
    vg_codigo = "": vg_nombre = ""
    B_Receta.Show 1, Me
    If vg_codigo = "" Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow

    vaSpread1.Col = J - 1
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 2
    vaSpread1.Value = "R"
    vaSpread1.ForeColor = &HFF&
    vaSpread1.BackColor = &H80FF80
    
    vaSpread1.Col = J
    ' Limpiar Datos y Formato Celda
    vaSpread1.Action = 3
    ' Retorna Modo de la columna
    vaSpread1.BlockMode = False
    vaSpread1.Font.Bold = False
    vaSpread1.Font.Size = 8
    vaSpread1.Text = vg_nombre
              
    vaSpread1.Col = J + 1
    If Trim(vaSpread1.Text) = "" Then
       vaSpread1.CellType = 3
       vaSpread1.TypeIntegerMin = 1
       vaSpread1.TypeIntegerMax = 9999999
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.Text = 0
       vaSpread1.ForeColor = &HFF0000
    End If
    
    vaSpread1.Col = J + 2
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 1
    vaSpread1.Text = Format(fg_CalCtoRecInv(Val(vg_codigo)), fg_Pict(6, 2))
   
    vaSpread1.Col = J + 3
    If Trim(vaSpread1.Text) <> "" And Trim(vaSpread1.Text) <> Val(vg_codigo) Then vaSpread1.Col = (maxcolumna * 5 + 1) + ((J + 2) / 5): vaSpread1.Text = 1: If indcos = True Then vaSpread1.Col = J + 2: veccos((Int(J / 5) + 1), 1) = Round(veccos((Int(J / 5) + 1), 1) - vaSpread1.Text, vg_DCa)
    vaSpread1.Col = J + 3
    vaSpread1.Text = Val(vg_codigo)
    indgrabado = 1
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
Case 5
    '------- Insertar linea
    indcol = iblockcol
    
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then J = (vectorcol(i) - 1): Exit For
    Next i
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i: vaSpread1.Col = J
        If Trim(vaSpread1.Text) <> "" Then wsmaxfilas = vaSpread1.Row
    Next i
    If vaSpread1.MaxRows > 100 Then Del_Row = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
    wsmaxfilas = (wsmaxfilas + (iblockrow2 - iblockrow) + 1)
    If wsmaxfilas > vaSpread1.MaxRows Then Exit Sub
    
    iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
    
    vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, (100 - ((iblockrow2 - iblockrow) + 1)), iblockcol, iblockrow2 + 1
    vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False

    For i = 3 To (vaSpread1.MaxCols - maxcolumna) Step 5
        vaSpread1.Row = 0: vaSpread1.Col = i
        If CDate(Mid(Trim(vaSpread1.Text), 5, Len(Trim(vaSpread1.Text)))) < Format(Date - 1, "d/mm/yyyy") Then
            Dim f As Long, c As Long
                For c = i - 1 To i + 2
                    vaSpread1.Row = iblockrow: vaSpread1.Col = c
                    vaSpread1.BackColor = Shape1(1).FillColor
                Next c
        End If
    Next i
    '------- Validar días modificados
    For J = iblockrow To (100 - ((iblockrow2 - iblockrow) + 1))
        For i = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
            vaSpread1.Row = J
            vaSpread1.Col = i + 1
            If Trim(vaSpread1.Text) <> "" Then
               vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
               If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
            End If
        Next i
    Next J
    '------- Fin validar días modificados
    
    'For i = 3 To vaSpread1.MaxCols Step 5
    '    vaSpread1.Row = 0: vaSpread1.Col = i
    '    If InStr(1, Trim(vaSpread1.Text), "Día " & Format(Date, "d/mm/yyyy")) = 1 Then
    '        For Col = 0 To i - 4
    '            vaSpread1.Row = iblockrow: vaSpread1.Col = Col + 2
    '            vaSpread1.BackColor = Shape1(1).FillColor
    '        Next Col
    '    End If
    'Next i
    iblockcol = indcol
    indgrabado = 1
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
Case 6
    '------- Eliminar Linea
    indcol = iblockcol
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor And Trim(vaSpread1.Text) <> "" Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Or Trim(vaSpread1.Text) = "" Then GoTo Paso
    J = 0
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then J = (vectorcol(i) - 1): Exit For
    Next i
    If J = 0 Then Exit Sub
    If vaSpread1.MaxRows > 100 Then delrow = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
    If indactivo = 0 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.MaxCols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
    Next i
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
        If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
    Next i
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    If indcos = True Then
       For i = iblockcol To iblockcol2 Step 5
           For J = iblockrow To iblockrow2
               vaSpread1.Row = J: vaSpread1.Col = i + 1
               If Trim(vaSpread1.Text) <> "" Then vaSpread1.Col = i + 3: veccos((Int((i + 1) / 5) + 1), 1) = Round(veccos((Int((i + 1) / 5) + 1), 1) - vaSpread1.Text, vg_DCa): vaSpread1.Col = i + 2: veccos((Int((i + 1) / 5) + 1), 4) = Round(veccos((Int((i + 1) / 5) + 1), 4) - vaSpread1.Text, vg_DPr)
           Next J
       Next i
    End If
    vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False
    '------- Validar días modificados
    For J = iblockrow To (100 - ((iblockrow2 - iblockrow) + 1))
        For i = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
            vaSpread1.Row = J
            vaSpread1.Col = i + 1
            If Trim(vaSpread1.Text) <> "" Then
               vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
               If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
            End If
        Next i
    Next J
    '------- Fin validar días modificados
    iblockcol = auxcol
    vaSpread1.BlockMode = False
    indgrabado = 1
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indactivo = 0
Paso:
    vaSpread1.Row = vaSpread1.ActiveRow
    For i = 1 To vaSpread1.MaxCols
        vaSpread1.Col = i
        If Trim(vaSpread1.Text) <> "" Then MsgBox "Existe mas información en la linea, no puede eliminarla completamente", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    If vaSpread1.MaxRows > 100 Then Del_Row = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
    vaSpread1.Row = iblockrow2
    vaSpread1.Col = iblockcol
    vaSpread1.DeleteRows iblockrow, 1
    indgrabado = 1
    iblockcol = indcol
    For i = 3 To (vaSpread1.MaxCols - maxcolumna) Step 5
        vaSpread1.Row = 0: vaSpread1.Col = i
        If InStr(1, Trim(vaSpread1.Text), "Día " & Format(Date, "d/mm/yyyy")) = 1 Then
            For Col = 0 To i - 4
                vaSpread1.Row = vaSpread1.MaxRows: vaSpread1.Col = Col + 2
                vaSpread1.BackColor = Shape1(1).FillColor
            Next Col
        End If
    Next i
    '------- Validar días modificados
    For J = iblockrow To (100 - ((iblockrow2 - iblockrow) + 1))
        For i = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
            vaSpread1.Row = J
            vaSpread1.Col = i + 1
            If Trim(vaSpread1.Text) <> "" Then
               vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
               If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
            End If
        Next i
    Next J
    '------- Fin validar días modificados
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
'Case 8
'    '*** Subir Linea ***'
'    vaSpread1.Row = vaSpread1.ActiveRow
'    vaSpread1.Col = vaSpread1.ActiveCol
'    If vaSpread1.Row = 1 Then Exit Sub
'    If vaSpread1.Col = 1 Then Exit Sub
'    indcol = iblockcol
'    If iblockcol < 1 Then
'       For i = 1 To maxcolumna
'           vaSpread1.Col = vectorcol(i)
'           vaSpread1.Row = 1
'           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'       Next i
'    Else
'       For i = iblockcol To iblockcol2
'           vaSpread1.Col = i
'           For j = iblockrow To iblockrow2
'              vaSpread1.Row = j
'              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'           Next j
'       Next i
'    End If
'
'    vaSpread1.Col = 1
'    If Trim(vaSpread1.Text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'
'    If (iblockrow - ((iblockrow2 - iblockrow) + 1)) < 1 Then
'       MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
'       Exit Sub
'    End If
'    If vaSpread1.MaxRows > 100 Then Del_Row = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
'    If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
'    For i = 1 To maxcolumna
'        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
'    Next i
'    For i = 1 To maxcolumna
'        If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
'        If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
'    Next i
'
'    ' *** Copiar Datos Ultima Ultima Fila *** '
'    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'    vaSpread1.MoveRange iblockcol, (iblockrow - 1), iblockcol2, (iblockrow - 1), iblockcol, vaSpread1.MaxRows
'
'    ' *** Copiar Datos a la fila Seleccionada *** '
'    vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow - 1), False
'    vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow - 1)
'
'    ' ***  Devolver Datos a la fila y restar ultima fila *** '
'    vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
'    vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
'    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
'
'    vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
'    vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
'    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
'    indgrabado = 1
'    Toolbar1.Buttons(1).Visible = False
'    Toolbar1.Buttons(2).Visible = True
'    Toolbar1.Buttons(6).Visible = True
'    Toolbar1.Buttons(7).Visible = False
'Case 9
'    '*** Bajar Linea ***'
'    vaSpread1.Row = vaSpread1.ActiveRow
'    vaSpread1.Col = vaSpread1.ActiveCol
'    If vaSpread1.Row = 100 Then Exit Sub
'    If vaSpread1.Col = 1 Then Exit Sub
'    If iblockcol < 1 Then
'       For i = 1 To maxcolumna
'           vaSpread1.Col = vectorcol(i)
'           vaSpread1.Row = 1
'           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'       Next i
'    Else
'       For i = iblockcol To iblockcol2
'           vaSpread1.Col = i
'           For j = iblockrow To iblockrow2
'              vaSpread1.Row = j
'              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'           Next j
'       Next i
'    End If
'    vaSpread1.Col = 1
'    vaSpread1.Row = vaSpread1.ActiveRow + 1
'    If Trim(vaSpread1.Text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'    vaSpread1.Row = vaSpread1.ActiveRow - 1
'    If (iblockrow2 + ((iblockrow2 - iblockrow) + 1)) > 100 Then
'       MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
'       Exit Sub
'    End If
'    indcol = iblockcol
'    If vaSpread1.MaxRows > 100 Then
'       Del_Row = vaSpread1.MaxRows - 100
'       vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
'    End If
'    If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
'    For i = 1 To maxcolumna
'        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
'    Next i
'    For i = 1 To maxcolumna
'        If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
'        If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
'    Next i
'    ' ***      Copiar Datos Ultima Ultima Fila *** '
'    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
'    vaSpread1.MoveRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), iblockcol, vaSpread1.MaxRows
'
'    ' ***      Copiar Datos a la fila Seleccionada *** '
'    vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
'    vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
'
'    ' ***      Devolver Datos a la fila y restar ultima fila *** '
'    vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
'    vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
'    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
'
'    vaSpread1.Row = iblockrow + 1: vaSpread1.Col = iblockcol
'    vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
'    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
'    indgrabado = 1
'    Toolbar1.Buttons(1).Visible = False
'    Toolbar1.Buttons(2).Visible = True
'    Toolbar1.Buttons(6).Visible = True
'    Toolbar1.Buttons(7).Visible = False

Case 8
    '------- Subir linea
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = 1 Then Exit Sub
    If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.Text) <> "") Then
       For i = 1 To maxcolumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For J = iblockrow To iblockrow2
              vaSpread1.Row = J
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
           Next J
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col > 1 Then
        indcol = iblockcol
        vaSpread1.Col = 1
        If Trim(vaSpread1.Text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        
        If (iblockrow - ((iblockrow2 - iblockrow) + 1)) < 1 Then
           MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
           Exit Sub
        End If
        If vaSpread1.MaxRows > 100 Then Del_Row = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
        If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
        For i = 1 To maxcolumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To maxcolumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
        Next i
        '------- Validar días modificados
        For J = (iblockrow - 1) To vaSpread1.MaxRows
            For i = iblockcol To iblockcol2 Step 5
                vaSpread1.Row = J
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.Text) <> "" Then
                   vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
                   If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
                End If
            Next i
        Next J
        '------- Fin validar días modificados
        
        '------- Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.MoveRange iblockcol, (iblockrow - 1), iblockcol2, (iblockrow - 1), iblockcol, vaSpread1.MaxRows
        '------- Copiar datos fila seleccionada
        vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow - 1), False
        vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow - 1)
        '------- Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
        vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.Text) = "" Then Exit Sub
        For i = iblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.Text) <> "" Then Exit For
        Next i
        For z = iblockrow + 1 To 100 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.Text) <> "" Then Exit For
        Next z
        If z > 100 Then
            For fil = 100 To 1 Step -1
                For colu = 1 To vaSpread1.MaxCols
                    vaSpread1.Col = colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.Text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next colu
                If z <= 100 Then Exit For
            Next fil
        End If
        filaAct = iblockrow         'Fila actual
        filaAnt = IIf(i < 1, 1, i)  'Fila anterior
        filaPos = z                 'Fila posterior
        
        '------- Validar días modificados
        For J = filaAnt To vaSpread1.MaxRows '(filaAct - 1)
            For i = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
                vaSpread1.Row = J
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.Text) <> "" Then
                   vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
                   If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
                End If
            Next i
        Next J
        '------- Fin validar días modificados
        
        '------- Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (filaAct - filaAnt)
        vaSpread1.MoveRange 1, filaAnt, vaSpread1.MaxCols, (filaAct - 1), 1, 101
    
        '------- Mover estructura
        vaSpread1.MoveRange 1, filaAct, vaSpread1.MaxCols, (filaPos - 1), 1, filaAnt
        '------- Devolver respaldo
        vaSpread1.MoveRange 1, 101, vaSpread1.MaxCols, 101 + (filaAct - filaAnt - 1), 1, filaAnt + (filaPos - filaAct)
        vaSpread1.SetActiveCell 1, filaAnt
    End If
    vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
    indgrabado = 1
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    vaSpread1.MaxRows = 100
    vaSpread1.Col = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.BackColor = Shape1(2).FillColor
    Next i
Case 9
    '------- Bajar linea
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = 100 Then Exit Sub
'    If iblockcol < 1 Then
'       For i = 1 To maxcolumna
'           vaSpread1.Col = vectorcol(i)
'           vaSpread1.Row = 1
'           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'       Next i
'    Else
'       For i = iblockcol To iblockcol2
'           vaSpread1.Col = i
'           For J = iblockrow To iblockrow2
'              vaSpread1.Row = J
'              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
'           Next J
'       Next i
'    End If
    If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.Text) <> "") Then
       For i = 1 To maxcolumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For J = iblockrow To iblockrow2
              vaSpread1.Row = J
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
           Next J
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col > 1 Then
        vaSpread1.Col = 1
        vaSpread1.Row = vaSpread1.ActiveRow + 1
        If Trim(vaSpread1.Text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow - 1
        If (iblockrow2 + ((iblockrow2 - iblockrow) + 1)) > 100 Then
           MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
           Exit Sub
        End If
        indcol = iblockcol
        If vaSpread1.MaxRows > 100 Then
           Del_Row = vaSpread1.MaxRows - 100
           vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
        End If
        If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
        For i = 1 To maxcolumna
            If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
        Next i
        For i = 1 To maxcolumna
            If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
            If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
        Next i
        
        '------- Validar días modificados
        For J = iblockrow To vaSpread1.MaxRows
            For i = iblockcol To iblockcol2 Step 5
                vaSpread1.Row = J
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.Text) <> "" Then
                   vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
                   If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
                End If
            Next i
        Next J
        '------- Fin validar días modificados
        
        '------- Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.MoveRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), iblockcol, vaSpread1.MaxRows
    
        '------- Copiar datos fila Seleccionada
        vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
        vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
    
        '------- Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
        vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
            
        vaSpread1.Row = iblockrow + 1: vaSpread1.Col = iblockcol
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.Text) = "" Then Exit Sub
        For z = iblockrow + 1 To 100 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.Text) <> "" Then Exit For
        Next z
        If z > 100 Then Exit Sub
'            For fil = 100 To 1 Step -1
'                For colu = 1 To vaSpread1.MaxCols
'                    vaSpread1.Col = colu: vaSpread1.Row = fil
'                    If Trim(vaSpread1.Text) <> "" Then
'                        z = fil + 1: Exit For
'                    End If
'                Next colu
'                If z <= 100 Then Exit For
'            Next fil
'        End If
        vaSpread1.Col = vaSpread1.ActiveCol
        auxIblockrow = z
        For i = auxIblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.Text) <> "" Then Exit For
        Next i
        For z = auxIblockrow + 1 To 100 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.Text) <> "" Then Exit For
        Next z
        If z > 100 Then
            For fil = 100 To 1 Step -1
                For colu = 1 To vaSpread1.MaxCols
                    vaSpread1.Col = colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.Text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next colu
                If z <= 100 Then Exit For
            Next fil
        End If
        filaAct = auxIblockrow         'Fila actual
        filaAnt = IIf(i < 1, 1, i)  'Fila anterior
        filaPos = z                 'Fila posterior
        '------- Validar días modificados
        For J = filaAnt To vaSpread1.MaxRows '(filaAct - 1)
            For i = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
                vaSpread1.Row = J
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.Text) <> "" Then
                   vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
                   If Trim(vaSpread1.Text) = "" Then vaSpread1.Text = 2
                End If
            Next i
        Next J
        '------- Fin validar días modificados
        
        '------- Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (filaAct - filaAnt)
        vaSpread1.MoveRange 1, filaAnt, vaSpread1.MaxCols, (filaAct - 1), 1, 101
        
        '------- Mover estructura
        vaSpread1.MoveRange 1, filaAct, vaSpread1.MaxCols, (filaPos - 1), 1, filaAnt

        '------- Devolver respaldo
        vaSpread1.MoveRange 1, 101, vaSpread1.MaxCols, 101 + (filaAct - filaAnt - 1), 1, filaAnt + (filaPos - filaAct)
        vaSpread1.SetActiveCell 1, filaAnt + (filaPos - filaAct)
    End If
    iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
    indgrabado = 1
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    vaSpread1.MaxRows = 100
    vaSpread1.Col = 1
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.BackColor = Shape1(2).FillColor
    Next i

Case 11, 12
    '------- Copiar y pegar linea
    If Index = 11 Then
       If iblockcol < 1 Then
          For i = 1 To maxcolumna
              vaSpread1.Col = vectorcol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
          Next i
       Else
          For i = iblockcol To iblockcol2
              vaSpread1.Col = i
              For J = iblockrow To iblockrow2
                 vaSpread1.Row = J
                 If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
              Next J
          Next i
       End If
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    aiblockrow = iblockrow: aiblockrow2 = iblockrow2
    aiblockcol = iblockcol: aiblockcol2 = iblockcol2
    If vaSpread1.Col = 1 Then Exit Sub
    If vaSpread1.MaxRows > 100 Then Del_Row = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
    Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = True
    If iblockcol < 1 Then aiblockcol = 1: aiblockcol2 = vaSpread1.MaxCols
    indcortarpegar = 1
    If Index = 11 Then indcortarpegar = 0
Case 13
    '------- copiar y pegar
    If indcortarpegar = 0 Then
       If (iblockcol2 - iblockcol) > (aiblockcol2 - aiblockcol) Or (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then
          MsgBox "Imposible Pegar la infomación ya que el área de Cortar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
          Exit Sub
       End If
'      If IBlockCol2 > AIBlockCol2 Then
'         MsgBox "Imposible Cortar la infomación ya que el área de Cortar y el área de Pegado tienen formas distintas", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
'         Exit Sub
 '     End If
       indcortarpegar = 0
    Else
       If (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then
          MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
          Exit Sub
       End If
       If aiblockcol <> iblockcol2 And aiblockcol = 1 Then
          MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
          Exit Sub
       End If
    End If
    If iblockcol < 1 Then
       For i = 1 To maxcolumna
           vaSpread1.Col = vectorcol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
       Next i
    Else
       For i = iblockcol To iblockcol2
           vaSpread1.Col = i
           For J = iblockrow To iblockrow2
              vaSpread1.Row = J
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
           Next J
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    If indcortarpegar = 0 Then Toolbar1.Buttons(6).Visible = True: Toolbar1.Buttons(7).Visible = False
    If vaSpread1.MaxRows > 100 Then Del_Row = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row

    '------- destinacion de copiar y pegar datos
    If iblockcol < 1 Then iblockcol = 1: iblockcol2 = vaSpread1.MaxCols
    If aiblockcol2 = vaSpread1.MaxCols Then aiblockcol2 = vaSpread1.MaxCols - 1
    vaSpread1.Row = 0: vaSpread1.Col = iblockcol
    If vaSpread1.Text <> "N.Rac." Then
       For i = 1 To maxcolumna
           If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
       Next i
       For i = 1 To maxcolumna
           If (vectorcol(i) - 1) = aiblockcol Or vectorcol(i) = aiblockcol Or (vectorcol(i) + 1) = aiblockcol Or (vectorcol(i) + 2) = aiblockcol Then aiblockcol = (vectorcol(i) - 1): Exit For
       Next i
       For i = 1 To maxcolumna
           If (vectorcol(i) - 1) = iblockcol2 Or vectorcol(i) = iblockcol2 Or (vectorcol(i) + 1) = iblockcol2 Or (vectorcol(i) + 2) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 3)): Exit For
       Next i
       For i = 1 To maxcolumna
           If (vectorcol(i) - 1) = aiblockcol2 Or vectorcol(i) = aiblockcol2 Or (vectorcol(i) + 1) = aiblockcol2 Or (vectorcol(i) + 2) = aiblockcol2 Then aiblockcol2 = (vectorcol(i) + 3): Exit For
       Next i
    End If
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    For i = iblockcol To iblockcol2 Step 5
        If indcortarpegar = 1 Then
           vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
           If indcos = True Then
              veccos((Int((i + 1) / 5) + 1), 1) = 0: veccos((Int((i + 1) / 5) + 1), 4) = 0
              For J = aiblockrow To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
                  vaSpread1.Row = J: vaSpread1.Col = i + 1
                  If Trim(vaSpread1.Text) <> "" Then vaSpread1.Col = i + 3: veccos((Int((i + 1) / 5) + 1), 1) = Round(veccos((Int((i + 1) / 5) + 1), 1) + vaSpread1.Text, vg_DCa): vaSpread1.Col = i + 2: veccos((Int((i + 1) / 5) + 1), 4) = Round(veccos((Int((i + 1) / 5) + 1), 4) + vaSpread1.Text, vg_DPr)
              Next J
           End If
        ElseIf indcortarpegar = 0 Then
           vaSpread1.MoveRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
        End If
        '------- Validar días modificados
        For J = aiblockrow To vaSpread1.ActiveRow + (aiblockrow2 - aiblockrow)
            vaSpread1.Row = J
            vaSpread1.Col = i + 1
            If Trim(vaSpread1.Text) <> "" Then
               vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
               vaSpread1.Text = 1
            End If
        Next J
        '------- Fin validar días modificados
    Next i
    indgrabado = 1
    aiblockcol = indcol: iblockcol2 = indcol2
    aiblockrow = indrow: aiblockrow2 = indrow2
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
End Select
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
    Plato_Click (5)
Case 10
    Plato_Click (6)
Case 12
    Plato_Click (8)
Case 13
    Plato_Click (9)
Case 15
    Plantilla_Click (5)
Case 16
    Plantilla_Click (8)
Case 18
    Plantilla_Click (10)
Case 19
'    Plantilla_Click (3)
Case 20
    Plantilla_Click (11)
Case 22
    Plantilla_Click (20)
End Select
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
indactivo = 1
iblockrow = BlockRow
iblockrow2 = BlockRow2
iblockcol = BlockCol
iblockcol2 = BlockCol2
If BlockRow < 0 Then iblockrow = 1
If BlockRow2 < 0 Then iblockrow2 = 100
If BlockRow2 > 100 Then iblockrow2 = 100
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Then Exit Sub
OpGrilla(14).Enabled = IIf(Col = 1, True, False)
Plato(14).Enabled = IIf(Col = 1, True, False)
indactivo = 1
iblockrow = vaSpread1.ActiveRow
iblockrow2 = vaSpread1.ActiveRow
iblockcol = vaSpread1.ActiveCol
iblockcol2 = vaSpread1.ActiveCol
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row < 1 Or Col = 1 Then Exit Sub
Plato_Click (2)
End Sub
Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If Toolbar1.Buttons(2).Enabled = False Then indgrabado = 0: Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = Col
If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
If vaSpread1.ChangeMade = False Or Col = 1 Or Mode = 1 Then i = vaSpread1.Text: Exit Sub
If vaSpread1.ChangeMade = True Then vaSpread1.Col = (maxcolumna * 5 + 1) + (vaSpread1.Col / 5): vaSpread1.Text = 1: If indcos = True Then vaSpread1.Col = Col: J = Col - 1: veccos((Int(J / 5) + 1), 4) = Round(veccos((Int(J / 5) + 1), 4) - (i), vg_DPr): veccos((Int(J / 5) + 1), 4) = Round(veccos((Int(J / 5) + 1), 4) + (vaSpread1.Text), vg_DPr)
indgrabado = 1
Plantilla(0).Enabled = True
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
If Toolbar1.Buttons(2).Enabled = False Then indgrabado = 0: Exit Sub
Dim delrow As Integer, indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer
Select Case KeyCode
Case 86
    Exit Sub
Case 46
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    J = 0
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then J = (vectorcol(i) - 1): Exit For
    Next i
    If J = 0 Then Exit Sub
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    If vaSpread1.MaxRows > 100 Then delrow = vaSpread1.MaxRows - 100: vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
    If indactivo = 0 Then iblockcol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread1.MaxCols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Then iblockcol = (vectorcol(i) - 1): Exit For
    Next i
    For i = 1 To maxcolumna
        If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 4)): Exit For
        If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 3): Exit For
    Next i
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    If indcos = True Then
       For i = iblockcol To iblockcol2 Step 5
           For J = iblockrow To iblockrow2
               vaSpread1.Row = J: vaSpread1.Col = i + 1
               If Trim(vaSpread1.Text) <> "" Then vaSpread1.Col = i + 3: veccos((Int((i + 1) / 5) + 1), 1) = Round(veccos((Int((i + 1) / 5) + 1), 1) - vaSpread1.Text, vg_DCa)
           Next J
       Next i
    End If
    vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow2, False
    iblockcol = auxcol
    vaSpread1.BlockMode = False
    indgrabado = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indactivo = 0
End Select
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If indcos = False Then Exit Sub
Dim xcol As Long
Dim toapla As Double, toaesf As Double, toafoo As Double, totdia As Double, totesf As Double, nracre As Double, nracfo As Double
vaSpread1.Col = vaSpread1.ActiveCol
xcol = 0
For i = 1 To maxcolumna
    If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or vectorcol(i) = (vaSpread1.Col - 2)) Then xcol = vectorcol(i): Exit For
Next i
vaSpread1.Row = 0: vaSpread1.Col = xcol: Frame2(1).Caption = vaSpread1.Text: Frame2(2).Caption = vaSpread1.Text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.Text
toapla = 0: toaesf = 0: toafoo = 0: totdia = 0: totesf = 0: nracre = 0: nracfo = 0
For i = 1 To UBound(veccos)
    If i <= (Int(xcol / 5) + 1) Then toapla = CCur(toapla + veccos(i, 1)): toaesf = CCur(toaesf + veccos(i, 2)): toafoo = CCur(toafoo + veccos(i, 3)): nracre = CCur(nracre + veccos(i, 4)): nracfo = CCur(nracfo + veccos(i, 5))
    totdia = CCur(totdia + veccos(i, 1))
    totesf = CCur(totesf + veccos(i, 2))
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
Label1(8).Caption = Format(veccos((Int(xcol / 5) + 1), 1), fg_Pict(6, 2))
Label1(20).Caption = Format(veccos((Int(xcol / 5) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format(veccos((Int(xcol / 5) + 1), 2), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(veccos((Int(xcol / 5) + 1), 1) + veccos((Int(xcol / 5) + 1), 2)), fg_Pict(6, 2))
Label1(23).Caption = Format(veccos((Int(xcol / 5) + 1), 3), fg_Pict(6, 2))
Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
Label1(32).Caption = Format(toaesf, fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(toapla + toaesf), fg_Pict(6, 2))
Label1(34).Caption = Format(veccos((Int(xcol / 5) + 1), 4), fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((toapla + toaesf) / nracre), fg_Pict(6, 2))
Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2))
End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
Select Case Button
Case 2
    If vaSpread1.Visible <> True Then Exit Sub
    Indvaspread1 = 0
    PopupMenu MenuDetalle
    
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
End Select
End Sub

Private Sub GrabarPlantillaMinuta()
Dim desc As String
Dim codrec As Long, numrac As Long, estser As Long, fecha As Long, conregdet As Long, indice As Long, existedat As Long, inddia As Long
Dim fechasis As Long, fecini As Long, fecfin As Long
Dim cosrec As Double, cospro As Double
On Error GoTo Man_Error
inddia = 1: conregdet = 0: gauge1.Value = 0: gauge.Value = 0: fecha = 0: fecini = 0: fecfin = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh
fechasis = 0: fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
fg_carga (ss)

'------- Grabar Estructura Servicio
vg_db.BeginTrans
  For i = 2 To (vaSpread1.MaxCols - maxcolumna - 1) Step 5
        gauge1.Value = Val((inddia / maxcolumna) * 100)
        Label3.Caption = ""
        Label3.Caption = "Día : " & inddia
        existedat = 0
        vaSpread1.Row = 1: vaSpread1.Col = i
      
        For J = 1 To 100
            vaSpread1.Row = J
            If inddia < 10 Then
               fecha = Val(vg_fecha) & "0" & inddia
            Else
               fecha = Val(vg_fecha) & inddia
            End If
            vaSpread1.Col = i + 1
            If Trim(vaSpread1.Text) <> "" Then existedat = 1: Exit For
        Next J
        indice = 0
        RS.Open "select b_minuta.min_codigo from  b_minuta, b_minutadet " & _
                "where  b_minuta.min_codigo=b_minutadet.mid_codigo " & _
                "and    b_minuta.min_cencos='" & vg_codcasino & "' " & _
                "and    b_minuta.min_codreg=" & vg_codregimen & " " & _
                "and    b_minuta.min_codser=" & vg_codservicio & " " & _
                "and    b_minuta.min_fecmin=" & Val(fecha) & " " & _
                "and    b_minutadet.mid_tipmin='2'", vg_db, adOpenStatic
        If Not RS.EOF Then
           indice = RS!min_codigo
           RS.Close: Set RS = Nothing
           If indice > 0 And existedat = 0 Then vg_db.Execute "delete b_minutadet from b_minutadet where mid_codigo=" & indice & " and mid_tipmin='2'"
        Else
           RS.Close: Set RS = Nothing
           If existedat > 0 Then
                RS.Open "select min_codigo from b_minuta order by min_codigo desc", vg_db, adOpenStatic
                If Not RS.EOF Then
                   RS.MoveFirst
                   indice = RS!min_codigo + 1
                Else
                   indice = 1
                End If
                RS.Close: Set RS = Nothing
                vg_db.Execute "insert into b_minuta (min_codigo, min_cencos, min_codreg, " & _
                              "min_codser, min_fecmin, min_indblo) values (" & indice & ", " & _
                              "'" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & _
                              "" & Val(fecha) & ", 0)"
           End If
        End If
        gauge.Value = 0: conregdet = 0: estser = 0
        If existedat > 0 Then
           '------- Actualizar detalle minutas
           For J = 1 To 100
              conregdet = conregdet + 1
              gauge.Value = Val((conregdet / 100) * 100)
              desc = "": codrec = 0: numrec = 0: cosrec = 0
              vaSpread1.Row = J
              vaSpread1.Col = vaSpread1.MaxCols
              If Trim(vaSpread1.Text) <> "" Then estser = vaSpread1.Text
              vaSpread1.Col = i + 1: desc = Trim(vaSpread1.Text)
              
              If desc <> "" Then
                 vaSpread1.Col = i + 2: numrac = vaSpread1.Text
                 vaSpread1.Col = i + 3: cosrec = vaSpread1.Text
                 vaSpread1.Col = i + 4: codrec = vaSpread1.Text
                 RS.Open "select * from b_minutadet where mid_codigo=" & indice & " and mid_numlin=" & J & " and mid_tipmin='2' ", vg_db, adOpenStatic
                 If Not RS.EOF Then
                    RS.Close: Set RS = Nothing
                    vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
                    If Trim(vaSpread1.Text) <> "" And (Trim(vaSpread1.Text) = 1 Or Trim(vaSpread1.Text) = 2) Then
                       vg_db.Execute "update b_minutadet " & _
                                     "set    mid_codrec=" & codrec & " " & _
                                     "where  mid_codigo=" & indice & " " & _
                                     "and    mid_numlin=" & J & " " & _
                                     "and    mid_tipmin='2' " & _
                                     "and    mid_codrec<>" & codrec & ""
               
                       vg_db.Execute "update b_minutadet inner join b_minuta " & _
                                     "on     b_minutadet.mid_codigo=b_minuta.min_codigo " & _
                                     "set    mid_numrac=" & numrac & ", " & _
                                     "       mid_descri='" & desc & "', " & _
                                     "       mid_estser=" & estser & ", " & _
                                     "       mid_cosrec=" & cosrec & " " & _
                                     "where  mid_codigo=" & indice & " " & _
                                     "and    mid_numlin=" & J & " " & _
                                     "and    mid_tipmin='2' " & _
                                     "and   (mid_numrac<>" & numrac & " " & _
                                     "or     mid_descri<>'" & desc & "' " & _
                                     "or     mid_estser<>" & estser & " " & _
                                     "or     mid_cosrec<>" & cosrec & ")"
                    End If
'                    vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
                    If Trim(vaSpread1.Text) <> "" And Trim(vaSpread1.Text) = 1 Then
                       If fecini < fecha And fecini = 0 Then fecini = fecha
                    End If
                 Else
                    RS.Close: Set RS = Nothing
                    vg_db.Execute "insert into b_minutadet (mid_codigo, mid_tipmin, " & _
                                  "mid_numlin, mid_estser, mid_codrec, mid_numrac, mid_descri, mid_cosrec) " & _
                                  "values (" & indice & ", '2', " & J & ", " & estser & ", " & codrec & ", " & numrac & ", '" & Trim(desc) & "', " & cosrec & ")"
                 End If
              Else
                 vg_db.Execute "Delete b_minutadet from b_minutadet " & _
                               "where mid_codigo=" & indice & " " & _
                               "and   mid_numlin=" & J & " " & _
                               "and   mid_tipmin='2'"
              End If
           Next J
           For J = 1 To 100
               vaSpread1.Row = J
               vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
               If Trim(vaSpread1.Text) <> "" And Trim(vaSpread1.Text) = 1 Then
                  '------- Borrar minuta cambios
                  vg_db.Execute "delete b_minutacambios from b_minutacambios where cam_fecmin=" & Val(fecha) & ""
                  
                  RS.Open "select b_ingrediente.ing_codigo, " & _
                          "sum(b_minutadet.mid_numrac*(b_recetadet.red_canpro/b_receta.rec_basrac)) as cantidad " & _
                          "from   b_receta, b_recetadet, b_minutadet, b_ingrediente " & _
                          "where  b_minutadet.mid_codrec=b_recetadet.red_codigo " & _
                          "and    b_minutadet.mid_codrec=b_receta.rec_codigo " & _
                          "and    b_recetadet.red_codpro=b_ingrediente.ing_codigo " & _
                          "and    b_minutadet.mid_codigo=" & indice & " " & _
                          "and    b_minutadet.mid_tipmin='2' " & _
                          "group by b_ingrediente.ing_codigo", vg_db, adOpenStatic
                  If Not RS.EOF Then
                     Do While Not RS.EOF
                        vg_db.Execute "insert into b_minutacambios (cam_feccam, cam_codmin, cam_codpro, cam_canpro, cam_fecmin, cam_fecped) " & _
                                      "values (" & fechasis & ", " & indice & ", '" & RS!ing_codigo & "', " & RS!cantidad & ", " & Val(fecha) & ", 0)"
                         RS.MoveNext
                     Loop
                  End If
                  RS.Close: Set RS = Nothing
                  Exit For
               End If
           Next J
        End If
      inddia = inddia + 1
  Next i
  fecfin = fecha
  If fecini > 0 Then
     
     '------- Buscar por rango de fecha los producto incluido en mes y luego eliminar y grabar minuta costo
     RS.Open "select distinct b_ingrediente.ing_codigo, b_ingrediente.ing_precos " & _
             "From  b_receta, b_recetadet, b_minuta, b_minutadet, b_ingrediente " & _
             "Where b_minuta.min_codigo = b_minutadet.mid_codigo " & _
             "and   b_minutadet.mid_codrec=b_recetadet.red_codigo " & _
             "and   b_recetadet.red_codigo=b_receta.rec_codigo " & _
             "and   b_recetadet.red_codpro=b_ingrediente.ing_codigo " & _
             "and   b_minuta.min_cencos='" & vg_codcasino & "' " & _
             "and   b_minuta.min_fecmin>=" & fecini & " " & _
             "and   b_minuta.min_fecmin<=" & fecfin & " " & _
             "and   b_minutadet.mid_tipmin='2'", vg_db, adOpenStatic
     If Not RS.EOF Then
        Do While Not RS.EOF
           vg_db.Execute "delete b_minutacosto from b_minutacosto " & _
                         "where mic_fecval=" & fechasis & " " & _
                         "and   mic_tipmin='2' " & _
                         "and   mic_codpro='" & RS!ing_codigo & "'"
           vg_db.Execute "insert into b_minutacosto(mic_fecval, mic_tipmin, mic_codpro, mic_cospro) " & _
                         "values (" & fechasis & ", '2', '" & RS!ing_codigo & "', " & RS!ing_precos & ")"
           RS.MoveNext
        Loop
     End If
     RS.Close: Set RS = Nothing
     
     '------- Actualizar costo teórica a partir de fecha modificación
     RS.Open "select b_minutadet.* " & _
             "from   b_minuta, b_minutadet " & _
             "where  b_minuta.min_codigo=b_minutadet.mid_codigo " & _
             "and    b_minuta.min_cencos='" & vg_codcasino & "' " & _
             "and    b_minuta.min_fecmin>=" & fecini & " " & _
             "and    b_minuta.min_fecmin<=" & fecfin & " " & _
             "and    b_minutadet.mid_tipmin='2'", vg_db, adOpenStatic
     If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
     Do While Not RS.EOF
        cosrec = 0: cosrec = Format(fg_CalCtoRecPlan(fechasis, 2, RS!mid_codrec), fg_Pict(6, 2))
        vg_db.Execute "update b_minutadet " & _
                      "set    mid_fecval=" & fechasis & ", " & _
                      "       mid_cosrec=" & cosrec & " " & _
                      "where  mid_codigo=" & RS!mid_codigo & " " & _
                      "and    mid_tipmin='2' " & _
                      "and    mid_numlin=" & RS!mid_numlin & " " & _
                      "and    mid_codrec=" & RS!mid_codrec & ""
        RS.MoveNext
     Loop
     RS.Close: Set RS = Nothing
  End If
vg_db.CommitTrans
Picture1.Visible = False: gauge.Visible = False
vaSpread1.Refresh
fg_descarga

Exit Sub
Man_Error:
If Err = -2147467259 Then
    MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error"
    vg_db.RollbackTrans
    Exit Sub
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub DetallePlantillaMinuta()
fg_carga ""

wsmaxfilas = 0: SwSalir = 0: maxcolumna = 0
iblockrow = 0: iblockrow2 = 0: iblockcol = 0: iblockcol2 = 0: SwSalir = 0
aiblockrow = 0: aiblockrow2 = 0: aiblockcol = 0: aiblockcol2 = 0
indactivo = 0: indgrabado = 0

'------- formatear columna
maxcolumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))
vaSpread1.MaxRows = 100
vaSpread1.MaxCols = 0: vaSpread1.MaxCols = 5 * maxcolumna + 1: vaSpread1.Row = 0
vaSpread1.Col = 1
vaSpread1.ColsFrozen = 1
vaSpread1.VisibleCols = 1
vaSpread1.ColWidth(1) = 15
vaSpread1.Text = "Estructura Servicio"
ReDim Preserve vectorcol(0)
For i = 2 To vaSpread1.MaxCols Step 5
    
    vaSpread1.Col = i
    vaSpread1.ColWidth(i) = 1.5
    vaSpread1.Text = " "
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 1
    vaSpread1.ColWidth(i + 1) = 21
    If i = 2 Then
       ReDim Preserve vectorcol(1)
       vectorcol(1) = 3
       vaSpread1.Text = " Día " & (i - 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
    Else
       vaSpread1.Text = " Día " & CLng((i / 5) + 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
       ReDim Preserve vectorcol(CLng((i / 5) + 1))
       vectorcol(CLng((i / 5) + 1)) = i + 1
    End If
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 2
    vaSpread1.ColWidth(i + 2) = 6
    vaSpread1.Text = "Rac."
    vaSpread1.ColHidden = False
   
    vaSpread1.Col = i + 3
    vaSpread1.ColWidth(i + 3) = 9
    vaSpread1.Text = "Costo"
    vaSpread1.ColHidden = False
    
    vaSpread1.Col = i + 4
    vaSpread1.Text = "Cod. Receta"
    vaSpread1.ColHidden = True
    
    For J = 1 To vaSpread1.MaxRows
        vaSpread1.Row = J

        vaSpread1.Col = i
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 0
        vaSpread1.Text = ""

        vaSpread1.Col = i + 1
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 0
        vaSpread1.Text = " "

        vaSpread1.Col = i + 2
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 0
        vaSpread1.Text = " "

        vaSpread1.Col = i + 3
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 0
        vaSpread1.Text = " "

        vaSpread1.Col = i + 4
        vaSpread1.CellType = 5
        vaSpread1.TypeHAlign = 0
        vaSpread1.Text = " "

        vaSpread1.Col = i + 5
        vaSpread1.CellType = 1
        vaSpread1.TypeHAlign = 0
        vaSpread1.Text = " "

    Next J
    vaSpread1.Row = 0
Next i

vaSpread1.Row = 0
For i = 1 To maxcolumna
   vaSpread1.MaxCols = vaSpread1.MaxCols + 1
   vaSpread1.Col = vaSpread1.MaxCols
   vaSpread1.Text = "Estado"
   vaSpread1.ColHidden = True
Next i
vaSpread1.MaxCols = vaSpread1.MaxCols + 1
vaSpread1.Col = vaSpread1.MaxCols
vaSpread1.ColWidth(vaSpread1.MaxCols) = 5
vaSpread1.Text = "Còd. Est."
vaSpread1.ColHidden = True

vaSpread1.Row = -1: vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
vaSpread1.Row = -1: vaSpread1.Col = 1
vaSpread1.Font.Bold = True
vaSpread1.Font.Size = 9
vaSpread1.BackColor = Shape1(2).FillColor 'Verde


J = 0: i = 0
RS.Open "select b_minutadet.mid_tipmin, b_minutadet.mid_numlin, b_minutadet.mid_codrec, " & _
        "b_minutadet.mid_descri, b_minutadet.mid_cosrec, b_minuta.min_fecmin, b_minuta.min_indblo, " & _
        "b_receta.rec_nombre, b_minutadet.mid_numrac, b_minutadet.mid_estser, a_estservicio.ess_nombre " & _
        "from  b_receta, b_minuta, b_minutadet, a_estservicio " & _
        "where b_minuta.min_codigo=b_minutadet.mid_codigo " & _
        "and   b_minutadet.mid_codrec=b_receta.rec_codigo " & _
        "and   b_minutadet.mid_estser=a_estservicio.ess_codigo " & _
        "and   b_minuta.min_cencos='" & vg_codcasino & "' " & _
        "and   b_minuta.min_codreg=" & vg_codregimen & " " & _
        "and   b_minuta.min_codser=" & vg_codservicio & " " & _
        "and   val(mid(b_minuta.min_fecmin,1,6))=" & Val(vg_fecha) & " " & _
        "and   b_minutadet.mid_tipmin='2' " & _
        "order by b_minutadet.mid_estser, b_minuta.min_fecmin, b_minutadet.mid_numlin", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      J = (((Val(Mid(RS!min_fecmin, 7, 2)) * 5) - 5) + 1) + 1
      vaSpread1.Row = RS!mid_numlin
      If RS!mid_estser <> i Then
         vaSpread1.Col = 1
         vaSpread1.Text = RS!ess_nombre
         
         vaSpread1.Col = vaSpread1.MaxCols
         vaSpread1.CellType = 5
         vaSpread1.TypeHAlign = 2
         vaSpread1.Text = RS!mid_estser
         i = RS!mid_estser
      End If
      vaSpread1.Col = J
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 2
      vaSpread1.Value = "R"
      vaSpread1.ForeColor = &HFF&
      vaSpread1.BackColor = &H80FF80
           
      vaSpread1.Col = J + 1
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 0
      vaSpread1.Text = Trim(RS!rec_nombre)
                         
      vaSpread1.Col = J + 2
      vaSpread1.CellType = 3
      vaSpread1.TypeIntegerMin = 1
      vaSpread1.TypeIntegerMax = 9999999
      vaSpread1.TypeHAlign = 1
      vaSpread1.TypeSpin = False
      vaSpread1.TypeIntegerSpinInc = 1
      vaSpread1.TypeIntegerSpinWrap = False
      vaSpread1.Value = RS!mid_numrac
      vaSpread1.ForeColor = &HFF0000
                       
      vaSpread1.Col = J + 3
      vaSpread1.CellType = 5
      vaSpread1.TypeHAlign = 1
      vaSpread1.Text = Format(RS!mid_cosrec, fg_Pict(6, 2))
      
      vaSpread1.Col = J + 4: vaSpread1.Text = RS!mid_codrec
      'If RS!min_indblo > 0 Then vaSpread1.Row = -1: vaSpread1.Col = j: vaSpread1.BackColor = Shape1(1).FillColor: vaSpread1.Col = j + 1: vaSpread1.BackColor = Shape1(1).FillColor: vaSpread1.Col = j + 2: vaSpread1.CellType = 5: vaSpread1.TypeHAlign = 1: vaSpread1.BackColor = Shape1(1).FillColor: vaSpread1.Col = j + 3: vaSpread1.BackColor = Shape1(1).FillColor
     
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing: fg_descarga
Else
   RS.Close: Set RS = Nothing: fg_descarga
   RS.Open "select a_estservicio.* from a_estservicio where ess_codser=" & vg_codservicio & " order by ess_orden", vg_db, adOpenStatic
   If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
   Do While Not RS.EOF
      vaSpread1.Row = RS!ess_orden
      vaSpread1.Col = 1
      vaSpread1.Text = RS!ess_nombre
      For i = 2 To vaSpread1.MaxCols Step 5
          vaSpread1.Col = vaSpread1.MaxCols
          vaSpread1.Text = RS!ess_codigo
      Next i
      RS.MoveNext
   Loop
   RS.Close: Set RS = Nothing
End If

For i = 3 To (vaSpread1.MaxCols - maxcolumna) Step 5
    vaSpread1.Row = 0: vaSpread1.Col = i
    If CDate(Mid(Trim(vaSpread1.Text), 5, Len(Trim(vaSpread1.Text)))) < Format(Date - 1, "d/mm/yyyy") Then
        Dim fil As Long, Col As Long
        For fil = 1 To vaSpread1.MaxRows
            For Col = i - 1 To i + 2
                vaSpread1.Row = fil: vaSpread1.Col = Col
                vaSpread1.BackColor = Shape1(1).FillColor
            Next Col
        Next fil
    End If
Next i
vaSpread1.Row = 1: vaSpread1.Col = 1
iblockrow = vaSpread1.Row: aiblockrow = vaSpread1.Row
iblockrow2 = vaSpread1.Row: aiblockrow2 = vaSpread1.Row
iblockcol = vaSpread1.Col: aiblockcol = vaSpread1.Col
iblockcol2 = vaSpread1.Col: aiblockcol2 = vaSpread1.Col
End Sub
