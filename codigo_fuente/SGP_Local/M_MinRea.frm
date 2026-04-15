VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_MinRea 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Real"
   ClientHeight    =   8310
   ClientLeft      =   2775
   ClientTop       =   2040
   ClientWidth     =   11685
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11685
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   2625
      Index           =   0
      Left            =   0
      TabIndex        =   13
      Top             =   5490
      Visible         =   0   'False
      Width           =   15195
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
         TabIndex        =   52
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
            Index           =   38
            Left            =   2370
            TabIndex        =   67
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
            Index           =   37
            Left            =   2370
            TabIndex        =   66
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
            Index           =   36
            Left            =   2370
            TabIndex        =   65
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
            Index           =   35
            Left            =   960
            TabIndex        =   64
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
            Index           =   34
            Left            =   960
            TabIndex        =   63
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
            Index           =   33
            Left            =   960
            TabIndex        =   62
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
            Index           =   32
            Left            =   960
            TabIndex        =   61
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
            Index           =   31
            Left            =   960
            TabIndex        =   60
            Top             =   465
            Width           =   1320
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
            Index           =   30
            Left            =   2790
            TabIndex        =   59
            Top             =   210
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Planificado"
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
            TabIndex        =   58
            Top             =   210
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Band."
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
            TabIndex        =   57
            Top             =   1650
            Width           =   855
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
            TabIndex        =   56
            Top             =   1350
            Width           =   420
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
            TabIndex        =   55
            Top             =   1050
            Width           =   795
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
            TabIndex        =   54
            Top             =   750
            Width           =   645
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
            Index           =   24
            Left            =   90
            TabIndex        =   53
            Top             =   465
            Width           =   855
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
         TabIndex        =   36
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
            Index           =   23
            Left            =   2370
            TabIndex        =   51
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
            Index           =   22
            Left            =   960
            TabIndex        =   50
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
            Index           =   21
            Left            =   960
            TabIndex        =   49
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
            Index           =   20
            Left            =   960
            TabIndex        =   48
            Top             =   525
            Width           =   1320
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
            Index           =   19
            Left            =   2790
            TabIndex        =   47
            Top             =   240
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Planificado"
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
            TabIndex        =   46
            Top             =   240
            Width           =   960
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
            TabIndex        =   45
            Top             =   1095
            Width           =   795
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
            TabIndex        =   44
            Top             =   810
            Width           =   645
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
            TabIndex        =   43
            Top             =   525
            Width           =   855
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
            TabIndex        =   42
            Top             =   1380
            Width           =   420
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Band."
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
            TabIndex        =   41
            Top             =   1680
            Width           =   855
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
            TabIndex        =   40
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            Index           =   47
            Left            =   2370
            TabIndex        =   37
            Top             =   1680
            Width           =   1320
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
         TabIndex        =   16
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
            Index           =   14
            Left            =   2370
            TabIndex        =   35
            Top             =   1830
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
            Index           =   13
            Left            =   2370
            TabIndex        =   34
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
            Index           =   12
            Left            =   2370
            TabIndex        =   33
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
            Index           =   11
            Left            =   2370
            TabIndex        =   32
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
            Index           =   10
            Left            =   1020
            TabIndex        =   31
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
            Index           =   9
            Left            =   1020
            TabIndex        =   30
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
            Index           =   8
            Left            =   1020
            TabIndex        =   29
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
            Index           =   7
            Left            =   2370
            TabIndex        =   28
            Top             =   525
            Width           =   1320
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
            TabIndex        =   27
            Top             =   1830
            Visible         =   0   'False
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
            TabIndex        =   26
            Top             =   1770
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            TabIndex        =   25
            Top             =   1095
            Width           =   450
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
            TabIndex        =   24
            Top             =   810
            Width           =   645
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
            TabIndex        =   23
            Top             =   525
            Width           =   855
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
            Index           =   0
            Left            =   2520
            TabIndex        =   22
            Top             =   240
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto. Bandeja"
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
            TabIndex        =   21
            Top             =   240
            Width           =   1110
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
            TabIndex        =   20
            Top             =   525
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
            Index           =   41
            Left            =   1020
            TabIndex        =   19
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
            Index           =   48
            Left            =   2370
            TabIndex        =   18
            Top             =   1380
            Width           =   1320
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
            Index           =   49
            Left            =   120
            TabIndex        =   17
            Top             =   1380
            Width           =   420
         End
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   15
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
            Picture         =   "M_MinRea.frx":0000
            Top             =   150
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   3360
         Picture         =   "M_MinRea.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Costo Bandeja Planificado"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   2760
         Picture         =   "M_MinRea.frx":0614
         Stretch         =   -1  'True
         ToolTipText     =   "Costo Totales"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
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
      Left            =   3900
      ScaleHeight     =   1035
      ScaleWidth      =   6675
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   6735
      Begin MSComctlLib.ProgressBar gauge 
         Height          =   210
         Left            =   120
         TabIndex        =   9
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
         TabIndex        =   10
         Top             =   240
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   370
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Día"
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
         Left            =   120
         TabIndex        =   12
         Top             =   0
         Width           =   330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Detalle"
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
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   45
      TabIndex        =   0
      Top             =   840
      Width           =   10905
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
         TabIndex        =   4
         Top             =   150
         Width           =   1215
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
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         Height          =   195
         Index           =   0
         Left            =   9405
         TabIndex        =   3
         Top             =   135
         Width           =   1155
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
         Caption         =   "Celda Bloqueada"
         Height          =   195
         Index           =   1
         Left            =   7650
         TabIndex        =   2
         Top             =   135
         Width           =   1215
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
         Caption         =   "Estructura de Servicio"
         Height          =   195
         Index           =   2
         Left            =   5505
         TabIndex        =   1
         Top             =   135
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   5
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
      DragIcon        =   "M_MinRea.frx":091E
      Height          =   4140
      Left            =   0
      TabIndex        =   6
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
      SpreadDesigner  =   "M_MinRea.frx":0D60
      UserResize      =   1
      VisibleCols     =   1
      VisibleRows     =   100
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
      Left            =   15
      TabIndex        =   7
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
         Caption         =   "Costo Receta"
         Index           =   11
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Frecuencia Recetas"
         Index           =   12
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Actualizar Costo Recetas"
         Index           =   13
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Exportar Recetas"
         Index           =   14
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   18
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
         Caption         =   "&Insertar Línea"
         Index           =   5
      End
      Begin VB.Menu Plato 
         Caption         =   "&Eliminar Línea"
         Index           =   6
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Plato 
         Caption         =   "&Subir Línea"
         Index           =   8
      End
      Begin VB.Menu Plato 
         Caption         =   "&Bajar Línea"
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
         Caption         =   "&Buscar Receta"
         Index           =   15
      End
      Begin VB.Menu Plato 
         Caption         =   "&Agregar Estructura"
         Index           =   16
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
         Caption         =   "&Insertar Línea"
         Index           =   5
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Eliminar Línea"
         Index           =   6
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Subir Línea"
         Index           =   8
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Bajar Línea"
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
         Caption         =   "Pegado &Especial"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Buscar Receta"
         Index           =   15
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Agrega Estructura"
         Index           =   16
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
Option Explicit
Option Compare Text

Private RS              As New ADODB.Recordset
Private RS1             As New ADODB.Recordset
Private i               As Long
Private j               As Long
Private IndCortarPegar  As Long
Private Fecha           As Long
Private MaxColumna      As Long
Private MaxFila         As Long
Private AddReceta       As Long
Private IblockRow       As Integer
Private IblockRow2      As Integer
Private IblockCol       As Integer
Private iblockcol2      As Integer
Private SwSalir         As Integer
Private AiBlockRow      As Integer
Private AiBlockRow2     As Integer
Private AiBlockCol      As Integer
Private AiBlockCol2     As Integer
Private indactivo       As Integer
Private IndCos          As Boolean
Private etapa5          As Boolean
Private indgri          As Boolean
Private vCtoPis         As Double
Private vCtoTec         As Double
Private VecCos()        As Variant
Private VecCosenc()     As Variant
Private VectorCol()     As Long
Private MsgTitulo       As String

Private BtnX            As Variant
Private Cancel          As Boolean
Private ExisteDat       As Long
Private ExisteDatMinuta As Boolean
Private AuxCol          As Long
Private X               As Long
Private NumRac          As Long
Private ToaPla          As Double
Private ToaEsf          As Double
Private ToaFoo          As Double
Dim xColIni As Variant, xRowIni As Variant, xColFin As Variant, xRowFin As Variant
Dim TipoCopia As String
Dim numracionesanterior As Long
Dim EstMinBlo           As String

Private Sub Estructura1_Click(Index As Integer)
    LlenaSubMenu Estructura1, Index
End Sub

Private Sub LlenaSubMenu(SubMenu As Object, Index As Integer)
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    If Trim(vaSpread1.text) <> "" Then MsgBox "Seleccione una celda de estructura que este vacía...", vbInformation, MsgTitulo: Exit Sub
    vaSpread1.text = SubMenu(Index).Caption
    vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.text = SubMenu(Index).HelpContextID
    Estructura1(Index).Enabled = False: Estructura2(Index).Enabled = False
    
    For j = vaSpread1.ActiveRow To (vaSpread1.MaxRows - 1)
        
        For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
            
            vaSpread1.Row = j
            vaSpread1.Col = i + 1
        
        Next i
        
        vaSpread1.Col = 1
        vaSpread1.Row = j + 1
        If Trim(vaSpread1.text) <> "" Then Exit For
    
    Next j
    
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True

End Sub

Private Sub Estructura2_Click(Index As Integer)

If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
   MsgBox "Minuta, sin acceso a modificaciones", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
End If
LlenaSubMenu Estructura2, Index

End Sub

Private Sub Form_Activate()
    
    fg_descarga
    TraerFechaCierre

End Sub

Private Sub Form_Load()

Dim X       As Long
Dim nomser  As String
Dim nomreg  As String
Dim RS      As New ADODB.Recordset

    Me.HelpContextID = vg_OpcM
    Me.Height = 6765
    Me.Width = 11055
    fg_centra Me
    MsgTitulo = "Planificación Real"
    fg_carga ""
    
    '-------> Traer nombre regimen
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Regimen(2, vg_codregimen, ""), vg_db, adOpenStatic
    If Not RS.EOF Then nomreg = Trim(RS!reg_nombre)
    RS.Close: Set RS = Nothing
    '-------> Traer nombre servicio
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Servicio(8, vg_codservicio, ""), vg_db, adOpenStatic
    If Not RS.EOF Then nomser = Trim(RS!ser_nombre)
    RS.Close: Set RS = Nothing
    
    Label4.Caption = M_Plami1.fpayuda(0).Caption & "(" & M_Plami1.fpText.text & ")" & " - " & nomreg & " - " & nomser
    Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "
    IndCos = False: etapa5 = False: indgri = False
    
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = " "
    Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = "Grabar Datos": BtnX.Enabled = IIf(Mid(ValidarUsuario(M_Plami1), 2, 2) = "0", False, True)
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Cortar", , tbrDefault, "A_Cortar"): BtnX.Visible = True: BtnX.ToolTipText = "Cortar"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Copiar", , tbrDefault, "A_Copiar"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar"
    Set BtnX = Toolbar1.Buttons.Add(, "I_Pegar", , tbrDefault, "I_Pegar"): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Toolbar1.Buttons.Add(, "A_Pegar", , tbrDefault, "A_Pegar"): BtnX.Visible = False: BtnX.ToolTipText = "Pegar"
    Set BtnX = Toolbar1.Buttons.Add(, "I_PegadoEspecial", , tbrDefault, "I_PegadoEspecial"): BtnX.Visible = True: BtnX.ToolTipText = ""
    Set BtnX = Toolbar1.Buttons.Add(, "A_PegadoEspecial", , tbrDefault, "A_PegadoEspecial"): BtnX.Visible = False: BtnX.ToolTipText = "Pegado Especial"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Buscar", , tbrDefault, "A_Buscar"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar Receta"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_InsertarF", , tbrDefault, "A_InsertarF"): BtnX.Visible = True: BtnX.ToolTipText = "Insertar"
    Set BtnX = Toolbar1.Buttons.Add(, "A_EliminarF", , tbrDefault, "A_EliminarF"): BtnX.Visible = True: BtnX.ToolTipText = "Eliminar"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_SubirF", , tbrDefault, "A_SubirF"): BtnX.Visible = True: BtnX.ToolTipText = "Subir"
    Set BtnX = Toolbar1.Buttons.Add(, "A_BajarF", , tbrDefault, "A_BajarF"): BtnX.Visible = True: BtnX.ToolTipText = "Bajar"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_VerReceta", , tbrDefault, "A_VerReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Ver Recetas"
    Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Planificación Teórica" 'habilitar opción de copiar 5 etapas If (vg_codregimen > 9999 And "S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','"))) Then btnX.Enabled = False: btnX.ToolTipText = "" Else btnX.ToolTipText = "Copiar Planificación Teórica"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Aportes", , tbrDefault, "A_Aportes"): BtnX.Visible = True: BtnX.ToolTipText = "Aportes Nutricionales x Días"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Costo", , tbrDefault, "A_Costo"): BtnX.Visible = True: BtnX.ToolTipText = "Visualizar Costo"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Frecuencia", , tbrDefault, "A_Frecuencia"): BtnX.Visible = True: BtnX.ToolTipText = "Frecuencia Recetas"
    Set BtnX = Toolbar1.Buttons.Add(, "A_ActCostoReceta", , tbrDefault, "A_ActCostoReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Costo Receta"
    Set BtnX = Toolbar1.Buttons.Add(, "A_ExporReceta", , tbrDefault, "A_ExporReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Recetas"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    Call DetallePlantillaMinuta
    
    '-------> Llena Sub Menu estructura
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.EstServicio(1, vg_codservicio, 0), vg_db, adOpenStatic
    If Not RS.EOF Then
        X = 1
        Do While Not RS.EOF
            Load Estructura1(X): Load Estructura2(X)
            Estructura1(X).Caption = Trim(RS!ess_nombre): Estructura2(X).Caption = Trim(RS!ess_nombre)
            Estructura1(X).HelpContextID = RS!ess_codigo: Estructura2(X).HelpContextID = RS!ess_codigo
            Estructura1(X).Enabled = True: Estructura2(X).Enabled = True
            For i = 1 To vaSpread1.MaxRows
                vaSpread1.Col = vaSpread1.MaxCols: vaSpread1.Row = i
                If Trim(vaSpread1.text) <> "" Then
                    If Val(vaSpread1.text) = RS!ess_codigo Then
                        Estructura1(X).Enabled = False
                        Estructura2(X).Enabled = False
                    End If
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
If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 445
If Me.WindowState <> 1 Then vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380
End Sub

Private Sub Form_Unload(Cancel As Integer)
If SwSalir <> 0 Then Exit Sub
If Toolbar1.Buttons(1).Visible = True Then Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Cancel = -1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
If Toolbar1.Buttons(2).Visible = True Then GrabarPlantillaMinuta
Toolbar1.Buttons(1).Visible = True
Toolbar1.Buttons(2).Visible = False
SwSalir = 1
Me.Hide
Unload Me
M_Plami1.WindowState = 0
End Sub

Private Sub Image2_Click(Index As Integer)
Image2(0).Enabled = False
Image2(1).Enabled = False
fg_carga ""
G_TeoRea.LlenarGrafico vg_codcasino, vg_codregimen & ",", vg_codservicio & ",", Val(vg_fecha) & "01", Format(dEoM("01/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)), "yyyymmdd"), 1, 0, "'2'", 0, IIf(Index = 0, False, True)
G_TeoRea.Show 1
fg_descarga
Image2(0).Enabled = True
Image2(1).Enabled = True
End Sub

Private Sub Plantilla_Click(Index As Integer)

Dim RS                  As New ADODB.Recordset
Dim StrRec              As String
Dim StrRecb             As String
Dim aAp                 As String
Dim sql1                As String
Dim j                   As Long
Dim i                   As Long
Dim X                   As Long
Dim CodRec              As Long
Dim tiprec              As Long
Dim cosali              As Double
Dim CosDes              As Double
Dim SearchFlagsEqual    As Variant
Dim IndGrabado          As Boolean
Dim Colu                As Long
Dim xcol                As Integer
Dim Receta              As New M_Receta
Dim vecactrec           As Variant
Dim inddia              As Long

Select Case Index

Case 0 '-------> Grabar planificación minuta real
    
    If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
    If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Cancel = -1: Exit Sub
    If Toolbar1.Buttons(2).Visible = True Then GrabarPlantillaMinuta
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False

Case 5 '-------> Ver detalle receta
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Then vg_newestrec = True Else vg_newestrec = False
    xcol = 0
    For i = 1 To MaxColumna
        If (VectorCol(i) = vaSpread1.Col Or VectorCol(i) = (vaSpread1.Col + 1) Or VectorCol(i) = (vaSpread1.Col - 1) Or VectorCol(i) = (vaSpread1.Col - 2)) And Trim(vaSpread1.text) <> "" Then xcol = VectorCol(i): Exit For
    Next i
    If xcol = 0 Then MsgBox "No existe receta ha vizualizar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If vg_newestrec = True Then
       
       vg_fecval = 0: vg_fecval = Val(vg_fecha) & Right("0" & (Int(xcol / 5) + 1), 2)
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open RutinaLectura.Minutas(7, vg_codregimen, vg_codservicio, vg_fecval, "2"), vg_db, adOpenStatic
       If Not RS.EOF Then vg_fecval = RS!mid_fecval: vg_opcion = 2
       RS.Close: Set RS = Nothing
    
    End If
    vaSpread1.Col = xcol + 3
    StrRec = vaSpread1.text
    If Len(StrRec) <> 0 Then
       Do While InStr(StrRec, ";") <> 0
          StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
          StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
          vg_newcodrec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
          vg_tiprec = Val(Mid(StrRecb, 1))
       Loop
    End If
    vg_auxtiprec = vg_tiprec
    vg_5etapas = IIf(vg_codregimen < 10000, False, True)

    Receta.Show 1, Me
    vg_newestrec = False
    If vg_newcodrec <> 0 And Trim(vg_newnomrec) <> "" And vaSpread1.BackColor <> Shape1(1).FillColor And vg_auxtiprec = vg_tiprec Then
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
        '-------> Calcular costo receta alimentación y desechables
        cosali = Format(fg_CalCtoRecInv(Val(vg_newcodrec), vg_tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
        CosDes = Format(fg_CalCtoRecInv(Val(vg_newcodrec), vg_tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
        vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
        
        vaSpread1.Col = xcol + 3
        vaSpread1.text = vg_newcodrec & "&" & vg_tiprec & "&;"
        
        '-------> revizar si existe receta iguales en el mes y actualizar
        For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
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
                   
'                      vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
'                      If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxCols Then vaSpread1.text = 2
                   End If
                End If
            Next j
        Next i
        If IndCos = True Then
           For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
               Calctodia 1, i + 1
           Next i
           MostrarCosto vaSpread1.ActiveCol
        End If
        For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
            CalctodiaEnc 1, i + 1
        Next i
        '-------> Actualizar Lista receta
        If B_Receta.vaSpread1.MaxRows > 0 Then
            B_Receta.vaSpread1.Row = B_Receta.vaSpread1.SearchCol(1, 1, B_Receta.vaSpread1.MaxRows, Val(vg_newcodrec), SearchFlagsEqual)
            B_Receta.vaSpread1.Col = 3: B_Receta.vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
            B_Receta.vaSpread1.Col = 4
            If vg_tiprec = -1 Then
               B_Receta.vaSpread1.text = "Local"
            ElseIf vg_tiprec = 0 Then
               B_Receta.vaSpread1.text = "Patrón"
            ElseIf vg_tiprec > 0 Then
               B_Receta.vaSpread1.text = "x Regimen"
            End If
            B_Receta.vaSpread1.Col = 5: vaSpread1.text = vg_tiprec
        End If
        vg_newcodrec = 0: vg_newnomrec = "": vg_tiprec = -2
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        Toolbar1.Buttons(6).Visible = True
        Toolbar1.Buttons(7).Visible = False
    End If
    vg_newcodrec = 0
    vg_5etapas = IIf("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")), False, True)
    If IndCos = True Then Me.Refresh: Toolbar1.Refresh: Frame2(0).Refresh: Frame2(1).Refresh: Frame2(2).Refresh: Frame2(3).Refresh: Frame2(4).Refresh
Case 8 '-------> Copiar planificación minuta teórica
    '-------> habilitar opción 5 etapas jpaz If (vg_codregimen > 9999 And "S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','"))) Then Exit Sub
    M_CPlaTe.Show 1, Me
Case 10 '-------> Aporte Nutricional
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    j = 0
    For i = 1 To MaxColumna
        If (VectorCol(i) = vaSpread1.Col Or VectorCol(i) = (vaSpread1.Col + 1) Or VectorCol(i) = (vaSpread1.Col - 1) Or VectorCol(i) = (vaSpread1.Col - 2)) Then j = VectorCol(i): Exit For
    Next i
    vaSpread1.Col = j: vaSpread1.Row = 0
    C_ApoPla.LlenarApoPlan M_MinRea, "Aporte Planificación Real " & vaSpread1.text, vg_codcasino, vg_codregimen, vg_codservicio, Val(vg_fecha), 2, j
    C_ApoPla.Show 1, Me
Case 11 '-------> Costo planificación minutas
    If Frame2(0).Visible = True Then Frame2(0).Visible = False: vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380: Image2(0).Visible = False: Image2(1).Visible = False: Exit Sub
    vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 4000
    Frame2(0).Move 0, ScaleHeight - 2600, ScaleWidth, ScaleHeight - 1200
    Frame2(0).Visible = True
    CargarCosto
    Image2(0).Visible = True
    Image2(1).Visible = True
Case 12 '-------> Frecuencia
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    C_FrePla.LlenarFrecPlan "Frecuencia Planificación Teórica ", vg_codcasino, Val(Vg_FechaHasta), vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), 2
    'C_FrePla.LlenarFrecPlan "Frecuencia Planificación Real " & Mid(Vg_FechaDesde, 5, 2) & "/" & Mid(Vg_FechaDesde, 1, 4), Mid(Vg_FechaHasta, 5, 2) & "/" & Mid(Vg_FechaHasta, 1, 4), vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), 2
    'C_FrePla.LlenarFrecPlan "Frecuencia Planificación Real " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codcasino, vg_codregimen, vg_codservicio, Val(vg_fecha), 2
    C_FrePla.Show 1, Me
Case 13 '-------> Actualizar costo recetas y planificación
    If IndGrabado = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    If vg_tipbase = "1" Then
       '-------> Insert tabla productospmpdia
       aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPActRecPlaMinRea"
       fg_CheckTmp aAp
       vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                     "INTO " & aAp & " " & _
                     "FROM b_productospmpdia " & _
                     "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                     "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                     "AND   ppd_propon > 0 " & _
                     "GROUP BY ppd_cencos, ppd_codpro"
       vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
       vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
       vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    End If
    sql1 = IIf(vg_tipbase = "1", " val(mid(d.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),d.min_fecmin),1,6)) ")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
      
    If vg_tipbase = "1" Then
       
       RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.uni_nombre " & _
               "FROM b_productos a, a_unidad c, b_minuta d, b_minutadet e, b_recetadet f, b_ingrediente g, " & aAp & " h, b_contlistpreing i " & _
               "WHERE d.min_codigo = e.mid_codigo " & _
               "AND   e.mid_codrec = f.red_codigo " & _
               "AND   e.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
               "AND   f.red_codpro = g.ing_codigo " & _
               "AND   g.ing_codigo = i.cpi_coding " & _
               "AND   a.pro_codigo = h.ppd_codpro " & _
               "AND   a.pro_codigo = i.cpi_codcom " & _
               "AND   i.cpi_cencos = '" & vg_codcasino & "' " & _
               "AND   h.ppd_cencos = '" & vg_codcasino & "' " & _
               "AND  (a.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 1) OR a.pro_codigo NOT IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & ")) " & _
               "AND   d.min_cencos = '" & vg_codcasino & "' " & _
               "AND   d.min_codreg = " & vg_codregimen & " " & _
               "AND   d.min_codser = " & vg_codservicio & " " & _
               "AND   " & sql1 & " = " & vg_fecha & " " & _
               "AND   e.mid_tipmin = '2'  " & _
               "AND   a.pro_coduni = c.uni_codigo " & _
               "AND   i.cpi_precos <= 0 AND h.ppd_propon = 0 " & _
               "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
    Else
       RS.Open "SELECT DISTINCT a.pro_codigo, a.pro_nombre, c.uni_nombre " & _
               "FROM b_productos a, a_unidad c, b_minuta d, b_minutadet e, b_recetadet f, b_ingrediente g, b_productospmpdia h, b_contlistpreing i " & _
               "WHERE d.min_codigo = e.mid_codigo " & _
               "AND   e.mid_codrec = f.red_codigo " & _
               "AND   e.mid_tiprec = f.red_tiprec AND ((f.red_tiprec <> 0 AND f.red_cencos = '" & MuestraCasino(1) & "') OR (f.red_tiprec = 0 AND f.red_cencos = '0')) " & _
               "AND   f.red_codpro = g.ing_codigo " & _
               "AND   g.ing_codigo = i.cpi_coding " & _
               "AND   a.pro_codigo = h.ppd_codpro " & _
               "AND   a.pro_codigo = i.cpi_codcom " & _
               "AND   i.cpi_cencos = '" & vg_codcasino & "' " & _
               "AND   h.ppd_cencos = '" & vg_codcasino & "' AND h.ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " " & _
               "AND  (a.pro_codigo IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 1) OR a.pro_codigo NOT IN (SELECT bod_codpro FROM b_bodegas WHERE bod_codbod = " & vg_codbod & ")) " & _
               "AND   d.min_cencos = '" & vg_codcasino & "' " & _
               "AND   d.min_codreg = " & vg_codregimen & " " & _
               "AND   d.min_codser = " & vg_codservicio & " " & _
               "AND   " & sql1 & " = " & vg_fecha & " " & _
               "AND   e.mid_tipmin = '2'  " & _
               "AND   a.pro_coduni = c.uni_codigo " & _
               "AND   i.cpi_precos <= 0 AND h.ppd_propon = 0 " & _
               "AND  (a.pro_fecven > " & Format(Date, "yyyymmdd") & " OR a.pro_fecven <= 0) ORDER BY a.pro_nombre", vg_db, adOpenStatic
    End If
    '-------> Borrar tablas temporales
    If RS.EOF Then
       RS.Close: Set RS = Nothing
       If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
       MsgBox "No existe productos, con valores ceros", vbCritical + vbOKOnly, MsgTitulo
    Else
       RS.Close: Set RS = Nothing
       If Trim(aAp) <> "" And vg_tipbase = "1" Then vg_db.Execute "DROP TABLE " & aAp & ""
       M_ProPre.LlenarListaPrecio vg_codcasino, vg_codregimen, vg_codservicio, Val(vg_fecha), 2, Val(Vg_FechaHasta)
       M_ProPre.Show 1, Me
    End If
    fg_carga ""

    '-------> Traer total de receta desde planificación de minutas y luego calcular costo
    sql1 = IIf(vg_tipbase = "1", " VAL(MID(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
   
    RS.Open "SELECT COUNT(b.mid_codrec) AS nreg FROM b_minuta a, b_minutadet b " & _
            "WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & vg_codcasino & "' AND a.min_codreg = " & vg_codregimen & " AND a.min_codser = " & vg_codservicio & " " & _
            "AND   " & sql1 & " = " & Val(vg_fecha) & " AND b.mid_tipmin = '2'", vg_db, adOpenStatic
    If RS.EOF Or RS!nreg < 1 Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    ReDim vecactrec(RS!nreg, 4)
    RS.Close: Set RS = Nothing
    For i = 1 To UBound(vecactrec)
        DoEvents
        vecactrec(i, 1) = 0 '-------> codigo receta
        vecactrec(i, 2) = 0 '-------> tipo receta
        vecactrec(i, 3) = 0 '-------> costo receta alimentación
        vecactrec(i, 4) = 0 '-------> costo receta desechable
    Next i
    i = 1
    
    gauge1.Value = 0: gauge.Value = 0: Fecha = 0: inddia = 1: Fecha = 0: cosali = 0: CosDes = 0
    Picture1.Visible = True: Label2.Visible = False: Label3.Visible = True: Label3.Caption = "Recopilando información, un momento....": gauge.Visible = True: gauge.Visible = False
    sql1 = IIf(vg_tipbase = "1", " val(mid(a.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),a.min_fecmin),1,6)) ")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
   
    
    RS.Open "SELECT DISTINCT b.mid_codrec, b.mid_tiprec FROM b_minuta a, b_minutadet b " & _
            "WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & vg_codcasino & "' AND a.min_codreg = " & vg_codregimen & " AND a.min_codser = " & vg_codservicio & " " & _
            "AND   " & sql1 & " = " & Val(vg_fecha) & " AND b.mid_tipmin = '2'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "No existe información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    Do While Not RS.EOF
       DoEvents
       vecactrec(i, 1) = RS!mid_codrec
       vecactrec(i, 2) = RS!mid_tiprec
       vecactrec(i, 3) = Format(fg_CalCtoRecInv(RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
       vecactrec(i, 4) = Format(fg_CalCtoRecInv(RS!mid_codrec, IIf(IsNull(RS!mid_tiprec), 0, RS!mid_tiprec), (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
       RS.MoveNext: i = i + 1
    Loop
    RS.Close: Set RS = Nothing
    
    gauge1.Value = 0: gauge.Value = 0: Fecha = 0: inddia = 1: Fecha = 0: cosali = 0: CosDes = 0
    Picture1.Visible = True: Label2.Visible = False: Label3.Visible = True: Label3.Caption = "Actualizando costo receta, en planificación": gauge.Visible = True: gauge.Visible = False
    For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
'    For i = 2 To (vaSpread1.MaxCols - 1) Step 5
        DoEvents
        gauge1.Value = Val((i / vaSpread1.MaxCols) * 100)
        ExisteDat = 0
        vaSpread1.Row = 1: vaSpread1.Col = i
        Fecha = Val(vg_fecha) & fg_pone_cero(inddia, 2)
        If vaSpread1.BackColor <> Shape1(1).FillColor Then
           For j = 1 To (vaSpread1.MaxRows - 1)
               vaSpread1.Row = j
               vaSpread1.Col = i + 1
               If Trim(vaSpread1.text) <> "" Then ExisteDat = 1: Exit For
           Next j
           If ExisteDat > 0 Then
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
'                    cosali = fg_CalCtoRecInv(codrec, tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")))
'                    cosdes = fg_CalCtoRecInv(codrec, tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")))
                    vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
                    If vg_tipbase = "1" Then
                       vg_db.Execute "UPDATE b_minutadet INNER JOIN b_minuta ON b_minutadet.mid_codigo=b_minuta.min_codigo SET b_minutadet.mid_cosrec = " & cosali & ", b_minutadet.mid_cosdes = " & CosDes & " " & _
                                     "WHERE b_minuta.min_cencos = '" & vg_codcasino & "' AND b_minuta.min_codreg = " & vg_codregimen & " AND b_minuta.min_codser = " & vg_codservicio & " AND b_minuta.min_fecmin = " & Val(Fecha) & " AND b_minutadet.mid_codrec = " & CodRec & " AND b_minutadet.mid_tiprec = " & tiprec & " AND b_minutadet.mid_numlin = " & j & " AND b_minutadet.mid_tipmin = '2'"
                    Else
                       vg_db.Execute "UPDATE b_minutadet SET b_minutadet.mid_cosrec = " & cosali & ", b_minutadet.mid_cosdes = " & CosDes & " FROM b_minuta, b_minutadet WHERE b_minutadet.mid_codigo = b_minuta.min_codigo " & _
                                     "AND b_minuta.min_cencos = '" & vg_codcasino & "' AND b_minuta.min_codreg = " & vg_codregimen & " AND b_minuta.min_codser = " & vg_codservicio & " AND b_minuta.min_fecmin = " & Val(Fecha) & " AND b_minutadet.mid_codrec = " & CodRec & " AND b_minutadet.mid_tiprec = " & tiprec & " AND b_minutadet.mid_numlin = " & j & " AND b_minutadet.mid_tipmin = '2'"
                    End If
'                    vaSpread1.Col = (maxcolumna * 5 + 1) + ((i + 3) / 5)
'                    If Trim(vaSpread1.text) = "" And vaSpread1.Col < vaSpread1.MaxRows Then vaSpread1.text = 2
                    IndGrabado = 1
                  End If
              Next j
           End If
        End If
        inddia = inddia + 1
    Next i
    Label2.Visible = True: Picture1.Visible = False: gauge.Visible = False
    vaSpread1.Refresh
    If IndGrabado = 1 Then fg_descarga: MsgBox "Actualización costo receta finalizado sin problema, luego grabe información", vbInformation + vbOKOnly, MsgTitulo: Plantilla(0).Enabled = True: Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = True: Exit Sub
    fg_descarga
Case 14 '-------> Exportar recetas
    If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
    C_ExpRec.LlenarExporReceta "Exportar Recetas " & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4), vg_codcasino, vg_codregimen, vg_codservicio, Val(vg_fecha), 2
    C_ExpRec.Show 1, Me
Case 20
    SwSalir = 0
'    If Toolbar1.Buttons(2).Enabled = False Then indgrabado = 0
    If Toolbar1.Buttons(1).Visible = True Then SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0: Exit Sub
    If MsgBox(" Actualiza planificación real...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Toolbar1.Buttons(2).Visible = False
    If Toolbar1.Buttons(2).Visible = True Then GrabarPlantillaMinuta
    SwSalir = 1: Me.Hide: Unload Me: M_Plami1.WindowState = 0
End Select
End Sub

Private Sub Plato_Click(Index As Integer)
On Error GoTo Man_Error

Dim Del_Row          As Integer
Dim indcol           As Integer
Dim indrow           As Integer
Dim IndCol2          As Integer
Dim IndRow2          As Integer
Dim indrow3          As Long
Dim XX               As Long
Dim Col              As Long
Dim fil              As Long
Dim AddRec           As Long
Dim cosali           As Double
Dim CosDes           As Double
Dim AadRec           As Long
Dim z                As Long
Dim Colu             As Long
Dim FilaAct          As Long
Dim FilaAnt          As Long
Dim FilaPos          As Long
Dim AuxIblockrow     As Long
Dim ValLcntH         As String 'Long
Dim L                As Long
Dim f                As Long
Dim c                As Long
Dim nrodia           As String
Dim vecdia()         As String
Dim xSer             As Long
Dim iSer             As Long
Dim cantCol          As Long
Dim cantCol1         As Long
Dim LargoVec         As Long
Dim d                As String
Dim contador         As Long
Dim contador_b       As Long
Dim ColumnaActiva    As Long
Dim FilaActiva       As Long
Dim ColumnaAntActiva As Long
Dim accion           As String
Dim n                As Long
Dim n1               As Long
Dim max              As Long
Dim max1             As Long
Dim ff               As Long
Dim g                As Long
Dim tope             As Long
Dim desc             As String
Dim NFilas           As Long
Dim FechaMin         As Long
Dim RS               As New ADODB.Recordset

contador = 0: contador_b = 0: cantCol = 0: LargoVec = 0:  accion = "": n1 = 0: n = 0: NFilas = 0

'If Index <> 2 And Index <> 15 And Index <> 12 And Index <> 13 And Index <> 5 Then
If Index <> 2 And Index <> 15 And Index <> 12 And Index <> 13 Then
    
    If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
       
       MsgBox "Minuta, sin acceso a modificaciones", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    
    End If

End If
    
If Toolbar1.Buttons(2).Enabled = False Or (vg_codregimen > 9999 And etapa5 And AddReceta = 0) Then
   
   Exit Sub

End If

Select Case Index

Case 2 '-------> Ingresar Recetas
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Or (vg_codregimen > 9999 And etapa5 And AddReceta = 0) Then Exit Sub
    IblockCol = vaSpread1.ActiveCol: AiBlockCol = vaSpread1.ActiveCol
    iblockcol2 = vaSpread1.ActiveCol: AiBlockCol2 = vaSpread1.ActiveCol
    IblockRow = vaSpread1.ActiveRow: AiBlockRow = vaSpread1.ActiveRow
    IblockRow2 = vaSpread1.ActiveRow: AiBlockRow2 = vaSpread1.ActiveRow
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    '-------> Validar dia bloqueado
    If vaSpread1.BackColor = Shape1(1).FillColor Then
       MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    
    j = 0
    For i = 1 To MaxColumna
        If vaSpread1.Col = VectorCol(i) Then j = VectorCol(i): Exit For
    Next i
    If j = 0 Then Exit Sub
    
    vg_codigo = "": vg_nombre = "": vg_tiprec = -2
    vaSpread1.Col = j - 1
    vaSpread1.Row = vaSpread1.ActiveRow
        
    '-------> Validar receta 5 etapa
    If vaSpread1.BackColor = &H80FF80 And etapa5 Then
       MsgBox "No puede modificar receta, corresponde receta centralizada", vbCritical + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    
    '-------> Validar minuta bloque si puede insertar recetas
    If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
       
       '-------> Sacar fecha de la grilla
       vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1)))
       vaSpread1.Col = vaSpread1.ActiveCol

       FechaMin = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
          
       Set RS = vg_db.Execute("sgp_Sel_ValidarMinBloque '" & vg_codcasino & "', " & FechaMin & "")
       If RS.EOF Then
          RS.Close: Set RS = Nothing
          MsgBox "No puede insertar recetas, para esta fecha", vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
       End If
       RS.Close: Set RS = Nothing
           
    End If
    
    vaSpread1.Col = j - 1
    vaSpread1.Row = vaSpread1.ActiveRow

    AddRec = 0
    If etapa5 And AddReceta > 0 And Trim(vaSpread1.text) = "" Then
       For i = 1 To vaSpread1.MaxRows - 1
           vaSpread1.Row = i
           If vaSpread1.BackColor = &HFFFF00 Then
              AddRec = AddRec + 1
              If AddRec >= AddReceta Then MsgBox "No puede ingresar más receta, tiene un maximo " & AddReceta & " por día", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
           End If
       Next i
    End If

    B_Receta.Show 1, Me
    If Trim(vg_codigo) = "" Or Trim(vg_nombre) = "" Or vg_tiprec < -1 Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow

    vaSpread1.Col = j - 1
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 2
    vaSpread1.Value = IIf(Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1), "A", "R")
    vaSpread1.ForeColor = IIf(Not etapa5, &HFF&, &H400000)
    vaSpread1.BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
    
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
       '-------> Asignar Raciones estimadas
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = j + 1
       vaSpread1.CellType = 3
       vaSpread1.TypeIntegerMin = 1
       vaSpread1.TypeIntegerMax = 9999999
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypeSpin = False
       vaSpread1.TypeIntegerSpinInc = 1
       vaSpread1.TypeIntegerSpinWrap = False
       vaSpread1.text = 0
       vaSpread1.ForeColor = &HFF0000
    End If
    
    vaSpread1.Col = j + 2
    vaSpread1.CellType = 5
    vaSpread1.TypeHAlign = 1
    cosali = Format(fg_CalCtoRecInv(Val(vg_codigo), vg_tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','"))), fg_Pict(6, 2))
    CosDes = Format(fg_CalCtoRecInv(Val(vg_codigo), vg_tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))), fg_Pict(6, 2))
    vaSpread1.text = Format((cosali + CosDes), fg_Pict(6, 2))
    
    vaSpread1.Col = j + 3
    
    If Trim(vaSpread1.text) <> Val(vg_codigo) And ((MaxColumna * 5 + 1) + ((j + 2) / 5)) < vaSpread1.MaxCols Then
       vaSpread1.Col = (MaxColumna * 5 + 1) + ((j + 2) / 5)
       vaSpread1.text = 1
       ': If indcos = True Then vaSpread1.col = J + 2: veccos((Int(J / 5) + 1), 1) = Round(veccos((Int(J / 5) + 1), 1) - vaSpread1.Text, vg_DCa)
    End If
    
    vaSpread1.Col = j + 3
    vaSpread1.text = Val(vg_codigo) & "&" & vg_tiprec & "&;"
    If IndCos = True Then Calctodia vaSpread1.Row, j
    CalctodiaEnc vaSpread1.Row, j
    
    vaSpread1.Row = vaSpread1.ActiveRow
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    
Case 5 '-------> Insertar linea
  
    vaSpread1.Visible = False
    indcol = IblockCol
    IblockCol = 1: iblockcol2 = vaSpread1.MaxCols
    vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
    vaSpread1.MaxRows = vaSpread1.MaxRows + ((xRowFin - xRowIni) + 1) '1
    vaSpread1.InsertRows xRowIni, ((xRowFin - xRowIni) + 1)
    
    Do While xRowIni <= xRowFin
    For i = 3 To (vaSpread1.MaxCols - MaxColumna) Step 5
        vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))): vaSpread1.Col = i
        If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
            For c = i - 1 To i + 2
                vaSpread1.Row = xRowIni: vaSpread1.Col = c
                vaSpread1.BackColor = Shape1(1).FillColor
            Next c
        End If
    Next i
      xRowIni = xRowIni + 1
    Loop
    vaSpread1.Visible = True
    
    '-------> Validar días modificados
    For j = IblockRow To ((vaSpread1.MaxRows - 1)) '- ((iblockrow2 - iblockrow) + 1))
        For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
            vaSpread1.Row = j
        Next i
    Next j
    '-------> Fin validar días modificados
    IblockCol = indcol
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indgri = True
    
Case 6 '-------> Eliminar Linea
    indcol = IblockCol
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor And Trim(vaSpread1.text) <> "" Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    If vaSpread1.BackColor = Shape1(1).FillColor Or Trim(vaSpread1.text) = "" Then GoTo paso
    j = 0
    For i = 1 To MaxColumna
        If (VectorCol(i) - 1) = vaSpread1.Col Or VectorCol(i) = vaSpread1.Col Then j = (VectorCol(i) - 1): Exit For
    Next i
    If j = 0 Then Exit Sub
    If indactivo = 0 Then IblockCol = vaSpread1.ActiveCol: iblockcol2 = vaSpread1.ActiveCol: IblockRow = vaSpread1.ActiveRow: IblockRow2 = vaSpread1.ActiveRow
    AiBlockCol = IblockCol
    AiBlockRow = IblockRow
    AiBlockCol2 = iblockcol2
    AiBlockRow2 = IblockRow2
    If IblockCol < 0 Then IblockCol = 2: iblockcol2 = vaSpread1.MaxCols
    AiBlockCol = IblockCol
    AiBlockRow = IblockRow
    AiBlockCol2 = iblockcol2
    AiBlockRow2 = IblockRow2
    For i = 1 To MaxColumna
        If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
    Next i
    For i = 1 To MaxColumna
        If (VectorCol(i) - 1) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 4)): Exit For
        If VectorCol(i) = iblockcol2 Then iblockcol2 = (VectorCol(i) + 3): Exit For
    Next i
    indcol = AiBlockCol: IndCol2 = iblockcol2
    indrow = AiBlockRow: IndRow2 = AiBlockRow2
    If IndCos = True Then
       For i = IblockCol To iblockcol2 Step 5
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
    End If
    For i = IblockCol To iblockcol2 Step 5
        CalctodiaEnc 1, i + 1
    Next i
    '-------> Validar días modificados
    For j = IblockRow To ((vaSpread1.MaxRows - 1) - ((IblockRow2 - IblockRow) + 1) + 1)
        For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
            vaSpread1.Row = j
            vaSpread1.Col = i + 1
        Next i
    Next j
    '-------> Fin validar días modificados
    IblockCol = AuxCol
    vaSpread1.BlockMode = False
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indactivo = 0
paso:
    vaSpread1.Row = vaSpread1.ActiveRow
    For i = 1 To vaSpread1.MaxCols
        vaSpread1.Col = i
        If Trim(vaSpread1.text) <> "" Then MsgBox "Existe mas información en la linea, no puede eliminarla completamente", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
    Next i
    vaSpread1.Row = IblockRow2
    vaSpread1.Col = IblockCol
    vaSpread1.Visible = False
    vaSpread1.DeleteRows IblockRow, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    vaSpread1.Visible = True
    IblockCol = indcol
    For i = 3 To (vaSpread1.MaxCols - MaxColumna) Step 5
        vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))): vaSpread1.Col = i
        If CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
            For Col = 0 To i - 4
                vaSpread1.Row = (vaSpread1.MaxRows - 1): vaSpread1.Col = Col + 2
                vaSpread1.BackColor = Shape1(1).FillColor
            Next Col
        End If
    Next i
    '-------> Validar días modificados
    For j = IblockRow To ((vaSpread1.MaxRows - 1))
        For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
            vaSpread1.Row = j
            vaSpread1.Col = i + 1
            If Trim(vaSpread1.text) <> "" Then
            End If
        Next i
    Next j
    '-------> Fin validar días modificados
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indgri = True
    
Case 8 '-------> Subir linea
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = 1 Or vaSpread1.Row = vaSpread1.MaxRows Then Exit Sub
    If IblockCol < 1 Or (IblockCol = 1 And Trim(vaSpread1.text) <> "") Then
       For i = 1 To MaxColumna
           vaSpread1.Col = VectorCol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
       Next i
    Else
       For i = IblockCol To iblockcol2
           vaSpread1.Col = i
           For j = IblockRow To IblockRow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
           Next j
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col > 1 Then
        indcol = IblockCol
        vaSpread1.Col = 1
        If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        
        If (IblockRow - ((IblockRow2 - IblockRow) + 1)) < 1 Then
           MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
        End If
        If IblockCol < 0 Then IblockCol = 1: iblockcol2 = vaSpread1.MaxCols
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Or (VectorCol(i) + 1) = IblockCol Or (VectorCol(i) + 2) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 4)): Exit For
            If VectorCol(i) = iblockcol2 Or VectorCol(i) + 1 = iblockcol2 Or VectorCol(i) + 2 = iblockcol2 Then iblockcol2 = (VectorCol(i) + 3): Exit For
        Next i
        
        '-------> Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = True
        vaSpread1.MoveRange IblockCol, (IblockRow - 1), iblockcol2, (IblockRow - 1), IblockCol, vaSpread1.MaxRows
        '-------> Copiar datos fila seleccionada
        vaSpread1.ClearRange IblockCol, (IblockRow + 1), iblockcol2, (IblockRow - 1), False
        vaSpread1.MoveRange IblockCol, IblockRow, iblockcol2, IblockRow, IblockCol, (IblockRow - 1)
        '-------> Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange IblockCol, IblockRow, iblockcol2, IblockRow, False
        vaSpread1.MoveRange IblockCol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, IblockCol, IblockRow
        vaSpread1.Row = IblockRow - 1: vaSpread1.Col = IblockCol
        vaSpread1.DeleteRows vaSpread1.MaxRows, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        '-------> Validar días modificados
        For j = (IblockRow - 1) To (vaSpread1.MaxRows - 1)
            For i = IblockCol To iblockcol2 Step 5
                vaSpread1.Row = j
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.text) <> "" Then
                End If
            Next i
        Next j
        '-------> Fin validar días modificados
        vaSpread1.Row = IblockRow - 1: vaSpread1.Col = IblockCol
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.text) = "" Then Exit Sub
        For i = IblockRow - 1 To 1 Step -1 '-------> Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next i
        For z = IblockRow + 1 To (vaSpread1.MaxRows - 1) '-------> Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then
            For fil = (vaSpread1.MaxRows - 1) To 1 Step -1
                For Colu = 1 To vaSpread1.MaxCols
                    vaSpread1.Col = Colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next Colu
                If z <= (vaSpread1.MaxRows - 1) Then Exit For
                If z <= (vaSpread1.MaxRows) Then Exit For
            Next fil
        End If
        FilaAct = IblockRow         '-------> Fila actual
        FilaAnt = IIf(i < 1, 1, i)  '-------> Fila anterior
        FilaPos = z                 '-------> Fila posterior
        
        '-------> Validar días modificados
        For j = FilaAnt To (vaSpread1.MaxRows - 1)
            For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
                vaSpread1.Row = j
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.text) <> "" Then
                End If
            Next i
        Next j
        '-------> Fin validar días modificados
        
        '-------> Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (FilaAct - FilaAnt)
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows + (FilaAct - FilaAnt)
            vaSpread1.Row = i
            vaSpread1.RowHidden = True
        Next i
        vaSpread1.MoveRange 1, FilaAnt, vaSpread1.MaxCols, (FilaAct - 1), 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1
        
        '-------> Mover estructura
        vaSpread1.MoveRange 1, FilaAct, vaSpread1.MaxCols, (FilaPos - 1), 1, FilaAnt
        '-------> Devolver respaldo
        vaSpread1.MoveRange 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1, vaSpread1.MaxCols, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 + (FilaAct - FilaAnt - 1), 1, FilaAnt + (FilaPos - FilaAct)
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows
            vaSpread1.DeleteRows i, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Next i
        vaSpread1.SetActiveCell 1, FilaAnt
    End If
    vaSpread1.Row = IblockRow - 1: vaSpread1.Col = IblockCol
    IblockRow = vaSpread1.ActiveRow: IblockRow2 = vaSpread1.ActiveRow: IblockCol = vaSpread1.ActiveCol
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
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
    indgri = True
    
Case 9 '-------> Bajar linea
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Row = vaSpread1.MaxRows Then Exit Sub
    If IblockCol < 1 Or (IblockCol = 1 And Trim(vaSpread1.text) <> "") Then
       For i = 1 To MaxColumna
           vaSpread1.Col = VectorCol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
       Next i
    Else
       For i = IblockCol To iblockcol2
           vaSpread1.Col = i
           For j = IblockRow To IblockRow2
              vaSpread1.Row = j
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
           Next j
       Next i
    End If
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col > 1 Then
        vaSpread1.Col = 1
        vaSpread1.Row = vaSpread1.ActiveRow + 1
        If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
        vaSpread1.Row = vaSpread1.ActiveRow - 1
        If (IblockRow2 + ((IblockRow2 - IblockRow) + 1)) > (vaSpread1.MaxRows - 1) Then
           MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
        End If
        indcol = IblockCol
        If IblockCol < 0 Then IblockCol = 1: iblockcol2 = vaSpread1.MaxCols
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Or (VectorCol(i) + 1) = IblockCol Or (VectorCol(i) + 2) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
        Next i
        For i = 1 To MaxColumna
            If (VectorCol(i) - 1) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 4)): Exit For
            If VectorCol(i) = iblockcol2 Or VectorCol(i) + 1 = iblockcol2 Or VectorCol(i) + 2 = iblockcol2 Then iblockcol2 = (VectorCol(i) + 3): Exit For
        Next i
        
        '-------> Validar días modificados
        For j = IblockRow To (vaSpread1.MaxRows - 1)
            For i = IblockCol To iblockcol2 Step 5
                vaSpread1.Row = j
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.text) <> "" Then
                End If
            Next i
        Next j
        '-------> Fin validar días modificados
        
        '-------> Copiar datos ultima fila
        vaSpread1.MaxRows = vaSpread1.MaxRows + 1
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.RowHidden = True
        vaSpread1.MoveRange IblockCol, (IblockRow + 1), iblockcol2, (IblockRow + 1), IblockCol, vaSpread1.MaxRows
    
        '-------> Copiar datos fila Seleccionada
        vaSpread1.ClearRange IblockCol, (IblockRow + 1), iblockcol2, (IblockRow + 1), False
        vaSpread1.MoveRange IblockCol, IblockRow, iblockcol2, IblockRow, IblockCol, (IblockRow + 1)
    
        '-------> Devolver datos fila y restar ultima fila
        vaSpread1.ClearRange IblockCol, IblockRow, iblockcol2, IblockRow, False
        vaSpread1.MoveRange IblockCol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, IblockCol, IblockRow
        vaSpread1.DeleteRows vaSpread1.MaxRows, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        vaSpread1.Row = IblockRow + 1: vaSpread1.Col = IblockCol
        vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
    ElseIf vaSpread1.Col = 1 Then
        If Trim(vaSpread1.text) = "" Then Exit Sub
        For z = IblockRow + 1 To (vaSpread1.MaxRows - 1) '-------> Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then Exit Sub
        vaSpread1.Col = vaSpread1.ActiveCol
        AuxIblockrow = z
        For i = AuxIblockrow - 1 To 1 Step -1 '-------> Recorre el espacio que hay entre la estructura seleccioneda y la anterior
            vaSpread1.Row = i
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next i
        For z = AuxIblockrow + 1 To (vaSpread1.MaxRows - 1) '-------> Recorre el espacio que hay entre la estructura seleccioneda y la posterior
            vaSpread1.Row = z
            If Trim(vaSpread1.text) <> "" Then Exit For
        Next z
        If z > (vaSpread1.MaxRows - 1) Then
            For fil = (vaSpread1.MaxRows - 1) To 1 Step -1
                For Colu = 1 To vaSpread1.MaxCols
                    vaSpread1.Col = Colu: vaSpread1.Row = fil
                    If Trim(vaSpread1.text) <> "" Then
                        z = fil + 1: Exit For
                    End If
                Next Colu
                If z <= (vaSpread1.MaxRows - 1) Then Exit For
            Next fil
        End If
        FilaAct = AuxIblockrow      '-------> Fila actual
        FilaAnt = IIf(i < 1, 1, i)  '-------> Fila anterior
        FilaPos = z                 '-------> Fila posterior
        '-------> Validar días modificados
        For j = FilaAnt To (vaSpread1.MaxRows - 1)
            For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
                vaSpread1.Row = j
                vaSpread1.Col = i + 1
                If Trim(vaSpread1.text) <> "" Then
                End If
            Next i
        Next j
        '-------> Fin validar días modificados
        
        '-------> Agregar filas temporales y respaldar
        vaSpread1.MaxRows = vaSpread1.MaxRows + (FilaAct - FilaAnt)
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows + (FilaAct - FilaAnt)
            vaSpread1.Row = i
            vaSpread1.RowHidden = True
        Next i
        vaSpread1.MoveRange 1, FilaAnt, vaSpread1.MaxCols, (FilaAct - 1), 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1
        
        '-------> Mover estructura
        vaSpread1.MoveRange 1, FilaAct, vaSpread1.MaxCols, (FilaPos - 1), 1, FilaAnt

        '-------> Devolver respaldo
        vaSpread1.MoveRange 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1, vaSpread1.MaxCols, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 + (FilaAct - FilaAnt - 1), 1, FilaAnt + (FilaPos - FilaAct)
        For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows
            vaSpread1.DeleteRows i, 1
            vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        Next i
        vaSpread1.SetActiveCell 1, FilaAnt + (FilaPos - FilaAct)
    End If
    IblockRow = vaSpread1.ActiveRow: IblockRow2 = vaSpread1.ActiveRow: IblockCol = vaSpread1.ActiveCol
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
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
    indgri = True
    
Case 11, 12 '-------> Copiar y pegar linea
    vaSpread1.Col = vaSpread1.ActiveCol
    vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1)))

    If vaSpread1.ActiveRow = vaSpread1.MaxRows And vaSpread1.text <> "N.Rac." Then Exit Sub
    If Index = 11 Then
       If IblockCol < 1 Then
          For i = 1 To MaxColumna
              vaSpread1.Col = VectorCol(i)
              vaSpread1.Row = 1
              If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Existen Días Bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
          Next i
       Else
          For i = IblockCol To iblockcol2
              vaSpread1.Col = i
              For j = IblockRow To IblockRow2
                 vaSpread1.Row = j
                 If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Bloque seleccionado existen días bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
              Next j
          Next i
       End If
       '-------> Validar recetas 5 etapas
       j = 0
       For i = 1 To MaxColumna
           If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Then j = (VectorCol(i) - 1): Exit For
       Next i
       If j = 0 Then Exit Sub
       If etapa5 And AddReceta > 0 Then
          For j = j To iblockcol2 Step 5
              vaSpread1.Col = j
              For i = IblockRow To (IblockRow2)
                  vaSpread1.Row = i
                  If vaSpread1.BackColor = &H80FF80 Then MsgBox "No puede cortar receta, corresponde 5 etapas", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
              Next i
          Next j
       End If
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    
    '------> Verificar si copiar receta o raciones solamente
    vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1)))
    If vaSpread1.text = "N.Rac." Then
      TipoCopia = "Copiar Raciones"
    Else
      TipoCopia = "Copiar Receta"
    End If
    
    AiBlockRow = IblockRow: AiBlockRow2 = IblockRow2
    AiBlockCol = IblockCol: AiBlockCol2 = iblockcol2
    If vaSpread1.Col = 1 Then Exit Sub
    Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(7).Visible = True
    Plato(13).Enabled = True: OpGrilla(13).Enabled = True
    Plato(14).Enabled = True: OpGrilla(14).Enabled = True
    If IblockCol < 1 Then AiBlockCol = 1: AiBlockCol2 = vaSpread1.MaxCols
    IndCortarPegar = 1
    If Index = 11 Then IndCortarPegar = 0: Toolbar1.Buttons(8).Visible = True: Toolbar1.Buttons(9).Visible = False: Plato(14).Enabled = False: OpGrilla(14).Enabled = False Else Toolbar1.Buttons(8).Visible = False: Toolbar1.Buttons(9).Visible = True: Plato(14).Enabled = True: OpGrilla(14).Enabled = True

Case 13, 14 'cortar y pegar
    

    '-------> Validar recetas 5 etapas
    AddRec = 0
    If etapa5 And AddReceta > 0 And Index <> 14 And TipoCopia <> "Copiar Raciones" Then
       For X = 1 To MaxColumna
           If (VectorCol(X) - 1) = IblockCol Or VectorCol(X) = IblockCol Or (VectorCol(X) + 1) = IblockCol Or (VectorCol(X) + 2) = IblockCol Then IblockCol = (VectorCol(X) - 1): Exit For
       Next X
        For j = IblockCol To iblockcol2
           vaSpread1.Col = j - 1
           For i = IblockRow To (IblockRow2 + (AiBlockRow2 - AiBlockRow))
              vaSpread1.Row = i
              If vaSpread1.BackColor = &H80FF80 Then
                 MsgBox "No puede modificar receta, corresponde 5 etapas", vbCritical + vbOKOnly, MsgTitulo
                 Exit Sub
              End If
          Next i
          For i = 1 To vaSpread1.MaxRows - 1
              vaSpread1.Row = i
              If vaSpread1.BackColor = &HFFFF00 Then
                 AddRec = AddRec + 1
                 If AddRec >= AddReceta Then
                    MsgBox "No puede ingresar más receta, tiene un maximo " & AddReceta & " por día", vbCritical + vbOKOnly, MsgTitulo
                    Exit Sub
                 End If
              End If
          Next i
        Next j
    End If
    
    '-------> copiar y pegar
    If IndCortarPegar = 0 Then
       If (iblockcol2 - IblockCol) > (AiBlockCol2 - AiBlockCol) Or (IblockRow2 - IblockRow) > (AiBlockRow2 - AiBlockRow) Then MsgBox "Imposible Pegar la infomación ya que el área de Cortar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
       IndCortarPegar = 0
    Else
       If (IblockRow2 - IblockRow) > (AiBlockRow2 - AiBlockRow) Then
          MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo
          Exit Sub
       End If
       If AiBlockCol <> iblockcol2 And AiBlockCol = 1 Then
          MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo
          Exit Sub
       End If
    End If
    
    If IblockCol < 1 Then
       
       For i = 1 To MaxColumna
           vaSpread1.Col = VectorCol(i)
           vaSpread1.Row = 1
           
           If vaSpread1.BackColor = Shape1(1).FillColor Then
              MsgBox "Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo
              vaSpread1.Row = vaSpread1.ActiveRow
              vaSpread1.Col = vaSpread1.ActiveCol
              Exit Sub
           End If
       
       Next i
    
    Else
       For i = IblockCol To iblockcol2
           vaSpread1.Col = i
           For j = IblockRow To IblockRow2
              
              vaSpread1.Row = j
              If TipoCopia = "Copiar Raciones" And MaxFila = j Then
                 vaSpread1.Row = j - 1
              End If

              If vaSpread1.BackColor = Shape1(1).FillColor And Index <> 14 Then
                 MsgBox "Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo
                 Exit Sub
              End If
           
           Next j
       Next i
    End If
    
    '-------> Validar minuta bloque si puede insertar recetas
    If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
       
       vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1)))
       vaSpread1.Col = IblockCol + 1

       FechaMin = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")

       Set RS = vg_db.Execute("sgp_Sel_ValidarMinBloque '" & vg_codcasino & "', " & FechaMin & "")
       If RS.EOF Then
          RS.Close: Set RS = Nothing
          MsgBox "No puede insertar recetas, para esta fecha", vbCritical + vbOKOnly, MsgTitulo
          vaSpread1.Row = vaSpread1.ActiveRow
          vaSpread1.Col = vaSpread1.ActiveCol
          Exit Sub
       End If
       RS.Close: Set RS = Nothing
           
    End If
        
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    If IndCortarPegar = 0 Then
       Toolbar1.Buttons(6).Visible = True
       Toolbar1.Buttons(7).Visible = False
    End If
    
    '-------> destinacion de copiar y pegar datos
    If IblockCol < 1 Then
       IblockCol = 1
       iblockcol2 = vaSpread1.MaxCols
    End If
    
    If AiBlockCol2 = vaSpread1.MaxCols Then
       AiBlockCol2 = vaSpread1.MaxCols - 1
    End If
    
    vaSpread1.Row = 0: vaSpread1.Col = IblockCol
    vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1)))
    If vaSpread1.text = "N.Rac." And TipoCopia = "Copiar Raciones" Then
       cantCol = AiBlockCol2 - AiBlockCol
       cantCol1 = iblockcol2 - IblockCol
    ElseIf vaSpread1.text <> "N.Rac." And TipoCopia = "Copiar Raciones" Then
        MsgBox "Imposible Pegar la infomación ya que tiene una columna distinta N.Raciones", vbInformation + vbOKOnly, MsgTitulo:  Exit Sub
    Else
       vaSpread1.Row = 0
       For i = 1 To MaxColumna
           If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Or (VectorCol(i) + 1) = IblockCol Or (VectorCol(i) + 2) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
       Next i
       For i = 1 To MaxColumna
           If (VectorCol(i) - 1) = AiBlockCol Or VectorCol(i) = AiBlockCol Or (VectorCol(i) + 1) = AiBlockCol Or (VectorCol(i) + 2) = AiBlockCol Then AiBlockCol = (VectorCol(i) - 1): Exit For
       Next i
       For i = 1 To MaxColumna
           If (VectorCol(i) - 1) = iblockcol2 Or VectorCol(i) = iblockcol2 Or (VectorCol(i) + 1) = iblockcol2 Or (VectorCol(i) + 2) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 3)): Exit For
       Next i
       For i = 1 To MaxColumna
           If (VectorCol(i) - 1) = AiBlockCol2 Or VectorCol(i) = AiBlockCol2 Or (VectorCol(i) + 1) = AiBlockCol2 Or (VectorCol(i) + 2) = AiBlockCol2 Then AiBlockCol2 = (VectorCol(i) + 3): Exit For
       Next i
    End If
    
    '-----> Llena vectores con las raciones
    LargoVec = AiBlockRow2 - AiBlockRow + 1

    If AiBlockCol > 1 And AiBlockRow > 0 Then
       ReDim VecSelGrid(0)
       ReDim VecSelGrid(20000)
       For i = AiBlockCol To AiBlockCol2
           vaSpread1.Col = i
           vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))) '0
           d = vaSpread1.text
           If vaSpread1.text = "N.Rac." Then
              For j = AiBlockRow To AiBlockRow + LargoVec - 1
                  vaSpread1.Col = i
                  vaSpread1.Row = j
                  d = vaSpread1.text
                  contador = contador + 1
                  If Trim(vaSpread1.text) <> "" Then VecSelGrid(contador) = vaSpread1.text   ' Almacena las raciones a copiar
              Next j
           End If
       Next i
    End If
    
    If vaSpread1.ActiveCol > 1 And vaSpread1.ActiveRow > 0 Then
'       ReDim VecRacPegar(0)
       Dim VecRacPegar() As Variant
       ReDim VecRacPegar(20000, 2)
       For i = 1 To UBound(VecRacPegar)
           VecRacPegar(i, 1) = 0
           VecRacPegar(i, 2) = 0
       Next
       
       For i = IblockCol To iblockcol2
           vaSpread1.Col = i
           vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))) '0
           If vaSpread1.text = "N.Rac." Then
              For j = vaSpread1.ActiveRow To vaSpread1.ActiveRow + contador - 1
                  vaSpread1.Col = i: vaSpread1.Row = j
                  contador_b = contador_b + 1
                  If Trim(vaSpread1.text) <> "" Then VecRacPegar(contador_b, 1) = vaSpread1.text ' Almacena las raciones a reemplazar
                  vaSpread1.Col = i + 1: vaSpread1.Row = j
                  If Trim(vaSpread1.text) <> "" Then VecRacPegar(contador_b, 2) = vaSpread1.text ' Almacena las raciones a reemplazar
              Next j
           End If
       Next i
    End If
    
    indcol = AiBlockCol: IndCol2 = iblockcol2
    indrow = AiBlockRow: IndRow2 = AiBlockRow2
    If Index = 14 And IndCortarPegar = 1 Then
       
       If (AiBlockRow2 - AiBlockRow) <> 0 Or (AiBlockCol2 - AiBlockCol) <> 4 Then
          MsgBox "Por esta opción solamente puede copiar una receta", vbInformation + vbOKOnly, MsgTitulo
          IblockCol = vaSpread1.ActiveCol
          Exit Sub
       End If
       
       '-------> Rutina pegado especial
       vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))): nrodia = ""
       For i = AiBlockCol To AiBlockCol2 Step 5
           vaSpread1.Col = i + 1
           nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
       Next i
       For i = 1 To MaxColumna
           vaSpread1.Col = VectorCol(i)
           vaSpread1.Row = 1
           If vaSpread1.BackColor = Shape1(1).FillColor Then vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))): nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
       Next i
       '------> Validar receta 5 etapa
       AadRec = 0
       If etapa5 And AddReceta > 0 Then
          For j = 1 To UBound(VectorCol)
              vaSpread1.Col = VectorCol(j) - 1
              AddRec = 0
              For i = 1 To vaSpread1.MaxRows - 1
                  vaSpread1.Row = i
                  If vaSpread1.BackColor = &HFFFF00 Then
                     AddRec = AddRec + 1
                     If AddRec >= AddReceta Then
                        vaSpread1.Col = VectorCol(j)
                        vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))): nrodia = nrodia & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2)) & ";"
                        Exit For
                     End If
                  End If
              Next i
          Next j
       End If
       vg_codigo = ""
       Call M_CpRPla.Inicio("Copia Especial Recetas en Planificación Real", "PLAREA", Vg_FechaDesde, Vg_FechaHasta, nrodia, 1)
       M_CpRPla.Show 1
       If Trim(vg_codigo) = "" Then
          IblockCol = vaSpread1.ActiveCol
          Exit Sub
       End If
       
       '-------> mover días no permitidos
       ReDim Preserve vecdia(0)
       ValLcntH = "": i = 0
       For j = 1 To Len(vg_codigo)
           If Asc(Mid(vg_codigo, j, 1)) <> 59 Then
              ValLcntH = ValLcntH + Mid(vg_codigo, j, 1)
           Else
              ReDim Preserve vecdia(i): vecdia(i) = ValLcntH: ValLcntH = "": i = i + 1
           End If
       Next j
       If Trim(ValLcntH) <> "" Then
            ReDim Preserve vecdia(i)
            vecdia(i) = ValLcntH
        End If
       
       For i = 3 To (vaSpread1.MaxCols - MaxColumna) Step 5
           vaSpread1.Row = AiBlockRow
           vaSpread1.Col = vaSpread1.MaxCols
           iSer = Val(vaSpread1.text)
           vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1)))
           vaSpread1.Col = i
           L = 0
           nrodia = Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2))
           For j = 0 To UBound(vecdia)
               If nrodia = vecdia(j) Then
                  vaSpread1.Row = AiBlockRow: vaSpread1.Col = i - 1
                  If Trim(vaSpread1.text) <> "" Then
                     For X = AiBlockRow + 1 To vaSpread1.MaxRows
                         vaSpread1.Row = X: vaSpread1.Col = vaSpread1.MaxCols: xSer = Val(vaSpread1.text)
                         vaSpread1.Col = i + 1
                         If vaSpread1.Row = vaSpread1.MaxRows Then vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows X, 1: L = X: Exit For
                         If xSer <> iSer And xSer > 0 Then
                            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows X, 1: L = X: Exit For
                         ElseIf Trim(vaSpread1.text) <> "" And xSer > 0 Then
                            vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.InsertRows X + 1, 1: X = X + 1: L = X: Exit For
                         ElseIf Trim(vaSpread1.text) = "" Then
                            Exit For
                         End If
                     Next X
                     vaSpread1.CopyRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i - 1, X
                     vaSpread1.Row = X
                  Else
                     vaSpread1.CopyRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i - 1, AiBlockRow
                     vaSpread1.Row = AiBlockRow
                  End If
                  '-------> Asignar colores
                  For X = (i - 1) To (i - 1) + 4
                      
                      vaSpread1.Col = X
                      vaSpread1.BackColor = Shape1(0).FillColor
                      
                      For XX = 1 To MaxColumna
                          
                          If (VectorCol(XX) - 1) = vaSpread1.Col Then
                              vaSpread1.Col = X + 2
                              vaSpread1.CellType = CellTypeNumber
                              vaSpread1.TypeNumberDecPlaces = 0
'                              vaSpread1.TypeIntegerMin = 1
'                              vaSpread1.TypeIntegerMax = 9999999
                              vaSpread1.TypeNumberMin = 0
                              vaSpread1.TypeNumberMax = 9999999
                              vaSpread1.TypeHAlign = TypeHAlignRight
                              vaSpread1.TypeSpin = False
                              vaSpread1.TypeIntegerSpinInc = 1
                              vaSpread1.TypeIntegerSpinWrap = False
                              Exit For
                          End If
                      
                      Next XX
                      
                      vaSpread1.Col = X
                      
                      If X = (i - 1) Then
                         
                         vaSpread1.text = IIf(Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1), "A", "R")
                         vaSpread1.ForeColor = &HFF&
                         vaSpread1.BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
                      
                      End If
                  
                  Next X
                  If L > 0 Then
                     z = L
                     For L = 3 To (vaSpread1.MaxCols - MaxColumna) Step 5
                         
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
                  
                  '-------> Validar días modificados
                  For z = AiBlockRow To vaSpread1.ActiveRow + (AiBlockRow2 - AiBlockRow)
                      vaSpread1.Row = z
                      vaSpread1.Col = i ' + 1
                      If Trim(vaSpread1.text) <> "" And ((MaxColumna * 5 + 1) + ((i + 2) / 5)) < vaSpread1.MaxCols Then
                         vaSpread1.Col = (MaxColumna * 5 + 1) + ((i + 2) / 5)
                         vaSpread1.text = 1
                      End If
                  Next z
                  '-------> Fin validar días modificados
                  Exit For
               End If
           Next j
       Next i
       indgri = True
    Else
       indrow3 = vaSpread1.MaxRows
       vaSpread1.GetSelection 1, xColIni, xRowIni, xColFin, xRowFin
       For i = IblockCol To iblockcol2 Step 5
           If IndCortarPegar = 1 Then
              vaSpread1.Row = AiBlockRow: vaSpread1.Col = AiBlockCol
'              If vaSpread1.BackColor = Shape1(1).FillColor Then
                 vaSpread1.MaxRows = vaSpread1.MaxRows + (AiBlockRow2 - AiBlockRow) + 1
                 vaSpread1.CopyRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow)
                 '-------> Asignar colores
                 For j = vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow) To vaSpread1.MaxRows
                     vaSpread1.Row = j
                     For X = (i) To i + (AiBlockCol2 - AiBlockCol) Step 5 '(i) + 4
                         vaSpread1.Col = X + 1
                         vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
                         vaSpread1.Col = X + 2
                         vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
                         vaSpread1.Col = X + 3
                         vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
                         
                         
                         If AiBlockRow <> MaxFila Then
                            vaSpread1.BackColor = Shape1(0).FillColor
                         End If
                         If TipoCopia = "Copiar Raciones" Then
                             vaSpread1.CellType = CellTypeNumber
                         End If
                         For XX = 1 To MaxColumna
                             If (VectorCol(XX) - 1) = vaSpread1.Col Then
                                vaSpread1.Col = X + 2
                                vaSpread1.CellType = CellTypeNumber
                                vaSpread1.TypeNumberDecPlaces = 0
                                vaSpread1.TypeNumberMin = 0
                                vaSpread1.TypeNumberMax = 9999999
                                vaSpread1.TypeHAlign = TypeHAlignRight
                                vaSpread1.TypeSpin = False
                                vaSpread1.TypeIntegerSpinInc = 1
                                vaSpread1.TypeIntegerSpinWrap = False
                                Exit For
                             End If
                         Next XX
                         vaSpread1.Col = X
                         
'                         If X = (i) And Trim(vaSpread1.text) <> "" And TipoCopia <> "Copiar Raciones" Then
                         If (Trim(vaSpread1.text) = "R" Or Trim(vaSpread1.text) = "A") And Trim(vaSpread1.text) <> "" And TipoCopia <> "Copiar Raciones" Then
                            
                            vaSpread1.text = IIf(Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1), "A", "R")
                            vaSpread1.ForeColor = &HFF&
                            vaSpread1.BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
                         
                         End If
                     
                     Next X
                 Next j
                 '-------> Fin asignar colores
              If TipoCopia = "Copiar Raciones" Then
                 vaSpread1.CopyRange i, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow), i, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
              Else
'                 vaSpread1.CopyRange IblockCol, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
'                 a = IblockCol
                 If (AiBlockCol2 - AiBlockCol) = 4 Then
                    vaSpread1.CopyRange i, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow), i + 4, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 Else
                    vaSpread1.CopyRange i, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow), i + (AiBlockCol2 - AiBlockCol), vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 End If
              End If
              vaSpread1.MaxRows = indrow3
           ElseIf IndCortarPegar = 0 Then
              vaSpread1.Row = AiBlockRow: vaSpread1.Col = AiBlockCol
              If vaSpread1.BackColor = Shape1(1).FillColor Then
                 vaSpread1.MaxRows = vaSpread1.MaxRows + (AiBlockRow2 - AiBlockRow) + 1
                 vaSpread1.MoveRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow)
                 '-------> Asignar colores
                 For j = vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow) To vaSpread1.MaxRows
                     vaSpread1.Row = j
                     For X = (i) To (i) + 4
                         vaSpread1.Col = X
                         vaSpread1.BackColor = Shape1(0).FillColor
                         For XX = 1 To MaxColumna
                             If (VectorCol(XX) - 1) = vaSpread1.Col Then
                                vaSpread1.Col = X + 2
                                vaSpread1.CellType = CellTypeNumber
                                vaSpread1.TypeNumberDecPlaces = 0
'                                vaSpread1.TypeIntegerMin = 1
'                                vaSpread1.TypeIntegerMax = 9999999
                                vaSpread1.TypeNumberMin = 0
                                vaSpread1.TypeNumberMax = 9999999
                                vaSpread1.TypeHAlign = TypeHAlignRight
                                vaSpread1.TypeSpin = False
                                vaSpread1.TypeIntegerSpinInc = 1
                                vaSpread1.TypeIntegerSpinWrap = False
                                Exit For
                             End If
                         Next XX
                         vaSpread1.Col = X
                         If X = (i) And Trim(vaSpread1.text) <> "" Then
                            
                            vaSpread1.text = IIf(Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1), "A", "R")
                            vaSpread1.ForeColor = &HFF&
                            vaSpread1.BackColor = IIf(Not etapa5, &H80FF80, &HFFFF00)
                         
                         End If
                     Next X
                 Next j
                 '-------> Fin asignar colores
                 vaSpread1.MoveRange IblockCol, vaSpread1.MaxRows - (AiBlockRow2 - AiBlockRow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                 vaSpread1.MaxRows = indrow3
              Else
                 vaSpread1.MoveRange AiBlockCol, AiBlockRow, AiBlockCol2, AiBlockRow2, i, vaSpread1.ActiveRow
              End If
           End If
           For j = vaSpread1.ActiveRow To vaSpread1.ActiveRow + (AiBlockRow2 - AiBlockRow)
               vaSpread1.Row = j
               For X = AiBlockCol To AiBlockCol2 Step 5
                   vaSpread1.Col = X + 1
                   If Trim(vaSpread1.text) <> "" And ((MaxColumna * 5 + 1) + ((X + 3) / 5)) < vaSpread1.MaxCols Then
                      vaSpread1.Col = (MaxColumna * 5 + 1) + ((X + 3) / 5)
                      vaSpread1.text = 1
                   End If
               Next X
           Next j
           '-------> Fin validar días modificados
       Next i
    End If
    
    '------> Se trabaja como excel las raciones
    ColumnaActiva = vaSpread1.ActiveCol: FilaActiva = vaSpread1.ActiveRow: ColumnaAntActiva = ColumnaActiva - 1
    vaSpread1.Col = ColumnaActiva: vaSpread1.Row = 0
    
    If IndCos = True Then
       For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
    End If
    For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
        CalctodiaEnc 1, i + 1
    Next i
    AiBlockCol = indcol: iblockcol2 = IndCol2
    AiBlockRow = indrow: AiBlockRow2 = IndRow2
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    
Case 15
    B_BusVas.Partidas Me
    B_BusVas.Show 1
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Man_Error

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
'    If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
'        MsgBox "Minuta, sin acceso a modificaciones", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
'    End If
    Plantilla_Click (8)
Case 21 '--------> Aporte nutricional x día
    Plantilla_Click (10)
Case 22 '-------> Costo planificación minutas
    Plantilla_Click (11)
Case 23 '-------> Frecuencia
    Plantilla_Click (12)
Case 24 '-------> Actualizar costo recetas y planificación
    Plantilla_Click (13)
Case 25 '-------> Exportar recetas
    Plantilla_Click (14)
Case 27
    Plantilla_Click (20)
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
On Error GoTo Man_Error

indactivo = 1
IblockRow = BlockRow
IblockRow2 = BlockRow2
IblockCol = BlockCol
iblockcol2 = BlockCol2
If BlockRow < 0 Then IblockRow = 1
'jpaz If BlockRow2 < 0 Then iblockrow2 = 100
'jpaz If BlockRow2 > 100 Then iblockrow2 = 100
If BlockRow2 < 0 Then IblockRow2 = (vaSpread1.MaxRows - 1)
If BlockRow2 >= vaSpread1.MaxRows Then IblockRow2 = (vaSpread1.MaxRows - 1)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
On Error GoTo Man_Error

If Row < 1 Then Exit Sub
OpGrilla(15).Enabled = IIf(Col = 1 And (vg_codregimen > 9999 And etapa5 And AddReceta = 0), False, True)
Plato(15).Enabled = IIf(Col = 1 And (vg_codregimen > 9999 And etapa5 And AddReceta = 0), False, True)
'-------> bloquear opción ingreso receta
If (vg_codregimen > 9999 And etapa5) And AddReceta = 0 Then
   Plantilla(8).Enabled = False
   For i = 0 To 15
       If Plato(i).Visible = True And Plato(i).Caption <> "-" Then Plato(i).Enabled = False
       If OpGrilla(i).Visible = True And OpGrilla(i).Caption <> "-" Then OpGrilla(i).Enabled = False
   Next i
   Exit Sub
End If
indactivo = 1
IblockRow = vaSpread1.ActiveRow
IblockRow2 = vaSpread1.ActiveRow
IblockCol = vaSpread1.ActiveCol
iblockcol2 = vaSpread1.ActiveCol
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.ActiveCol

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
On Error GoTo Man_Error

If Row < 1 Or Col = 1 Then Exit Sub
Plato_Click (2)

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
On Error GoTo Man_Error

If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
vaSpread1.Row = Row: vaSpread1.Col = Col
If vaSpread1.BackColor = Shape1(1).FillColor Then MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
If Mode = 1 And Not ChangeMade Then
   numracionesanterior = Val(vaSpread1.text)
End If
If vaSpread1.ChangeMade = False Or Col = 1 Or Mode = 1 Then i = vaSpread1.text: Exit Sub
''-------> Inicio MVI JPAZ 20130226
'If Mode = 0 And ChangeMade Then
'   If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) And Not ExisteDatMinuta Then
'      Let vaSpread1.Row = Row
'      Let vaSpread1.Col = Col
'      Let vaSpread1.text = numracionesanterior
'      MsgBox "Minuta, sin acceso a modificaciones", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
'   End If
'End If
''-------> Fin MVI JPAZ 20130226
If vaSpread1.ChangeMade = True Then
   vaSpread1.Col = (MaxColumna * 5 + 1) + (vaSpread1.Col / 5): vaSpread1.text = 1
   If IndCos = True Then
      vaSpread1.Col = Col: j = Col - 1:  Calctodia vaSpread1.Row, j
   End If
   vaSpread1.Col = Col: j = Col - 1: CalctodiaEnc vaSpread1.Row, j
End If
vaSpread1.Row = Row
Plantilla(0).Enabled = True
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Man_Error

If Toolbar1.Buttons(2).Enabled = False Or (vg_codregimen > 9999 And etapa5 And AddReceta = 0) Then Exit Sub
Dim DelRow As Integer, indcol As Integer, indrow As Integer, IndCol2 As Integer, IndRow2 As Integer
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
    'If Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1) Then
    '
    '    MsgBox "Minuta, sin acceso a modificaciones", vbCritical + vbOKOnly, Msgtitulo: Exit Sub
    'End If
    
    If vaSpread1.MaxRows = vaSpread1.ActiveRow Or vaSpread1.MaxRows = IblockRow Or vaSpread1.MaxRows = IblockRow2 Then
       Exit Sub
    End If
    
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.BackColor = Shape1(1).FillColor Then
       MsgBox "Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo
       Exit Sub
    End If
    
    j = 0
    For i = 1 To MaxColumna
        If (VectorCol(i) - 1) = vaSpread1.Col Or VectorCol(i) = vaSpread1.Col Then
           j = (VectorCol(i) - 1)
           Exit For
        End If
    Next i
    
    If j = 0 Then Exit Sub
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    If indactivo = 0 Or IblockCol < 1 Or IblockRow < 1 Then
       IblockCol = vaSpread1.ActiveCol
       iblockcol2 = vaSpread1.ActiveCol
       IblockRow = vaSpread1.ActiveRow
       IblockRow2 = vaSpread1.ActiveRow
    End If
    
    AiBlockCol = IblockCol
    AiBlockRow = IblockRow
    AiBlockCol2 = iblockcol2
    AiBlockRow2 = IblockRow2
    If IblockCol < 0 Then
       IblockCol = 2
       iblockcol2 = vaSpread1.MaxCols
    End If
    
    AiBlockCol = IblockCol
    AiBlockRow = IblockRow
    AiBlockCol2 = iblockcol2
    AiBlockRow2 = IblockRow2
    For i = 1 To MaxColumna
        If (VectorCol(i) - 1) = IblockCol Or VectorCol(i) = IblockCol Then IblockCol = (VectorCol(i) - 1): Exit For
    Next i
    For i = 1 To MaxColumna
        If (VectorCol(i) - 1) = iblockcol2 Then iblockcol2 = ((VectorCol(i) + 3)): Exit For
        If VectorCol(i) = iblockcol2 Then iblockcol2 = (VectorCol(i) + 3): Exit For
    Next i
    indcol = AiBlockCol: IndCol2 = iblockcol2
    indrow = AiBlockRow: IndRow2 = AiBlockRow2
    
    '-------> Validar recetas 5 etapas
    If etapa5 And AddReceta > 0 Then
        For j = IblockCol To iblockcol2 Step 5
           vaSpread1.Col = j
           For i = IblockRow To IblockRow2
              vaSpread1.Row = i
              If vaSpread1.BackColor = &H80FF80 Then
                 MsgBox "No puede eliminar receta, corresponde centralizado", vbCritical + vbOKOnly, MsgTitulo
                 Exit Sub
              End If
          Next i
        Next j
    End If
    
    vaSpread1.ClearRange IblockCol, IblockRow, iblockcol2, IblockRow2, False
    If IndCos = True Then
       For i = IblockCol To iblockcol2 Step 5
           Calctodia 1, i + 1
       Next i
       MostrarCosto vaSpread1.ActiveCol
    End If
    
    For i = IblockCol To iblockcol2 Step 5
        CalctodiaEnc 1, i + 1
    Next i
    IblockCol = AuxCol
    vaSpread1.BlockMode = False
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    indactivo = 0
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
IblockRow = NewRow
IblockRow2 = NewRow
IblockCol = NewCol
iblockcol2 = NewCol
If NewRow < 0 Then IblockRow = 1
If NewRow < 0 Then IblockRow2 = (vaSpread1.MaxRows - 1)
If NewRow >= vaSpread1.MaxRows Then IblockRow2 = (vaSpread1.MaxRows - 1)
If IndCos = False Or NewCol < 1 Then Exit Sub
MostrarCosto NewCol

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error GoTo Man_Error

Select Case Button
Case 2
    If vaSpread1.Visible <> True Then Exit Sub
    'Indvaspread1 = 0
    PopupMenu MenuDetalle
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub Opgrilla_Click(Index As Integer)
On Error GoTo Man_Error

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
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Private Sub GrabarPlantillaMinuta()

Dim RS              As New ADODB.Recordset
Dim desc            As String
Dim StrRec          As String
Dim StrRecb         As String
Dim CodRec          As Long
Dim NumRac          As Long
Dim estser          As Long
Dim Fecha           As Long
Dim conregdet       As Long
Dim indice          As Long
Dim ExisteDat       As Long
Dim inddia          As Long
Dim tiprec          As Long
Dim fechasis        As Long
Dim fecini          As Long
Dim fecfin          As Long
Dim totrac          As Long
Dim fecinm          As Long
Dim cosali          As Double
Dim CosDes          As Double
Dim cospro          As Double
Dim rec5eta         As String
Dim sql2            As String
Dim aAp             As String
Dim estgra          As Boolean
Dim indcosto        As Boolean
Dim MensajeCosto    As Variant
Dim MensajeRaciones As Variant
Dim CostoMinutaDia  As Double
Dim CostoBandeja    As Double
Dim FecMinuta       As Long
Dim MyBuffer        As String
Dim color           As String
Dim EstGraba        As Boolean

On Error GoTo Man_Error

fg_carga ""

'-------> Grabar Estructura Servicio
vaSpread1.Enabled = False
Toolbar1.Enabled = False
Main(0).Enabled = False
Main(1).Enabled = False
'-------> Validar si existe costo techo > costo minuta
'-------> cargar costo
indcosto = False
If Frame2(0).Visible = False Then CargarCosto

If Val(Label1(8).Caption) > 0 Then
   
   CostoBandeja = Label1(8).Caption
   
   If vCtoTec > 0 And CostoBandeja > 0 And CostoBandeja > (((5 / 100) * vCtoTec) + vCtoTec) Then
      
      MensajeCosto = "Costo minuta día (" & Format(CostoBandeja, fg_Pict(6, 2)) & ") es mayor costo techo (" & Format((((5 / 100) * vCtoTec) + vCtoTec), fg_Pict(6, 2)) & "), verificar los siguientes días : " & Chr(13) & Chr(13)
      
      For i = 2 To (vaSpread1.MaxCols - 1) Step 5
          
          DoEvents
          vaSpread1.Row = SpreadHeader + IIf(vCtoTec > 0, 1, 0)
          vaSpread1.Col = i
          CostoMinutaDia = IIf(vaSpread1.text = "", 0, IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0))
          
          If vCtoTec > 0 And CostoMinutaDia > 0 And CostoMinutaDia > vCtoTec Then
             
             indcosto = True
             MensajeCosto = MensajeCosto & "Día " & fg_pone_cero(inddia, 2) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4) & " : " & Format(CostoMinutaDia, fg_Pict(6, 2)) & Chr(13)
          
          End If
          
          inddia = inddia + 1
      
      Next
      
      If indcosto Then MsgBox MensajeCosto, vbInformation, MsgTitulo
   
   End If

End If
   
'-------> validar si existe numero raciones mayores que total comensales x día
'-------> Traer estructuras seleccionada y mueve vector
Dim VectorEstructura() As Variant

j = 1
For i = 1 To vaSpread1.MaxRows - 1
    
    DoEvents
    vaSpread1.Row = i
    vaSpread1.Col = vaSpread1.MaxCols
    
    If Trim(vaSpread1.text) <> "" Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open "select ess_codigo, ess_marcaplatos from a_estservicio where ess_cencos = '" & MuestraCasino(1) & "' and ess_codigo = " & Val(vaSpread1.text) & " and ess_marcaplatos = '1'", vg_db, adOpenStatic
       
       If Not RS.EOF Then
          
          ReDim Preserve VectorEstructura(j)
          VectorEstructura(j) = RS!ess_codigo
          j = j + 1
       
       End If
       
       RS.Close: Set RS = Nothing
    
    End If

Next i

Dim TotalComensales As Double
Dim RacionesProteicosDia As Double
Dim CodigoEstructura As Long
   
'-------> Validar cantidad proteicos si supera total comensales dìa
If j > 1 Then
  
  indcosto = False
  inddia = 1
  MensajeRaciones = "La cantidad de raciones proteicos supera el total comensales día, verificar los siguientes días : " & Chr(13) & Chr(13)
  
  For i = 2 To (vaSpread1.MaxCols - 1) Step 5
      
      DoEvents
      vaSpread1.Row = vaSpread1.MaxRows
      vaSpread1.Col = i + 2
      TotalComensales = IIf(vaSpread1.text = "", 0, IIf(Val(vaSpread1.text) > 0, vaSpread1.text, 0))
      RacionesProteicosDia = 0
      
      For j = 1 To vaSpread1.MaxRows - 1
          
          vaSpread1.Row = j
          vaSpread1.Col = vaSpread1.MaxCols
          
          If Trim(vaSpread1.text) <> "" Then
             
             CodigoEstructura = vaSpread1.text
             '-------> Validar si estructura esta seleccionada como plato proteicos
             
             For X = 1 To UBound(VectorEstructura)
                 
                 If VectorEstructura(X) = CodigoEstructura Then
                    
                    vaSpread1.Col = i + 2: RacionesProteicosDia = RacionesProteicosDia + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
                    Exit For
                 
                 End If
             
             Next X
          
          End If
      
      Next j
      
      If TotalComensales > 0 And RacionesProteicosDia > 0 And RacionesProteicosDia > TotalComensales Then
         
         indcosto = True
         MensajeRaciones = MensajeRaciones & "Día " & fg_pone_cero(inddia, 2) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4) & " : " & RacionesProteicosDia & "   Comensales total dìa : " & TotalComensales & Chr(13)
      
      End If
      inddia = inddia + 1
  
  Next i
  
  If indcosto Then MsgBox MensajeRaciones, vbInformation, MsgTitulo

End If

'-------> Grabar datos
inddia = 1
conregdet = 0
gauge1.Value = 0
gauge.Value = 0
Fecha = 0
fecini = 0
fecfin = 0
fecinm = 0
fechasis = 0
fechasis = Val(Mid(Date, 7, 4)) & Mid(Date, 4, 2) & Mid(Date, 1, 2)
estgra = False

Picture1.Visible = True
Label3.Visible = True
gauge.Visible = True
Picture1.Refresh
Label3.Refresh
gauge.Refresh
gauge1.Refresh

For i = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
    
    DoEvents
    
    gauge1.Value = Val((inddia / MaxColumna) * 100)
    Label3.Caption = "": Label3.Caption = "Día : " & inddia
    FecMinuta = Val(vg_fecha) & fg_pone_cero(inddia, 2)
    ExisteDat = 0
    vaSpread1.Row = 1
    vaSpread1.Col = i
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = i + 2
    totrac = Val(vaSpread1.text)
    
    EstGraba = False
    
    '-------> Grabar raciones en minutas raciones producidas
    vaSpread1.Row = 1
    '-------> Validar si minuta Esta cerrada y mover estado color 0 = Deshabilitada - 1 = Habilitada
    color = "0"
    If vaSpread1.BackColor <> Shape1(1).FillColor Then
       
       color = "1"
       
    End If
    
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<GrabaMinuta>"

    '-------> Mover total raciones del día
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = i + 2
    totrac = Val(vaSpread1.text)
    
    gauge.Value = 0
    conregdet = 0
    estser = 0
       
       '-------> Actualizar detalle minutas
       For j = 1 To (vaSpread1.MaxRows - 1)
           
           conregdet = conregdet + 1
           gauge.Value = Val((conregdet / (vaSpread1.MaxRows - 1)) * 100)
           desc = ""
           CodRec = 0
           cosali = 0
           CosDes = 0
           vaSpread1.Row = j
           vaSpread1.Col = vaSpread1.MaxCols
           If Trim(vaSpread1.text) <> "" Then estser = vaSpread1.text
           
           vaSpread1.Col = i
           rec5eta = IIf(Not etapa5, "0", IIf(vaSpread1.BackColor = &HFFFF00, "0", "1"))
           
           vaSpread1.Col = i + 1
           desc = Trim(Mid(vaSpread1.text, 1, 50))
           
           If desc <> "" And estser > 0 Then
              
              vaSpread1.Col = i + 2
              
              NumRac = vaSpread1.text
              
              vaSpread1.Col = i + 3 ': cosali = vaSpread1.text
              
              vaSpread1.Col = i + 4
              
              StrRec = vaSpread1.text
              
              If Len(StrRec) <> 0 Then
                 
                 Do While InStr(StrRec, ";") <> 0
                 
                    StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                    StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                    CodRec = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                    tiprec = Val(Mid(StrRecb, 1))
                 
                 Loop
              
              End If
              
              cosali = fg_CalCtoRecInv(CodRec, tiprec, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")))
              CosDes = fg_CalCtoRecInv(CodRec, tiprec, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")))
              MyBuffer = MyBuffer & " <Minuta"
              MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)

              desc = Replace(Trim(desc), Chr(34), "&quot;")
              desc = Replace(Trim(desc), Chr(38), "&amp;")
              desc = Replace(Trim(desc), Chr(39), "&apos;")
              desc = Replace(Trim(desc), Chr(60), "&lt;")
              desc = Replace(Trim(desc), Chr(62), "&gt;")
                                
              MyBuffer = MyBuffer & " NumRacion = " & Chr(34) & NumRac & Chr(34)
              MyBuffer = MyBuffer & " DescReceta = " & Chr(34) & desc & Chr(34)
              MyBuffer = MyBuffer & " CodEstructura = " & Chr(34) & estser & Chr(34)
              MyBuffer = MyBuffer & " TipoReceta = " & Chr(34) & tiprec & Chr(34)
              MyBuffer = MyBuffer & " Rec5eta = " & Chr(34) & rec5eta & Chr(34)
              MyBuffer = MyBuffer & " NumLin = " & Chr(34) & j & Chr(34)
              MyBuffer = MyBuffer & " CodReceta = " & Chr(34) & CodRec & Chr(34)
              MyBuffer = MyBuffer & " FecVal = " & Chr(34) & 0 & Chr(34)
              MyBuffer = MyBuffer & " CosAli = " & Chr(34) & cosali & Chr(34)
              MyBuffer = MyBuffer & " CosDes = " & Chr(34) & CosDes & Chr(34)
              MyBuffer = MyBuffer & "/>"
              
              EstGraba = True
              
              If fecini < Fecha And fecini = 0 Then
              
                 fecini = Fecha
                 
              End If
           
           End If
       
       Next j

    inddia = inddia + 1
    
'    If EstGraba Then
       
       MyBuffer = MyBuffer & "</GrabaMinuta>"
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("sgp_Ins_XmlMinutaReal '" & MyBuffer & "', '" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & FecMinuta & ", " & totrac & ", '" & color & "', '" & vg_NUsr & "'")
    
       If Not RS.EOF Then
       
          If RS(0) > 0 Then
          
             MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
       
          ElseIf Trim(RS(1)) <> "" Then
          
             MsgBox RS(1), vbCritical + vbOKOnly, MsgTitulo
          
          End If
    
       End If
    
       RS.Close
       Set RS = Nothing

'   End If
   
Next i

fecfin = Fecha
If fecini > 0 Then
   '-------> Buscar por rango de fecha los producto incluido en mes y luego eliminar y grabar minuta costo
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   RS.Open "SELECT DISTINCT e.ing_codigo, f.cpi_precos " & _
           "FROM  b_receta a, b_recetadet b, b_minuta c, b_minutadet d, b_ingrediente e, b_contlistpreing f " & _
           "WHERE c.min_codigo = d.mid_codigo " & _
           "AND   d.mid_codrec = b.red_codigo " & _
           "AND   d.mid_tiprec = b.red_tiprec AND ((b.red_tiprec <> 0 AND b.red_cencos = '" & MuestraCasino(1) & "') OR (b.red_tiprec = 0 AND b.red_cencos = '0')) " & _
           "AND   b.red_codigo = a.rec_codigo " & _
           "AND   b.red_codpro = e.ing_codigo " & _
           "AND   c.min_cencos = '" & vg_codcasino & "' " & _
           "AND   c.min_fecmin >= " & fecini & " " & _
           "AND   c.min_fecmin <= " & fecfin & " " & _
           "AND   d.mid_tipmin = '2' AND f.cpi_coding = e.ing_codigo AND f.cpi_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
   
   If Not RS.EOF Then
      
      Do While Not RS.EOF
         
         DoEvents
         
         vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto " & _
                       "WHERE mic_cencos = '" & MuestraCasino(1) & "' AND mic_fecval = " & fechasis & " " & _
                       "AND   mic_tipmin = '2' " & _
                       "AND   mic_codpro = '" & RS!ing_codigo & "'"
         
         vg_db.Execute "INSERT INTO b_minutacosto(mic_cencos, mic_fecval, mic_tipmin, mic_codpro, mic_cospro) " & _
                       "VALUES ('" & MuestraCasino(1) & "', " & fechasis & ", '2', '" & RS!ing_codigo & "', " & RS!cpi_precos & ")"
         
         RS.MoveNext
      
      Loop
   
   End If
   
   RS.Close: Set RS = Nothing
   
   Dim vecrec As Variant
End If
'-------> Grabar recetas que fueron modificada solo en un sola minuta y que no se actualizo en las demas minutas,
'-------> Solo va grabar las recetas que pasaron de a local y todavia estan como patron
vg_db.Execute "UPDATE b_minutadet SET mid_tiprec = -1 " & _
              "WHERE  mid_codigo IN (SELECT DISTINCT a.min_codigo FROM b_minuta a WHERE a.min_cencos = '" & vg_codcasino & "' AND a.min_fecmin >= " & fecinm & " AND a.min_fecmin <= " & fecfin & ") " & _
              "AND    mid_codrec IN (SELECT DISTINCT b.mid_codrec FROM b_minuta a, b_minutadet b WHERE a.min_codigo = b.mid_codigo AND a.min_cencos = '" & vg_codcasino & "' AND a.min_fecmin >= " & fecinm & " AND a.min_fecmin <= " & fecfin & " AND b.mid_tipmin = '2' AND b.mid_tiprec = -1) " & _
              "AND    mid_tiprec = 0 " & _
              "AND    mid_tipmin = '2'"

Picture1.Visible = False: gauge.Visible = False
Toolbar1.Enabled = True
Main(0).Enabled = True
Main(1).Enabled = True
vaSpread1.Enabled = True
vaSpread1.Refresh
indgri = False
fg_descarga

Exit Sub
Man_Error:

Picture1.Visible = False: gauge.Visible = False
vaSpread1.Enabled = True
Main(0).Enabled = True
Main(1).Enabled = True
Toolbar1.Enabled = True
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Sub DetallePlantillaMinuta()

On Error GoTo Man_Error

Dim RS      As New ADODB.Recordset
Dim indrow3 As Long
Dim inddia  As Long
Dim Fecha   As String
Dim sql1    As String
Dim cosali  As Double
Dim CosDes  As Double
Dim fecesf  As Long
Dim estfij  As Boolean
Dim fil     As Long
Dim Col     As Long
'Dim fecesf  As Long
'Dim estfij  As Boolean
Dim aAp     As String
Dim vTotRac As Long
Dim vCosVec As Double

fg_carga ""
SwSalir = 0: MaxColumna = 0: indactivo = 0: vCtoPis = 0: vCtoTec = 0
IblockRow = 0: IblockRow2 = 0: IblockCol = 0: iblockcol2 = 0: SwSalir = 0
AiBlockRow = 0: AiBlockRow2 = 0: AiBlockCol = 0: AiBlockCol2 = 0
etapa5 = IIf("S" = fg_CambiaChar(GetParametro("5etapas"), ";", "','") And vg_codregimen > 9999, True, False)

'-------> Formatear columna
MaxColumna = Val(fg_mes(Mid(vg_fecha, 5, 2) & Mid(vg_fecha, 1, 4)))

'-------> Defenir vector costo encabezado
ReDim VecCosenc(MaxColumna, 2)
For i = 1 To UBound(VecCosenc)
    VecCosenc(i, 1) = 0
    VecCosenc(i, 2) = 0
Next i

'-------> Cargar adicional receta 5 etapas
AddReceta = 0
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open "SELECT par_valor, par_codigo FROM a_param WHERE par_cencos = '" & MuestraCasino(1) & "' AND par_codigo = 'addreceta'", vg_db, adOpenStatic
If Not RS.EOF Then AddReceta = RS!par_valor
RS.Close: Set RS = Nothing

'-------> Validar si existe costo patron
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open RutinaLectura.CostoPatron(1, vg_codregimen, vg_codservicio, Val(vg_fecha)), vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      If Trim(RS!cpa_descripcion) = "PISO" Then vCtoPis = IIf(IsNull(RS!cpa_valor), 0, RS!cpa_valor)
      If Trim(RS!cpa_descripcion) = "TECHO" Then vCtoTec = IIf(IsNull(RS!cpa_valor), 0, RS!cpa_valor)
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing

EstMinBlo = IIf(Not ValidarAccesoMinutaBloqueyBloqueo(vg_codcasino, 1), "A", "R")

With vaSpread1
    .MaxRows = 1000
    .MaxCols = 0: .MaxCols = 5 * MaxColumna + 1: .Row = 0
    '-------> turn off display of row headers
    '.RowHeadersShow = False
    '-------> Set up column headers
    .ColHeaderRows = IIf(vCtoPis > 0 And vCtoTec > 0, 4, IIf(vCtoPis > 0 And vCtoTec = 0, 3, IIf(vCtoPis = 0 And vCtoTec > 0, 3, 2)))
    '.ShadowColor = &HFFC0C0
    .ShadowColor = &H8000000F
    .ShadowText = &H800000
    
    For i = 2 To .MaxCols Step 5
       If vCtoTec > 0 Then
          .AddCellSpan i, SpreadHeader, 5, 1
          .Col = i
          .Row = SpreadHeader
          .TypeHAlign = TypeHAlignRight
          .text = Format(vCtoTec, fg_Pict(6, 2))
       End If
       .AddCellSpan i, SpreadHeader + IIf(vCtoTec > 0, 1, 0), 5, 1
       If vCtoPis > 0 Then
          .AddCellSpan i, SpreadHeader + IIf(vCtoTec = 0, 1, 2), 5, 1
          .Col = i
          .Row = SpreadHeader + IIf(vCtoTec = 0, 1, 2)
          .TypeHAlign = TypeHAlignRight
          .text = Format(vCtoPis, fg_Pict(6, 2))
       End If
    Next i
    .Col = 1
    .ColsFrozen = 1
    .VisibleCols = 1
    .ColWidth(1) = 15
    If vCtoTec > 0 Then .Row = SpreadHeader: .TypeHAlign = TypeHAlignLeft: .text = "Costo Patrón Techo"
    .Row = SpreadHeader + IIf(vCtoTec > 0, 1, 0)
    .TypeHAlign = TypeHAlignLeft
    .text = "Costo Minuta Día"
    If vCtoPis > 0 Then .Row = SpreadHeader + IIf(vCtoTec > 0, 2, 1): .TypeHAlign = TypeHAlignLeft: .text = "Costo Patrón Piso"
    .Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoTec < 1 And vCtoPis > 0, 2, IIf(vCtoTec > 0 And vCtoPis = 0, 2, 1)))
    .TypeHAlign = TypeHAlignLeft
    .text = "Estructura Servicio"
    
    ReDim Preserve VectorCol(0)
    For i = 2 To .MaxCols Step 5
        .Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoTec < 1 And vCtoPis > 0, 2, IIf(vCtoTec > 0 And vCtoPis = 0, 2, 1)))
        .Col = i
        .ColWidth(i) = 1.5
        .text = " "
        .ColHidden = False
        
        .Col = i + 1
        .ColWidth(i + 1) = 21
        If i = 2 Then
           ReDim Preserve VectorCol(1)
           VectorCol(1) = 3
           .text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & (i - 1), 2), 1), 1, 3) & " " & (i - 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
        Else
           .text = " " & Mid(fg_Fecha_Dia(vg_fecha & Right("0" & CLng((i / 5) + 1), 2), 1), 1, 3) & " " & CLng((i / 5) + 1) & "/" & Mid(vg_fecha, 5, 2) & "/" & Mid(vg_fecha, 1, 4)
           ReDim Preserve VectorCol(CLng((i / 5) + 1))
           VectorCol(CLng((i / 5) + 1)) = i + 1
        End If
        .ColHidden = False
        
        .Col = i + 2
        .ColWidth(i + 2) = 6
        .text = "N.Rac."
        .ColHidden = False
       
        .Col = i + 3
        .ColWidth(i + 3) = 9
        .text = "Costo"
        .ColHidden = False
        
        .Col = i + 4
        .text = "Cod. Receta"
        .ColHidden = True
        For j = 1 To .MaxRows
            .Row = j
    
            .Col = i
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = ""
    
            .Col = i + 1
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 2
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 3
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 4
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
            .Col = i + 5
            .CellType = CellTypeDate '= 1
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
        Next j
    Next i
    .Row = 0
    
    For i = 1 To MaxColumna
       .MaxCols = .MaxCols + 1
       .Col = .MaxCols
       .text = "Estado"
       .ColHidden = True
    Next i
    
    .MaxCols = .MaxCols + 1
    .Col = .MaxCols
    .ColWidth(.MaxCols) = 5
    .text = "Còd. Est."
    .ColHidden = True
    
    .Row = -1: .Col = -1: .BackColor = Shape1(0).FillColor  'Amarillo
    .Row = -1: .Col = 1
    .Font.Bold = True
    .Font.Size = 9
    .BackColor = Shape1(2).FillColor 'Verde
    
    j = 0: i = 0: indrow3 = 0: cosali = 0: CosDes = 0
    
    '-------> Cargar minutas
    ExisteDatMinuta = False
    sql1 = IIf(vg_tipbase = "1", " val(mid(b.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),b.min_fecmin),1,6)) ")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Minutas(5, vg_codregimen, vg_codservicio, Val(vg_fecha), "2"), vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          ExisteDatMinuta = True
          j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 5) - 5) + 1) + 1
          .Row = RS!mid_numlin
          
          If indrow3 < .Row Then
             indrow3 = .Row
          End If
          
          If RS!mid_estser <> i Then
             .Col = 1
             .text = IIf(IsNull(RS!ess_nombre), "No existe estructura servicio", RS!ess_nombre)
             
             .Col = .MaxCols
             .CellType = CellTypeStaticText
             .TypeHAlign = TypeHAlignCenter
             .text = IIf(IsNull(RS!mid_estser), 0, RS!mid_estser)
             i = RS!mid_estser
          End If
          
          .Col = j
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignCenter
          If EstMinBlo = "R" Then
             .Value = "R"
          Else
             .Value = IIf(RS!mid_rec5eta = "1", "R", EstMinBlo)
          End If
          
          .ForeColor = IIf(Not etapa5, &HFF&, IIf(RS!mid_rec5eta = "1" Or IsNull(RS!mid_rec5eta), &HFF&, &H400000))
          .BackColor = IIf(Not etapa5, &H80FF80, IIf(RS!mid_rec5eta = "0", &HFFFF00, &H80FF80))
               
          .Col = j + 1
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!rec_nombre), "No existe receta", Trim(RS!rec_nombre))
                             
          .Col = j + 2
          .CellType = CellTypeNumber
          .TypeNumberDecPlaces = 0
          .TypeNumberMin = 0
          .TypeNumberMax = 9999999
          .TypeHAlign = TypeHAlignRight
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
          .Value = IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac)
          .ForeColor = &HFF0000
                           
          '-------> Mover costo alimentación y desachable
          cosali = IIf(IsNull(RS!mid_cosrec), 0, RS!mid_cosrec)
          CosDes = IIf(IsNull(RS!mid_cosdes), 0, RS!mid_cosdes)
          .Col = j + 3
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignRight
          .text = Format((cosali + CosDes), fg_Pict(6, 2))
          .Col = j + 4: .text = RS!mid_codrec & "&" & RS!mid_tiprec & "&;"
          '-------> Mover costo minuta dia
          VecCosenc(Val(Mid(RS!min_fecmin, 7, 2)), 1) = (VecCosenc(Val(Mid(RS!min_fecmin, 7, 2)), 1) + ((cosali + CosDes) * IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac)))
          'If RS!min_indblo > 0 Then .Row = -1: .Col = j: .BackColor = Shape1(1).FillColor: .Col = j + 1: .BackColor = Shape1(1).FillColor: .Col = j + 2: .CellType = 5: .TypeHAlign = 1: .BackColor = Shape1(1).FillColor: .Col = j + 3: .BackColor = Shape1(1).FillColor
          
          RS.MoveNext
       
       Loop
       
       RS.Close
       Set RS = Nothing
       fg_descarga
    
    Else
       
       RS.Close
       Set RS = Nothing
       fg_descarga
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open RutinaLectura.EstServicio(1, vg_codservicio, 0), vg_db, adOpenStatic
       If RS.EOF Then RS.Close: Set RS = Nothing: Exit Sub
       Do While Not RS.EOF
          
          .Row = RS!ess_orden
          If indrow3 < .Row Then indrow3 = .Row
          .Col = 1
          .text = IIf(IsNull(RS!ess_nombre), "No existe estructura servicio", RS!ess_nombre)
          
          For i = 2 To .MaxCols Step 5
              
              .Col = .MaxCols
              .text = IIf(IsNull(RS!ess_codigo), 0, RS!ess_codigo)
          
          Next i
          
          RS.MoveNext
       
       Loop
       
       RS.Close
       Set RS = Nothing
    
    End If
    
    For i = 3 To (.MaxCols - MaxColumna) Step 5
        
        .Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoPis < 1 And vCtoTec > 0, 2, IIf(vCtoPis > 0 And vCtoTec = 0, 2, 1))): .Col = i
        
        If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Or CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < CDate(vg_ciedia) Then
           
           For fil = 1 To (.MaxRows - 1)
               
               For Col = i - 1 To i + 2
                   .Row = fil: .Col = Col
                   If .CellType = CellTypeNumber Then
                      .CellType = CellTypeStaticText
                      .TypeHAlign = TypeHAlignRight
                   End If
                   .BackColor = Shape1(1).FillColor
               
               Next Col
           
           Next fil
        
        End If
    
    Next i
    
    .MaxRows = indrow3 + 1
    .Row = .MaxRows
    MaxFila = .MaxRows
    .Col = 1
    .text = "Comensales"
    .Col = -1: .BackColor = &HE0E0E0
    
    '-------> Formatear ultima columna
    For i = 2 To (.MaxCols - MaxColumna) Step 5
        .Row = .MaxRows
        .Col = i + 2
        .CellType = CellTypeNumber
        .TypeNumberDecPlaces = 0
        .TypeNumberMin = 0
        .TypeNumberMax = 9999999
        .TypeHAlign = TypeHAlignRight
        .TypeSpin = False
        .TypeIntegerSpinInc = 1
        .TypeIntegerSpinWrap = False
        .Value = Format(0, fg_Pict(6, 0))
        .ForeColor = &HFF0000
    Next i
    
    '-------> Mover comensales
    sql1 = IIf(vg_tipbase = "1", " val(mid(b_minuta.min_fecmin,1,6)) ", " convert(int,substring(convert(varchar(8),b_minuta.min_fecmin),1,6)) ")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open RutinaLectura.Minutas(6, vg_codregimen, vg_codservicio, Val(vg_fecha), ""), vg_db, adOpenStatic
    If Not RS.EOF Then
       Do While Not RS.EOF
          j = (((Val(Mid(RS!min_fecmin, 7, 2)) * 5) - 5) + 1) + 1
          .Row = .MaxRows
          .Col = j + 2
          .CellType = CellTypeNumber
          .TypeNumberDecPlaces = 0
          .TypeNumberMin = 0
          .TypeNumberMax = 9999999
          .TypeHAlign = TypeHAlignRight
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
          .Value = IIf(IsNull(RS!min_racrea), 0, RS!min_racrea)
          .ForeColor = &HFF0000
          RS.MoveNext
       Loop
       RS.Close: Set RS = Nothing
    Else
       RS.Close: Set RS = Nothing
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open RutinaLectura.ServicioRaciones(1, vg_codservicio), vg_db, adOpenStatic
       
       If Not RS.EOF Then
          
          Do While Not RS.EOF
             
             inddia = 1
             
             For i = 2 To (.MaxCols - MaxColumna - 1) Step 5
                 
                 If RS!sra_serdia = IIf(fg_Dia(vg_fecha & fg_pone_cero(inddia, 2)) = 1, 7, Val(fg_Dia(vg_fecha & fg_pone_cero(inddia, 2)) - 1)) Then
                    
                    .Col = i + 2
                    .CellType = CellTypeNumber
                    .TypeNumberDecPlaces = 0
                    .TypeNumberMin = 0
                    .TypeNumberMax = 9999999
                    .TypeHAlign = TypeHAlignRight
                    .TypeSpin = False
                    .TypeIntegerSpinInc = 1
                    .TypeIntegerSpinWrap = False
                    .Value = IIf(IsNull(RS!raciones), 0, RS!raciones)
                    .ForeColor = &HFF0000
                 
                 End If
                 
                 inddia = inddia + 1
             
             Next i
             
             RS.MoveNext
          
          Loop
       
       End If
       
       RS.Close: Set RS = Nothing
    
    End If
    
    For i = 3 To (.MaxCols - MaxColumna) Step 5
        
        .Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoTec < 1 And vCtoPis > 0, 2, IIf(vCtoTec > 0 And vCtoPis = 0, 2, 1))): .Col = i
        
        If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) < Format(Date - IIf((fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 6 Or fg_Dia(Format(CDate(Mid(Trim(.text), 5, Len(Trim(.text)))), "yyyymmdd")) = 7) And fg_Dia(Format(Date, "yyyymmdd")) = 2, 4, 2), "d/mm/yyyy") Then
            
            For fil = 1 To (.MaxRows - 1)
                
                For Col = i - 1 To i + 2
                    
                    .Row = .MaxRows: .Col = Col
                    
                    If .CellType = CellTypeNumber Then
                       
                       .CellType = CellTypeStaticText
                       .TypeHAlign = TypeHAlignRight
                    
                    End If
                
                Next Col
            
            Next fil
        
        End If
    
    Next i
    '-------> Buscar fecha estructura fija
    
    inddia = 1: estfij = False
    '-------> Buscar datos estructura fija día
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.MinutaFijaDia(1, vg_codregimen, vg_codservicio, Val(vg_fecha), "2", ""), vg_db, adOpenStatic
    If Not RS.EOF Then estfij = True
    RS.Close: Set RS = Nothing
    fecesf = 0
    If Not estfij Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open RutinaLectura.MinutaFija(1, vg_codregimen, vg_codservicio, 0, 0, "", "", ""), vg_db, adOpenStatic
       If Not RS.EOF Then fecesf = IIf(IsNull(RS!fecval), 0, RS!fecval)
       RS.Close: Set RS = Nothing
    
    End If
    
    If Not estfij And fecesf > 0 And vg_tipbase = "1" Then
        '-------> Insert tabla productospmpdia
        aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPDetPlaMinRea"
        fg_CheckTmp aAp
        
        vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                      "INTO   " & aAp & " " & _
                      "FROM   b_productospmpdia " & _
                      "WHERE  ppd_cencos = '" & MuestraCasino(1) & "' " & _
                      "AND    ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                      "AND    ppd_propon > 0 " & _
                      "GROUP BY ppd_cencos, ppd_codpro"
        vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
        vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
        vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    
    End If
    
    For j = 2 To (.MaxCols - MaxColumna - 1) Step 5
        
        .Row = .MaxRows
        .Col = j + 2
        
        If Val(.text) > 0 Then
            
            Fecha = Val(vg_fecha) & Right("0" & inddia, 2)
            
            If estfij Then
               '-------> Calcular datos desde tabla estructura fija día
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               RS.Open RutinaLectura.MinutaFijaDia(2, vg_codregimen, vg_codservicio, Val(vg_fecha) & Right("0" & inddia, 2), "2", ""), vg_db, adOpenStatic
               If Not RS.EOF And Not IsNull(RS!cosesf) Then VecCosenc(inddia, 2) = VecCosenc(inddia, 2) + RS!cosesf
               RS.Close: Set RS = Nothing
            
            ElseIf Not estfij And fecesf > 0 Then
               
               '-------> Calcular datos desde tabla estructura fija
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
   
               If vg_tipbase = "1" Then
                  RS.Open RutinaLectura.MinutaFija(2, vg_codregimen, vg_codservicio, fecesf, fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))), aAp, "", ""), vg_db, adOpenStatic
               Else
                  RS.Open RutinaLectura.MinutaFija(3, vg_codregimen, vg_codservicio, fecesf, fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))), "", "", ""), vg_db, adOpenStatic
               End If
               If Not RS.EOF And Not IsNull(RS!cosesf) Then VecCosenc(inddia, 2) = VecCosenc(inddia, 2) + RS!cosesf
               RS.Close: Set RS = Nothing
            
            End If
        
        End If
        
        inddia = inddia + 1
    
    Next j
    
    '-------> Borrar tablas temporales
    If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
    
    '-------> Mostrar costo minuta dia
    j = 2
    For i = 1 To UBound(VecCosenc)
        .Row = .MaxRows
        .Col = j + 2
        vTotRac = 0
        If Trim(.text) <> "" And Val(.text) <> 0 Then vTotRac = .text
        .Row = SpreadHeader + IIf(vCtoTec > 0, 1, 0)
        .Col = j
        .TypeHAlign = TypeHAlignRight
        vCosVec = 0
        vCosVec = Round(VecCosenc(i, 1) + VecCosenc(i, 2), 2)
        If vTotRac > 0 And vCosVec > 0 Then
           .text = Format(Round(vCosVec / vTotRac, 2), fg_Pict(6, 2))
        Else
           .text = ""
        End If
        j = j + 5
    Next i
    
    .Row = 1: .Col = 1
    IblockRow = .Row: AiBlockRow = .Row
    IblockRow2 = .Row: AiBlockRow2 = .Row
    IblockCol = .Col: AiBlockCol = .Col
    iblockcol2 = .Col: AiBlockCol2 = .Col
End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
End Sub

Sub CalctodiaEnc(Row As Long, Col As Long)
Dim X       As Long
Dim NumRac  As Long
Dim vCosVec As Double
Dim vTotRac As Double
Dim cosdia  As Double

VecCosenc((Int(Col / 5) + 1), 1) = 0
For X = 1 To (vaSpread1.MaxRows - 1)
    vaSpread1.Row = X
    vaSpread1.Col = Col + 1: NumRac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
    vaSpread1.Col = Col + 2: cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
    vaSpread1.Col = Col + 3
    If Trim(vaSpread1.text) <> "" And NumRac > 0 Then
       vaSpread1.Col = Col + 2: VecCosenc((Int(Col / 5) + 1), 1) = Round(VecCosenc((Int(Col / 5) + 1), 1) + (cosdia * NumRac), vg_DCa)
    End If
Next X
vaSpread1.Row = vaSpread1.MaxRows
vaSpread1.Col = Col + 1
vTotRac = 0
If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) <> 0 Then vTotRac = vaSpread1.text
vaSpread1.Row = SpreadHeader + IIf(vCtoTec > 0, 1, 0)
vaSpread1.Col = Col - 1
vaSpread1.TypeHAlign = TypeHAlignRight
vCosVec = 0
vCosVec = Round(VecCosenc((Int(Col / 5) + 1), 1) + VecCosenc((Int(Col / 5) + 1), 2), 2)
If vTotRac > 0 And vCosVec > 0 Then vaSpread1.text = Format(Round(vCosVec / vTotRac, 2), fg_Pict(6, 2)) Else vaSpread1.text = ""
End Sub

Sub Calctodia(Row As Long, Col As Long)
Dim X       As Long
Dim NumRac  As Long
Dim cosdia  As Double

VecCos((Int(Col / 5) + 1), 1) = 0: VecCos((Int(Col / 5) + 1), 4) = 0
With vaSpread1
    For X = 1 To (.MaxRows - 1)
        .Row = X
        .Col = Col + 1: NumRac = IIf(Val(.text) = 0, 0, .text)
        .Col = Col + 2: cosdia = IIf(Val(.text) = 0, 0, .text)
        .Col = Col + 3
        If Trim(.text) <> "" And NumRac > 0 Then
           .Col = Col + 2: VecCos((Int(Col / 5) + 1), 1) = Round(VecCos((Int(Col / 5) + 1), 1) + (cosdia * NumRac), vg_DCa)
        End If
    Next X
    .Row = .MaxRows
    .Col = Col + 1: VecCos((Int(Col / 5) + 1), 4) = Round(VecCos((Int(Col / 5) + 1), 4) + IIf(Val(.text) = 0, 0, .text), vg_DCa)
End With
End Sub

Sub MostrarCosto(Col As Long)
Dim xcol    As Long
Dim ToaPla  As Double
Dim ToaEsf  As Double
Dim ToaFoo  As Double
Dim totdia  As Double
Dim totesf  As Double
Dim nracre  As Double
Dim nracfo  As Double
Dim totrac  As Double

vaSpread1.Col = Col
xcol = 0
For i = 1 To MaxColumna
    If (VectorCol(i) = vaSpread1.Col Or VectorCol(i) = (vaSpread1.Col + 1) Or VectorCol(i) = (vaSpread1.Col - 1) Or VectorCol(i) = (vaSpread1.Col - 2)) Then xcol = VectorCol(i): Exit For
Next i
'vaSpread1.Row = 0
vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoTec < 1 And vCtoPis > 0, 2, IIf(vCtoTec > 0 And vCtoPis = 0, 2, 1)))
vaSpread1.Col = xcol: Frame2(2).Caption = vaSpread1.text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
ToaPla = 0: ToaEsf = 0: ToaFoo = 0: totdia = 0: totesf = 0: nracre = 0: nracfo = 0: totrac = 0
For i = 1 To UBound(VecCos)
    If i <= (Int(xcol / 5) + 1) Then
       ToaPla = CCur(ToaPla + VecCos(i, 1))
       ToaEsf = CCur(ToaEsf + VecCos(i, 2))
       ToaFoo = CCur(ToaFoo + VecCos(i, 3))
       nracre = CCur(nracre + VecCos(i, 4))
       nracfo = CCur(nracfo + VecCos(i, 5))
    End If
    totrac = CCur(totrac + VecCos(i, 4))
    totdia = CCur(totdia + VecCos(i, 1))
    totesf = CCur(totesf + VecCos(i, 2))
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
If totrac > 0 Then Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2)) Else Label1(40).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(41).Caption = Format(CCur(totesf / totrac), fg_Pict(6, 2)) Else Label1(41).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(8).Caption = Format(CCur((totdia + totesf) / totrac), fg_Pict(6, 2)) Else Label1(8).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(48).Caption = Format(totrac, fg_Pict(6, 2)) Else Label1(48).Caption = Format(0, fg_Pict(6, 2))
Label1(20).Caption = Format(VecCos((Int(xcol / 5) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format(VecCos((Int(xcol / 5) + 1), 2), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(VecCos((Int(xcol / 5) + 1), 1) + (VecCos((Int(xcol / 5) + 1), 2))), fg_Pict(6, 2))
Label1(23).Caption = Format(VecCos((Int(xcol / 5) + 1), 3), fg_Pict(6, 2))
Label1(44).Caption = Format(VecCos((Int(xcol / 5) + 1), 4), fg_Pict(6, 2))
If VecCos((Int(xcol / 5) + 1), 4) > 0 Then Label1(45).Caption = Format(CCur((VecCos((Int(xcol / 5) + 1), 1) + (VecCos((Int(xcol / 5) + 1), 2))) / VecCos((Int(xcol / 5) + 1), 4)), fg_Pict(6, 2)) Else Label1(45).Caption = Format(0, fg_Pict(6, 2))
Label1(46).Caption = Format(VecCos((Int(xcol / 5) + 1), 5), fg_Pict(6, 2))
If VecCos((Int(xcol / 5) + 1), 5) > 0 Then Label1(47).Caption = Format(CCur(VecCos((Int(xcol / 5) + 1), 3) / VecCos((Int(xcol / 5) + 1), 5)), fg_Pict(6, 2)) Else Label1(47).Caption = Format(0, fg_Pict(6, 2))
Label1(31).Caption = Format(ToaPla, fg_Pict(6, 2))
Label1(32).Caption = Format((ToaEsf), fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(ToaPla + (ToaEsf)), fg_Pict(6, 2))
Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((ToaPla + ToaEsf) / nracre), fg_Pict(6, 2)) Else Label1(35).Caption = Format(0, fg_Pict(6, 2))
Label1(36).Caption = Format(ToaFoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(ToaFoo / nracfo), fg_Pict(6, 2)) Else Label1(38).Caption = Format(0, fg_Pict(6, 2))
End Sub

Sub CargarCosto()

Dim cosdia As Double, totdia As Double, totesf As Double, totrac As Double, estfij As Boolean
Dim Fecha As Long, xcol As Long, inddia As Long, fecesf As Double, nracre As Long, nracfo As Long
Dim aAp As String, sql1 As String, sql2 As String, sql3 As String

    fg_carga ""
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then vaSpread1.Col = 3
    DoEvents
    Label1(7).Caption = Format(0, fg_Pict(6, 2))
    Label1(8).Caption = Format(0, fg_Pict(6, 2))
    Label1(9).Caption = Format(0, fg_Pict(6, 2))
    Label1(11).Caption = Format(0, fg_Pict(6, 2))
    Label1(12).Caption = Format(0, fg_Pict(6, 2))
    Label1(13).Caption = Format(0, fg_Pict(6, 2))
    Label1(20).Caption = Format(0, fg_Pict(6, 2))
    Label1(21).Caption = Format(0, fg_Pict(6, 2))
    Label1(22).Caption = Format(0, fg_Pict(6, 2))
    Label1(23).Caption = Format(0, fg_Pict(6, 2))
    Label1(31).Caption = Format(0, fg_Pict(6, 2))
    Label1(32).Caption = Format(0, fg_Pict(6, 2))
    Label1(33).Caption = Format(0, fg_Pict(6, 2))
    Label1(34).Caption = Format(0, fg_Pict(6, 2))
    Label1(35).Caption = Format(0, fg_Pict(6, 2))
    Label1(36).Caption = Format(0, fg_Pict(6, 2))
    Label1(37).Caption = Format(0, fg_Pict(6, 2))
    Label1(38).Caption = Format(0, fg_Pict(6, 2))
    Label1(40).Caption = Format(0, fg_Pict(6, 2))
    Label1(41).Caption = Format(0, fg_Pict(6, 2))
    Label1(44).Caption = Format(0, fg_Pict(6, 2))
    Label1(45).Caption = Format(0, fg_Pict(6, 2))
    Label1(46).Caption = Format(0, fg_Pict(6, 2))
    Label1(47).Caption = Format(0, fg_Pict(6, 2))
    Label1(48).Caption = Format(0, fg_Pict(6, 2))
    j = 0: cosdia = 0: totdia = 0: totesf = 0: fecesf = 0: inddia = 1: NumRac = 0: totrac = 0
    For i = 1 To MaxColumna
        DoEvents
        If (VectorCol(i) = vaSpread1.Col Or VectorCol(i) = (vaSpread1.Col + 1) Or VectorCol(i) = (vaSpread1.Col - 1) Or VectorCol(i) = (vaSpread1.Col - 2)) Then xcol = VectorCol(i): Exit For
    Next i
    vaSpread1.Row = SpreadHeader + IIf(vCtoPis > 0 And vCtoTec > 0, 3, IIf(vCtoTec < 1 And vCtoPis > 0, 2, IIf(vCtoTec > 0 And vCtoPis = 0, 2, 1)))
    vaSpread1.Col = xcol: Frame2(2).Caption = vaSpread1.text: Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
    ReDim VecCos(MaxColumna, 5)
    For j = 1 To MaxColumna
        DoEvents
        VecCos(j, 1) = 0: VecCos(j, 2) = 0: VecCos(j, 3) = 0: VecCos(j, 4) = 0: VecCos(j, 5) = 0
    Next j
    estfij = False
    '-------> Buscar datos estructura fija día
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open RutinaLectura.MinutaFijaDia(1, vg_codregimen, vg_codservicio, Val(vg_fecha), "2", ""), vg_db, adOpenStatic
    If Not RS.EOF Then estfij = True
    RS.Close: Set RS = Nothing
    fecesf = 0
    If Not estfij Then
       DoEvents
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       RS.Open RutinaLectura.MinutaFija(1, vg_codregimen, vg_codservicio, 0, 0, "", "", ""), vg_db, adOpenStatic
       If Not RS.EOF Then fecesf = IIf(IsNull(RS!fecval), 0, RS!fecval)
       RS.Close: Set RS = Nothing
    
    End If
    '-------> Calcular costo día planificado & estructura fija & salida
    Bar1(0).Min = 0: Bar1(0).Value = 0: Bar1(0).max = MaxColumna: Frame2(4).Visible = True: Bar1(0).Visible = True
    '-------> Mover salida producción
    sql1 = IIf(vg_tipbase = "1", " SUM(IIf(a.tov_tipdoc='SP',b.dev_ptotal,'-' & b.dev_ptotal)) AS totdoc ", " SUM(CASE WHEN a.tov_tipdoc = 'SP' THEN b.dev_ptotal ELSE (-1*b.dev_ptotal) END) AS totdoc ")
    sql2 = IIf(vg_tipbase = "1", " format(a.tov_fecpro,'mm/yyyy') ", " substring(convert(varchar(10), a.tov_fecpro,103),4,8) ")
    sql3 = IIf(vg_tipbase = "1", " format(('" & fg_Ctod1(Val(vg_fecha) & Right("01", 2)) & "'),'mm/yyyy') ", " substring(('" & fg_Ctod1(Val(vg_fecha) & Right("01", 2)) & "'),4,8) ")
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT a.tov_fecpro, a.tov_codreg, a.tov_codser, " & sql1 & " " & _
            "FROM  b_totventas a, b_detventas b, b_productos c " & _
            "WHERE a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
            "AND   a.tov_codreg = " & vg_codregimen & " " & _
            "AND   a.tov_codser = " & vg_codservicio & " " & _
            "AND  (a.tov_tipdoc = 'SP' or a.tov_tipdoc = 'DP') " & _
            "AND   b.dev_canmer <> 0 " & _
            "AND   a.tov_codbod = " & vg_codbod & " AND a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' " & _
            "AND   " & sql2 & " = " & sql3 & " " & _
            "GROUP BY a.tov_fecpro, a.tov_codreg, a.tov_codser", vg_db, adOpenStatic
    
    Do While Not RS.EOF
       
       DoEvents
       VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 1) = 0: VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 2) = 0: VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 4) = 0: VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 5) = 0
       VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 3) = 0
       VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 3) = Round(VecCos(Val(Mid(RS!tov_fecpro, 1, 2)), 3) + RS!totdoc, vg_DCa)
       RS.MoveNext
    
    Loop
    RS.Close: Set RS = Nothing

For j = 2 To (vaSpread1.MaxCols - MaxColumna - 1) Step 5
    
    DoEvents
    Bar1(0).Value = Bar1(0).Value + 1
    Fecha = Val(vg_fecha) & Right("0" & inddia, 2)
    VecCos(inddia, 1) = 0: VecCos(inddia, 2) = 0: VecCos(inddia, 4) = 0: VecCos(inddia, 5) = 0
    
    For i = 1 To (vaSpread1.MaxRows - 1)
        
        DoEvents
        vaSpread1.Row = i
        vaSpread1.Col = j + 2: NumRac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = j + 3: cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = j + 4
        
        If Trim(vaSpread1.text) <> "" And NumRac > 0 Then
           
           totdia = Round(totdia + (cosdia * NumRac), vg_DCa)
           VecCos(inddia, 1) = Round(VecCos(inddia, 1) + (cosdia * NumRac), vg_DCa)
        
        End If
    
    Next i
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = j + 2
    VecCos(inddia, 4) = Round(VecCos(inddia, 4) + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
    totrac = Round(totrac + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = j + 2
    If Not estfij And fecesf > 0 And vg_tipbase = "1" Then
       '-------> Insert tabla productospmpdia
       aAp = Trim(vg_NUsr) & "_tmp_ProductoPMPCargarCostoRea"
       fg_CheckTmp aAp
        vg_db.Execute "SELECT ppd_cencos, ppd_codpro, 0 AS ppd_propon, Max(ppd_fecdia) AS ppd_fecdia " & _
                      "INTO " & aAp & " " & _
                      "FROM b_productospmpdia " & _
                      "WHERE ppd_cencos = '" & MuestraCasino(1) & "' " & _
                      "AND   ppd_fecdia >= " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_fecdia <= " & Format(CDate(vg_ciedia), "yyyymmdd") & " " & _
                      "AND   ppd_propon > 0 " & _
                      "GROUP BY ppd_cencos, ppd_codpro"
        vg_db.Execute "ALTER TABLE " & aAp & " ADD Constraint pmp_pk Primary Key (ppd_cencos, ppd_codpro, ppd_fecdia)"
        vg_db.Execute "UPDATE " & aAp & " INNER JOIN b_productospmpdia ON (" & aAp & ".ppd_fecdia = b_productospmpdia.ppd_fecdia) AND (" & aAp & ".ppd_codpro = b_productospmpdia.ppd_codpro) AND (" & aAp & ".ppd_cencos = b_productospmpdia.ppd_cencos) SET " & aAp & ".ppd_propon=b_productospmpdia.ppd_propon"
        vg_db.Execute "INSERT INTO " & aAp & " (ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia) SELECT DISTINCT ppd_cencos, ppd_codpro, ppd_propon, ppd_fecdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & " AND ppd_codpro NOT IN (SELECT DISTINCT ppd_codpro FROM " & aAp & ")"
    
    End If
    
    If Val(vaSpread1.text) > 0 Then
        If estfij Then
           '-------> Calcular datos desde tabla estructura fija día
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
           
           RS.Open RutinaLectura.MinutaFijaDia(2, vg_codregimen, vg_codservicio, Val(vg_fecha) & Right("0" & inddia, 2), "2", ""), vg_db, adOpenStatic
           If Not RS.EOF And Not IsNull(RS!cosesf) Then totesf = Round(totesf + RS!cosesf, vg_DCa): VecCos(inddia, 2) = Round(VecCos(inddia, 2) + RS!cosesf, vg_DCa)
           RS.Close: Set RS = Nothing
        
        ElseIf Not estfij And fecesf > 0 Then
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
     
            '-------> Calcular datos desde tabla estructura fija
            If vg_tipbase = "1" Then
               RS.Open RutinaLectura.MinutaFija(2, vg_codregimen, vg_codservicio, fecesf, fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))), aAp, "", ""), vg_db, adOpenStatic
            Else
               RS.Open RutinaLectura.MinutaFija(3, vg_codregimen, vg_codservicio, fecesf, fg_NumDia(Trim(Left(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2), Len(fg_Fecha_Dia(vg_fecha & Right("0" & inddia, 2), 2)) - 2))), "", "", ""), vg_db, adOpenStatic
            End If
            If Not RS.EOF And Not IsNull(RS!cosesf) Then totesf = Round(totesf + RS!cosesf, vg_DCa): VecCos(inddia, 2) = Round(VecCos(inddia, 2) + RS!cosesf, vg_DCa)
            RS.Close: Set RS = Nothing
        
        End If
    
    End If
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    RS.Open RutinaLectura.MinutaRaciones(1, vg_codregimen, vg_codservicio, "PRODUCIDAS", Val(vg_fecha) & Right("0" & inddia, 2)), vg_db, adOpenStatic
    If Not RS.EOF And Not IsNull(RS!mir_nrorac) Then VecCos(inddia, 5) = Round(VecCos(inddia, 5) + RS!mir_nrorac, vg_DPr) Else VecCos(inddia, 5) = 0
    RS.Close: Set RS = Nothing
    
    inddia = inddia + 1

Next j
'-------> Borrar tablas temporales
If Trim(aAp) <> "" Then vg_db.Execute "DROP TABLE " & aAp & ""
Frame2(4).Visible = False
Bar1(0).Visible = False
'-------> Fin Calcular costo día
ToaPla = 0: ToaEsf = 0: ToaFoo = 0: NumRac = 0: nracfo = 0
For i = 1 To (Int(xcol / 5) + 1)
    DoEvents
    ToaPla = Round(ToaPla + VecCos(i, 1), vg_DCa)
    ToaEsf = Round(ToaEsf + VecCos(i, 2), vg_DCa)
    ToaFoo = Round(ToaFoo + VecCos(i, 3), vg_DCa)
    nracre = Round(nracre + VecCos(i, 4), vg_DPr)
    nracfo = Round(nracfo + VecCos(i, 5), vg_DPr)
Next i
Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
Label1(11).Caption = Format(totesf, fg_Pict(6, 2))
Label1(12).Caption = Format(CCur(totdia + totesf), fg_Pict(6, 2))
If totrac > 0 Then Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2)) Else Label1(40).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(41).Caption = Format(CCur(totesf / totrac), fg_Pict(6, 2)) Else Label1(41).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(8).Caption = Format(CCur((totdia + totesf) / totrac), fg_Pict(6, 2)) Else Label1(8).Caption = Format(0, fg_Pict(6, 2))
If totrac > 0 Then Label1(48).Caption = Format(totrac, fg_Pict(6, 2)) Else Label1(48).Caption = Format(0, fg_Pict(6, 2))
Label1(20).Caption = Format(VecCos((Int(xcol / 5) + 1), 1), fg_Pict(6, 2))
Label1(21).Caption = Format((VecCos((Int(xcol / 5) + 1), 2)), fg_Pict(6, 2))
Label1(22).Caption = Format(CCur(VecCos((Int(xcol / 5) + 1), 1) + (VecCos((Int(xcol / 5) + 1), 2))), fg_Pict(6, 2))
Label1(23).Caption = Format(VecCos((Int(xcol / 5) + 1), 3), fg_Pict(6, 2))
Label1(44).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(45).Caption = Format(CCur((VecCos((Int(xcol / 5) + 1), 1) + (VecCos((Int(xcol / 5) + 1), 2))) / nracre), fg_Pict(6, 2)) Else Label1(45).Caption = Format(0, fg_Pict(6, 2))
Label1(46).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(47).Caption = Format(CCur(VecCos((Int(xcol / 5) + 1), 3) / nracfo), fg_Pict(6, 2)) Else Label1(47).Caption = Format(0, fg_Pict(6, 2))
Label1(31).Caption = Format(ToaPla, fg_Pict(6, 2))
Label1(32).Caption = Format((ToaEsf), fg_Pict(6, 2))
Label1(33).Caption = Format(CCur(ToaPla + (ToaEsf)), fg_Pict(6, 2))
Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
If nracre > 0 Then Label1(35).Caption = Format(CCur((ToaPla + ToaEsf) / nracre), fg_Pict(6, 2)) Else Label1(35).Caption = Format(0, fg_Pict(6, 2))
Label1(36).Caption = Format(ToaFoo, fg_Pict(6, 2))
Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
If nracfo > 0 Then Label1(38).Caption = Format(CCur(ToaFoo / nracfo), fg_Pict(6, 2)) Else Label1(38).Caption = Format(0, fg_Pict(6, 2))
IndCos = True
fg_descarga
End Sub
