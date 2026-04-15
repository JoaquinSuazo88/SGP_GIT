VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form M_MinSR2 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Minuta Bloque"
   ClientHeight    =   7860
   ClientLeft      =   2430
   ClientTop       =   2490
   ClientWidth     =   14040
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7860
   ScaleWidth      =   14040
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Height          =   3255
      Index           =   0
      Left            =   0
      TabIndex        =   22
      Top             =   4620
      Visible         =   0   'False
      Width           =   15315
      Begin VB.CommandButton Actualizar_LYD 
         Caption         =   "Actualizar LYD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   11760
         TabIndex        =   82
         Top             =   630
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         Caption         =   "Acumulado hasta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Index           =   3
         Left            =   7800
         TabIndex        =   64
         Top             =   480
         Width           =   3795
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   81
            Top             =   2280
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   80
            Top             =   1980
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   79
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   990
            TabIndex        =   78
            Top             =   2280
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   990
            TabIndex        =   77
            Top             =   1980
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   990
            TabIndex        =   76
            Top             =   1680
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   990
            TabIndex        =   75
            Top             =   750
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   990
            TabIndex        =   74
            Top             =   465
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Realizado"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   73
            Top             =   210
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Planificado"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   72
            Top             =   210
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Band."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   71
            Top             =   2280
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rac."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   70
            Top             =   1980
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Total"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   69
            Top             =   1680
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LYD"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   68
            Top             =   750
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alimentos"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   67
            Top             =   465
            Width           =   855
         End
         Begin VB.Line Line3 
            X1              =   105
            X2              =   3675
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LYD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   52
            Left            =   90
            TabIndex        =   66
            Top             =   1300
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   53
            Left            =   990
            TabIndex        =   65
            Top             =   1260
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Día 01/08/2004"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Index           =   2
         Left            =   3960
         TabIndex        =   46
         Top             =   480
         Width           =   3795
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   63
            Top             =   1725
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   990
            TabIndex        =   62
            Top             =   1725
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   990
            TabIndex        =   61
            Top             =   810
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   990
            TabIndex        =   60
            Top             =   525
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Realizado"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   59
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Planificado"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   58
            Top             =   240
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Total"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   57
            Top             =   1725
            Width           =   765
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LYD"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   56
            Top             =   810
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alimentos"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   55
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rac."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   54
            Top             =   2010
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto.Band."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   53
            Top             =   2310
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   44
            Left            =   990
            TabIndex        =   52
            Top             =   2010
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   45
            Left            =   990
            TabIndex        =   51
            Top             =   2310
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   50
            Top             =   2010
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   49
            Top             =   2310
            Width           =   1320
         End
         Begin VB.Line Line2 
            X1              =   105
            X2              =   3675
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LYD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   50
            Left            =   105
            TabIndex        =   48
            Top             =   1300
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   51
            Left            =   990
            TabIndex        =   47
            Top             =   1260
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Total Mes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2685
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   3795
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   45
            Top             =   2295
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   44
            Top             =   1620
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   43
            Top             =   810
            Width           =   1320
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   42
            Top             =   2295
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   41
            Top             =   1620
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   40
            Top             =   525
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Patrón"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   39
            Top             =   2295
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Total"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   38
            Top             =   1620
            Width           =   435
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LYD"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   37
            Top             =   810
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Alimentos"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   36
            Top             =   525
            Width           =   855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo Total"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   35
            Top             =   240
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cto. Bandeja"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   34
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   33
            Top             =   525
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   32
            Top             =   810
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   31
            Top             =   1905
            Width           =   1320
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rac."
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   30
            Top             =   1905
            Width           =   360
         End
         Begin VB.Line Line1 
            X1              =   105
            X2              =   3675
            Y1              =   1155
            Y2              =   1155
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "LYD"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   29
            Top             =   1300
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   28
            Top             =   1260
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
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
            TabIndex        =   27
            Top             =   1260
            Visible         =   0   'False
            Width           =   1320
         End
      End
      Begin VB.Frame Frame2 
         Height          =   375
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   120
         Visible         =   0   'False
         Width           =   2775
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   25
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
            Picture         =   "M_MinSR2.frx":0000
            Top             =   150
            Visible         =   0   'False
            Width           =   480
         End
      End
      Begin VB.CommandButton CmdCerrar 
         Caption         =   "Cerrar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   11760
         TabIndex        =   23
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   2760
         Picture         =   "M_MinSR2.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Costo Totales"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   3360
         Picture         =   "M_MinSR2.frx":0614
         Stretch         =   -1  'True
         ToolTipText     =   "Costo Bandeja Planificado"
         Top             =   240
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
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
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Día"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   0
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1035
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   13545
      Begin VB.CheckBox Check1 
         Caption         =   "Todas "
         Height          =   255
         Index           =   6
         Left            =   12390
         TabIndex        =   20
         Top             =   510
         Value           =   1  'Checked
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "%Pond. Diaria"
         Height          =   255
         Index           =   5
         Left            =   10350
         TabIndex        =   19
         Top             =   510
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "%Pond.xEstr. "
         Enabled         =   0   'False
         Height          =   255
         Index           =   4
         Left            =   8310
         TabIndex        =   18
         Top             =   510
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calorías"
         Height          =   255
         Index           =   3
         Left            =   6630
         TabIndex        =   17
         Top             =   510
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.CheckBox Check1 
         Caption         =   "N.Rac."
         Height          =   195
         Index           =   0
         Left            =   2550
         TabIndex        =   16
         Top             =   510
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Costo"
         Height          =   255
         Index           =   1
         Left            =   3915
         TabIndex        =   15
         Top             =   510
         Value           =   1  'Checked
         Width           =   795
      End
      Begin VB.CheckBox Check1 
         Caption         =   "R"
         Height          =   255
         Index           =   2
         Left            =   5490
         TabIndex        =   14
         Top             =   510
         Value           =   1  'Checked
         Width           =   495
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00FFFFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   3
         Left            =   11040
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cambio Minuta"
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
         Left            =   11400
         TabIndex        =   13
         Top             =   135
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Semana Nº"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   12
         Top             =   150
         Width           =   930
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   0
         Left            =   9120
         Top             =   165
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Celda Habilitada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   9480
         TabIndex        =   11
         Top             =   135
         Width           =   1365
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H008080FF&
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
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7665
         TabIndex        =   10
         Top             =   135
         Width           =   1395
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         FillColor       =   &H00C0FFC0&
         FillStyle       =   0  'Solid
         Height          =   135
         Index           =   2
         Left            =   5010
         Top             =   165
         Width           =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Estructura de Servicio"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5370
         TabIndex        =   9
         Top             =   135
         Width           =   1860
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14040
      _ExtentX        =   24765
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread1 
      DragIcon        =   "M_MinSR2.frx":091E
      Height          =   3015
      Left            =   -90
      TabIndex        =   8
      Top             =   2040
      Width           =   11055
      _Version        =   393216
      _ExtentX        =   19500
      _ExtentY        =   5318
      _StockProps     =   64
      BackColorStyle  =   1
      ColsFrozen      =   1
      EditEnterAction =   2
      EditModeReplace =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   1
      MaxRows         =   1
      ProcessTab      =   -1  'True
      Protect         =   0   'False
      RestrictRows    =   -1  'True
      SpreadDesigner  =   "M_MinSR2.frx":0D60
      UserResize      =   1
      VisibleCols     =   1
      VisibleRows     =   1
      TextTip         =   2
      TextTipDelay    =   0
      ScrollBarTrack  =   3
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Espere. Buscando Información..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   90
      TabIndex        =   21
      Top             =   1800
      Width           =   2685
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planificación Minutas Teórica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
         Begin VB.Menu Aportes 
            Caption         =   "Sin % P-G-Cho-Agrs"
            Index           =   10
         End
         Begin VB.Menu Aportes 
            Caption         =   "Con % P-G-Cho-Agrs"
            Index           =   20
         End
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
         Caption         =   "Ac&tualizar Costo Recetas"
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
         Caption         =   "Informe Matriz de Precios"
         Index           =   17
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
         Caption         =   "&Agrega Estructura"
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
         Caption         =   "Pegado Especial"
         Enabled         =   0   'False
         Index           =   14
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "&Buscar Receta"
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
Attribute VB_Name = "M_MinSR2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private RS                  As New ADODB.Recordset
Private RS1                 As New ADODB.Recordset
Private i                   As Long
Private j                   As Long
Private indcortarpegar      As Long
Private fechadesde          As Long
Private fechahasta          As Long
Private MaxColumna          As Long
Private maxfila             As Long
Private AddReceta           As Long
Private iblockrow           As Integer
Private iblockrow2          As Integer
Private iblockcol           As Integer
Private iblockcol2          As Integer
Private SwSalir             As Integer
Private aiblockrow          As Integer
Private aiblockrow2         As Integer
Private aiblockcol          As Integer
Private aiblockcol2         As Integer
Private indactivo           As Integer
Private IndGrabado          As Integer
Private VecDia()            As Long
Private vCtoPis             As Double
Private vCtoTec             As Double
Private indcos              As Boolean
Private veccos()            As Variant
Private VecCosenc()         As Variant
Private vectorcol()         As Long
Private MsgTitulo           As String

Dim xColIni As Variant, xRowIni As Variant, xcolfin As Variant, xRowFin As Variant

Private BtnX                As Variant
Private Cancel              As Variant
Private SearchFlagsEqual    As Variant
Private existedat           As Variant
Private AadRec              As Variant
Private AuxCol              As Variant
Private z                   As Long
Private fil                 As Long
Private Colu                As Long
Private X                   As Long
Private ValLcntH            As Variant
Private L                   As Variant
Private xx                  As Long
Private Indvaspread1        As Variant
Private NumRec              As Variant
Private StrRec              As Variant
Private StrRecb             As Variant
Private numrac              As Variant
Private toapla              As Variant
Private toaesf              As Variant
Private toafoo              As Variant
Private nrodia              As String
Private iser                As Long
Private xser                As Long
Private CosDes              As Double
Private EstCheck            As Boolean
Dim SpresdText              As String
Dim CellTex                 As String
Dim SpreadClon              As New M_MinSR2
Dim ContadorDeshacer        As Long
Dim TipoMinuta              As Boolean
Dim AuxTipoMinuta           As Boolean
Dim TipoCopia               As String
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
'Enum DeshacerType
'    AddFile = 1
'    DelFile = 2
'End Enum

Private Sub Actualizar_LYD_Click()
On Error GoTo Man_Error

'-------> Cargar costo
Call CargarCosto(True)

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Aportes_Click(Index As Integer)

Dim i As Long
Dim j As Long
Dim FechaGrilla As Long

Select Case Index

Case 10
     
     vaSpread1.Row = vaSpread1.ActiveRow
     vaSpread1.Col = vaSpread1.ActiveCol
     If vaSpread1.Col = 1 Or vaSpread1.Col = 2 Then Exit Sub
     j = 0
            
     For i = 1 To MaxColumna
         
         If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or _
            vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 3) Or vectorcol(i) = (vaSpread1.Col - 4)) Then
            
            j = vectorcol(i): Exit For
         
         End If
         
     Next i
     vaSpread1.Col = j
     vaSpread1.Row = 0
 
     Let VarSitioRemoto = True
     Call C_ApoPla.LlenarApoPlan(M_MinSR2, "Aporte Planificación Teórica " & vaSpread1.text, vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), 1, j)
     C_ApoPla.Show 1, Me
     Let VarSitioRemoto = False

Case 20

     vaSpread1.Row = vaSpread1.ActiveRow
     vaSpread1.Col = vaSpread1.ActiveCol
     
     If vaSpread1.Col = 1 Or vaSpread1.Col = 2 Then
        
        Exit Sub
     
     End If
     j = 0
            
     For i = 1 To MaxColumna
         
         If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or _
            vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 3) Or vectorcol(i) = (vaSpread1.Col - 4)) Then
            
            j = vectorcol(i): Exit For
         
         End If
         
     Next i
         
     '-------> Mover Fecha de grilla
     vaSpread1.Row = SpreadHeader + 3
     vaSpread1.Col = j
     FechaGrilla = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")
     
     vaSpread1.Col = j
     vaSpread1.Row = 0
 
     Let VarSitioRemoto = True
     Call C_AporteSansis.LlenarApoPlan(M_MinSR2, "Aporte Planificación Bloque " & vaSpread1.text, vg_codcasino, vg_codregimen, vg_codservicio, Val(FechaGrilla), 1, j)
     C_AporteSansis.Show 1, Me
     Let VarSitioRemoto = False

End Select

End Sub

Private Sub Check1_Click(Index As Integer)

If EstCheck Then Exit Sub
Dim i As Integer
If Index = 6 Then
   
   For i = 0 To 5
      
      EstCheck = True
      
      If TipoMinuta = True And (i = 4 Or i = 5) Then
      
      Else
         
         Check1(i).Value = Check1(6).Value
         HabilitaCol i
      
      End If
   
   Next i

Else
   
   HabilitaCol Index

End If
EstCheck = False

End Sub

Sub HabilitaCol(op As Integer)

Dim icol As Long
icol = IIf(op = 0, 3, IIf(op = 1, 4, IIf(op = 2, 0, IIf(op = 3, 6, IIf(op = 4, 0, 2)))))
If op = 4 Then

'   vaSpread1.Row = 0
'   vaSpread1.Col = 2
'   If vaSpread1.ColHidden = True Then
'      vaSpread1.ColHidden = False
'      DoEvents
'   Else
'      vaSpread1.ColHidden = True
'      DoEvents
'   End If

Else
    
    For i = 3 To (vaSpread1.maxcols - 2) Step 7
      
      DoEvents
      vaSpread1.Row = 0
      vaSpread1.Col = i + icol
      
      If vaSpread1.ColHidden = True Then
         
         vaSpread1.ColHidden = False
         DoEvents
      
      Else
         
         vaSpread1.ColHidden = True
         DoEvents
      
      End If
    
    Next i

End If

vaSpread1.Visible = True

End Sub

Private Sub CmdCerrar_Click()

On Error GoTo Man_Error

Frame2(0).Visible = False
vaSpread1.Move 0, 1760, ScaleWidth, ScaleHeight - 1760
Image2(0).Visible = False: Image2(1).Visible = False
indcos = False
Exit Sub

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Estructura1_Click(Index As Integer)

On Error GoTo Man_Error

'-------> Validar si minuta esta bloqueada
If ValidarBloqueoMinuta Then Exit Sub

If Not TipoMinuta Then
    
    LlenaSubMenu Estructura1, Index

Else
    
    LlenarSubMenuBloque Estructura1, Index

End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub LlenarSubMenuBloque(SubMenu As Object, Index As Integer)

On Error GoTo Man_Error

Dim auxest As Long
Dim xrow As Long, i As Long

If vaSpread1.MaxRows = vaSpread1.ActiveRow Then
   
   Call MsgBox("No puede insertar estructura fija ultima fila", vbCritical + vbOKOnly, MsgTitulo)
   Exit Sub

End If

GrabarCambios 1, 1, "Estructura Servicio"
xrow = vaSpread1.ActiveRow
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = vaSpread1.maxcols - 1
If Trim(vaSpread1.text) <> "" Then auxest = Val(vaSpread1.text)
vaSpread1.Col = 1

DesqloqSubMenu vaSpread1.text
vaSpread1.text = SubMenu(Index).Caption

ActualizaEstructuraInferior vaSpread1, SubMenu(Index).Caption
vaSpread1.Col = vaSpread1.maxcols - 1: vaSpread1.text = SubMenu(Index).HelpContextID
'------->
For i = xrow + 1 To vaSpread1.MaxRows - 1
    
    vaSpread1.Row = i
    vaSpread1.Col = vaSpread1.maxcols - 1
    
    If Val(vaSpread1.text) = auxest Then
       
       vaSpread1.text = SubMenu(Index).HelpContextID
    
    Else
       
       Exit For
    
    End If

Next i

Estructura1.item(Index).Enabled = False: Estructura2.item(Index).Enabled = False
Plantilla(0).Enabled = True
IndGrabado = 1
Toolbar1.Buttons(1).Visible = False
Toolbar1.Buttons(2).Visible = True

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Sub DesqloqSubMenu(OpcioneMenu As String)

On Error GoTo Man_Error

Dim iA As Long
For iA = 1 To Estructura2.count - 1
    
    If Trim(Estructura2.item(iA).Caption) = Trim(OpcioneMenu) Then
       
       Estructura1.item(iA).Enabled = True
       Estructura2.item(iA).Enabled = True
    
    End If

Next iA

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Sub ActualizaEstructuraInferior(ByVal Spread As vaSpread, ByVal NameEstruct As String, Optional UltimoCambio As Long)

On Error GoTo Man_Error

Dim xRespRow As Long, xRespCol As Long, EstructuraAnterior As String
Dim x1 As Long
xRespRow = Spread.Row
xRespCol = Spread.Col
Spread.Col = Spread.maxcols - 1
EstructuraAnterior = Spread.text
For x1 = Spread.Row To Spread.MaxRows - 1
    
    Spread.Row = x1
    
    If Trim(Spread.text) = "" Or Spread.Row = Spread.MaxRows Or Trim(EstructuraAnterior) <> Trim(Spread.text) Then
       
       Exit For
    
    End If
    
    Spread.text = NameEstruct

Next x1
Spread.Row = xRespRow
Spread.Col = xRespCol

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Sub LlenaSubMenu(SubMenu As Object, Index As Integer)

On Error GoTo Man_Error
    
    Dim i  As Long
    Dim j As Long
    Dim colgrupo As Long
    Dim CodigoGrupo As Long
    Dim RowGrupoMin As Long
    Dim RowGrupoMax As Long
    
    If vaSpread1.ActiveRow = vaSpread1.MaxRows Then
       
       Call MsgBox("No puede insertar estructura fija ultima fila", vbCritical + vbOKOnly, MsgTitulo)
       Exit Sub
    
    End If
    
    '-------> Rescata El Codigo de Agrupacion
    Dim RS2  As New ADODB.Recordset
    
    If RS2.State = 1 Then RS2.Close
    RS2.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS2 = vg_db.Execute("sgpadm_Sel_CodigodeAgrupacionServicio " & vg_codservicio & ", " & SubMenu(Index).HelpContextID & "")
    CodigoGrupo = RS2!ess_agrupacionestructura
    RS2.Close
    Set RS2 = Nothing
    
    GrabarCambios 1, 1, "Estructura Servicio"
    'columna de grupo de estructura y Encabezado
    colgrupo = vaSpread1.GetColFromID("Grupo") + 1
    '-------> Buscar grupo estructura servicio
    RowGrupoMin = vaSpread1.SearchCol(colgrupo, 0, -1, CodigoGrupo, SearchFlagsValue)
    If RowGrupoMin > 0 Then
       
       For i = RowGrupoMin To vaSpread1.MaxRows - 1
           
           vaSpread1.Row = i
           vaSpread1.Col = vaSpread1.maxcols
           
           If Val(vaSpread1.text) <> CodigoGrupo And Trim(vaSpread1.text) <> "" Then
              
              Exit For
           
           End If
           
           RowGrupoMax = i
       
       Next i
       
       If vaSpread1.ActiveRow >= RowGrupoMin And vaSpread1.ActiveRow <= RowGrupoMax Then
          
          vaSpread1.Row = vaSpread1.ActiveRow
          
          If Trim(vaSpread1.text) = "" Then
             
             RowGrupoMax = vaSpread1.ActiveRow
          
          Else
             
             RowGrupoMax = RowGrupoMax + 1
             vaSpread1.MaxRows = vaSpread1.MaxRows + 1
             vaSpread1.InsertRows RowGrupoMax, 1
          
          End If
       
       Else
          
          RowGrupoMax = RowGrupoMax + 1
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows RowGrupoMax, 1
       
       End If
    
    Else
       
       vaSpread1.Row = vaSpread1.ActiveRow
       
       If Val(vaSpread1.text) >= 0 Then
          
          RowGrupoMax = vaSpread1.Row
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows RowGrupoMax, 1
       
       Else
          
          'columna de grupo de estructura y Encabezado
          colgrupo = vaSpread1.GetColFromID("Grupo") + 1
          '-------> Buscar grupo estructura servicio
          vaSpread1.Col = colgrupo
          RowGrupoMin = vaSpread1.SearchCol(colgrupo, 0, -1, Trim(vaSpread1.text), SearchFlagsValue)
          RowGrupoMax = 0
          
          For i = RowGrupoMin To vaSpread1.MaxRows - 1
              
              vaSpread1.Row = i
              vaSpread1.Col = vaSpread1.maxcols
              
              If Val(vaSpread1.text) <> CodigoGrupo And Trim(vaSpread1.text) <> "" Then
                 
                 Exit For
              
              End If
              
              RowGrupoMax = i
          
          Next i
          
          RowGrupoMax = RowGrupoMax + 1
          vaSpread1.MaxRows = vaSpread1.MaxRows + 1
          vaSpread1.InsertRows RowGrupoMax, 1
       
       End If

    End If

    vaSpread1.Row = RowGrupoMax
    '-------> Mover codigo estructura
    vaSpread1.Col = vaSpread1.maxcols - 1
    vaSpread1.text = SubMenu(Index).HelpContextID
    '-------> Mover grupo estructura
    vaSpread1.Col = vaSpread1.maxcols
    vaSpread1.text = CodigoGrupo
    '-------> Mover descripción estructura servico
    vaSpread1.Col = 1
    vaSpread1.text = SubMenu(Index).Caption
    vaSpread1.Col = vaSpread1.maxcols: vaSpread1.text = CodigoGrupo
    vaSpread1.Col = vaSpread1.maxcols - 1: vaSpread1.text = SubMenu(Index).HelpContextID

    If RowGrupoMin < 0 Then
       
       vaSpread1.Col = 2
       vaSpread1.CellType = CellTypePercent
       vaSpread1.text = 0
       vaSpread1.CellType = CellTypePercent
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
       vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
       vaSpread1.TypePercentDecPlaces = 0
       vaSpread1.TypePercentMax = 1000
       vaSpread1.TypeNegRed = True
    
    End If
    '-------> Mover color a las lineas nuevas
    vaSpread1.Row = RowGrupoMax
    vaSpread1.Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor
    '-------> Mover color a la columna estructura servicio
    vaSpread1.Col = 1
    vaSpread1.BackColor = Shape1(2).FillColor

    Estructura1(Index).Enabled = False: Estructura2(Index).Enabled = False
    IndGrabado = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Sub

Private Sub Estructura2_Click(Index As Integer)

On Error GoTo Man_Error

'20211209 If Not TipoMinuta Then
'
'    LlenaSubMenu Estructura2, Index
'
'Else
    
    LlenarSubMenuBloque Estructura2, Index

'20211209 End If

'    LlenaSubMenu Estructura2, Index

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()
    
    Call fg_descarga

End Sub

Private Sub Form_Load()

Dim X       As Long
Dim nomser  As String
Dim nomreg  As String
Dim RS      As New ADODB.Recordset
Dim RS1     As New ADODB.Recordset

    Me.Height = 6765
    Me.Width = 11055
    Me.HelpContextID = vg_OpcM
    fg_centra Me
    MsgTitulo = M_MinSR1.Caption
    fg_carga ""
    EstCheck = False
    indcos = False
    ContadorDeshacer = 0
    Toolbar1.ImageList = Partida.IL1
    Set BtnX = Toolbar1.Buttons.Add(, "I_Grabar   ", , tbrDefault, "I_Grabar   "): BtnX.Visible = True: BtnX.ToolTipText = " "
    Set BtnX = Toolbar1.Buttons.Add(, "A_Grabar   ", , tbrDefault, "A_Grabar   "): BtnX.Visible = False: BtnX.ToolTipText = "Grabar Datos": BtnX.Enabled = IIf(Mid(ValidarUsuario(M_MinSR1), 2, 2) = "0", False, True)
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
    Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Copiar Planificación Teórica"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Enabled = False
   ' Set BtnX = Toolbar1.Buttons.Add(, "A_ExportaPlanif", , tbrDefault, "A_ExportaPlanif"): BtnX.Visible = False: Me.HelpContextID = 1032000: BtnX.Enabled = IIf(Mid(ValidarUsuarioAcceso(M_MinSR2), 1, 1) = "0", False, True): BtnX.ToolTipText = IIf(Mid(ValidarUsuarioAcceso(M_MinSR2), 1, 1) = "0", "", "Exportar Planificación Teórica")
'    Set BtnX = Toolbar1.Buttons.Add(, "A_ImportaPlanif", , tbrDefault, "A_ImportaPlanif"): BtnX.Visible = False: Me.HelpContextID = 1034000: BtnX.Enabled = IIf(Mid(ValidarUsuarioAcceso(M_MinSR2), 1, 1) = "0", False, True): BtnX.ToolTipText = IIf(Mid(ValidarUsuarioAcceso(M_MinSR2), 1, 1) = "0", "", "Importar Planificación")

    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False

    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False
'    Set BtnX = Toolbar1.Buttons.Add(, "A_Aportes", , tbrDefault, "A_Aportes"): BtnX.Visible = True: BtnX.ToolTipText = "Aportes Nutricionales x Días"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Aportes", , tbrDropdown, "A_Aportes"): BtnX.Visible = True: BtnX.ToolTipText = "Aportes Nutricionales x Días": BtnX.ButtonMenus.Add text:="Sin % P-G-Cho-Agrs": BtnX.ButtonMenus.Add text:="Con % P-G-Cho-Agrs"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Costo", , tbrDefault, "A_Costo"): BtnX.Visible = True: BtnX.ToolTipText = "Visualizar Costo"
    Set BtnX = Toolbar1.Buttons.Add(, "A_BuscarPro", , tbrDefault, "A_BuscarPro"): BtnX.Visible = False: BtnX.ToolTipText = "Gramos Productos Mensual"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Frecuencia", , tbrDefault, "A_Frecuencia"): BtnX.Visible = True: BtnX.ToolTipText = "Frecuencia Recetas"
    Set BtnX = Toolbar1.Buttons.Add(, "A_ActCostoReceta", , tbrDefault, "A_ActCostoReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Actualizar Costo Receta"
    Set BtnX = Toolbar1.Buttons.Add(, "A_ExporReceta", , tbrDefault, "A_ExporReceta"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar Recetas Excel"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
'    Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Planificación Minuta a Excel "
    Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDropdown, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exporta Minuta Bloque a Excel ": BtnX.ButtonMenus.Add text:="Formato I": BtnX.ButtonMenus.Add text:="Formato II Resumido"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "Ingrediente", , tbrDefault, "Ingrediente"): BtnX.Visible = True: BtnX.ToolTipText = "Frecuencia de Ingrediente"
    Set BtnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): BtnX.Visible = True: BtnX.Enabled = False: BtnX.ToolTipText = "Deshacer"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Retrocede", , tbrDefault, "A_Retrocede"): BtnX.Visible = True: BtnX.Enabled = True: BtnX.ToolTipText = "Retrocede Minuta Bloque"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Avanza", , tbrDefault, "A_Avanza"): BtnX.Visible = True: BtnX.Enabled = True: BtnX.ToolTipText = "Avanza Minuta Bloque"
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrSeparator, 0): BtnX.Visible = False: BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    Me.HelpContextID = vg_OpcM

   ' Label4.Caption = Trim(M_MinSR1.fpayuda(0).Caption) & "(" & M_MinSR1.fpText.text & ")" & " - " & Trim(M_MinSR1.fpayuda(1).Caption) & " - " & Trim(M_MinSR1.fpayuda(2).Caption) & " - " & IIf(vg_IDBloque = 0, "", vg_IDBloque)
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute(" sgpadm_sel_OfertasAsociadaalCeco " & M_MinSR1.fpText.text & "")
    If Not RS1.EOF Then
    
        Label4.Caption = Trim(M_MinSR1.fpayuda(0).Caption) & "(" & M_MinSR1.fpText.text & ")" & " - " & Trim(M_MinSR1.fpayuda(1).Caption) & " - " & Trim(M_MinSR1.fpayuda(2).Caption) & " - " & IIf(vg_IDBloque = 0, "", vg_IDBloque) & "   Ofertas Asociadas " & RS1!ofertas_asoc
     
    End If
    RS1.Close
    Set RS1 = Nothing
    
    Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "
    FormatearGrilla
    DetallePlantillaMinuta
    LlenarEstructuraServicio

End Sub

Sub LlenarEstructuraServicio()

Dim RS As New ADODB.Recordset
Dim i As Long
Dim X As Long
'------> Borrar opciones de menú
If Estructura1(0).Visible = False Then Estructura1(0).Visible = True: Estructura2(0).Visible = True
For i = 1 To Estructura1.count - 1
    
    Unload Estructura1(i)

Next i
For i = 1 To Estructura2.count - 1
    
    Unload Estructura2(i)

Next i
'-------> Mover datos menu estructura
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ListaEstservicioMinBloque_V02 " & vg_codservicio & "")
If Not RS.EOF Then
       
       X = 1
       
       Do While Not RS.EOF
          
          Load Estructura1(X)
          Load Estructura2(X)
          Estructura1(X).Caption = Trim(RS!ess_nombre)
          Estructura2(X).Caption = Trim(RS!ess_nombre)
          Estructura1(X).HelpContextID = RS!ess_codigo
          Estructura2(X).HelpContextID = RS!ess_codigo
          
          Estructura1(X).Enabled = True
          Estructura2(X).Enabled = True
          
          For i = 1 To vaSpread1.MaxRows
              
              vaSpread1.Col = vaSpread1.maxcols - 1: vaSpread1.Row = i
              If Trim(vaSpread1.text) <> "" Then
                 
                 If Val(vaSpread1.text) = RS!ess_codigo Then Estructura1(X).Enabled = False: Estructura2(X).Enabled = False
              
              End If
          
          Next
          
          X = X + 1
          RS.MoveNext
      
      Loop

End If
RS.Close
Set RS = Nothing

If Estructura1(0).Visible = True Then

   Estructura1(0).Visible = False
   Estructura2(0).Visible = False

End If

End Sub

Private Sub Form_Resize()

If Me.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
If Me.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 800 '675
If Me.WindowState <> 1 Then vaSpread1.Move 0, 1760, ScaleWidth, ScaleHeight - 1760

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If SwSalir <> 0 Then Exit Sub
    If IndGrabado <> 1 Then Me.Hide: Unload Me: M_MinSR1.WindowState = 0: Exit Sub
    If MsgBox(" Actualiza planificación bloque...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then IndGrabado = 0: Cancel = -1: Me.Hide: Unload Me: M_MinSR1.WindowState = 0
    If IndGrabado = 1 Then GrabarPlantillaMinuta
    IndGrabado = 0
    Plantilla(0).Enabled = False
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
    SwSalir = 1
    Me.Hide
    Unload Me
    M_MinSR1.WindowState = 0
    Set SpreadClon = Nothing

End Sub

Private Sub Image2_Click(Index As Integer)
    
    Image2(0).Enabled = False
    Image2(1).Enabled = False
    fg_carga ""
    Call fg_descarga
    Image2(0).Enabled = True
    Image2(1).Enabled = True

End Sub

Private Sub Plantilla_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS          As New ADODB.Recordset
Dim StrRec      As String
Dim StrRecb     As String
Dim aAp         As String
Dim Sql1        As String
Dim i           As Long
Dim j           As Long
Dim X           As Long
Dim CodRec      As Long
Dim tiprec      As Long
Dim cosali      As Double

Dim IndDia      As Long
Dim xcol        As Variant
Dim vecactrec   As Variant
Dim nrodia      As String
Dim MyBuffer    As String
Dim Cabecera As Long
Dim T As Long
Dim dato As Variant

'Dim NroDia      As Long
    Select Case Index
        
        Case 0 '-------> Grabar planificación
            
            If Toolbar1.Buttons(2).Enabled = False Then IndGrabado = 0: Exit Sub
            If MsgBox(" Actualiza planificación bloque...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
               
               IndGrabado = 0
               Plantilla(0).Enabled = False
               Toolbar1.Buttons(1).Visible = True
               Toolbar1.Buttons(2).Visible = False
               Toolbar1.Buttons(31).Enabled = False
               Cancel = -1
               Exit Sub
            
            End If
            
            If IndGrabado = 1 Then GrabarPlantillaMinuta
            IndGrabado = 0
            Plantilla(0).Enabled = False
            Toolbar1.Buttons(1).Visible = True
            Toolbar1.Buttons(2).Visible = False
            Toolbar1.Buttons(31).Enabled = False
        
        Case 5 '-------> Ver detalle recetas

            If vaSpread1.MaxRows = vaSpread1.ActiveRow Then Exit Sub 'jpaz
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            If vaSpread1.Col = 1 Or vaSpread1.Col = 2 Then Exit Sub
            If vaSpread1.BackColor = Shape1(1).FillColor Then
                
                vg_newestrec = True
            
            Else
                
                vg_newestrec = False
            
            End If
        
            xcol = 0
            For i = 1 To MaxColumna
                
                If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or _
                   vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 3) Or vectorcol(i) = (vaSpread1.Col - 4) Or vectorcol(i) = (vaSpread1.Col - 5)) _
                   And Trim(vaSpread1.text) <> "" Then
                   
                   xcol = vectorcol(i): Exit For
                
                End If
            
            Next i
            
            If xcol <= 0 Then MsgBox "Debe seleccionar la columna de receta", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
            vg_newnomrec = ""
            vaSpread1.Col = xcol + 4
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
            Dim Receta As New M_Receta
            Let VarSitioRemoto = True
            vaSpread1.Col = xcol: vaSpread1.SetActiveCell xcol, vaSpread1.Row
            vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin

            Receta.Show 1, Me
            Me.Refresh
            Let VarSitioRemoto = False
            Toolbar1.Refresh
            vg_newestrec = False
            If vg_newcodrec <> 0 And Trim(vg_newnomrec) <> "" And (vg_auxtiprec = vg_tiprec) Then
                
                vaSpread1.Col = xcol + 4
                vaSpread1.Row = vaSpread1.ActiveRow
                vaSpread1.Col = xcol
                '-------> Limpiar Datos y Formato Celda
                vaSpread1.Action = 3
                '------- Retorna Modo de la columna
                vaSpread1.BlockMode = False
                vaSpread1.Font.Bold = False
                vaSpread1.Font.Size = 8
                vaSpread1.text = vg_newnomrec
                '-------> Mover codigo receta
                vaSpread1.Col = xcol + 4
                vaSpread1.text = vg_newcodrec & "&" & vg_tiprec & "&;"
                
                If indcos = True Then
                   
                   For i = 2 To (vaSpread1.maxcols - 2) Step 7
                       
                       Calctodia 1, i + 1
                   
                   Next i
                   MostrarCosto vaSpread1.ActiveCol
                
                End If
                
                For i = 3 To (vaSpread1.maxcols - 2) Step 7
                    
                    CalctodiaEnc 1, i + 1
                
                Next i
                IndGrabado = 1
                vg_newcodrec = 0: vg_newnomrec = "": vg_tiprec = -2
                Plantilla(0).Enabled = True
                Toolbar1.Buttons(1).Visible = False
                Toolbar1.Buttons(2).Visible = True
                Toolbar1.Buttons(6).Visible = True
                Toolbar1.Buttons(7).Visible = False
            
            End If
            vg_newcodrec = 0
            If indcos = True Then Me.Refresh: Toolbar1.Refresh: Frame2(0).Refresh: Frame2(1).Refresh: Frame2(2).Refresh: Frame2(3).Refresh: Frame2(4).Refresh
        
        Case 8 '-------> Copiar planificación
            M_CPlaTe.Show 1, Me
        
        Case 10 '-------> Calcular Aporte Día
        
        Case 11 '-------> Costo recetas
            
            If Frame2(0).Visible = True Then
               
               Frame2(0).Visible = False
               vaSpread1.Move 0, 1760, ScaleWidth, ScaleHeight - 1760
               Image2(0).Visible = False: Image2(1).Visible = False
               indcos = False
               Exit Sub
            
            End If
            vaSpread1.Move 0, 1760, ScaleWidth, ScaleHeight - 5000
            Frame2(0).Move 0, ScaleHeight - 3200, ScaleWidth, ScaleHeight - 1200
            Frame2(0).Visible = True
'            CmdCerrar.Left = 19900
            
            '-------> Ocultar boton calculo Lyd
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("select ser_LYD from a_servicio WITH ( NOLOCK ) where ser_codigo = " & vg_codservicio & " and isnull(ser_LYD,0) = 1")
            If Not RS.EOF Then
               
               Actualizar_LYD.Visible = False
            
            Else
               
               Actualizar_LYD.Visible = True
            
            End If
            RS.Close
            Set RS = Nothing
            
            '-------> Cargar costo
            CargarCosto False
        
        Case 12 '-------> Frecuencia recetas en planificación
            
            If IndGrabado = 1 Then
                
                MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo
                Exit Sub
            
            End If
            
'            Let VarSitioRemoto = True
'            Call C_FrePla.LlenarFrecPlan("Frecuencia Minuta Bloque ", vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), 1, Val(Vg_FechaHasta))
'            C_FrePla.Show 1, Me
'            Let VarSitioRemoto = False
            
            Let VarSitioRemoto = True
            Call C_FreMinBlo.LlenarFrecPlan("Frecuencia Minuta Bloque ", vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), 1, Val(Vg_FechaHasta))
            C_FreMinBlo.Show 1, Me
            Let VarSitioRemoto = False
            
        Case 13 '-------> Actualizar Costo recetas
            
            Dim estact As Boolean
            estact = False
            If IndGrabado = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
            Toolbar1.Enabled = False
            vaSpread1.Refresh
            DetallePlantillaMinuta
            If IndGrabado = 1 Then
                
                Call fg_descarga
                MsgBox "Actualización costo receta finalizado sin problema, luego grabe información", vbInformation + vbOKOnly, MsgTitulo
                Plantilla(0).Enabled = True
                Toolbar1.Buttons(1).Visible = False
                Toolbar1.Buttons(2).Visible = True
                Toolbar1.Enabled = True
                Exit Sub
            
            End If
            Toolbar1.Enabled = True
            Call fg_descarga
        
        Case 14 '-------> Exportar recetas
            
            If IndGrabado = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo: Exit Sub
            Call C_ExpRecMBloque.LlenarExporRecetaBloque("Exportar Recetas " & fg_Ctod1(Vg_FechaDesde) & " - " & fg_Ctod1(Vg_FechaHasta), vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), Val(Vg_FechaHasta))
            C_ExpRecMBloque.Show 1, Me
        
        Case 17
            
            If IndGrabado = 1 Then
               
               MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo
               Exit Sub
            
            End If
            Call I_matrizdeprecios(vg_codcasino, Vg_FechaDesde, Vg_FechaHasta, vg_IDBloque)
        
        Case 20 '-------> Salir de planificación
            
            SwSalir = 0
            If Toolbar1.Buttons(2).Enabled = False Then IndGrabado = 0
            If IndGrabado <> 1 Then SwSalir = 1: Me.Hide: Unload Me: M_MinSR1.WindowState = 0: Exit Sub
            If MsgBox(" Actualiza planificación teórica...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then IndGrabado = 0
            If IndGrabado = 1 Then GrabarPlantillaMinuta
            IndGrabado = 0
            SwSalir = 1: Me.Hide: Unload Me: M_MinSR1.WindowState = 0
    
        Case 19 'MVA - MVI - ACTIVAR EL FORMULARIO QUE COPIA LA MINUTA CON ENCABEZADO CCOSTO, REGIMEN Y SERVICIO
            
            m_copia_min_seg.Show 1, Me
    
    End Select

Exit Sub
Man_Error:
    If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
    If Err = 3034 Then Exit Sub
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Plato_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS                  As New ADODB.Recordset
Dim Del_Row             As Integer
Dim c                   As Long
Dim IndCol              As Integer
Dim indrow              As Integer
Dim indcol2             As Integer
Dim indrow2             As Integer
Dim indrow3             As Integer
Dim FilaAct             As Long
Dim FilaAnt             As Long
Dim FilaPos             As Long
Dim AuxIblockrow        As Integer
Dim addrec              As Long
Dim codest              As Long
Dim cosali              As Double
Dim CosDes              As Double
Dim NroMes              As String
Dim FinGrilla           As String
Dim MesInicio           As String
Dim FechaBusqueda       As String
Dim SumaMes             As String
Dim MesInicio3          As String
Dim xx                  As Long
Dim xp                  As Long
Dim FechaDia            As Long
Dim SeleccionOpt        As Long
Dim CodGrupoEstBaj      As Long
Dim CantTotalPorcentaje As Double
            
Dim VecSelGrid          As Variant
Dim VecRacPegar         As Variant
Dim contador            As Long
Dim contador_b          As Long
Dim cantCol             As Long
Dim LargoVec            As Long
Dim accion              As String
Dim ColumnaActiva       As Long
Dim FilaActiva          As Long
Dim ColumnaAntActiva    As Long
Dim n                   As Long
Dim n1                  As Long
Dim NFilas              As Long
Dim CantCol1            As Long
Dim d                   As Variant
Dim Max                 As Long
Dim max1                As Long
Dim ff                  As Long
Dim f                   As Long
Dim desc                As String
Dim g                   As Long
Dim j                   As Long
Dim tope                As Long
Dim jjj                 As Long

contador = 0
contador_b = 0
cantCol = 0
LargoVec = 0
accion = ""
n1 = 0
n = 0
NFilas = 0

'Fila Maestra de Grupo
    
    
    Select Case Index
        
        Case 2 '-------> Ingresar recetas
            
            iblockcol = vaSpread1.ActiveCol: aiblockcol = vaSpread1.ActiveCol
            iblockcol2 = vaSpread1.ActiveCol: aiblockcol2 = vaSpread1.ActiveCol
            iblockrow = vaSpread1.ActiveRow: aiblockrow = vaSpread1.ActiveRow
            iblockrow2 = vaSpread1.ActiveRow: aiblockrow2 = vaSpread1.ActiveRow
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
                
            '-------> Validar si minuta esta bloqueada
            If ValidarBloqueoMinutaDetalle(vaSpread1.Row, vaSpread1.Col) Then Exit Sub
            
            vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
            '-------> Validar que no sea la ultima columna
            If xRowIni = vaSpread1.MaxRows Then Exit Sub
            
            vaSpread1.Col = xColIni
            j = 0
            
            For i = 1 To MaxColumna
                
                If vaSpread1.Col = vectorcol(i) Then j = vectorcol(i): Exit For
            
            Next i
            
            If j = 0 Then Exit Sub
            vg_codigo = "": vg_nombre = "": vg_tiprec = -2
            '-------> Validar receta 5 etapa
        
            vaSpread1.Col = j - 1
            vaSpread1.Row = vaSpread1.ActiveRow
            AadRec = 0
            Let VarSitioRemoto = True
            B_RecMBi.Show 1, Me
            Let VarSitioRemoto = False
            
           'INI ARI
           ' Asocia digo de Grupo y Estructura de Servicio segun la estructura anterior
            
            Dim rowposicion As Long
            Dim nombreEstructura As String
            Dim colgrupo As Long
            Dim CodigoGrupo As Long
            Dim estructuraservicio As Long
            Dim NomReceta As String
            Dim LyD As Boolean
            
            rowposicion = vaSpread1.ActiveRow
            
            For i = vaSpread1.ActiveRow To 1 Step -1
                
                vaSpread1.Row = i
                vaSpread1.Col = 1
                nombreEstructura = vaSpread1.text
              
              If nombreEstructura <> "" Then
                    
                    'columna de grupo de estructura y Encabezado
                    colgrupo = vaSpread1.GetColFromID("Grupo") + 1
                    'cual es el grupo actual y si el Encabezado del Grupo
                    vaSpread1.Col = colgrupo
                    CodigoGrupo = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
                    
                    vaSpread1.Col = vaSpread1.maxcols - 1
                    estructuraservicio = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
                    
                    'Fila Maestra de Grupo
                    'RowGrupo = vaSpread1.SearchCol(colgrupo, 0, -1, CodigoGrupo, SearchFlagsValue)
                    vaSpread1.Row = rowposicion
                    vaSpread1.Col = colgrupo
                    vaSpread1.text = CodigoGrupo
                    
                    vaSpread1.Col = vaSpread1.maxcols - 1
                    vaSpread1.text = estructuraservicio
                    Exit For
                   
              End If
                 
            Next
            vaSpread1.Row = rowposicion
               
           'FIN ARI
            
            If Trim(vg_codigo) = "" Or Trim(vg_nombre) = "" Then Exit Sub
            vg_tiprec = IIf(vg_codregimen > 9999 And vg_codservicio > 9999, vg_codregimen, -1)
             
            GrabarCambios 1, 1, ""
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = j - 1
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignCenter
            vaSpread1.Value = "R"
            vaSpread1.ForeColor = &HFF&
            vaSpread1.BackColor = &H80FF80
        
            vaSpread1.Col = j

            '-------> Limpiar Datos y Formato Celda
            vaSpread1.Action = 3
            '-------> Retorna Modo de la columna
            vaSpread1.BlockMode = False
            vaSpread1.Font.Bold = False
            vaSpread1.Font.Size = 8
            vaSpread1.text = vg_nombre
            vaSpread1.BackColor = Shape1(0).FillColor
                      
            '-------> Calcular costo receta
            vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
            vaSpread1.Col = xColIni
            vaSpread1.Row = SpreadHeader + 3
            LyD = False
            NomReceta = ""
            FechaDia = CLng(Format(Mid(vaSpread1.text, 5, Len(vaSpread1.text)), "yyyymmdd"))
            SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_ResumenCostoxReceta_V02 '" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & vg_codigo & ", " & FechaDia & ", " & SeleccionOpt & "")
            If Not RS.EOF Then
               
               LyD = RS!rec_LYD
               vg_Valor = RS!promedioreceta
               NomReceta = RS!rec_nombre
            
            End If
            RS.Close
            Set RS = Nothing
            
            vaSpread1.Row = vaSpread1.ActiveRow
            
            '-------> Limpiar Datos y Formato Celda
            vaSpread1.Action = 3
            '-------> Retorna Modo de la columna
            vaSpread1.BlockMode = False
            vaSpread1.Font.Bold = False
            vaSpread1.Font.Size = 8
            vaSpread1.text = IIf(LyD, "[*] ", "") & Trim(NomReceta) 'vg_nombre
            vaSpread1.BackColor = Shape1(0).FillColor
             
            '-------> Porcentaje del dia
            vaSpread1.Col = j + 1
'            If Trim(vaSpread1.text) = "" And Not LyD Then
            If Trim(vaSpread1.text) = "" And Not LyD Then
               
               vaSpread1.CellType = CellTypePercent
               vaSpread1.text = 0
               vaSpread1.ForeColor = &HFF0000
               vaSpread1.CellType = CellTypePercent
               vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
               vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
               vaSpread1.TypePercentDecPlaces = 0
               vaSpread1.TypePercentMax = 1000
               ' display negative numbers as red
               vaSpread1.TypeNegRed = True
            
            ElseIf LyD Then
               
               vaSpread1.CellType = CellTypeStaticText
               vaSpread1.text = ""
            
            End If
            
            '-------> Mover raciones
            vaSpread1.Col = j + 2
            If Trim(vaSpread1.text) = "" Then
               
               vaSpread1.Row = vaSpread1.ActiveRow
               vaSpread1.Col = j + 2
               vaSpread1.CellType = IIf(TipoMinuta Or LyD, CellTypeNumber, CellTypeStaticText)
               vaSpread1.TypeNumberDecPlaces = 0
               vaSpread1.TypeNumberMin = 0
               vaSpread1.TypeNumberMax = 9999999
               vaSpread1.TypeHAlign = 1
               vaSpread1.TypeSpin = False
               vaSpread1.TypeIntegerSpinInc = 1
               vaSpread1.TypeIntegerSpinWrap = False
               vaSpread1.text = 0
               vaSpread1.ForeColor = &HFF0000
            
            End If
            vaSpread1.BackColor = Shape1(0).FillColor
            
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = j + 3
            vaSpread1.CellType = CellTypeStaticText
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.text = Format(vg_Valor, fg_Pict(6, 2))
            vaSpread1.BackColor = Shape1(0).FillColor
            
            vaSpread1.Col = j + 4
            '-------> Traer codigo recetas
            If vaSpread1.text <> Val(vg_codigo) & "&" & vg_tiprec & "&;" Or Trim(vaSpread1.text) = "" Then
               
               '-------> Mover tipo celda nota cuando sucede un cambio
               vaSpread1.TextTip = TextTipFloating
               vaSpread1.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
               vaSpread1.CellNote = "Cambio"
            
            End If
            
            vaSpread1.text = Val(vg_codigo) & "&" & vg_tiprec & "&;"
            
            '-------> Mover calorias
            vaSpread1.Col = j + 5
            vaSpread1.TypeHAlign = TypeHAlignRight
            vaSpread1.text = Format(vg_Calorias, fg_Pict(6, 2))
            
            If indcos = True Then Calctodia vaSpread1.Row, IIf(TipoMinuta, j, j + 1)
            CalctodiaEnc vaSpread1.Row, IIf(TipoMinuta, j, j + 1)
            
            vaSpread1.Row = vaSpread1.ActiveRow
            IndGrabado = 1
            Plato(0).Enabled = True: OpGrilla(0).Enabled = True
            Plato(13).Enabled = False: OpGrilla(13).Enabled = False
            Plato(14).Enabled = False: OpGrilla(14).Enabled = False
            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False

        Case 5 '-------> Insertar linea

            '-------> Validar si día esta bloqueado
            If ValidarBloqueoMinuta Then Exit Sub
            
            IndCol = iblockcol
            iblockcol = 1: iblockcol2 = vaSpread1.maxcols
            vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
            GrabarCambios Val(xRowIni), Val(xRowFin), "Insertar"
            vaSpread1.MaxRows = vaSpread1.MaxRows + ((xRowFin - xRowIni) + 1) '1
            vaSpread1.InsertRows xRowIni, ((xRowFin - xRowIni) + 1)
'            '-------> Rescatar codigo estructura y grupo anterior
'            Dim CodEstrIns As Long
'            Dim CodGrupoEstrIns As Long
'            CodEstrIns = 0
'            CodGrupoEstrIns = 0
'            If xRowIni - 1 > 0 Then
'               vaSpread1.Row = xRowIni - 1
'               vaSpread1.Col = vaSpread1.maxcols - 1
'               CodEstrIns = vaSpread1.text
'               vaSpread1.Col = vaSpread1.maxcols
'               CodGrupoEstrIns = vaSpread1.text
'            End If
            Do While xRowIni <= xRowFin
               
               vaSpread1.Row = xRowIni: vaSpread1.Col = -1
               vaSpread1.BackColor = Shape1(0).FillColor
            
               vaSpread1.Row = xRowIni
               vaSpread1.Col = 1
               vaSpread1.Font.Bold = True
               vaSpread1.Font.Size = 9
               vaSpread1.BackColor = Shape1(2).FillColor
               
'               '-------> Mover valores codigo estructura y grupo
'               If CodEstrIns > 0 Then
'                  vaSpread1.Col = vaSpread1.maxcols - 1
'                  vaSpread1.text = CodEstrIns
'                  vaSpread1.Col = vaSpread1.maxcols
'                  vaSpread1.text = CodGrupoEstrIns
'               End If
                
                For i = 3 To (vaSpread1.maxcols) Step 7
                    
                    vaSpread1.Row = vaSpread1.MaxRows
                    vaSpread1.Col = i + 1
                    
                    If vaSpread1.Lock = True Then
                        
                        For c = i - 1 To i + 2
                            
                            vaSpread1.Row = xRowIni: vaSpread1.Col = c
                            vaSpread1.BackColor = Shape1(1).FillColor
                        
                        Next c
                    
                    End If
                
                Next i
               
               xRowIni = xRowIni + 1
            
            Loop
            iblockcol = IndCol
            IndGrabado = 1
            Plato(0).Enabled = True: OpGrilla(0).Enabled = True
            Plato(13).Enabled = False: OpGrilla(13).Enabled = False
            Plato(14).Enabled = False: OpGrilla(14).Enabled = False
            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
            
        Case 6 '-------> Eliminar linea
            
            vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
            IndCol = iblockcol
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            
            '-------> Validar si minuta esta bloqueada
            If ValidarBloqueoMinuta Then Exit Sub
            
            '-------> Validar si va eliminar ultima columna
            If xRowIni = vaSpread1.MaxRows Then
               
               Call MsgBox("No puede eliminar ultima fila", vbCritical + vbOKOnly, MsgTitulo)
               Exit Sub
            
            End If
            '-------> Validar columna
            Dim Porcentaje As Double
            Dim CodGrupoEstructura As Long
            If iblockcol = -1 Then
               
               '-------> Validar que no existan datos para los días siguientes
               For j = xRowIni To xRowFin
               
               For i = 2 To vaSpread1.maxcols - 2 Step 7
                   
                   vaSpread1.Row = j 'xRowIni
                   vaSpread1.Col = i + 1
                   
                   If Trim(vaSpread1.text) <> "" Then
                      
                      If MsgBox("Existen recetas los demas días, desea eliminar fila?...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                         
                         Exit Sub
                      
                      Else
                         
                         Exit For
                      
                      End If
                   
                   End If
               
               Next i
               
               Next j
               vaSpread1.Visible = False
               '-------> Grabar evento de cambio
               GrabarCambios Val(iblockrow), Val(1), "Eliminar"
               For i = xRowIni To IIf(vaSpread1.MaxRows = xRowFin, xRowFin - 1, xRowFin)
                   
                   vaSpread1.Row = i
                   vaSpread1.Col = 2
                   
                   If vaSpread1.text <> "" Then
                      
                      Porcentaje = Val(vaSpread1.text)
                      vaSpread1.Col = vaSpread1.maxcols
                      CodGrupoEstructura = vaSpread1.text
                      MoverPorcentaje Porcentaje, CodGrupoEstructura
                   
                   End If
               
               Next i
               
               vaSpread1.DeleteRows xRowIni, (IIf(vaSpread1.MaxRows = xRowFin, xRowFin - 1, xRowFin) - xRowIni + 1)
               vaSpread1.MaxRows = vaSpread1.MaxRows - (IIf(vaSpread1.MaxRows = xRowFin, xRowFin - 1, xRowFin) - xRowIni + 1)
               vaSpread1.Visible = True
               IndGrabado = 1
'               Exit Sub
            
            ElseIf Not TipoMinuta Then
               
               '-------> Validar que no seleccione la estructura y % porcentaje estructura
               If (xColIni = 1 Or xColIni = 2) Then
                  
                  Call MsgBox("No puede eliminar estructura o bien % estructura", vbCritical + vbOKOnly, MsgTitulo)
                  Exit Sub
               
               End If
               '-------> Grabar evento de cambio
               GrabarCambios Val(iblockrow), Val(1), "Eliminar"
               '-------> Traer primera columna
               aiblockcol = PrimeraColumna(xColIni, MaxColumna)
               '-------> Traer fin columna
               iblockcol2 = FinalColumna(xcolfin, MaxColumna)
               vaSpread1.ClearRange iblockcol, xRowIni, IIf(Not TipoMinuta, iblockcol2, iblockcol), xRowFin, False
           
           ElseIf TipoMinuta Then
               
               '-------> Validar que no existan datos para los días siguientes
               
               For j = xRowIni To xRowFin
               
               For i = 2 To vaSpread1.maxcols - 2 Step 7
                   
                   vaSpread1.Row = j 'xRowIni
                   vaSpread1.Col = i + 1
                   
                   If Trim(vaSpread1.text) <> "" Then
                      
                      If MsgBox("Existen recetas los demas días, desea eliminar fila?...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
                         
                         Exit Sub
                      
                      Else
                         
                         Exit For
                      
                      End If
                   
                   End If
               
               Next i
               
               Next j
               vaSpread1.DeleteRows xRowIni, (IIf(vaSpread1.MaxRows = xRowFin, xRowFin - 1, xRowFin) - xRowIni + 1)
               vaSpread1.MaxRows = vaSpread1.MaxRows - (IIf(vaSpread1.MaxRows = xRowFin, xRowFin - 1, xRowFin) - xRowIni + 1)
               vaSpread1.Col = vaSpread1.Row
               vaSpread1.Visible = True
           
           End If
           
           If iblockcol <> -1 Then
              
              If indcos = True Then
                  
                  For i = iblockcol To iblockcol2 Step 7
                      
                      Calctodia 1, i + 1
                  
                  Next i
                  
                  MostrarCosto vaSpread1.ActiveCol
              
              End If
              
              For i = iblockcol To iblockcol2 Step 7
                  
                  CalctodiaEnc 1, i + 1
              
              Next i
          
          Else
              
              If indcos = True Then
                  
                  For i = 3 To vaSpread1.maxcols - 2 Step 7
                      
                      Calctodia 1, i + 2 '1
                  
                  Next i
                  
                  MostrarCosto vaSpread1.ActiveCol
              
              End If
              
              For i = 3 To vaSpread1.maxcols - 2 Step 7
                  
                  CalctodiaEnc 1, i + 1
              
              Next i
          
          End If
            
            indactivo = 0
            IndGrabado = 1
            Plato(0).Enabled = True: OpGrilla(0).Enabled = True
            Plato(13).Enabled = False: OpGrilla(13).Enabled = False
            Plato(14).Enabled = False: OpGrilla(14).Enabled = False
            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
        
        Case 8 '-------> Subir linea
            
            iblockcol = vaSpread1.ActiveCol: aiblockcol = vaSpread1.ActiveCol
            iblockcol2 = vaSpread1.ActiveCol: aiblockcol2 = vaSpread1.ActiveCol
            iblockrow = vaSpread1.ActiveRow: aiblockrow = vaSpread1.ActiveRow
            iblockrow2 = vaSpread1.ActiveRow: aiblockrow2 = vaSpread1.ActiveRow
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            If vaSpread1.Row = 1 Or vaSpread1.Col = 2 Or vaSpread1.Row = vaSpread1.MaxRows Then Exit Sub 'jpaz
            If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
               
               For i = 1 To MaxColumna
                   
                   vaSpread1.Col = vectorcol(i)
                   vaSpread1.Row = 1
                    
                    If vaSpread1.BackColor = Shape1(1).FillColor Then
                        
                        Call MsgBox("Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo)
                        Exit Sub
                    
                    End If
               
               Next i
            
            Else
                
                For i = iblockcol To iblockcol2
                    
                    vaSpread1.Col = i
                    
                    For j = iblockrow To iblockrow2
                        
                        vaSpread1.Row = j
                        
                        If vaSpread1.BackColor = Shape1(1).FillColor Then
                            
                            Call MsgBox("Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo)
                            Exit Sub
                        
                        End If
                   
                   Next j
               Next i
            End If
            
            vaSpread1.Col = vaSpread1.ActiveCol
            
            If vaSpread1.Col > 1 Then
                
                IndCol = iblockcol
                vaSpread1.Col = 1
                
                If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
                
                If (iblockrow - ((iblockrow2 - iblockrow) + 1)) < 1 Then
                   
                   MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
                   Exit Sub
                
                End If
                
                If vaSpread1.MaxRows > 1000 Then Del_Row = vaSpread1.MaxRows - 1000: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
                If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.maxcols
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or _
                       (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Or _
                       (vectorcol(i) + 3) = iblockcol Or (vectorcol(i) + 4) = iblockcol Or (vectorcol(i) + 5) = iblockcol Then
                       
                       iblockcol = (vectorcol(i) - 1)
                       Exit For
                    
                    End If
                
                Next i
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = iblockcol2 Then
                       
                       iblockcol2 = ((vectorcol(i) + 4))
                       Exit For
                    
                    End If
                    
                    If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Or _
                      vectorcol(i) + 3 = iblockcol2 Or vectorcol(i) + 4 = iblockcol2 Or vectorcol(i) + 5 = iblockcol2 Then
                       
                       iblockcol2 = (vectorcol(i) + 5)
                       Exit For
                    
                    End If
                
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
                '-------> Devolver datos fila y restar ultima fila
                vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
                vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
                vaSpread1.Row = iblockrow - 1: vaSpread1.Col = iblockcol
                vaSpread1.DeleteRows vaSpread1.MaxRows, 1
                vaSpread1.MaxRows = vaSpread1.MaxRows - 1
                vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
            
            ElseIf vaSpread1.Col = 1 Then
                
                If Trim(vaSpread1.text) = "" Then Exit Sub
                
                For i = iblockrow - 1 To 1 Step -1 '-------> Recorre el espacio que hay entre la estructura seleccioneda y la anterior
                    
                    vaSpread1.Row = i
                    If Trim(vaSpread1.text) <> "" Then Exit For
                
                Next i
                
                For z = iblockrow + 1 To (vaSpread1.MaxRows - 1) '-------> 100 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
                    
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
        '                If z <= (vaSpread1.MaxRows - 1) Then Exit For
                        If z <= (vaSpread1.MaxRows) Then Exit For
                    
                    Next fil
                
                End If
                
                FilaAct = iblockrow         'Fila actual
                FilaAnt = IIf(i < 1, 1, i)  'Fila anterior
                FilaPos = z                 'Fila posterior
                
                If Not TipoMinuta Then

'                   Dim CodGrupoEstBaj As Long
'                   Dim CantTotalPorcentaje As Double
                   vaSpread1.Row = FilaAct
                   vaSpread1.Col = vaSpread1.maxcols
                   CodGrupoEstBaj = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
                   vaSpread1.Row = FilaAnt
                   
                   If CodGrupoEstBaj <> Val(vaSpread1.text) Then
                      
                      Call MsgBox("Grupo estructura es distinto, no puede mover estructura", vbCritical + vbOKOnly, MsgTitulo)
                      Exit Sub
                   
                   End If
                   
                   vaSpread1.Row = FilaAnt
                   vaSpread1.Col = 2
                   
                   If vaSpread1.CellType = CellTypePercent Then
                      
                      If vaSpread1.BackColor = Shape1(2).FillColor Then
                         
                         vaSpread1.BackColor = Shape1(0).FillColor
                      
                      End If
                      
                      CantTotalPorcentaje = IIf(Trim(vaSpread1.text) = "", 0, Val(vaSpread1.text))
                      vaSpread1.Row = FilaAct
                      vaSpread1.CellType = CellTypePercent
                      vaSpread1.TypeHAlign = 1
                      vaSpread1.TypePercentDecPlaces = 0
                      vaSpread1.ForeColor = &HFF0000
                      vaSpread1.text = CantTotalPorcentaje
                      vaSpread1.TypeNegRed = True
                      
                      If vaSpread1.BackColor = Shape1(0).FillColor Then
                         
                         vaSpread1.BackColor = Shape1(2).FillColor
                      
                      End If
                      
                      '-------> Mover valor cero
                      vaSpread1.Row = FilaAnt
                      vaSpread1.Col = 2
                      vaSpread1.text = ""
                      vaSpread1.CellType = CellTypeStaticText
                      
                   End If
                
                End If
                
                '-------> Agregar filas temporales y respaldar
                vaSpread1.MaxRows = vaSpread1.MaxRows + (FilaAct - FilaAnt)
                For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows + (FilaAct - FilaAnt)
                    
                    vaSpread1.Row = i
                    vaSpread1.RowHidden = True
                
                Next i
                vaSpread1.MoveRange 1, FilaAnt, vaSpread1.maxcols, IIf((FilaAct - 1) = 0, 1, (FilaAct - 1)), 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1
                
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
            IndGrabado = 1
            Plato(0).Enabled = True: OpGrilla(0).Enabled = True
            Plato(13).Enabled = False: OpGrilla(13).Enabled = False
            Plato(14).Enabled = False: OpGrilla(14).Enabled = False
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
        
        Case 9 '-------> Bajar linea
            
            iblockcol = vaSpread1.ActiveCol: aiblockcol = vaSpread1.ActiveCol
            iblockcol2 = vaSpread1.ActiveCol: aiblockcol2 = vaSpread1.ActiveCol
            iblockrow = vaSpread1.ActiveRow: aiblockrow = vaSpread1.ActiveRow
            iblockrow2 = vaSpread1.ActiveRow: aiblockrow2 = vaSpread1.ActiveRow
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            If vaSpread1.Row = vaSpread1.MaxRows Or vaSpread1.Col = 2 Then Exit Sub
            
            If iblockcol < 1 Or (iblockcol = 1 And Trim(vaSpread1.text) <> "") Then
                
                For i = 1 To MaxColumna
                    
                    vaSpread1.Col = vectorcol(i)
                    vaSpread1.Row = 1
                    
                    If vaSpread1.BackColor = Shape1(1).FillColor Then
                        
                        Call MsgBox("Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo)
                        Exit Sub
                    
                    End If
               
               Next i
            
            Else
               
               For i = iblockcol To iblockcol2
                   
                   vaSpread1.Col = i
                   
                   For j = iblockrow To iblockrow2
                        
                        vaSpread1.Row = j
                        
                        If vaSpread1.BackColor = Shape1(1).FillColor Then
                            Call MsgBox("Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo)
                            Exit Sub
                        
                        End If
                   
                   Next j
               
               Next i
            
            End If
            vaSpread1.Col = vaSpread1.ActiveCol
            GrabarCambios vaSpread1.Row, j, "Bajar Linea"
            
            If vaSpread1.Col > 1 Then
                
                vaSpread1.Col = 1
                vaSpread1.Row = vaSpread1.ActiveRow + 1
                If Trim(vaSpread1.text) <> "" Then MsgBox "No puede salirse del rango...", vbCritical + vbOKOnly, MsgTitulo: Exit Sub
                vaSpread1.Row = vaSpread1.ActiveRow - 1
                If (iblockrow2 + ((iblockrow2 - iblockrow) + 1)) > (vaSpread1.MaxRows - 1) Then '100 Then
                   
                   MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
                   Exit Sub
                
                End If
                IndCol = iblockcol
                If iblockcol < 0 Then iblockcol = 1: iblockcol2 = vaSpread1.maxcols
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or _
                       (vectorcol(i) + 2) = iblockcol Or (vectorcol(i) + 3) = iblockcol Or (vectorcol(i) + 4) = iblockcol Or (vectorcol(i) + 5) = iblockcol Then
                       
                       iblockcol = (vectorcol(i) - 1)
                       Exit For
                    
                    End If
                
                Next i
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = iblockcol2 Then
                       
                       iblockcol2 = ((vectorcol(i) + 4))
                       Exit For
                    
                    End If
                    
                    If vectorcol(i) = iblockcol2 Or vectorcol(i) + 1 = iblockcol2 Or vectorcol(i) + 2 = iblockcol2 Or vectorcol(i) + 3 = iblockcol2 Or _
                       vectorcol(i) + 4 = iblockcol2 Or vectorcol(i) + 5 = iblockcol2 Then
                       
                       iblockcol2 = (vectorcol(i) + 5)
                       Exit For
                    
                    End If
                
                Next i
                '-------> Copiar datos ultima fila
                vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                vaSpread1.Row = vaSpread1.MaxRows
                vaSpread1.RowHidden = True 'jpaz
                vaSpread1.MoveRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), iblockcol, vaSpread1.MaxRows
            
                '-------> Copiar datos fila Seleccionada
                vaSpread1.ClearRange iblockcol, (iblockrow + 1), iblockcol2, (iblockrow + 1), False
                vaSpread1.MoveRange iblockcol, iblockrow, iblockcol2, iblockrow, iblockcol, (iblockrow + 1)
            
                '-------> Devolver datos fila y restar ultima fila
                vaSpread1.ClearRange iblockcol, iblockrow, iblockcol2, iblockrow, False
                vaSpread1.MoveRange iblockcol, vaSpread1.MaxRows, iblockcol2, vaSpread1.MaxRows, iblockcol, iblockrow
                vaSpread1.DeleteRows vaSpread1.MaxRows, 1
                vaSpread1.MaxRows = vaSpread1.MaxRows - 1
                vaSpread1.Row = iblockrow + 1: vaSpread1.Col = iblockcol
                vaSpread1.SetActiveCell vaSpread1.Col, vaSpread1.Row
            
            ElseIf vaSpread1.Col = 1 Then
                
                If Trim(vaSpread1.text) = "" Then Exit Sub
                
                For z = iblockrow + 1 To (vaSpread1.MaxRows - 1) '100 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
                    
                    vaSpread1.Row = z
                    If Trim(vaSpread1.text) <> "" Then Exit For
                
                Next z
                
                If z > (vaSpread1.MaxRows - 1) Then Exit Sub
                vaSpread1.Col = vaSpread1.ActiveCol
                AuxIblockrow = z
                
                For i = AuxIblockrow - 1 To 1 Step -1 'Recorre el espacio que hay entre la estructura seleccioneda y la anterior
                    
                    vaSpread1.Row = i
                    If Trim(vaSpread1.text) <> "" Then Exit For
                
                Next i
                
                For z = AuxIblockrow + 1 To (vaSpread1.MaxRows - 1) '100 'Recorre el espacio que hay entre la estructura seleccioneda y la posterior
                    
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
                
                FilaAct = AuxIblockrow      'Fila actual
                FilaAnt = IIf(i < 1, 1, i)  'Fila anterior
                FilaPos = z                 'Fila posterior
                
                If Not TipoMinuta Then
                   
                   vaSpread1.Row = FilaAct
                   vaSpread1.Col = vaSpread1.maxcols
                   CodGrupoEstBaj = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
                   vaSpread1.Row = FilaAnt
                   
                   If CodGrupoEstBaj <> Val(vaSpread1.text) Then
                      
                      Call MsgBox("Grupo estructura es distinto, no puede mover estructura", vbCritical + vbOKOnly, MsgTitulo)
                      Exit Sub
                   
                   End If
                   vaSpread1.Row = FilaAnt
                   vaSpread1.Col = 2
                   
                   If vaSpread1.CellType = CellTypePercent Then
                      
                      If vaSpread1.BackColor = Shape1(2).FillColor Then
                         
                         vaSpread1.BackColor = Shape1(0).FillColor
                      
                      End If
                      
                      CantTotalPorcentaje = IIf(Trim(vaSpread1.text) = "", 0, Val(vaSpread1.text))
                      vaSpread1.Row = FilaAct
                      vaSpread1.CellType = CellTypePercent
                      vaSpread1.TypeHAlign = 1
                      vaSpread1.TypePercentDecPlaces = 0
                      vaSpread1.ForeColor = &HFF0000
                      vaSpread1.text = CantTotalPorcentaje
                      vaSpread1.TypeNegRed = True
                      If vaSpread1.BackColor = Shape1(0).FillColor Then
                         
                         vaSpread1.BackColor = Shape1(2).FillColor
                      
                      End If
                      
                      '-------> Mover valor cero
                      vaSpread1.Row = FilaAnt
                      vaSpread1.Col = 2
                      vaSpread1.text = ""
                      vaSpread1.CellType = CellTypeStaticText
                      
                   End If
                
                End If
                
                'Agregar filas temporales y respaldar
                vaSpread1.MaxRows = vaSpread1.MaxRows + (FilaAct - FilaAnt)
                For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows + (FilaAct - FilaAnt) 'jpaz
                    
                    vaSpread1.Row = i
                    vaSpread1.RowHidden = True
                
                Next i
                vaSpread1.MoveRange 1, FilaAnt, vaSpread1.maxcols, (FilaAct - 1), 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1
                'Mover estructura
                vaSpread1.MoveRange 1, FilaAct, vaSpread1.maxcols, IIf((FilaPos - 1) < FilaAct, FilaAct, (FilaPos - 1)), 1, FilaAnt
                'Devolver respaldo
                vaSpread1.MoveRange 1, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1, vaSpread1.maxcols, vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 + (FilaAct - FilaAnt - 1), 1, FilaAnt + IIf((FilaPos - FilaAct) <= 0, 1, (FilaPos - FilaAct))
                For i = vaSpread1.MaxRows - (FilaAct - FilaAnt) + 1 To vaSpread1.MaxRows
                    
                    vaSpread1.DeleteRows i, 1
                    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
                
                Next i
                vaSpread1.SetActiveCell 1, FilaAnt + (FilaPos - FilaAct)
            
            End If
            
            iblockrow = vaSpread1.ActiveRow: iblockrow2 = vaSpread1.ActiveRow: iblockcol = vaSpread1.ActiveCol
            IndGrabado = 1
            Plato(0).Enabled = True: OpGrilla(0).Enabled = True
            Plato(13).Enabled = False: OpGrilla(13).Enabled = False
            Plato(14).Enabled = False: OpGrilla(14).Enabled = False
            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
            vaSpread1.Col = 1
            For i = 1 To vaSpread1.MaxRows - 1
                
                vaSpread1.Row = i
                vaSpread1.BackColor = Shape1(2).FillColor
            
            Next i
        
        Case 11, 12 'copiar y pegar linea

'            If vaSpread1.ActiveRow = vaSpread1.MaxRows Then Exit Sub
            If Index = 11 Then
                
                If iblockcol < 1 Then
                    
                    For i = 1 To MaxColumna
                        
                        vaSpread1.Col = vectorcol(i)
                        vaSpread1.Row = 1
                        
                        If vaSpread1.BackColor = Shape1(1).FillColor Then
                            
                            Call MsgBox("Existen Días Bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo)
                            Exit Sub
                        
                        End If
                    
                    Next i
                
                Else
                    
                    For i = iblockcol To iblockcol2
                        
                        vaSpread1.Col = i
                        
                        For j = iblockrow To iblockrow2
                            
                            vaSpread1.Row = j
                            
                            If vaSpread1.BackColor = Shape1(1).FillColor Then
                                
                                Call MsgBox("Bloque seleccionado existen días bloqueado, no puede usar cortar", vbCritical + vbOKOnly, MsgTitulo)
                                Exit Sub
                            
                            End If
                        
                        Next j
                    
                    Next i
               
               End If
               'Validar recetas 5 etapas
               j = 0
               For i = 1 To MaxColumna
                   
                   If (vectorcol(i) - 2) = iblockcol Or vectorcol(i) = iblockcol Then j = (vectorcol(i) - 2): Exit For
               
               Next i
               If j = 0 Then Exit Sub
            
            End If
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            '------> Verificar si copiar receta, raciones o bien Ponderaciones día
            vaSpread1.Row = SpreadHeader + 3 '0
            If vaSpread1.text = "N.Rac." And TipoMinuta Then
                
                TipoCopia = "Copia Raciones"
            
            ElseIf vaSpread1.text = "N.Rac." And Not TipoMinuta Then
                
                TipoCopia = "Copia Raciones"
            
            ElseIf vaSpread1.text = "% Pond." And Not TipoMinuta Then
                
                TipoCopia = "Copia Ponderaciones"
            
            ElseIf vaSpread1.text = "% Pond." And TipoMinuta Then
                
                TipoCopia = "Copia Ponderaciones"
            
            Else
                
                TipoCopia = "Copia Receta"
            
            End If
            
            aiblockrow = iblockrow: aiblockrow2 = iblockrow2
            aiblockcol = iblockcol: aiblockcol2 = iblockcol2
            If vaSpread1.Col = 1 Or vaSpread1.Col = 2 Then Exit Sub
            If vaSpread1.MaxRows > 1000 Then Del_Row = vaSpread1.MaxRows - 1000: vaSpread1.MaxRows = vaSpread1.MaxRows - Del_Row
            Plato(13).Enabled = True: OpGrilla(13).Enabled = True
            Plato(14).Enabled = True: OpGrilla(14).Enabled = True
            Toolbar1.Buttons(6).Visible = False: Toolbar1.Buttons(7).Visible = True
            If iblockcol < 1 Then aiblockcol = 1: aiblockcol2 = vaSpread1.maxcols
            indcortarpegar = 1
            If Index = 11 Then indcortarpegar = 0: Toolbar1.Buttons(8).Visible = True: Toolbar1.Buttons(9).Visible = False: Plato(14).Enabled = False: OpGrilla(14).Enabled = False Else Toolbar1.Buttons(8).Visible = False: Toolbar1.Buttons(9).Visible = True: Plato(14).Enabled = True: OpGrilla(14).Enabled = True
        
        Case 13, 14 'Validar recetas 5 etapas
            AadRec = 0
            'copiar y pegar
            
            If indcortarpegar = 0 Then
               
               If (iblockcol2 - iblockcol) > (aiblockcol2 - aiblockcol) Or (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then
                  
                  MsgBox "Imposible Pegar la infomación ya que el área de Cortar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
                  Exit Sub
               
               End If
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
               
               If (aiblockrow2 - aiblockrow) + 1 + vaSpread1.ActiveRow > vaSpread1.MaxRows Then
                  
                  If (aiblockrow2 - aiblockrow) + vaSpread1.ActiveRow > vaSpread1.MaxRows Then
                     
                     MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "sobre pasa el maximo de filas :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
                     Exit Sub
                  
                  End If
               
               End If
            
            End If
            
            If iblockcol < 1 Then
                
                For i = 1 To MaxColumna
                   
                   vaSpread1.Col = vectorcol(i)
                   vaSpread1.Row = 1
                    
                    If vaSpread1.BackColor = Shape1(1).FillColor Then
                        
                        Call MsgBox("Existen Días Bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo)
                        Exit Sub
                    
                    End If
               
               Next i
            
            Else
                
                For i = iblockcol To iblockcol2
                   
                   vaSpread1.Col = i
                    
                    For j = iblockrow To iblockrow2
                        
                        vaSpread1.Row = j
                        
                        If vaSpread1.BackColor = Shape1(1).FillColor Then
                            
                            Call MsgBox("Bloque seleccionado existen días bloqueado, no puede modificar la estructura", vbCritical + vbOKOnly, MsgTitulo)
                            Exit Sub
                        
                        End If
                    
                    Next j
               
               Next i
            
            End If
            vaSpread1.Col = 1
            If vaSpread1.text = "Comensales" Then
               
               If (aiblockrow2 - aiblockrow) + vaSpread1.ActiveRow > vaSpread1.MaxRows Then
                  
                  Call MsgBox("Imposible copiar o bien pegar ultima fila", vbCritical + vbOKOnly, MsgTitulo)
                  Exit Sub ' Valida que no se peguen recetas en la Línea de Comensales.
               
               End If
            
            End If
            
            vaSpread1.Col = vaSpread1.ActiveCol
            If vaSpread1.Col = 1 Then Exit Sub
            If indcortarpegar = 0 Then OpGrilla(13).Enabled = False: Toolbar1.Buttons(6).Visible = True: Toolbar1.Buttons(7).Visible = False
            Plato(0).Enabled = True
            OpGrilla(0).Enabled = True
            
            ' destinacion de copiar y pegar datos
            If iblockcol < 1 Then iblockcol = 1: iblockcol2 = vaSpread1.maxcols
            If aiblockcol2 = vaSpread1.maxcols Then aiblockcol2 = vaSpread1.maxcols - 2
            vaSpread1.Row = 0: vaSpread1.Col = vaSpread1.ActiveCol 'iblockcol
            vaSpread1.Row = SpreadHeader + 3 '0
            Dim GlosaTipoMinuta As String
            GlosaTipoMinuta = IIf(TipoMinuta, "N.Rac.", "% Pond.")
            
            If (vaSpread1.text = GlosaTipoMinuta) And (TipoCopia = "Copia Raciones" Or TipoCopia = "Copia Ponderaciones") Then
                
                If (aiblockrow2 - aiblockrow) + iblockrow2 > vaSpread1.MaxRows Then MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, MsgTitulo: IndGrabado = 0: Exit Sub
                
                If TipoMinuta Then
                   
                   For i = 1 To MaxColumna
                       
                       If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Or _
                          (vectorcol(i) + 3) = iblockcol Or (vectorcol(i) + 4) = iblockcol Or (vectorcol(i) + 5) = iblockcol Then
                          
                          iblockcol = (vectorcol(i) - 1): Exit For
                       
                       End If
                   
                   Next i
                   
                   For i = 1 To MaxColumna
                       
                       If (vectorcol(i) - 1) = aiblockcol Or vectorcol(i) = aiblockcol Or (vectorcol(i) + 1) = aiblockcol Or (vectorcol(i) + 2) = aiblockcol Or _
                          (vectorcol(i) + 3) = aiblockcol Or (vectorcol(i) + 4) = aiblockcol Or (vectorcol(i) + 5) = aiblockcol Then
                           
                           aiblockcol = (vectorcol(i) - 1)
                           Exit For
                       
                       End If
                   
                   Next i
                   
                   For i = 1 To MaxColumna
                       
                       If (vectorcol(i) - 1) = iblockcol2 Or vectorcol(i) = iblockcol2 Or (vectorcol(i) + 1) = iblockcol2 Or (vectorcol(i) + 2) = iblockcol2 Or _
                          (vectorcol(i) + 3) = iblockcol2 Or (vectorcol(i) + 4) = iblockcol2 Or (vectorcol(i) + 5) = iblockcol2 Then
                          
                          iblockcol2 = ((vectorcol(i) + 5)): Exit For
                       
                       End If
                   
                   Next i
                   
                   For i = 1 To MaxColumna
                       
                       If (vectorcol(i) - 1) = aiblockcol2 Or vectorcol(i) = aiblockcol2 Or (vectorcol(i) + 1) = aiblockcol2 Or (vectorcol(i) + 2) = aiblockcol2 Or _
                          (vectorcol(i) + 3) = aiblockcol2 Or (vectorcol(i) + 4) = aiblockcol2 Or (vectorcol(i) + 5) = aiblockcol2 Then
                          
                          aiblockcol2 = (vectorcol(i) + 5): Exit For
                       
                       End If
                   
                   Next i
                
                Else
                   
                   cantCol = aiblockcol2 - aiblockcol
                   CantCol1 = iblockcol2 - iblockcol
                
                End If
'                cantCol = aiblockcol2 - aiblockcol
'                CantCol1 = iblockcol2 - iblockcol
            
            ElseIf (vaSpread1.text <> GlosaTipoMinuta) And (TipoCopia = "Copia Raciones" Or TipoCopia = "Copia Ponderaciones") And (aiblockrow2 - aiblockrow) + vaSpread1.ActiveRow <> vaSpread1.MaxRows Then
                
                MsgBox "Imposible Pegar la infomación ya que tiene una columna distinta N.Raciones o bien % Ponderación", vbInformation + vbOKOnly, MsgTitulo:  Exit Sub
            
            Else
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = iblockcol Or vectorcol(i) = iblockcol Or (vectorcol(i) + 1) = iblockcol Or (vectorcol(i) + 2) = iblockcol Or _
                       (vectorcol(i) + 3) = iblockcol Or (vectorcol(i) + 4) = iblockcol Or (vectorcol(i) + 5) = iblockcol Then
                       
                       iblockcol = (vectorcol(i) - 1): Exit For
                    
                    End If
                
                Next i
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = aiblockcol Or vectorcol(i) = aiblockcol Or (vectorcol(i) + 1) = aiblockcol Or (vectorcol(i) + 2) = aiblockcol Or _
                       (vectorcol(i) + 3) = aiblockcol Or (vectorcol(i) + 4) = aiblockcol Or (vectorcol(i) + 5) = aiblockcol Then
                        
                        aiblockcol = (vectorcol(i) - 1)
                        Exit For
                    
                    End If
                
                Next i
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = iblockcol2 Or vectorcol(i) = iblockcol2 Or (vectorcol(i) + 1) = iblockcol2 Or (vectorcol(i) + 2) = iblockcol2 Or _
                       (vectorcol(i) + 3) = iblockcol2 Or (vectorcol(i) + 4) = iblockcol2 Or (vectorcol(i) + 5) = iblockcol2 Then
                       
                       iblockcol2 = ((vectorcol(i) + 5)): Exit For
                    
                    End If
                
                Next i
                
                For i = 1 To MaxColumna
                    
                    If (vectorcol(i) - 1) = aiblockcol2 Or vectorcol(i) = aiblockcol2 Or (vectorcol(i) + 1) = aiblockcol2 Or (vectorcol(i) + 2) = aiblockcol2 Or _
                       (vectorcol(i) + 3) = aiblockcol2 Or (vectorcol(i) + 4) = aiblockcol2 Or (vectorcol(i) + 5) = aiblockcol2 Then
                       
                       aiblockcol2 = (vectorcol(i) + 5): Exit For
                    
                    End If
                
                Next i
                cantCol = aiblockcol2 - aiblockcol
                CantCol1 = iblockcol2 - iblockcol
            
            End If
            
            IndCol = aiblockcol: indcol2 = iblockcol2
            indrow = aiblockrow: indrow2 = aiblockrow2
            If Index = 14 And indcortarpegar = 1 Then
               
               If (aiblockrow2 - aiblockrow) <> 0 Or (aiblockcol2 - aiblockcol) <> 6 Then MsgBox "Por esta opción solamente puede copiar una receta", vbInformation + vbOKOnly, "Detalle Planificación Minutas": iblockcol = vaSpread1.ActiveCol: Exit Sub
               'Rutina pegado especial

               vaSpread1.Row = SpreadHeader + 3: nrodia = ""
               '-------> Validar si existen días feriados
               For i = 4 To vaSpread1.maxcols Step 7
                   
                   vaSpread1.Row = 1
                   vaSpread1.Col = i
                   
                   If vaSpread1.BackColor = Shape1(1).FillColor Then
                      
                      vaSpread1.Row = SpreadHeader + 3
                      vaSpread1.Col = i
                      
                      If Trim(vaSpread1.text) <> "" Then
                         
                         nrodia = nrodia & Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 4) & ";"
                      
                      End If
                   
                   End If
               
               Next i
               
               vaSpread1.Row = SpreadHeader + 3: nrodia = ""
               For i = aiblockcol To aiblockcol2 Step 7
                   
                   vaSpread1.Col = i + 1
                   nrodia = nrodia & Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 4) & ";"
                   Let NroMes = NroMes & Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 2))
               
               Next i
               '-------> Validar si existen dias feriados
               For i = 4 To vaSpread1.maxcols - 2 Step 7
                   
                   vaSpread1.Row = 1
                   vaSpread1.Col = i
                   
                   If vaSpread1.BackColor = Shape1(1).FillColor Then
                      
                      vaSpread1.Row = SpreadHeader + 3
                      
                      If Trim(vaSpread1.text) <> "" Then
                         
                         nrodia = nrodia & Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 4) & ";"
                      
                      End If
                   
                   End If
               
               Next i
                
                'Validar receta 5 etapa
                AadRec = 0
                Dim SwFecha As Boolean
                Dim FechaPegado As Long

                vg_codigo = ""
                Vg_Codigo2 = ""
                Vg_Codigo3 = ""
                Vg_Codigo4 = ""
                Vg_Mes2 = 0
                Vg_Mes3 = 0
                Vg_Mes4 = 0
                
                SwFecha = 0
                FechaPegado = 0
                If Val(Mid(Trim(Vg_FechaDesde), 5, 2)) = NroMes Then
                    
                    Let SwFecha = 1
                    Let NroMes = 1
                    Call M_PeSSit.Inicio("Copia Especial Recetas Minuta Bloque", "PLATEO", Mid(Vg_FechaDesde, 1, 6), Mid(Vg_FechaHasta, 1, 6), nrodia, NroMes)
                    M_PeSSit.Show 1
                
                End If
                
                If SwFecha = 0 Then
                    
                    Let FechaPegado = Val(Mid(Trim(Vg_FechaDesde), 5, 2)) + 1
                    
                    If FechaPegado = 13 Then
                        
                        Let FechaPegado = 1
                    
                    End If
                    
                    If FechaPegado = NroMes Then
                        
                        Let SwFecha = 1
                        Let NroMes = 2
                        Call M_PeSSit.Inicio("Copia Especial Recetas Minuta Bloque", "PLATEO", Mid(Vg_FechaDesde, 1, 6), Mid(Vg_FechaHasta, 1, 6), nrodia, NroMes)
                        M_PeSSit.Show 1
                    
                    End If
                
                End If
                
                If SwFecha = 0 Then
                    
                    FechaPegado = Val(Mid(Trim(Vg_FechaDesde), 5, 2)) + 2
                    
                    If FechaPegado = 13 Then
                        
                        Let FechaPegado = 1
                    
                    End If
                        
                    If FechaPegado = NroMes Then
                        
                        Let SwFecha = 1
                        Let NroMes = 3
                        Call M_PeSSit.Inicio("Copia Especial Recetas Minuta Bloque", "PLATEO", Mid(Vg_FechaDesde, 1, 6), Mid(Vg_FechaHasta, 1, 6), nrodia, NroMes)
                        M_PeSSit.Show 1
                    
                    End If
                
                End If
                        
                If Trim(vg_codigo) = "" And Trim(Vg_Codigo2) = "" And Trim(Vg_Codigo3) = "" And Trim(Vg_Codigo4) = "" Then
                    
                    iblockcol = vaSpread1.ActiveCol
                    Exit Sub
                
                End If

       'Mover días no permitidos
       ReDim Preserve VecDia(0)
       ValLcntH = ""
       i = 0
       
       GrabarCambios 1, 1, "Pegado Especial"
       
       If Len(vg_codigo) > 0 Or Len(Vg_Codigo2) > 0 Or Len(Vg_Codigo3) > 0 Or Len(Vg_Codigo4) > 0 Then
           
            For j = 1 To Len(vg_codigo)
                
                If Asc(Mid(vg_codigo, j, 1)) <> 59 Then
                   
                   Let ValLcntH = ValLcntH + Mid(vg_codigo, j, 1)
                
                Else
                   
                   ReDim Preserve VecDia(i)
                   Let VecDia(i) = ValLcntH
                   Let ValLcntH = ""
                   Let i = i + 1
                
                End If
            
            Next j
            
            If Trim(ValLcntH) <> "" Then
                
                ReDim Preserve VecDia(i)
                VecDia(i) = ValLcntH
            
            End If
            Dim auxmes As Long
            Dim CodGrupoEst As Long
            Dim CodEstr As Long
            
            For i = 4 To (vaSpread1.maxcols - 2) Step 7
                
                vaSpread1.Row = aiblockrow
                vaSpread1.Col = vaSpread1.maxcols - 1
                CodEstr = Val(vaSpread1.text)
                vaSpread1.Col = vaSpread1.maxcols
                CodGrupoEst = Val(vaSpread1.text)

                vaSpread1.Row = SpreadHeader + 3
                vaSpread1.Col = i
                L = 0
                If Trim(vaSpread1.text) <> "" Then
                
                If auxmes <> Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 2)) Then
                    
                    If auxmes > 0 Then
                    
                       vg_codigo = ""
                    
                       If Trim(Vg_Codigo2) <> "" And Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 2)) = Vg_Mes2 Then
                       
                          vg_codigo = Vg_Codigo2
                          Vg_Codigo2 = ""
                    
                       ElseIf Trim(Vg_Codigo3) <> "" And Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 2)) = Vg_Mes3 Then
                       
                          vg_codigo = Vg_Codigo3
                          Vg_Codigo3 = ""
                    
                       ElseIf Trim(Vg_Codigo4) <> "" And Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 2)) = Vg_Mes4 Then
                       
                          vg_codigo = Vg_Codigo4
                          Vg_Codigo4 = ""
                    
                       End If
                    
                    End If
                   'Mover días no permitidos
                    ReDim Preserve VecDia(0)
                    For jjj = 0 To UBound(VecDia)
                        
                        Let VecDia(jjj) = 0
                    
                    Next jjj
                    
                    Dim iii As Long
                    ValLcntH = ""
                    iii = 0
                    For jjj = 1 To Len(vg_codigo)
                        
                        If Asc(Mid(vg_codigo, jjj, 1)) <> 59 Then
                           
                           Let ValLcntH = ValLcntH + Mid(vg_codigo, jjj, 1)
                        
                        Else
                           
                           ReDim Preserve VecDia(iii)
                           Let VecDia(iii) = ValLcntH
                           Let ValLcntH = ""
                           Let iii = iii + 1
                        
                        End If
                    
                    Next jjj
                    
                    If Trim(ValLcntH) <> "" Then
                        
                        ReDim Preserve VecDia(iii)
                        VecDia(iii) = ValLcntH
                    
                    End If
                    
                    auxmes = Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 5, 2))
                
                End If
                
                nrodia = Val(Mid(Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd"), 7, 2))
                
                For j = 0 To UBound(VecDia)
                    
                    If nrodia = VecDia(j) Then
                        
                        vaSpread1.Row = aiblockrow
                        vaSpread1.Col = i '- 1
                        
                        If Trim(vaSpread1.text) <> "" Then
                            
                            For X = aiblockrow + 1 To vaSpread1.MaxRows
                                
                                vaSpread1.Row = X
                                vaSpread1.Col = vaSpread1.maxcols - 1
                                xser = Val(vaSpread1.text)
                                vaSpread1.Col = i + 1
                                
                                If vaSpread1.Row = vaSpread1.MaxRows Then
                                    
                                    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                                    vaSpread1.InsertRows X, 1
                                    vaSpread1.Row = X
                                    
                                    For xx = 2 To vaSpread1.maxcols
                                       
                                       vaSpread1.Col = xx
                                       vaSpread1.BackColor = Shape1(0).FillColor
                                    
                                    Next xx
                                    
                                    L = X
                                    Exit For
                                
                                End If
                                
                                If xser <> CodEstr And xser > 0 Then
                                    
                                    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                                    vaSpread1.InsertRows X, 1
                                    vaSpread1.Row = X
                                    
                                    For xx = 2 To vaSpread1.maxcols
                                       
                                       vaSpread1.Col = xx
                                       vaSpread1.BackColor = Shape1(0).FillColor
                                    
                                    Next xx
                                    
                                    L = X
                                    Exit For
                                
                                ElseIf Trim(vaSpread1.text) <> "" And xser > 0 Then
                                    
                                    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                                    vaSpread1.InsertRows X + 1, 1
                                    X = X + 1
                                    vaSpread1.Row = X
                                    
                                    For xx = 2 To vaSpread1.maxcols
                                       
                                       vaSpread1.Col = xx
                                       vaSpread1.BackColor = Shape1(0).FillColor
                                    
                                    Next xx
                                    
                                    L = X
                                    Exit For
                                
                                ElseIf Trim(vaSpread1.text) = "" Then
                                    
                                    Exit For
                                
                                End If
                            
                            Next X
                            
                            Call vaSpread1.CopyRange(aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, X)
                            vaSpread1.Row = X
                        
                        Else
                            
                            Call vaSpread1.CopyRange(aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i - 1, aiblockrow)
                            vaSpread1.Row = aiblockrow
                        
                        End If
                       'Asignar colores
                        
                        For X = (i - 1) To (i - 1) + 4
                           
                           vaSpread1.Col = X
                           vaSpread1.BackColor = Shape1(0).FillColor
                           
                           For xx = 1 To MaxColumna
                               
                               If (vectorcol(xx) - 1) = vaSpread1.Col Then
                               
                                  '-------> Porcentaje del dia
                                  vaSpread1.Col = X + 2
                                  vaSpread1.CellType = CellTypePercent
                                  'vaSpread1.text = 0
                                  vaSpread1.ForeColor = &HFF0000
                                  vaSpread1.CellType = CellTypePercent
                                  vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
                                  vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
                                  vaSpread1.TypePercentDecPlaces = 0
                                  vaSpread1.TypePercentMax = 1000
                                  ' display negative numbers as red
                                  vaSpread1.TypeNegRed = True
            
                                  vaSpread1.Col = X + 3
                                  vaSpread1.CellType = IIf(TipoMinuta, CellTypeNumber, CellTypeStaticText) 'CellTypeNumber
                                  vaSpread1.TypeNumberDecPlaces = 0
'                                  vaSpread1.TypeIntegerMin = 1
'                                  vaSpread1.TypeIntegerMax = 9999999
                                  vaSpread1.TypeNumberMin = 0
                                  vaSpread1.TypeNumberMax = 9999999
                                  vaSpread1.TypeHAlign = TypeHAlignRight
                                  vaSpread1.TypeSpin = False
                                  vaSpread1.TypeIntegerSpinInc = 1
                                  vaSpread1.TypeIntegerSpinWrap = False

                                   '-------> Mover tipo celda nota cuando sucede un cambio
                                   vaSpread1.Col = X + 5
                                   vaSpread1.TextTip = TextTipFloating
                                   vaSpread1.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                                   vaSpread1.CellNote = "Cambio"
                                  
                                   vaSpread1.Col = X + 2
                                   
                                   vaSpread1.Col = vaSpread1.maxcols - 1
                                   vaSpread1.text = CodEstr
                                   
                                   vaSpread1.Col = vaSpread1.maxcols
                                   vaSpread1.text = CodGrupoEst
                                   
                                   Exit For
                               
                               End If
                           
                           Next xx
                           
                           vaSpread1.Col = X
                           
                           If X = (i - 1) Then
                                
                                vaSpread1.ForeColor = &HFF&
                                vaSpread1.BackColor = &H80FF80
                            
                            End If
                       
                       Next X
                       
                       If L > 0 Then
                          
                          z = L
                          
                          For L = 3 To (vaSpread1.maxcols - 2) Step 7
                              
                              vaSpread1.Row = 1
                              vaSpread1.Col = L
                              
                              If vaSpread1.BackColor = Shape1(1).FillColor Then
                                 
                                 vaSpread1.Row = z
                                 
                                 For X = (L - 1) To (L - 1) + 4
                                     
                                     vaSpread1.Col = X
                                     vaSpread1.BackColor = Shape1(1).FillColor
                                 
                                 Next X
                              
                              End If
                          
                          Next L
                       
                       End If
                       
                       'Fin asignar colores
                       Exit For
                       
                    End If
                
                Next j
                
                End If
            
            Next i
        
        End If

            Else
               
               'Validar receta 5 etapas
               GrabarCambios vaSpread1.Row, j, "Copiado y Pegado"
               indrow3 = vaSpread1.MaxRows
               vaSpread1.Visible = False
               Dim AuxGlosaTipoMinuta As String
               
               For i = iblockcol To iblockcol2 Step 7
                   
                   If indcortarpegar = 1 Then
                      
                      vaSpread1.MaxRows = vaSpread1.MaxRows + (aiblockrow2 - aiblockrow) + 1
                      vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow)
                      'Asignar colores
                      
                      For j = vaSpread1.MaxRows - (aiblockrow2 - aiblockrow) To vaSpread1.MaxRows
                          
                          vaSpread1.Row = j
                          
                          For X = (i) To (i) + 4 '6 '4
                              
                              vaSpread1.Col = X
                              
                              If aiblockrow <> maxfila Then
                                 
                                 vaSpread1.BackColor = Shape1(0).FillColor
                              
                              End If
                              
                              vaSpread1.Row = SpreadHeader + 3
                              AuxGlosaTipoMinuta = vaSpread1.text
                              vaSpread1.Row = j
                              
                              For xx = 1 To MaxColumna
                                  
                                  If (vectorcol(xx) - 1) = vaSpread1.Col And Trim(vaSpread1.text) <> "" Then
                                     
                                     '-------> Porcentaje del dia
                                     vaSpread1.Col = X + 2
                                     vaSpread1.CellType = CellTypePercent
                                     'vaSpread1.text = 0
                                     vaSpread1.ForeColor = &HFF0000
                                     vaSpread1.CellType = CellTypePercent
                                     vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
                                     vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
                                     vaSpread1.TypePercentDecPlaces = 0
                                     vaSpread1.TypePercentMax = 1000
                                     ' display negative numbers as red
                                     vaSpread1.TypeNegRed = True
                                     
'                                     If TipoMinuta Then
                                        vaSpread1.Col = X + 3
                                        vaSpread1.CellType = IIf(TipoMinuta, CellTypeNumber, CellTypeStaticText)
                                        vaSpread1.TypeNumberDecPlaces = 0
'                                        vaSpread1.TypeIntegerMin = 1
'                                        vaSpread1.TypeIntegerMax = 9999999
                                        vaSpread1.TypeNumberMin = 0
                                        vaSpread1.TypeNumberMax = 9999999
                                        vaSpread1.TypeHAlign = TypeHAlignRight
                                        vaSpread1.TypeSpin = False
                                        vaSpread1.TypeIntegerSpinInc = 1
                                        vaSpread1.TypeIntegerSpinWrap = False
'                                     End If
                                     Exit For
                                  
                                  End If
                              
                              Next xx
                              vaSpread1.Row = SpreadHeader + 3
                              AuxGlosaTipoMinuta = vaSpread1.text
                              vaSpread1.Row = j
                              vaSpread1.Col = X
                              
                              If X = (i) And Trim(vaSpread1.text) <> "" And TipoCopia <> "Copia Raciones" And AuxGlosaTipoMinuta <> GlosaTipoMinuta Then
                                 
                                 vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                              
                              ElseIf X = (i) And Trim(vaSpread1.text) <> "" And TipoCopia <> "Copia Ponderaciones" And AuxGlosaTipoMinuta <> GlosaTipoMinuta Then
                                 
                                 vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                              
                              ElseIf X = (i) And Trim(vaSpread1.text) = "R" Then
                                 
                                 vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                              
                              End If
                              
'                              If x = (i) And Trim(vaSpread1.text) <> "" Then vaSpread1.ForeColor = &HFF&: vaSpread1.BackColor = &H80FF80
                          Next X
                      
                      Next j
                      
                      'Fin asignar colores
                      '-------> Mover tipo celda nota cuando sucede un cambio
                      
                      For xp = vaSpread1.MaxRows - (aiblockrow2 - aiblockrow) To vaSpread1.MaxRows
                          
                          vaSpread1.Row = xp
                          vaSpread1.Col = i + 5
                          vaSpread1.TextTip = TextTipFloating
                          vaSpread1.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                          vaSpread1.CellNote = "Cambio"
                      
                      Next xp
                     
                     If TipoCopia = "Copia Raciones" Then
                        
                        vaSpread1.CopyRange i + 3, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), i + 3, vaSpread1.MaxRows, i + 3, vaSpread1.ActiveRow
                     
                     ElseIf TipoCopia = "Copia Ponderaciones" Then
                        
                        vaSpread1.CopyRange i, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), i, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                     
                     Else
                        
                        vaSpread1.CopyRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
'                         vaSpread1.CopyRange iblockcol, vaSpread1.MaxRows - (aiblockrow2 - aiblockrow), iblockcol2, vaSpread1.MaxRows, i, vaSpread1.ActiveRow
                     
                     End If
                     accion = "Copiar"
                     vaSpread1.MaxRows = indrow3
                   
                   ElseIf indcortarpegar = 0 Then
                      
                      '-------> Mover tipo celda nota cuando sucede un cambio
                      For xp = aiblockrow To aiblockrow2
                          
                          vaSpread1.Row = xp
                          vaSpread1.Col = aiblockcol + 5
                          vaSpread1.TextTip = TextTipFloating
                          vaSpread1.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
                          vaSpread1.CellNote = "Cambio"
                      
                      Next xp
                      vaSpread1.MoveRange aiblockcol, aiblockrow, aiblockcol2, aiblockrow2, i, vaSpread1.ActiveRow
                      accion = "Pegar"
                   
                   End If
               
               Next i
               
               vaSpread1.Visible = True
            
            End If
            
            If indcos = True Then
               
               For i = 4 To (vaSpread1.maxcols - 2) Step 7
                   
                   Calctodia 1, i + 1
               
               Next i
               MostrarCosto vaSpread1.ActiveCol
            
            End If
            
            For i = 3 To (vaSpread1.maxcols - 2) Step 7
                
                CalctodiaEnc 1, i + 1
            
            Next i
            
            ColumnaActiva = vaSpread1.ActiveCol
            FilaActiva = vaSpread1.ActiveRow
            ColumnaAntActiva = ColumnaActiva - 1
            vaSpread1.Col = ColumnaActiva
            vaSpread1.Row = 0
            IndGrabado = 1
            aiblockcol = IndCol
            iblockcol2 = indcol2
            aiblockrow = indrow
            aiblockrow2 = indrow2
            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            vaSpread1.SetFocus
        
        Case 15
            
            B_IngRecMinBloNormal.Partidas Me
            B_IngRecMinBloNormal.Show 1
    
        Case 16
        
'20211209            Dim tipoceco As String
'
'            tipoceco = ""
'            If RS.State = 1 Then RS.Close
'            RS.CursorLocation = adUseClient
'            vg_db.CursorLocation = adUseClient
'
'            Set RS = vg_db.Execute("sgpadm_Sel_TipoCeco'" & vg_codcasino & "'")
'            If Not RS.EOF Then
'
'               tipoceco = RS!cli_tipoceco
'
'            End If
'            RS.Close
'20211209            Set RS = Nothing
            
            If vaSpread1.ActiveCol = 1 And vaSpread1.ActiveRow <> vaSpread1.MaxRows And Trim(vaSpread1.text) <> "" Then '20211209 And tipoceco = "1" Then
               
               G_Proc.CellEdite B_CelEdi, "Editar Estructura", "Nombre Estructura", vaSpread1, "2"
               Toolbar1.Buttons(1).Visible = False
               Toolbar1.Buttons(2).Visible = True
               IndGrabado = 1
               
            End If
        
    End Select

ResivarFilasTengaAsigEstGrupo

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

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
            
            Call Plantilla_Click(5)
        
        'Case 19
            
            'Plantilla_Click (8)
        
        Case 21 'Exportar Planificación Minuta
        
        Case 22 'Importación Planificación Minuta
        
        Case 24
            
            Call Plantilla_Click(10)
        
        Case 25
            
            Plantilla_Click (11)
        
        Case 27
            
            Call Plantilla_Click(12)
        
        Case 28
            
            Call Plantilla_Click(13)
        
        Case 29
            
            Plantilla_Click (14)
        
        Case 31
            
            'ExportarExcel
        
        Case 33 'Frecuencia ingrediente
            
'            If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, Msgtitulo:  Exit Sub
'            C_FreIngMinBlo.LlenarFrecIngMinBlo "Frecuencia Ingrediente Minuta Bloque " & Vg_FechaDesde & " - " & Vg_FechaHasta, vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), Val(Vg_FechaHasta)
'            C_FreIngMinBlo.Show 1, Me
        
            If Toolbar1.Buttons(2).Visible = True Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, MsgTitulo:  Exit Sub
            C_FreIngMinBlo.LlenarFrecIngMinBlo_Inicio "Frecuencia Ingrediente Minuta Bloque " & Vg_FechaDesde & " - " & Vg_FechaHasta, vg_codcasino, vg_codregimen, vg_codservicio, Val(Vg_FechaDesde), Val(Vg_FechaHasta)
            C_FreIngMinBlo.Show 1, Me
            
        Case 34 'Deshacer
            
            Deshacer "SpreadMBloque" & vg_NUsr & ContadorDeshacer & ".ss6"
            If ContadorDeshacer < 1 Then: Toolbar1.Buttons(34).Visible = True: Toolbar1.Buttons(34).Enabled = False
        
        Case 19 'MVA - MVI - ACTIVAR EL FORMULARIO QUE COPIA LA MINUTA CON ENCABEZADO CCOSTO, REGIMEN Y SERVICIO
            
            Plantilla_Click (19)
        
        Case 36 'Retroceder
            
            If Toolbar1.Buttons(2).Visible = True Then MsgBox "Debe grabar antes de avanzar o retroceder bloque", vbInformation + vbOKOnly, MsgTitulo:  Exit Sub
            RetrocederAvanzar_MinutaBloque "1"
        
        Case 38 'Avanzar
            
            If Toolbar1.Buttons(2).Visible = True Then MsgBox "Debe grabar antes de avanzar o retroceder bloque", vbInformation + vbOKOnly, MsgTitulo:  Exit Sub
            RetrocederAvanzar_MinutaBloque "2"
        
        Case 40 'Salir
            
            Plantilla_Click (20)
    
    End Select

End Sub


Sub RetrocederAvanzar_MinutaBloque(ByVal op As String)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Sql As String

'-------> Cerrar hoja calculo
If Frame2(0).Visible = True Then
   
   Frame2(0).Visible = False
   vaSpread1.Move 0, 1760, ScaleWidth, ScaleHeight - 1760
   Image2(0).Visible = False: Image2(1).Visible = False
   indcos = False

End If

Sql = ""
Sql = IIf(op = "1", "sgpadm_Sel_TraerMinutaBloqueRetroceder", "sgpadm_Sel_TraerMinutaBloqueAvanzar")

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("" & Sql & " '" & vg_codcasino & "', " & vg_IDBloque & "")
If Not RS.EOF Then
    
    vg_codcasino = RS!Ceco
    vg_codregimen = RS!Regimen
    vg_codservicio = RS!Servicio
    Vg_FechaDesde = RS!fechadesde
    Vg_FechaHasta = RS!fechahasta
    vg_IDBloque = RS!Id_Bloque
    Label4.Caption = Trim(RS!Cli_nombre) & "(" & RS!Ceco & ")" & " - " & Trim(RS!reg_nombre) & " - " & Trim(RS!ser_nombre) & " - Bloque " & IIf(vg_IDBloque = 0, "", vg_IDBloque)
    Label1(1).Caption = "Nota : Las raciones debe incluir las raciones del personal "

    vaSpread1.MaxRows = 0
    vaSpread1.Visible = False
    FormatearGrilla
    DetallePlantillaMinuta
    LlenarEstructuraServicio
    vaSpread1.Visible = True
    Toolbar1.Buttons(34).Visible = True
    Toolbar1.Buttons(34).Enabled = False

Else
    
    RS.Close
    Set RS = Nothing
    MsgBox "No existe más información que mostrar", vbInformation + vbOKOnly, MsgTitulo:  Exit Sub

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

Select Case ButtonMenu

Case "Formato I"
    
    ExportarExcelMenuI

Case "Formato II Resumido"
    
    ExportarExcelMenuII

Case "Sin % P-G-Cho-Agrs"

    Aportes_Click 10

Case "Con % P-G-Cho-Agrs"
    
    Aportes_Click 20

End Select

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
    
    indactivo = 1
    iblockrow = BlockRow
    iblockrow2 = BlockRow2
    iblockcol = BlockCol
    iblockcol2 = BlockCol2
    If BlockRow < 0 Then iblockrow = 1
    If BlockRow2 < 0 Then iblockrow2 = (vaSpread1.MaxRows - 1)
    If BlockRow2 >= vaSpread1.MaxRows Then iblockrow2 = (vaSpread1.MaxRows - 1)

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)
    
    If Row < 1 Then Exit Sub
                                                                   
    Let OpGrilla(15).Enabled = IIf(Col = 1 And (vg_codregimen > 9999 And AddReceta = 0), False, True)
    Let Plato(15).Enabled = IIf(Col = 1 And (vg_codregimen > 9999 And AddReceta = 0), False, True)
    Let indactivo = 1
    Let iblockrow = vaSpread1.ActiveRow
    Let iblockrow2 = vaSpread1.ActiveRow
    Let iblockcol = vaSpread1.ActiveCol
    Let iblockcol2 = vaSpread1.ActiveCol
    Let vaSpread1.Row = vaSpread1.ActiveRow
    Let vaSpread1.Col = vaSpread1.ActiveCol
     
    If Col = 1 Then Plato_Click (16): Exit Sub
     
End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

Dim cod As Variant
Dim estructuraservicio As Long
Dim nombreEstructura As String
If Row = 1 Then
    
    vaSpread1.Col = 1
    nombreEstructura = vaSpread1.text
    
    If nombreEstructura = "" Then
       
       MsgBox "No Se puede ingresar Receta sin No tener Estructura de Servicio", 16
       Exit Sub
    
    End If

End If

If Row < 1 Or Col = 1 Then Exit Sub
ws_respuesta = ""
Let ColumnaReceta = Col
Call Plato_Click(2)

End Sub

Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

If vaSpread1.MaxRows < 1 Then Exit Sub
Dim codest As Long
Dim i As Long
Dim icol As Long
Dim LyD As Boolean

If vaSpread1.Col <> 2 Then
   
   'If Toolbar1.Buttons(2).Enabled = False Then Exit Sub
    Let vaSpread1.Row = vaSpread1.ActiveRow
    Let vaSpread1.Col = Col
    
    If vaSpread1.BackColor = Shape1(1).FillColor Then
        
        Call MsgBox("Día Bloqueado", vbCritical + vbOKOnly, MsgTitulo)
        vaSpread1.Lock = True
        Exit Sub
    
    End If
    
    If vaSpread1.ChangeMade = False Or Col = 1 Then
       
       Exit Sub
    
    End If
    
    If vaSpread1.ChangeMade = True And Mode = 0 Then
       
       If vaSpread1.text = "0.00%" Then
          
          codest = 0
          vaSpread1.Row = Row
          
          For i = (IIf(vaSpread1.Row = 1, 1, vaSpread1.Row + 1 - 1)) To 1 Step -1
              
              vaSpread1.Row = i
              vaSpread1.Col = 1
              
              If Trim(vaSpread1.text) <> "" Then
                 
                 vaSpread1.Col = vaSpread1.maxcols
                 codest = Val(vaSpread1.text)
                 Exit For
              
              End If
          
          Next i
       
       End If
       
       GrabarCambios 1, 1, "% "
       
       If Row = vaSpread1.MaxRows Then
          
          vaSpread1.Row = vaSpread1.ActiveRow
          vaSpread1.Col = Col
          j = Col - 1
          CalctodiaEnc vaSpread1.Row, j
       
       Else
          
          vaSpread1.Row = vaSpread1.ActiveRow
          vaSpread1.Col = Col
          j = Col - 1: CalctodiaEnc vaSpread1.Row, j
       
       End If
       
       If indcos = True Then
          
          If Row = vaSpread1.MaxRows Then
             
             vaSpread1.Row = vaSpread1.ActiveRow
             vaSpread1.Col = Col
             j = Col - 1
             Calctodia vaSpread1.Row, j
          
          Else
             
             If (Col <> 1 And Col <> 2) Then
                vaSpread1.Row = vaSpread1.ActiveRow
                vaSpread1.Col = Col - 1
                LyD = False
                
                If Mid(Trim(vaSpread1.text), 1, 3) = "[*]" Then
                   
                   LyD = True
                
                End If
                
                vaSpread1.Col = Col
'                j = IIf(TipoMinuta, Col - 1, Col)
                j = IIf(TipoMinuta, Col - 1, IIf(Not LyD, Col - 1, Col))
                Calctodia vaSpread1.Row, j
             
             End If
          
          End If
       
       End If
       
       If indcos = True Then
          
          If Row = vaSpread1.MaxRows Then
             
             vaSpread1.Row = vaSpread1.ActiveRow
             vaSpread1.Col = Col
             j = Col - 1
             Calctodia vaSpread1.Row, j
          
          End If
       
       End If
       
          
       '-------> Mover tipo celda nota cuando sucede un cambio
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Row = Row
       
       If Col = 2 Then
          
          vaSpread1.Col = Col - 1
       
       Else
          
          vaSpread1.Col = Col - 4
       
       End If
    
       vaSpread1.TextTip = TextTipFloating
       vaSpread1.CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
       vaSpread1.CellNote = "Cambio"
       
    End If
    
    vaSpread1.Col = Col
    vaSpread1.Row = Row
    IndGrabado = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True

ElseIf Col = 2 And vaSpread1.ChangeMade = True And Mode = 0 Then
    
    GrabarCambios 1, 1, "% Total"
    
    For i = 4 To vaSpread1.maxcols - 2 Step 7
        
        icol = i
        Call Cal_PorSugerido_Raciones(Row, icol)
    
    Next i

End If

End Sub

Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)

Dim delrow  As Integer
Dim IndCol  As Integer
Dim indrow  As Integer
Dim indcol2 As Integer
Dim indrow2 As Integer
Dim CodEstructura As Long
Dim auxest As Long
Dim xrow As Long
    'Esta combinación corresponde a las tecla control + v pegar
    If KeyCode = 86 And Shift = 2 Then
       
       Plato_Click 13
       Exit Sub
    
    ElseIf KeyCode = 67 And Shift = 2 Then
       
       Plato_Click 11
       Exit Sub
    
    End If
    
    Select Case KeyCode
        
        Case 65 To 90
           
           ws_respuesta = ""
           ws_respuesta = Chr(KeyCode)
           Plato_Click (2)
        
        Case 86
            
            Exit Sub
        
        Case 46
            
            If vaSpread1.MaxRows = vaSpread1.ActiveRow Or vaSpread1.MaxRows = iblockrow Or vaSpread1.MaxRows = iblockrow2 Then Exit Sub
            'Validar si minuta esta bloqueada
            If ValidarBloqueoMinuta Then Exit Sub
            
            vaSpread1.Row = vaSpread1.ActiveRow
            vaSpread1.Col = vaSpread1.ActiveCol
            If vaSpread1.Col = 1 Then 'Suprimir estructura servicio
               
'20211209               vaSpread1.Col = vaSpread1.ActiveCol + 1
'               If vaSpread1.text <> "" Then
                  
'                  MsgBox "No puede eliminar una estructura que contenga % Ponderación x Estructura " & VgLinea & "Elimine la linea completa", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
               
'               Else 'Mover codigo estructura anterior a la estructura eliminada
                  
                  GrabarCambios 1, 1, "Modificar Estructura"
                  vaSpread1.Col = vaSpread1.maxcols - 1
                  vaSpread1.Row = IIf(vaSpread1.ActiveRow - 1 < 1, 1, vaSpread1.ActiveRow - 1)
                  CodEstructura = IIf(Trim(vaSpread1.text) = "", 0, vaSpread1.text)
                  vaSpread1.Row = vaSpread1.ActiveRow
                  vaSpread1.text = CodEstructura 'IIf(CodEstructura = 0, "", CodEstructura)
                  vaSpread1.Col = vaSpread1.ActiveCol
                  vaSpread1.text = " "
               
'20211209               End If
            
            Else
               
               j = 0
               For i = 1 To MaxColumna
                   
                   If (vectorcol(i) - 1) = vaSpread1.Col Or vectorcol(i) = vaSpread1.Col Then j = (vectorcol(i) - 1): Exit For
               
               Next i
               
               If j = 0 Then Exit Sub
               
               '-------> Grabar Evento Eliminación recetas
               GrabarCambios 1, 1, "Eliminación Recetas"
               vaSpread1.Col = j
               vaSpread1.Row = vaSpread1.ActiveRow
               Plato(0).Enabled = True: OpGrilla(0).Enabled = True
               Plato(13).Enabled = False: OpGrilla(13).Enabled = False
               Plato(14).Enabled = False: OpGrilla(14).Enabled = False
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
                   
                   If (vectorcol(i) - 1) = iblockcol2 Then iblockcol2 = ((vectorcol(i) + 5)): Exit For
                   If vectorcol(i) = iblockcol2 Then iblockcol2 = (vectorcol(i) + 5): Exit For
               
               Next i
               
               IndCol = aiblockcol: indcol2 = iblockcol2
               indrow = aiblockrow: indrow2 = aiblockrow2
               vaSpread1.ClearRange iblockcol, (iblockrow), iblockcol2, iblockrow2, False
'               For i = iblockcol To iblockcol2
'                   vaSpread1.Col = i
'                   vaSpread1.BackColor = Shape1(0).FillColor
'               Next i
               
               If indcos = True Then
                  
                  For i = iblockcol To iblockcol2 Step 7
                      
                      Calctodia 1, i + 2
                  
                  Next i
                  
                  MostrarCosto vaSpread1.ActiveCol
               
               End If
               
               For i = iblockcol To iblockcol2 Step 7
                   
                   CalctodiaEnc 1, i + 2
               
               Next i
               iblockcol = AuxCol
               vaSpread1.BlockMode = False
            
            End If
            IndGrabado = 1
            Plantilla(0).Enabled = True
            Toolbar1.Buttons(1).Visible = False
            Toolbar1.Buttons(2).Visible = True
            Toolbar1.Buttons(6).Visible = True
            Toolbar1.Buttons(7).Visible = False
            indactivo = 0
    
    End Select

    ResivarFilasTengaAsigEstGrupo
 
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
    '-------> Validar si minuta esta bloqueada
    If ValidarBloqueoMinuta Then Exit Sub
    Indvaspread1 = 0
    PopupMenu MenuDetalle

End Select

End Sub

Private Sub Opgrilla_Click(Index As Integer)
    
    Select Case Index
        
        Case 0
            
            Call Plato_Click(0)
        
        Case 2
            
            Call Plato_Click(2)
        
        Case 3
            
            Call Plato_Click(3)
        
        Case 5
            
            Call Plato_Click(5)
        
        Case 6
            
            Call Plato_Click(6)
        
        Case 8
            
            Call Plato_Click(8)
        
        Case 9
            
            Call Plato_Click(9)
        
        Case 11
            Call Plato_Click(11)
        Case 12
            
            Call Plato_Click(12)
        
        Case 13
            
            Call Plato_Click(13)
        
        Case 14
            
            Call Plato_Click(14)
        
        Case 15
            
            Call Plato_Click(15)
    
    End Select

End Sub

Private Sub GrabarPlantillaMinuta()

On Error GoTo Man_Error
    
Dim RS               As New ADODB.Recordset
Dim FecMinuta        As Long
Dim IndDia           As Long
Dim MyBuffer         As String
Dim DescReceta       As String
Dim CodReceta        As Long
Dim NumRacion        As Long
Dim NumRacTotal      As Long
Dim PorcenDiario     As Long
Dim PorcenTotal      As Long
Dim CodEstructura    As Long
Dim NumLin           As Long
Dim CodigoAgrupacion As Long
Dim Fecha            As Long
Dim NameEstructura   As String

Dim dato             As Variant
Dim Pos              As Variant
Dim Cabecera         As Long

Dim EstGrpEst        As Boolean
Dim ModMinB          As String
Dim Sql              As String
    
IndDia = 1
gauge1.Value = 0
gauge.Value = 0
Picture1.Visible = True
Label3.Visible = True
gauge.Visible = True
Picture1.Refresh
Label3.Refresh
gauge.Refresh
gauge1.Refresh
    
fg_carga ""
    
vaSpread1.Enabled = False
Toolbar1.Enabled = False
Main(0).Enabled = False
Main(1).Enabled = False
Let IndDia = 1
Let EstGrpEst = True
        
        For i = 3 To (vaSpread1.maxcols - 2) Step 7
          
            DoEvents
            gauge1.Value = Val((IndDia / MaxColumna) * 100)
            
            Call vaSpread1.GetText(i + 1, SpreadHeader + 3, dato)
            Label3.Caption = "": Label3.Caption = "Día : " & dato
            FecMinuta = Format(Mid(dato, 5, Len(dato)), "yyyymmdd")
                    
            Let MyBuffer = ""
            Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
            Let MyBuffer = MyBuffer & "<GrabaMinuta>"
            Let NumLin = 1
            Let NumRacTotal = 0

            For j = 1 To (vaSpread1.MaxRows - 1)
                       
                gauge.Value = Val((j / (vaSpread1.MaxRows - 1)) * 100)
                DescReceta = ""
                NameEstructura = ""
                CodReceta = 0
                NumRacion = 0
                        
                PorcenDiario = 0
                PorcenTotal = 0
                vaSpread1.Row = j
                '-------> Sacar codigo estructura servicio
                vaSpread1.Col = vaSpread1.maxcols - 1
                If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) > 0 Then CodEstructura = vaSpread1.text
                '-------> Sacar codigo grupo estructura
                If CodEstructura <> 13197 Then
                   
                   PorcenTotal = 0
                
                End If
                vaSpread1.Col = vaSpread1.maxcols
                                                
                If Trim(vaSpread1.text) <> "" Then CodigoAgrupacion = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
                
                vaSpread1.Col = i + 1
                DescReceta = vaSpread1.text
                
                If Trim(DescReceta) <> "" And CodEstructura > 0 Then
                   
                   vaSpread1.Col = 1
                   NameEstructura = IIf(Trim(vaSpread1.text) <> "", Trim(vaSpread1.text), "")
                   vaSpread1.Col = i + 2
                   PorcenDiario = IIf(vaSpread1.text = "0.00%", 0, Val(vaSpread1.text))
                   vaSpread1.Col = i + 3
                   NumRacion = IIf(Val(vaSpread1.text) = 0, 0, Val(vaSpread1.text))
                   vaSpread1.Col = i + 5
                   '-------> Sacar codigo receta
                   vg_tiprec = 0
                   StrRec = vaSpread1.text
                   If Len(StrRec) <> 0 Then
                      
                      Do While InStr(StrRec, ";") <> 0
                         
                         StrRecb = Mid(StrRec, 1, InStr(StrRec, ";") - 1)
                         StrRec = IIf(Len(StrRec) > InStr(StrRec, ";"), Mid(StrRec, InStr(StrRec, ";") + 1), "")
                         CodReceta = Val(Mid(StrRecb, 1, InStr(StrRecb, "&") - 1)): StrRecb = Mid(StrRecb, InStr(StrRecb, "&") + 1)
                         vg_tiprec = Val(Mid(StrRecb, 1))
                      
                      Loop
                   
                   End If
                   '-------> Mover cambio
                   ModMinB = "0"
                   If vaSpread1.CellNote = "Cambio" Then
                      
                      ModMinB = "1"
                   
                   End If
                   vaSpread1.Col = 2
                   PorcenTotal = IIf(Trim(vaSpread1.text) = "", -1, Val(vaSpread1.text))
'                   vg_tiprec = vaSpread1.GetColFromID("Grupo") + 1
                            
                   MyBuffer = MyBuffer & " <Minuta"
                   MyBuffer = MyBuffer & " Op = " & Chr(34) & 0 & Chr(34)

                   DescReceta = Replace(Trim(DescReceta), Chr(34), "&quot;")
                   DescReceta = Replace(Trim(DescReceta), Chr(38), "&amp;")
                   DescReceta = Replace(Trim(DescReceta), Chr(39), "&apos;")
                   DescReceta = Replace(Trim(DescReceta), Chr(60), "&lt;")
                   DescReceta = Replace(Trim(DescReceta), Chr(62), "&gt;")
                   
                   NameEstructura = Replace(Trim(NameEstructura), Chr(34), "&quot;")
                   NameEstructura = Replace(Trim(NameEstructura), Chr(38), "&amp;")
                   NameEstructura = Replace(Trim(NameEstructura), Chr(39), "&apos;")
                   NameEstructura = Replace(Trim(NameEstructura), Chr(60), "&lt;")
                   NameEstructura = Replace(Trim(NameEstructura), Chr(62), "&gt;")
                   
                   MyBuffer = MyBuffer & " PorcenDiario = " & Chr(34) & PorcenDiario & Chr(34)
                   MyBuffer = MyBuffer & " NumRacion = " & Chr(34) & NumRacion & Chr(34)
                   MyBuffer = MyBuffer & " DescReceta = " & Chr(34) & DescReceta & Chr(34)
                   MyBuffer = MyBuffer & " CodEstructura = " & Chr(34) & CodEstructura & Chr(34)
                   MyBuffer = MyBuffer & " TipoReceta = " & Chr(34) & vg_tiprec & Chr(34)
                   MyBuffer = MyBuffer & " Rec5eta = " & Chr(34) & 0 & Chr(34)
                   MyBuffer = MyBuffer & " NumLin = " & Chr(34) & NumLin & Chr(34)
                   MyBuffer = MyBuffer & " CodReceta = " & Chr(34) & CodReceta & Chr(34)
                   MyBuffer = MyBuffer & " FecVal = " & Chr(34) & 0 & Chr(34)
                   MyBuffer = MyBuffer & " ModMinB = " & Chr(34) & ModMinB & Chr(34)
                   MyBuffer = MyBuffer & " CodAgrupacion = " & Chr(34) & CodigoAgrupacion & Chr(34)
                   MyBuffer = MyBuffer & " PorcenTotal = " & Chr(34) & PorcenTotal & Chr(34)
                   MyBuffer = MyBuffer & " NEst = " & Chr(34) & NameEstructura & Chr(34)
                   MyBuffer = MyBuffer & "/>"
                   
                End If
                NumLin = NumLin + 1
            
            Next j
            IndDia = IndDia + 1
            MyBuffer = MyBuffer & "</GrabaMinuta>"
            'Sav_Iu_GrabaMinutaSitioRemoto
            '-------> mover raciones totales
            vaSpread1.Col = i + 3
            vaSpread1.Row = vaSpread1.MaxRows
            NumRacTotal = Val(vaSpread1.text)

            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient

            Set RS = vg_db.Execute("sgpadm_Ins_XmlMinutaBloque_V01 '" & MyBuffer & "', '" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & Vg_FechaDesde & ", " & Vg_FechaHasta & ", " & FecMinuta & ", " & NumRacTotal & ", '" & IIf(EstGrpEst, "A", "N") & "'")
            If Not RS.EOF Then
            
               If RS(0) > 0 Then
                  
                  MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
               
               End If
            
            End If
            RS.Close
            Set RS = Nothing
            Let EstGrpEst = False
        
        Next i
        
        Picture1.Visible = False: gauge.Visible = False
        vaSpread1.Enabled = True
        Main(0).Enabled = True
        Main(1).Enabled = True
        vaSpread1.Refresh
        Toolbar1.Enabled = True
        '-------> Mover Datos grilla principal del formulario m_minsr1
        Sql = ""
        Sql = LimpiaDato(Trim(M_MinSR1.fpText.text))
        Sql = Sql & ", " & M_MinSR1.fpLongInteger1(0).Value & ", " & M_MinSR1.fpLongInteger1(1).Value & ""
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgpadm_Sel_ListarMinutaBloquexCeco " & Sql & "")
        M_MinSR1.vaSpread1.MaxRows = 0
        M_MinSR1.vaSpread1.Row = -1: M_MinSR1.vaSpread1.Col = -1: M_MinSR1.vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
        Do While Not RS.EOF = True
            
            M_MinSR1.vaSpread1.MaxRows = M_MinSR1.vaSpread1.MaxRows + 1
            M_MinSR1.vaSpread1.Row = M_MinSR1.vaSpread1.MaxRows
            
            If RS!IdEstadoMinuta <> 11 Then
                
                M_MinSR1.vaSpread1.Col = -1
                M_MinSR1.vaSpread1.BackColor = M_MinSR1.Shape1(1).FillColor ' Rojo
            
            End If
            
            M_MinSR1.vaSpread1.Col = 2
            M_MinSR1.vaSpread1.text = CStr(RS!Id_Bloque)
            M_MinSR1.vaSpread1.Col = 3
            M_MinSR1.vaSpread1.text = RS!Reg_Codigo & " - " & Trim(RS!reg_nombre)
            M_MinSR1.vaSpread1.Col = 4
            M_MinSR1.vaSpread1.text = RS!Ser_codigo & " - " & Trim(RS!ser_nombre)
            M_MinSR1.vaSpread1.Col = 5
            M_MinSR1.vaSpread1.text = Format(RS!fechadesde, "dd/mm/yyyy")
            M_MinSR1.vaSpread1.Col = 6
            M_MinSR1.vaSpread1.text = Format(RS!fechahasta, "dd/mm/yyyy")
            M_MinSR1.vaSpread1.Col = 7
            M_MinSR1.vaSpread1.text = RS!Reg_Codigo
            M_MinSR1.vaSpread1.Col = 8
            M_MinSR1.vaSpread1.text = RS!Ser_codigo
            M_MinSR1.vaSpread1.Col = 9
            M_MinSR1.vaSpread1.text = Trim(RS!reg_nombre)
            M_MinSR1.vaSpread1.Col = 10
            M_MinSR1.vaSpread1.text = Trim(RS!ser_nombre)
            RS.MoveNext
        
        Loop
        RS.Close
        Set RS = Nothing
        M_MinSR1.FpFecDesde.Enabled = True
        M_MinSR1.FpFecHasta.Enabled = True
        fg_descarga

Exit Sub
Man_Error:
    Picture1.Visible = False: gauge.Visible = False
    vaSpread1.Enabled = True
    Main(0).Enabled = True
    Main(1).Enabled = True
    Toolbar1.Enabled = True
    Call fg_descarga
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, MsgTitulo)
    Call ins_log_error(Date & Time & Err & ":  " & Error$(Err))

End Sub

Sub FormatearGrilla()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim AgrTit          As Long
Dim i               As Long
Dim FecDif          As Variant
Dim FechaInicial    As String
Dim sw              As Boolean
Dim SwPriVez        As Boolean
Dim dato            As Variant
Dim indrow3         As Long
Dim IndDia          As Long
Dim Fecha           As String
Dim Sql1            As String
Dim indblo          As Boolean
Dim estfij          As Boolean
Dim Cabecera        As Long
Dim T               As Long
Dim fecesf          As Long
Dim aAp             As String
Dim vTotRac         As Long
Dim vCosVec         As Double
Dim cosali          As Variant
fg_carga ""

M_MinSR2.Caption = M_MinSR1.Caption
Me.Caption = "Minuta bloque " & IIf(vg_IDBloque = 0, "[Modalidad Nuevo]", "[Modalidad Modificación]")
    
With vaSpread1

    SwSalir = 0
    MaxColumna = 0
    indactivo = 0
    IndGrabado = 0
    vCtoPis = 0
    vCtoTec = 0
    indblo = False
    iblockrow = 0
    iblockrow2 = 0
    iblockcol = 0
    iblockcol2 = 0
    SwSalir = 0
    aiblockrow = 0
    aiblockrow2 = 0
    aiblockcol = 0
    aiblockcol2 = 0
    '-------> determinar la cuando días entre la fecha desde - hasta
    MaxColumna = DateDiff("d", CDate(fg_Ctod1(Vg_FechaDesde)), CDate(fg_Ctod1(Vg_FechaHasta))) + 1
'    '------- Defenir vector costo encabezado
'    ReDim VecCosenc(maxColumna, 2)
'    For i = 1 To UBound(VecCosenc)
'        VecCosenc(i, 1) = 0
'        VecCosenc(i, 2) = 0
'    Next i

    '-------> Traer tipo minuta
    TipoMinuta = False
    AuxTipoMinuta = False
'    Dim AuxTipoMinuta As Boolean
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("select isnull(cli_tipominuta,0) as cli_tipominuta from b_clientes with (nolock) where cli_codigo = '" & vg_codcasino & "' and cli_tipo = 0")
    If Not RS.EOF Then
       
'       TipoMinuta = IIf(IsNull(RS!cli_tipominuta) Or RS!cli_tipominuta = 1, True, False)
'       AuxTipoMinuta = IIf(IsNull(RS!cli_tipominuta) Or RS!cli_tipominuta = 1, True, False)
       TipoMinuta = IIf(RS!cli_tipominuta = 4, True, False)
       AuxTipoMinuta = IIf(RS!cli_tipominuta = 4, True, False)
    
    End If
    RS.Close
    Set RS = Nothing
    
    '-------> Ocultar % si servicio Lyd
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
   
    Set RS = vg_db.Execute("select ser_LYD from a_servicio WITH ( NOLOCK ) where ser_codigo = " & vg_codservicio & " and isnull(ser_LYD,0) = 1")
    If Not RS.EOF Then
       
       TipoMinuta = True
'    Else
'       TipoMinuta = False
    
    End If
    RS.Close
    Set RS = Nothing
    
    .MaxRows = 1000
    .maxcols = 0: .maxcols = 7 * MaxColumna + 2: .Row = 0
    '-------> turn off display of row headers
    'vaSpread1.RowHeadersShow = False
    '-------> Set up column headers
    .ColHeaderRows = 4
    .ShadowColor = &H8000000F
    
    .Col = 1
    .ColsFrozen = 1
    .VisibleCols = 1
    .ColWidth(1) = 25
    .Row = SpreadHeader
    .TypeHAlign = TypeHAlignLeft
    .text = "Costo Patrón Piso"
    .Row = SpreadHeader + 1
    .TypeHAlign = TypeHAlignLeft
    .text = "Costo Minuta Día"
    .Row = SpreadHeader + 2
    .TypeHAlign = TypeHAlignLeft
    .text = "Costo Patrón Techo"
    .Row = SpreadHeader + 3
    .TypeHAlign = TypeHAlignLeft
    .text = "Estructura Servicio"
     
    ' Columna Nueva del Porcentaje por Alejandro Riquelme
    
    .Col = 2
    .ColsFrozen = 2
    .VisibleCols = 2
    .Col = 2
    .Row = SpreadHeader + 3
    .text = "% Pond.Estructura"
    .ColWidth(2) = 13
    .ColHidden = True 'TipoMinuta 'False
    .TypeHAlign = TypeHAlignLeft
    
    '-----------------------------------------
'    ReDim Preserve vectorcol(0)
    ReDim vectorcol(0)
    Dim FechaIni As Date
    Let FechaIni = fg_Ctod1(Vg_FechaDesde)
    Let AgrTit = 2
    Let sw = False
    Let SwPriVez = True
    
    For i = 3 To .maxcols Step 7
        
        .Col = i
        .Row = SpreadHeader + 3
        .ColWidth(i) = 1.5
        .text = " "
        .ColHidden = False
        
        .Col = i + 1
        .ColWidth(i + 1) = 21
        
        If sw = True Then
            
            Let AgrTit = AgrTit + 7
        
        Else
            
            If SwPriVez = True Then
                
                Let sw = True
            
            Else
                
                Let sw = True
                Let AgrTit = AgrTit + 7
            
            End If
        
        End If
        
        Let SwPriVez = False
        If Not IsDate(FechaIni) Then
           
           Let FechaIni = FechaIni + 1
           Let AgrTit = 2
           Let sw = False
        
        Else
            
            Let sw = True
        
        End If
        
        
        If i = 3 Then
           
           ReDim Preserve vectorcol(1)
           vectorcol(1) = 4
          .text = " " & Mid(fg_Fecha_Dia(Format(FechaIni, "yyyymmdd"), 1), 1, 3) & " " & FechaIni
        
        Else
           
           .text = " " & Mid(fg_Fecha_Dia(Format(FechaIni, "yyyymmdd"), 1), 1, 3) & " " & FechaIni
            ReDim Preserve vectorcol(CLng((i / 7) + 1))
            vectorcol(CLng((i / 7) + 1)) = i + 1
        
        End If
        
        Let FechaIni = FechaIni + 1
        .ColHidden = False
        
        .Col = i + 2
        .text = "% Pond."
        .ColHidden = TipoMinuta
        
        .Col = i + 3
        .ColWidth(i + 2) = 6
        .text = "N.Rac."
        .ColHidden = False
       
        .Col = i + 4
        .ColWidth(i + 3) = 6
        .text = "Cto.Plato"
        .ColHidden = False
        
        .Col = i + 5
        .text = "Cod. Receta"
        .ColHidden = True
        
        .Col = i + 6
        .ColWidth(i + 3) = 6
        .text = "Calorías"
        .ColHidden = False

        For j = 1 To .MaxRows
            
            .Row = j
    
            .Col = i
            .CellType = CellTypeStaticText
            .TypeHAlign = TypeHAlignLeft
            .text = " "
    
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
    
    ' Columna Nueva Alejandro Riquelme % Porcentaje Diario
    
             .Col = i + 5
             .CellType = CellTypeStaticText
             .TypeHAlign = TypeHAlignLeft
             .text = " "
             
             .Col = i + 6
             .CellType = CellTypeStaticText
             .TypeHAlign = TypeHAlignLeft
             .text = " "
    
        Next j
    
    Next i

    .maxcols = .maxcols + 1
    .Col = .maxcols
    .ColWidth(.maxcols) = 5
    .text = "Còd. Est."
    .ColHidden = True
    
    ' INI ARI
    ' Se creo dos Culmnas para los GRrpos de Estyructuras ARI
    
    .maxcols = .maxcols + 1
    .ColID = "Grupo"
    .Col = .maxcols
    .ColWidth(.maxcols) = 5
    .text = "Còd. Grupo."
    .ColHidden = True
    
    ' FIN ARI
   
    For i = 3 To .maxcols - 2 Step 7
        
        .AddCellSpan i, SpreadHeader, 6, 1
        .Col = i
        .Row = 0 'SpreadHeader
        .TypeHAlign = TypeHAlignRight
        .text = Format(0, fg_Pict(6, 2))
        .AddCellSpan i, SpreadHeader + 1, 6, 1
        .Col = i
        .Row = SpreadHeader + 1
        .TypeHAlign = TypeHAlignRight
        .text = Format(0, fg_Pict(6, 2))
        .AddCellSpan i, SpreadHeader + 2, 6, 1
        .Col = i
        .Row = SpreadHeader + 2
        .TypeHAlign = TypeHAlignRight
        .text = Format(0, fg_Pict(6, 2))
    
    Next i

    '-------> Mover costo patron y patron
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Sel_DetCostoPisoPatron '" & vg_codcasino & "', " & vg_codregimen & ", " & vg_codservicio & ", " & Mid(Vg_FechaDesde, 1, 6) & ", " & Mid(Vg_FechaHasta, 1, 6) & "")
    Do While Not RS.EOF
        
        .ShadowText = &H800000
        
        For i = 3 To .maxcols - 2 Step 7
            
            If Trim(RS!pcp_descripcion) = "PISO" Then
               
               'Poner costo de acuerdo al periodo
               .Row = SpreadHeader + 3
               .Col = i + 1
               
               If RS!pcp_anomes = Val(Format(Mid(Trim(.text), 5, Len(Trim(.text))), "yyyymm")) Then
                  
                  .Col = i
                  .Row = SpreadHeader
                  .text = Format(RS!pcp_valor, fg_Pict(6, 2))
               
               End If
            
            End If
            
            If Trim(RS!pcp_descripcion) = "TECHO" Then
               
               'Poner costo de acuerdo al periodo
               .Row = SpreadHeader + 3
               .Col = i + 1
               
               If RS!pcp_anomes = Val(Format(Mid(Trim(.text), 5, Len(Trim(.text))), "yyyymm")) Then
                  
                  .Col = i
                  .Row = SpreadHeader + 2
                  .text = Format(RS!pcp_valor, fg_Pict(6, 2))
               
               End If
            
            End If
        
        Next i
        
        RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    '-------> Fin mover costo piso y patron
    
    .Row = -1: .Col = -1
    .BackColor = Shape1(0).FillColor  'Amarillo
    .Row = -1: vaSpread1.Col = 1
    .Font.Bold = True
    .Font.Size = 9
    .BackColor = Shape1(2).FillColor 'Verde
    
    '-------> ocultar check
    Check1(4).Visible = IIf(TipoMinuta, False, True)
    Check1(5).Visible = IIf(TipoMinuta, False, True)
    fg_descarga

End With

    ReDim veccos(MaxColumna, 6)

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub DetallePlantillaMinuta()

On Error GoTo Man_Error

Dim RS              As New ADODB.Recordset
Dim RS1             As New ADODB.Recordset
Dim AgrTit          As Long
Dim i               As Long
Dim FecDif          As Variant
Dim FechaInicial    As String
Dim sw              As Boolean
Dim SwPriVez        As Boolean
Dim dato            As Variant
Dim indrow3         As Long
Dim IndDia          As Long
Dim Fecha           As String
Dim Sql1            As String
Dim indblo          As Boolean
Dim estfij          As Boolean
Dim Cabecera        As Long
Dim T               As Long
Dim fecesf          As Long
Dim aAp             As String
Dim vTotRac         As Long
Dim vCosVec         As Double
Dim cosali          As Variant
Dim j               As Long
Dim ExitMinuta      As Boolean

fg_carga ""
    
'-------> Defenir vector costo encabezado
ReDim VecCosenc(MaxColumna, 2)
For i = 1 To UBound(VecCosenc)
    VecCosenc(i, 1) = 0
    VecCosenc(i, 2) = 0
Next i

With vaSpread1

    SwSalir = 0
    indactivo = 0: IndGrabado = 0: vCtoPis = 0: vCtoTec = 0
    indblo = False
    iblockrow = 0: iblockrow2 = 0: iblockcol = 0: iblockcol2 = 0: SwSalir = 0
    aiblockrow = 0: aiblockrow2 = 0: aiblockcol = 0: aiblockcol2 = 0
    .Row = -1: .Col = -1
    vaSpread1.BackColor = Shape1(0).FillColor  'Amarillo
    vaSpread1.Row = -1: vaSpread1.Col = 1
    vaSpread1.Font.Bold = True
    vaSpread1.Font.Size = 9
    vaSpread1.BackColor = Shape1(2).FillColor 'Verde
    
    j = 0: i = 0: indrow3 = 0

    'MVA - MVI - ACA ES DONDE SE DEBE CAMBIAR PARA COSTEO DE MINUTA - 2013-01-18
    Dim ContadorLinea As Long
    Dim agrupacion As Long
    Dim fecha_emi As Long
    Dim contadoractual As Long
    Dim cantidaGrupo As Long
    Dim cantidaFecha As Long
    Dim row_anterior As Long
    Dim nombre_estructura As String
    
    Dim colgrupo As Long
    Dim CodigoGrupo As Long
    Dim RowGrupo As Long
    Dim GrupoActual As Long
    Dim Col_Porc As Long
    Dim AuxGrupoEst As Long
    Dim CosEst As Long
    cantidaGrupo = 0
    contadoractual = 0
    CosEst = 0
    ContadorLinea = 1
    agrupacion = -1
    
    Dim SeleccionOpt As Integer
    
    ExitMinuta = False
    SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_Sel_DetalleMinutaBloque_V05 " & vg_IDBloque & ", '" & vg_codcasino & "', " & SeleccionOpt & "")

    If Not RS.EOF Then
        
        AuxGrupoEst = -1
        
        Do While Not RS.EOF
         
           .Row = RS!mid_numlin
           
           '------> Buscar numero linea
           For i = 3 To (vaSpread1.maxcols - 2) Step 7
               
               Call vaSpread1.GetText(i + 1, SpreadHeader + 3, dato)
               If RS!min_fecmin = Val(Format(Mid(dato, 5, Len(dato)), "yyyymmdd")) Then j = i: Exit For
           
           Next i
           
           'columna de grupo de estructura y Encabezado
           If Not TipoMinuta Then
           colgrupo = vaSpread1.GetColFromID("Grupo") ' + 1
           '-------> Buscar grupo estructura servicio
           RowGrupo = vaSpread1.SearchCol(colgrupo, 0, -1, RS!ess_codigo, SearchFlagsValue)
           If RowGrupo = -1 Then
              
              .Col = 1
              .CellType = CellTypeStaticText
              .text = RS!ess_nombre
              
              If RS!ess_agrupacionestructura <> AuxGrupoEst Then
                 
                 If RS!mge_ponderaciontotal > -1 Then
                    
                    .Col = 2
                    .CellType = CellTypePercent
                    .TypeHAlign = 1
                    .TypePercentDecPlaces = 0
                    .ForeColor = &HFF0000
                    .text = RS!mge_ponderaciontotal
                    .TypeNegRed = True
                 
                 Else
                    
                    .Col = 2
                    
                    If TipoMinuta Then
                       
                       .CellType = CellTypeStaticText
                       .TypeHAlign = TypeHAlignCenter
                    
                    Else
                       
                       .CellType = CellTypePercent
                       .TypeHAlign = 1
                       .TypePercentDecPlaces = 0
                       .ForeColor = &HFF0000
                       .text = 0
                       .TypeNegRed = True
                    
                    End If
                 
                 End If
                 
                 AuxGrupoEst = RS!ess_agrupacionestructura
              
              Else
                 
                 .Col = 2
                 .CellType = CellTypeStaticText
                 .TypeHAlign = TypeHAlignCenter
              
              End If
            
            Else
            
           End If
           
           Else
              
              If RS!ess_codigo <> CosEst Then
                 
                 vaSpread1.Col = 1
                 
                 If IIf(IsNull(RS!ess_codigo), "", RS!ess_codigo) <> "" Then
                    
                    .text = RS!ess_nombre
                 
                 Else
                    
                    vaSpread1.text = Trim(RS!ess_nombre)
                 
                 End If
                 
                 CosEst = RS!ess_codigo
              
              End If
           
           End If
           
           .Col = j
           .CellType = CellTypeStaticText
           .TypeHAlign = TypeHAlignCenter
           .Value = "R"
           .ForeColor = &HFF&
           .BackColor = &H80FF80
           
           '-------> Descripción receta
           .Col = j + 1
           .CellType = CellTypeStaticText
           .TypeHAlign = TypeHAlignLeft
'           .text = RS!rec_nombre
           .text = IIf(RS!rec_LYD, "[*] ", "") & RS!rec_nombre
                             
           '-------> porcentaje diario
           .Col = j + 2
           If RS!rec_LYD Then
              
              .CellType = CellTypeStaticText
              .TypeHAlign = TypeHAlignRight
              .text = ""
           
           Else
              
              .CellType = CellTypePercent
              .TypePercentLeadingZero = TypeLeadingZeroYes
              .TypePercentNegStyle = TypePercentNegStyle8
              .TypePercentDecPlaces = 0
              .TypePercentMax = 1000
              ' display negative numbers as red
              .TypeNegRed = True
              .ForeColor = &HFF0000
              .text = RS!mid_porrac
           
           End If
'           .ColHidden = False

           .Col = j + 3
           If RS!MIN_INDBLO = 11 Then
              
              .CellType = IIf(TipoMinuta Or RS!rec_LYD, CellTypeNumber, CellTypeStaticText)
              .TypeNumberDecPlaces = 0
              .TypeNumberMin = 0
              .TypeNumberMax = 9999999
              .TypeHAlign = 1
              .TypeSpin = False
              .TypeIntegerSpinInc = 1
              .TypeIntegerSpinWrap = False
           
           Else
              
              .CellType = CellTypeStaticText
              .TypeHAlign = TypeHAlignRight
           
           End If
          .text = Format(RS!mid_numrac, fg_Pict(6, 0))
          .ForeColor = &HFF0000

          '-------> Mover costo alimentación y desechable
          .Col = j + 4
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignRight
          .text = Format(RS!promedioreceta, fg_Pict(6, 2))

          '-------> Mover codigo recetas
          .Col = j + 5
          .text = RS!rec_codigo & "&" & RS!mid_tiprec & "&;"
          If RS!mid_modminb = "1" Then
             
             '-------> Mover tipo celda nota cuando sucede un cambio
             .TextTip = TextTipFloating
             .CellNoteIndicator = CellNoteIndicatorShowAndFireEvent
             .CellNote = "Cambio"
          
          End If

          '-------> Mover aporte calorias
          .Col = j + 6
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignRight
          .text = Format(RS!AporteNut, fg_Pict(6, 2))
          
          '-------> Mover costo minuta dia alimentación
          If Not RS!rec_LYD Then
             
             VecCosenc((Int((j + 1) / 7) + 1), 1) = (VecCosenc((Int((j + 1) / 7) + 1), 1) + (RS!promedioreceta * IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac)))
          
          ElseIf RS!rec_LYD Then
             
             VecCosenc((Int((j + 1) / 7) + 1), 2) = (VecCosenc((Int((j + 1) / 7) + 1), 2) + (RS!promedioreceta * IIf(IsNull(RS!mid_numrac), 0, RS!mid_numrac)))
          
          End If
          
          If RS!MIN_INDBLO <> 11 Then
             
             .Row = -1: .Col = j: .BackColor = Shape1(1).FillColor: .Lock = True
             .Row = RS!mid_numlin
             .Col = j + 1: .BackColor = Shape1(1).FillColor: .Lock = True
             .Col = j + 2: .BackColor = Shape1(1).FillColor: .Lock = True
             .Col = j + 3: .BackColor = Shape1(1).FillColor: .Lock = True
             .Col = j + 4: .BackColor = Shape1(1).FillColor: .Lock = True
             .Col = j + 5: .BackColor = Shape1(1).FillColor: .Lock = True
             .Col = j + 6: .BackColor = Shape1(1).FillColor: .Lock = True
          
          ElseIf RS!mid_modmina = "1" Then
             
             .Col = j: .BackColor = Shape1(3).FillColor
             .Col = j + 1: .BackColor = Shape1(3).FillColor
             .Col = j + 2: .BackColor = Shape1(3).FillColor
             .Col = j + 3: .BackColor = Shape1(3).FillColor
             .Col = j + 4: .BackColor = Shape1(3).FillColor
             .Col = j + 5: .BackColor = Shape1(3).FillColor
             .Col = j + 6: .BackColor = Shape1(3).FillColor
          
          End If
        
          .Row = RS!mid_numlin
          
          .Col = .maxcols - 1
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignCenter
          .text = RS!ess_codigo

          .Col = .maxcols
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignCenter
          .text = IIf(TipoMinuta, 0, RS!ess_agrupacionestructura)
          
           
          If ContadorLinea < .Row Then ContadorLinea = .Row
          RS.MoveNext
       
       Loop
       
       ExitMinuta = True
       ContadorLinea = ContadorLinea + 1
       RS.Close
       Set RS = Nothing
       fg_descarga
    
    Else
       
       RS.Close
       Set RS = Nothing
       fg_descarga
       ExitMinuta = False
       ContadorLinea = 1
       agrupacion = -1
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       Set RS = vg_db.Execute("sgpadm_Sel_PonderacionTotalAgrupacion '" & vg_codcasino & "'," & vg_codservicio & "")
       
       If RS.EOF Then
       
          RS.Close
          Set RS = Nothing
          Exit Sub
       
       End If
       
       Do While Not RS.EOF
            
            .Row = ContadorLinea
     
            If agrupacion <> RS!ess_agrupacionestructura Then
          
               .Col = 2
               .CellType = CellTypePercent
               .BackColor = Shape1(2).FillColor
               .text = RS!mge_ponderaciontotal
          
               vaSpread1.CellType = CellTypePercent
               vaSpread1.TypePercentLeadingZero = TypeLeadingZeroYes
               vaSpread1.TypePercentNegStyle = TypePercentNegStyle8
               vaSpread1.TypePercentDecPlaces = 0
               vaSpread1.TypePercentMax = 1000
               ' display negative numbers as red
               vaSpread1.TypeNegRed = True
               agrupacion = RS!ess_agrupacionestructura
           
            End If
            
            '------->Mover descripcion estructura servicio
            .Col = 1
            .text = IIf(IsNull(RS!ess_nombre), "No existe estructura servicio", RS!ess_nombre)
            '------->Mover codigo estructura servicio
            .Col = .maxcols - 1
            .text = IIf(IsNull(RS!ess_codigo), 0, RS!ess_codigo)
            '------->Mover codigo agrupacion estructura servicio
            .Col = .maxcols
            .text = IIf(TipoMinuta, 0, IIf(IsNull(RS!ess_agrupacionestructura), 0, RS!ess_agrupacionestructura))

            RS.MoveNext
            ContadorLinea = ContadorLinea + 1
      
       Loop
       
       RS.Close
       Set RS = Nothing
    
    End If
    .Row = -1: .Col = -1
    '------->Mover color a la columna estructura servicio
    .Row = -1: .Col = 1
    .Font.Bold = True
    .Font.Size = 9
    .BackColor = Shape1(2).FillColor
    
    .MaxRows = ContadorLinea
    .Row = .MaxRows
    maxfila = .MaxRows
    .Col = 1
    .text = "Comensales"
    .Col = -1: .BackColor = &HE0E0E0
    
    '-------> Formatear ultima columna
    For i = 3 To (.maxcols - 1) Step 7
        
        .Row = .MaxRows - 1
        .Col = i + 3
        
        If .BackColor <> Shape1(1).FillColor Then
           
           .Row = .MaxRows
           .CellType = CellTypeNumber
           .TypeNumberDecPlaces = 0
           .TypeNumberMin = 0
           .TypeNumberMax = 9999999
           .TypeHAlign = TypeHAlignRight
           .TypeSpin = False
           .TypeIntegerSpinInc = 1
           .TypeIntegerSpinWrap = False
        
        Else
           
           .Row = .MaxRows
           .Lock = True
           .CellType = CellTypeStaticText
           .TypeHAlign = TypeHAlignRight
        
        End If
        .Value = Format(0, fg_Pict(6, 0))
        .ForeColor = &HFF0000
    
    Next i
    
    .Col = .maxcols - 1
    .text = 0
    .Col = .maxcols
    .text = 999
             
    '-------> Mover comensales
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("SELECT DISTINCT min_racteo, min_fecmin, min_indblo " & _
            "FROM cas_b_minuta With(NoLock) " & _
            "WHERE  min_cecori = '" & vg_codcasino & "' " & _
            " AND   min_cencos = '" & vg_codcasino & "' " & _
            " AND   min_codreg = " & vg_codregimen & " " & _
            " AND   min_codser = " & vg_codservicio & " " & _
            " AND   min_fecmin >= " & Val(Vg_FechaDesde) & " " & _
            " AND   min_fecmin <= " & Val(Vg_FechaHasta) & " AND id_bloque = " & vg_IDBloque & "" & _
            " ORDER BY min_fecmin")
    
    Do While Not RS.EOF
       
       '-------> Buscar día
       For i = 4 To (.maxcols - 1) Step 7
           
           .Row = SpreadHeader + 3
           .Col = i
           
           If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) = fg_Ctod1(RS!min_fecmin) Then
              
              j = .Col + 1
              Exit For
           
           End If
       
       Next i
       
       .Row = .MaxRows
       .Col = j + 1
       
       If RS!MIN_INDBLO = 11 Then
          
          .CellType = CellTypeNumber
          .TypeNumberDecPlaces = 0
          .TypeNumberMin = 0
          .TypeNumberMax = 9999999
          .TypeHAlign = TypeHAlignRight
          .TypeSpin = False
          .TypeIntegerSpinInc = 1
          .TypeIntegerSpinWrap = False
       
       Else
         
         .Lock = True
         .CellType = CellTypeStaticText
         .TypeHAlign = TypeHAlignRight
         '-------> Poner color a la primera columna
         .Row = .MaxRows
         .Col = 1
         .BackColor = Shape1(2).FillColor
         .Row = .MaxRows
         .Col = j + 1
       
       End If
       
       .Value = IIf(IsNull(RS!min_racteo), 0, RS!min_racteo)
       .ForeColor = &HFF0000
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing

    '-------> Bloquear dias
    Dim FecMin As Date
    Dim FecMax As Date
    j = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    Set RS = vg_db.Execute("sgpadm_Sel_TraerFechaMaxMinBloque '" & vg_codcasino & "', " & vg_IDBloque & "")
    If Not RS.EOF Then
       
       If Not IsNull(RS!fechamax) Then
          
          FecMin = fg_Ctod1(Vg_FechaDesde)
          FecMax = fg_Ctod1(RS!fechamax)
          
          Do While FecMin <= FecMax
             
             For i = 4 To (.maxcols - 1) Step 7
                 
                 .Row = SpreadHeader + 3
                 .Col = i
                 
                 If CDate(Mid(Trim(.text), 5, Len(Trim(.text)))) = FecMin Then
                    
                    j = .Col + 1
                    Exit For
                 
                 End If
             
             Next i
          
             For i = 1 To .MaxRows - 1
                
                .Row = i
                
                If FecMin = fg_Ctod1(Vg_FechaHasta) Then
                   
                   .Col = 2: .BackColor = Shape1(1).FillColor: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .Lock = True
                
                End If
                
                .Col = j - 2: .BackColor = Shape1(1).FillColor: .Lock = True
                .Col = j - 1: .BackColor = Shape1(1).FillColor: .Lock = True
                .Col = j: .BackColor = Shape1(1).FillColor: .CellType = CellTypeStaticText: .TypeHAlign = TypeHAlignRight: .Lock = True
                .Col = j + 1: .BackColor = Shape1(1).FillColor: .Lock = True
                .Col = j + 2: .BackColor = Shape1(1).FillColor: .Lock = True
                .Col = j + 3: .BackColor = Shape1(1).FillColor: .Lock = True
                .Col = j + 4: .BackColor = Shape1(1).FillColor: .Lock = True
             
             Next i
             
             FecMin = FecMin + 1
          
          Loop
       
       End If
    
    End If
    
    RS.Close
    Set RS = Nothing
    '-------> Mostrar costo minuta dia
    j = 3
    
    For i = 1 To UBound(VecCosenc)
        
        .Row = vaSpread1.MaxRows
        .Col = j + 3
        vTotRac = 0
        
        If Trim(.text) <> "" And Val(.text) <> 0 Then
            
            Let vTotRac = .text
        
        End If
        
        .Row = SpreadHeader + 1
        .Col = j
        .TypeHAlign = TypeHAlignRight
        vCosVec = 0
        vCosVec = Round(VecCosenc(i, 1) + VecCosenc(i, 2), 2)
        If vTotRac > 0 And vCosVec > 0 Then .text = Format(Round(vCosVec / vTotRac, 2), fg_Pict(6, 2)) Else .text = Format(0, fg_Pict(6, 2))
        j = j + 7
    
    Next i
    
    '------->
    If Not TipoMinuta And ExitMinuta Then

'       TipoCopia = "xxx"

       For j = 1 To vaSpread1.MaxRows - 1

           For i = 4 To vaSpread1.maxcols - 2 Step 7

               Call Cal_PorSugerido_Raciones(j, i)

           Next i

       Next j

    End If
    
    .Row = .MaxRows
    .Col = 1
    .text = "Comensales"
    .Col = -1: .BackColor = &HE0E0E0
    
    .Row = 1
    .Col = 1
    iblockrow = .Row: aiblockrow = .Row
    iblockrow2 = .Row: aiblockrow2 = .Row
    iblockcol = .Col: aiblockcol = .Col
    iblockcol2 = .Col: aiblockcol2 = .Col
    
    .SetActiveCell 1, 1

End With

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub CalctodiaEnc(Row As Long, Col As Long)

On Error GoTo Man_Error

Dim X       As Long
Dim numrac  As Double
Dim vCosVec As Double
Dim vTotRac As Double
Dim cosdia  As Double
Dim dato As Variant

'-------> calcular costo día

    Call vaSpread1.GetText(Col + 1, SpreadHeader + 3, dato)
    If "N.Rac." = dato And Row <> vaSpread1.MaxRows Then
        Col = (Col - 1)
    End If
    
    
    If Not TipoMinuta Then ' Solamente para tipo minuta simap
       
       '-------> Suma los Valores de Porcentaje Sugerido y calcula las Raciones de los Comensales
       Call Cal_PorSugerido_Raciones(Row, Col)
    
    End If
    
   If Col <> 1 Then
       VecCosenc((Int(Col / 7) + 1), 1) = 0
       VecCosenc((Int(Col / 7) + 1), 2) = 0
       
       For X = 1 To (vaSpread1.MaxRows - 1)
           
           vaSpread1.Row = X
           vaSpread1.Col = IIf(Row <> vaSpread1.MaxRows, Col + 2, Col + 1)
           
           If Trim(vaSpread1.text) <> "" Then
              
              numrac = IIf(Val(vaSpread1.text) = 0 Or Trim(vaSpread1.text) = "", 0, Val(vaSpread1.text))
           
           End If
           
           vaSpread1.Col = IIf(Row <> vaSpread1.MaxRows, Col + 3, Col + 2)
           
           If Trim(vaSpread1.text) <> "" Then
              
              cosdia = IIf(Val(vaSpread1.text) = 0 Or Trim(vaSpread1.text) = "", 0, CCur(vaSpread1.text))
           
           End If
           
           vaSpread1.Col = IIf(Row <> vaSpread1.MaxRows, Col + 4, Col + 3)
           
           If Trim(vaSpread1.text) <> "" And numrac > 0 Then
'              vaSpread1.Col = Col + 2: VecCosenc((Int(Col / 7) + 1), 1) = Round(VecCosenc((Int(Col / 7) + 1), 1) + (cosdia * numrac), vg_DCa)
              vaSpread1.Col = Col ' - 1
              
              If Mid(Trim(vaSpread1.text), 1, 3) <> "[*]" Then
                 
                 VecCosenc((Int(Col / 7) + 1), 1) = Round(VecCosenc((Int(Col / 7) + 1), 1) + (cosdia * numrac), vg_DCa)
              
              Else
                 
                 VecCosenc((Int(Col / 7) + 1), 2) = Round(VecCosenc((Int(Col / 7) + 1), 2) + (cosdia * numrac), vg_DCa)
              
              End If
           
           End If
        
        Next X
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = IIf(Row <> vaSpread1.MaxRows, Col + 2, Col + 1)
        vTotRac = 0
        
        If Trim(vaSpread1.text) <> "" And Val(vaSpread1.text) <> 0 Then
           
           vTotRac = vaSpread1.text
        
        End If
        
        vaSpread1.Row = SpreadHeader + 1
        vaSpread1.Col = IIf(Row <> vaSpread1.MaxRows, Col - 1, Col - 2)
        vaSpread1.TypeHAlign = TypeHAlignRight
        vCosVec = 0
        vCosVec = Round(VecCosenc((Int(Col / 7) + 1), 1) + VecCosenc((Int(Col / 7) + 1), 2), 2)
        If vTotRac > 0 And vCosVec > 0 Then
           
           vaSpread1.text = Format(Round(vCosVec / vTotRac, 2), fg_Pict(6, 2))
        
        Else
           
           vaSpread1.text = Format(0, fg_Pict(6, 2))
        
        End If
    
    End If
    
Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub Cal_PorSugerido_Raciones(Row As Long, Col As Long)

On Error GoTo Man_Error

Dim dato As Variant
Dim AuxCol As Long
Dim CodGrupoEst As Long
Dim RowGrupo As Long
Dim colgrupo  As Long
Dim i As Long
Dim j As Long
Dim ii As Long
Dim NumRacTotal As Long
Dim ValPorDiario As Long
Dim TotPorDiario As Long
Dim PorGrpEst As Long
Dim NomReceta As String

If Not TipoMinuta Then ' Solamente para tipo minuta simap
    
    With Me.vaSpread1
         Call .GetText(Col + 1, SpreadHeader + 3, dato)
         
         If "% Pond." = dato And TipoCopia <> "Copia Raciones" Then
            
            '-------> Mover codigo grupo estructura
            .Row = Row
            .Col = .maxcols
            CodGrupoEst = IIf(Trim(.text) = "", -1, .text)
            '-------> columna de grupo de estructura y Encabezado
            colgrupo = vaSpread1.GetColFromID("Grupo") + 1
            '-------> fila a buscar
            RowGrupo = .SearchCol(colgrupo, 0, -1, CodGrupoEst, SearchFlagsValue)
            
            If RowGrupo > -1 Then
               
               '-------> sacar porcentaje grupo estructura
               PorGrpEst = 0
               TotPorDiario = 0
               .Row = RowGrupo
               .Col = 2
               PorGrpEst = Val(.text)
               '-------> sacar Total raciones
               NumRacTotal = 0
               .Col = Col + 2
               .Row = .MaxRows
               NumRacTotal = IIf(Trim(.text) = "", 0, .text)
               '-------> sacar raciones % diario * raciones totales
               
               For i = RowGrupo To .MaxRows - 1
                   
                   .Row = i
                   '-------> sacar nombre de la receta
                   .Col = Col
                   NomReceta = .text
                   .Col = .maxcols
                   
                   If (CStr(CodGrupoEst) <> .text) And Trim(.text) <> "" Then
                      
                      '------> Poner color rojo cuando % diario > % grupo estructura en caso contrario pone azul
                      If TotPorDiario > PorGrpEst Then
                         
                         For ii = RowGrupo To .MaxRows - 1
                             
                             .Row = ii
                             '-------> sacar nombre de la receta
                             .Col = Col
                             NomReceta = .text
                             .Col = .maxcols
                             If (CStr(CodGrupoEst) <> .text) And Trim(.text) <> "" Then ValPorDiario = 0: TotPorDiario = 0: Exit For
                             
                             If Trim(NomReceta) <> "" Then
                                
                                .Col = Col + 1
                                .ForeColor = &HFF&
                             
                             End If
                         
                         Next ii
                         
                         TotPorDiario = 0
                         ValPorDiario = 0
                      
                      Else
                         
                         For ii = RowGrupo To .MaxRows - 1
                             
                             .Row = ii
                             .Col = Col
                             NomReceta = .text
                             .Col = .maxcols
                             If (CStr(CodGrupoEst) <> .text) And Trim(.text) <> "" Then TotPorDiario = 0: Exit For
                             
                             If Trim(NomReceta) <> "" Then
                                
                                .Col = Col + 1
                                .ForeColor = &HFF0000
                             
                             End If
                         
                         Next ii
                      
                      End If
                      '-------> Mover codigo grupo estructura
                      .Row = i
                      .Col = .maxcols
                      CodGrupoEst = IIf(Trim(.text) = "", -1, .text)
                      '-------> columna de grupo de estructura y Encabezado
                      colgrupo = vaSpread1.GetColFromID("Grupo") + 1
                      '-------> fila a buscar
                      RowGrupo = .SearchCol(colgrupo, 0, -1, CodGrupoEst, SearchFlagsValue)
                      
                      If RowGrupo < -1 Then
                         
                         Exit For
                      
                      End If
                   
                   End If
                   
                   If Trim(NomReceta) <> "" Then
                      
                      .Col = Col + 1
                      ValPorDiario = Val(.text)
                      .Col = Col + 2
                      
                      If ValPorDiario > 0 Then
                         
                         .text = Redondear(((ValPorDiario / 100) * NumRacTotal), 0)
                      
                      Else
                         
                         .Col = Col + 1
                         
                         If Trim(.text) <> "" Then
                            
                            .Col = Col + 2
                            .text = 0
                         
                         End If
                      
                      End If
                      
                      TotPorDiario = TotPorDiario + ValPorDiario
                   
                   End If
               
               Next i
               '------> Poner color rojo cuando % diario > % grupo estructura en caso contrario pone azul
               If TotPorDiario > PorGrpEst Then
                  
                  For i = RowGrupo To .MaxRows - 1
                      
                      .Row = i
                      '-------> sacar nombre de la receta
                      .Col = Col
                      NomReceta = .text
                      .Col = .maxcols
'                      If (CStr(CodGrupoEst) <> .text Or Trim(NomReceta) = "") And Trim(.text) <> "" Then Exit For
                      If (CStr(CodGrupoEst) <> .text) And Trim(.text) <> "" Then Exit For
                      
                      If Trim(NomReceta) <> "" Then
                         
                         .Col = Col + 1
                         .ForeColor = &HFF&
                      
                      End If
                  
                  Next i
               
               Else
                  
                  For i = RowGrupo To .MaxRows - 1
                      
                      .Row = i
                      .Col = Col
                      NomReceta = .text
                      .Col = .maxcols
'                      If (CStr(CodGrupoEst) <> .text Or Trim(NomReceta) = "") And Trim(.text) <> "" Then Exit For
                      If (CStr(CodGrupoEst) <> .text) And Trim(.text) <> "" Then Exit For
                      
                      If Trim(NomReceta) <> "" Then
                         
                         .Col = Col + 1
                         .ForeColor = &HFF0000
                      
                      End If
                  
                  Next i
               
               End If
            End If
         
         ElseIf "% Pond.Estructura" = dato Then
            
            '-------> Mover codigo grupo estructura
            .Row = Row
            .Col = .maxcols
            CodGrupoEst = IIf(Trim(.text) = "", -1, .text)
            '-------> columna de grupo de estructura y Encabezado
            colgrupo = vaSpread1.GetColFromID("Grupo") + 1
            '-------> fila a buscar
            RowGrupo = .SearchCol(colgrupo, 0, -1, CodGrupoEst, SearchFlagsValue)
            '-------> Mover codigo grupo estructura
            .Row = Row
            .Col = .maxcols
            CodGrupoEst = IIf(Trim(.text) = "", -1, .text)
            '-------> columna de grupo de estructura y Encabezado
            colgrupo = vaSpread1.GetColFromID("Grupo") + 1
            '-------> fila a buscar
            RowGrupo = .SearchCol(colgrupo, 0, -1, CodGrupoEst, SearchFlagsValue)
            AuxCol = Col
            
            If RowGrupo > -1 Then

               For j = 4 To vaSpread1.maxcols - 2 Step 7
               
               Col = j
               '-------> sacar porcentaje grupo estructura
               PorGrpEst = 0
               TotPorDiario = 0
               .Row = RowGrupo
               .Col = 2
               PorGrpEst = Val(.text)
               '-------> sacar Total raciones
               NumRacTotal = 0
               .Col = Col + 2
               .Row = .MaxRows
               NumRacTotal = IIf(Trim(.text) = "", 0, .text)
               '-------> sacar raciones % diario * raciones totales

               For i = RowGrupo To .MaxRows - 1
                   
                   .Row = i
                   '-------> sacar nombre de la receta
                   .Col = Col
                   NomReceta = .text
                   .Col = .maxcols
                   If (CStr(CodGrupoEst) <> .text Or Trim(NomReceta) = "") And Trim(.text) <> "" Then Exit For
                   If Trim(NomReceta) <> "" Then
                      
                      .Col = Col + 1
                      ValPorDiario = Val(.text)
                      .Col = Col + 2
                      
                      If Mid(Trim(NomReceta), 1, 3) <> "[*]" Then
                         
                         .text = ((ValPorDiario / 100) * NumRacTotal)
                      
                      End If
                      
                      TotPorDiario = TotPorDiario + ValPorDiario
                   
                   End If
               
               Next i


               '------> Poner color rojo cuando % diario > % grupo estructura en caso contrario pone azul
               If TotPorDiario > PorGrpEst Then

                  For i = RowGrupo To .MaxRows - 1

                      .Row = i
                      '-------> sacar nombre de la receta
                      .Col = Col
                      NomReceta = .text
                      .Col = .maxcols
                      If (CStr(CodGrupoEst) <> .text Or Trim(NomReceta) = "") And Trim(.text) <> "" Then Exit For

                      If Trim(NomReceta) <> "" Then
                         
                         .Col = Col + 1
                         .ForeColor = &HFF&
                      
                      End If

                  Next i

               Else

                  For i = RowGrupo To .MaxRows - 1
                      
                      .Row = i
                      .Col = Col
                      NomReceta = .text
                      .Col = .maxcols
                      If (CStr(CodGrupoEst) <> .text Or Trim(NomReceta) = "") And Trim(.text) <> "" Then Exit For
                      
                      If Trim(NomReceta) <> "" Then
                         
                         .Col = Col + 1
                         .ForeColor = &HFF0000
                      
                      End If
                  
                  Next i

               End If

               Next j

            End If
            
            Col = AuxCol
         
         ElseIf .Row = .MaxRows Or TipoCopia = "Copia Raciones" Then
               
               '-------> sacar raciones % diario * raciones totales
               .Col = IIf(TipoCopia = "Copia Raciones", Col + 2, Col + 1)
               .Row = .MaxRows
               NumRacTotal = IIf(Trim(.text) = "", 0, .text)
               
               For i = 1 To .MaxRows - 1
                   
                   .Row = i
                   '-------> sacar nombre de la receta
                   .Col = Col - 1
                   NomReceta = .text
                   .Col = .maxcols
                   
                   If Trim(NomReceta) <> "" Then
                      
                      .Col = IIf(TipoCopia = "Copia Raciones", Col + 1, Col)
                      ValPorDiario = Val(.text)
                      .Col = IIf(TipoCopia = "Copia Raciones", Col + 2, Col + 1)
                      
                      If ValPorDiario > 0 Then
                         
                         .text = Redondear(((ValPorDiario / 100) * NumRacTotal), 0)
                      
                      Else
                         
                         .Col = Col
                         If Mid(Trim(.text), 1, 3) <> "[*]" And Trim(.text) <> "" Then
                            
                            .Col = IIf(TipoCopia = "Copia Raciones", Col + 2, Col + 1) '.Col = Col + 2
                            .text = 0
                         
                         End If

'                         .text = 0
                      
                      End If
                   
                   End If
               
               Next i
         
         End If
    
    End With
    
End If

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub Calctodia(Row As Long, Col As Long)

On Error GoTo Man_Error

Dim X       As Long
Dim numrac  As Double
Dim cosdia  As Double
Dim dato    As Variant
    
    veccos((Int(Col / 7) + 1), 1) = 0
    veccos((Int(Col / 7) + 1), 2) = 0
    veccos((Int(Col / 7) + 1), 4) = 0
    
    Call vaSpread1.GetText(Col + 1, SpreadHeader + 3, dato)
    If "% Pond." = dato Then
        
        Col = (Col + 1)
    
    End If
        
    For X = 1 To (vaSpread1.MaxRows - 1)
        
        vaSpread1.Row = X
        vaSpread1.Col = Col + 1
        numrac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = Col + 2
        cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        vaSpread1.Col = Col + 3
        
        If Trim(vaSpread1.text) <> "" And numrac >= 0 Then
           
           vaSpread1.Col = Col - 1 '+ 2
           
           If Mid(Trim(vaSpread1.text), 1, 3) <> "[*]" Then
              
              veccos((Int(Col / 7) + 1), 1) = Round(veccos((Int(Col / 7) + 1), 1) + (cosdia * numrac), vg_DCa)
           
           Else
              
              veccos((Int(Col / 7) + 1), 2) = Round(veccos((Int(Col / 7) + 1), 2) + (cosdia * numrac), vg_DCa)
           
           End If
        
        End If
    
    Next X
    
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = Col + 1
    veccos((Int(Col / 7) + 1), 4) = Round(veccos((Int(Col / 7) + 1), 4) + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DCa)

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub MostrarCosto(Col As Long)
On Error GoTo Man_Error

Dim xcol    As Long
Dim toapla As Double
Dim toaesf As Double
Dim toafoo As Double
Dim totdia As Double
Dim totesf As Double
Dim nracre As Double
Dim nracfo As Double
Dim totrac As Double
Dim CostoDiaLyD As Double
Dim CostoAcumLyD As Double
Dim CostoTotLyD As Double

    vaSpread1.Col = Col
    xcol = 0
    For i = 1 To MaxColumna
        
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or _
           vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 3) Or vectorcol(i) = (vaSpread1.Col - 4) Or vectorcol(i) = (vaSpread1.Col - 5)) Then
           
           xcol = vectorcol(i)
           Exit For
        
        End If
    
    Next i
    vaSpread1.Row = 0
    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = xcol
    Frame2(2).Caption = vaSpread1.text
    Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
    
    toapla = 0
    toaesf = 0
    toafoo = 0
    totdia = 0
    totesf = 0
    nracre = 0
    nracfo = 0
    totrac = 0
    CostoDiaLyD = 0
    CostoAcumLyD = 0
    CostoTotLyD = 0
    
    For i = 1 To UBound(veccos)
        veccos(i, 6) = IIf(IsNull(veccos(i, 6)), 0, veccos(i, 6))
        
        If i <= (Int(xcol / 7) + 1) Then
           
           toapla = CCur(toapla + veccos(i, 1))
           toaesf = CCur(toaesf + veccos(i, 2))
           toafoo = CCur(toafoo + veccos(i, 3))
           nracre = CCur(nracre + veccos(i, 4))
           nracfo = CCur(nracfo + veccos(i, 5))
           CostoAcumLyD = CCur(CostoAcumLyD + IIf(IsNull(veccos(i, 6)), 0, veccos(i, 6)))
        
        End If
        
        totrac = CCur(totrac + veccos(i, 4))
        totdia = CCur(totdia + veccos(i, 1))
        totesf = CCur(totesf + veccos(i, 2))
        CostoTotLyD = CCur(CostoTotLyD + IIf(IsNull(veccos(i, 6)), 0, veccos(i, 6)))
    
    Next i
    
    Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
    Label1(11).Caption = Format(totesf + CostoTotLyD, fg_Pict(6, 2))
    Label1(12).Caption = Format(CCur(totdia + totesf + CostoTotLyD), fg_Pict(6, 2))
'    Label1(14).Caption = Format(CostoTotLyD, fg_Pict(6, 2))
    
    If totrac > 0 Then
       
       Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2))
    
    Else
       
       Label1(40).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    
    If totrac > 0 Then
       
       Label1(41).Caption = Format(CCur((totesf + CostoTotLyD) / totrac), fg_Pict(6, 2))
    
    Else
       
       Label1(41).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    
    If totrac > 0 Then
       
       Label1(8).Caption = Format(CCur((totdia + totesf + CostoTotLyD) / totrac), fg_Pict(6, 2))
    
    Else
       
       Label1(8).Caption = Format(0, fg_Pict(6, 2))
    
    End If
'    If totrac > 0 Then
'       Label1(10).Caption = Format(CCur((CostoTotLyD) / totrac), fg_Pict(6, 2))
'    Else
'       Label1(10).Caption = Format(0, fg_Pict(6, 2))
'    End If
    
    If totrac > 0 Then
       
       Label1(48).Caption = Format(totrac, fg_Pict(6, 2))
    
    Else
       
       Label1(48).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    Label1(20).Caption = Format(veccos((Int(xcol / 7) + 1), 1), fg_Pict(6, 2))
    Label1(21).Caption = Format(veccos((Int(xcol / 7) + 1), 2) + IIf(IsNull(veccos((Int(xcol / 7) + 1), 6)), 0, veccos((Int(xcol / 7) + 1), 6)), fg_Pict(6, 2))
    Label1(22).Caption = Format(CCur(veccos((Int(xcol / 7) + 1), 1) + (veccos((Int(xcol / 7) + 1), 2)) + IIf(IsNull((veccos((Int(xcol / 7) + 1), 6))), 0, (veccos((Int(xcol / 7) + 1), 6)))), fg_Pict(6, 2))
    Label1(23).Caption = Format(veccos((Int(xcol / 7) + 1), 3), fg_Pict(6, 2))
    Label1(44).Caption = Format(veccos((Int(xcol / 7) + 1), 4), fg_Pict(6, 2))
    
    If veccos((Int(xcol / 7) + 1), 4) > 0 Then
       
       Label1(45).Caption = Format(CCur((veccos((Int(xcol / 7) + 1), 1) + (veccos((Int(xcol / 7) + 1), 2)) + IIf(IsNull((veccos((Int(xcol / 7) + 1), 6))), 0, (veccos((Int(xcol / 7) + 1), 6)))) / veccos((Int(xcol / 7) + 1), 4)), fg_Pict(6, 2))
    
    Else
       
       Label1(45).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    
    Label1(46).Caption = Format(veccos((Int(xcol / 7) + 1), 5), fg_Pict(6, 2))
    
    If veccos((Int(xcol / 7) + 1), 5) > 0 Then
       
       Label1(47).Caption = Format(CCur(veccos((Int(xcol / 7) + 1), 3) / veccos((Int(xcol / 7) + 1), 5)), fg_Pict(6, 2))
    
    Else
        
        Label1(47).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    
    Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
    Label1(32).Caption = Format((toaesf + CostoAcumLyD), fg_Pict(6, 2))
    Label1(33).Caption = Format(CCur(toapla + (toaesf) + CostoAcumLyD), fg_Pict(6, 2))
    Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
    If nracre > 0 Then
       
       Label1(35).Caption = Format(CCur((toapla + toaesf + CostoAcumLyD) / nracre), fg_Pict(6, 2))
    
    Else
       
       Label1(35).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
    Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
    If nracfo > 0 Then
       
       Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2))
    
    Else
       
       Label1(38).Caption = Format(0, fg_Pict(6, 2))
    
    End If

'    Label1(51).Caption = Format(IIf(IsNull(veccos((Int(xcol / 7) + 1), 6)), 0, veccos((Int(xcol / 7) + 1), 6)), fg_Pict(6, 2))
'    Label1(53).Caption = Format(CostoAcumLyD, fg_Pict(6, 2))
Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub

Sub CargarCosto(OpLyd As Boolean)

On Error GoTo Man_Error

Dim cosdia  As Double
Dim totdia  As Double
Dim totesf  As Double
Dim totrac  As Double
Dim Sql1    As String
Dim Sql2    As String
Dim sql3    As String
Dim Fecha   As Long
Dim xcol    As Long
Dim IndDia  As Long
Dim fecesf  As Double
Dim nracre  As Long
Dim nracfo  As Double
Dim aAp     As String
Dim estfij  As Boolean
Dim numtor  As Long
Dim LyD As Boolean
Dim TotLyD As Double
Dim CostoDiaLyD As Double

    fg_carga ""
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Or vaSpread1.Col = 2 Then
       
       vaSpread1.Col = 3
    
    End If
    
    DoEvents
    Label1(7).Caption = Format(0, fg_Pict(6, 2))
    Label1(8).Caption = Format(0, fg_Pict(6, 2))
    Label1(9).Caption = Format(0, fg_Pict(6, 2))
'    Label1(10).Caption = Format(0, fg_Pict(6, 2))
    Label1(11).Caption = Format(0, fg_Pict(6, 2))
    Label1(12).Caption = Format(0, fg_Pict(6, 2))
    Label1(13).Caption = Format(0, fg_Pict(6, 2))
'    Label1(14).Caption = Format(0, fg_Pict(6, 2))
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
'    Label1(51).Caption = Format(0, fg_Pict(6, 2))
'    Label1(53).Caption = Format(0, fg_Pict(6, 2))
    
    j = 0: cosdia = 0: totdia = 0: totesf = 0: fecesf = 0: IndDia = 1: numrac = 0: totrac = 0: TotLyD = 0
    
    For i = 1 To MaxColumna
        
        If (vectorcol(i) = vaSpread1.Col Or vectorcol(i) = (vaSpread1.Col + 1) Or vectorcol(i) = (vaSpread1.Col - 1) Or _
        vectorcol(i) = (vaSpread1.Col - 2) Or vectorcol(i) = (vaSpread1.Col - 3) Or vectorcol(i) = (vaSpread1.Col - 4)) Then xcol = vectorcol(i): Exit For
    
    Next i
    
    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = xcol
    Frame2(2).Caption = vaSpread1.text
    Frame2(3).Caption = "Acumulado hasta " & vaSpread1.text
'    ReDim veccos(MaxColumna, 6)
    
    '-------> rutina calcular LyD
    If OpLyd Then
       
       Call Calcular_LyD
    
    End If
    
    estfij = False
    '-------> Calcular costo día planificado & estructura fija & salida
    For j = 1 To MaxColumna
        
        DoEvents
        veccos(j, 1) = 0 '-------> Costo Alimentación
        veccos(j, 2) = 0 '-------> recetas LYD
        veccos(j, 3) = 0
        veccos(j, 4) = 0
        veccos(j, 5) = 0
'        veccos(j, 6) = 0 '-------> Servicio LYD
    
    Next j
     
    Sql1 = " SUM(CASE WHEN a.tov_tipdoc = 'SP' THEN b.dev_ptotal ELSE (-1 * b.dev_ptotal) END) AS totdoc "
    Sql2 = " substring(convert(varchar(10), a.tov_fecpro,103),4,8) "
    sql3 = " substring(('" & fg_Ctod1(Val(Vg_FechaDesde) & Right("01", 2)) & "'),4,8) "
    RS.Open "SELECT a.tov_fecpro, a.tov_codreg, a.tov_codser, " & _
            "" & Sql1 & " " & _
            "FROM  cas_b_totventas a WITH ( NOLOCK ) , cas_b_detventas b WITH ( NOLOCK ), b_productos c WITH ( NOLOCK ) " & _
            "WHERE a.tov_cecori = b.dev_cecori AND a.tov_rutcli = b.dev_rutcli " & _
            "AND   a.tov_tipdoc = b.dev_tipdoc " & _
            "AND   a.tov_numdoc = b.dev_numdoc " & _
            "AND   b.dev_codmer = c.pro_codigo " & _
            "AND   c.pro_ctacon IN ('" & fg_CambiaChar(GetParametro("ctainsumo"), ";", "','") & "') " & _
            "AND   a.tov_codreg = " & vg_codregimen & " " & _
            "AND   a.tov_codser = " & vg_codservicio & " " & _
            "AND  (a.tov_tipdoc = 'SP' or a.tov_tipdoc = 'DP') " & _
            "AND   b.dev_canmer <> 0 " & _
            "AND   a.tov_estdoc <> 'A' AND a.tov_estdoc <> 'P' AND a.tov_cecori = '" & vg_codcasino & "' " & _
            "AND   " & Sql2 & " = " & sql3 & " " & _
            "GROUP BY a.tov_fecpro, a.tov_codreg, a.tov_codser", vg_db, adOpenStatic
    Do While Not RS.EOF
       
       DoEvents
       veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 1) = 0
       veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 2) = 0
       veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 4) = 0
       veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 5) = 0
       veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 3) = 0
       veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 3) = Round(veccos(Val(TraerPosicionDia(Format(RS!tov_fecpro, "yyyymmdd"))), 3) + RS!totdoc, vg_DCa)
       RS.MoveNext
    
    Loop
    RS.Close
    Set RS = Nothing
    
    Bar1(0).Min = 0
    Bar1(0).Value = 0
    Bar1(0).Max = MaxColumna
    Frame2(4).Visible = True
    Bar1(0).Visible = True
    
    For j = 3 To (vaSpread1.maxcols - 2) Step 7
        
        DoEvents
        Bar1(0).Value = Bar1(0).Value + 1
        vaSpread1.Row = SpreadHeader + 3
        vaSpread1.Col = j + 1
        Fecha = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")
        veccos(TraerPosicionDia(Fecha), 1) = 0
        veccos(TraerPosicionDia(Fecha), 2) = 0
        veccos(TraerPosicionDia(Fecha), 4) = 0
        veccos(TraerPosicionDia(Fecha), 5) = 0
        For i = 1 To (vaSpread1.MaxRows - 1)
            
            DoEvents
            vaSpread1.Row = i
            vaSpread1.Col = j + 1
            LyD = IIf(Mid(Trim(vaSpread1.text), 1, 3) = "[*]", True, False)
            '-------> mover raciones
            vaSpread1.Col = j + 3
            numrac = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
            '-------> mover costo día
            vaSpread1.Col = j + 4
            cosdia = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
            vaSpread1.Col = j + 5
            
            If Trim(vaSpread1.text) <> "" And numrac > 0 Then
               
               If LyD Then
                  
                  veccos(TraerPosicionDia(Fecha), 2) = Round(veccos(TraerPosicionDia(Fecha), 2) + (cosdia * numrac), vg_DCa)
                  totesf = Round(totesf + (cosdia * numrac), vg_DCa)
               
               Else
                  
                  veccos(TraerPosicionDia(Fecha), 1) = Round(veccos(TraerPosicionDia(Fecha), 1) + (cosdia * numrac), vg_DCa)
                  totdia = Round(totdia + (cosdia * numrac), vg_DCa)
               
               End If
            
            End If
        
        Next i
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = j + 3
        numtor = IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text)
        veccos(TraerPosicionDia(Fecha), 4) = Round(veccos(TraerPosicionDia(Fecha), 4) + numtor, vg_DPr)
        vaSpread1.Row = vaSpread1.MaxRows
        vaSpread1.Col = j + 3
        totrac = Round(totrac + IIf(Val(vaSpread1.text) = 0, 0, vaSpread1.text), vg_DPr)
        vaSpread1.Row = vaSpread1.MaxRows
        IndDia = IndDia + 1
    
    Next j
    Frame2(4).Visible = False
    Bar1(0).Visible = False
    
    '-------> Fin Calcular costo día
    toapla = 0
    toaesf = 0
    toafoo = 0
    numrac = 0
    nracfo = 0
    TotLyD = 0
    
    For i = 1 To (Int(xcol / 7) + 1)
        
        DoEvents
        toapla = Round(toapla + veccos(i, 1), vg_DCa)
        toaesf = Round(toaesf + veccos(i, 2), vg_DCa)
        toafoo = Round(toafoo + veccos(i, 3), vg_DCa)
        nracre = Round(nracre + veccos(i, 4), vg_DPr)
        nracfo = Round(nracfo + veccos(i, 5), vg_DPr)
        veccos(i, 6) = IIf(IsNull(veccos(i, 6)), 0, veccos(i, 6))
        TotLyD = Round(TotLyD + IIf(IsNull(veccos(i, 6)), 0, veccos(i, 6)), vg_DPr)
    
    Next i
    Label1(7).Caption = Format(totdia, fg_Pict(6, 2))
    Label1(11).Caption = Format(totesf + TotLyD, fg_Pict(6, 2))
    Label1(12).Caption = Format(CCur(totdia + totesf + TotLyD), fg_Pict(6, 2))
'    Label1(14).Caption = Format(TotLyD, fg_Pict(6, 2))
'    Label1(51).Caption = Format(veccos((Int(xcol / 7) + 1), 6), fg_Pict(6, 2))
    
    If totrac > 0 Then
       
       Label1(40).Caption = Format(CCur(totdia / totrac), fg_Pict(6, 2))
    
    Else
       
       Label1(40).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    
    If totrac > 0 Then
       
       Label1(41).Caption = Format(CCur((totesf + TotLyD) / totrac), fg_Pict(6, 2))
    
    Else
       
       Label1(41).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    
    If totrac > 0 Then
       
       Label1(8).Caption = Format(CCur((totdia + totesf + TotLyD) / totrac), fg_Pict(6, 2))
    
    Else
       
       Label1(8).Caption = Format(0, fg_Pict(6, 2))
    
    End If
'    If totrac > 0 Then
'       Label1(10).Caption = Format(CCur(TotLyD / totrac), fg_Pict(6, 2))
'    Else
'       Label1(10).Caption = Format(0, fg_Pict(6, 2))
'    End If
    If totrac > 0 Then
       
       Label1(48).Caption = Format(totrac, fg_Pict(6, 2))
    
    Else
       
       Label1(48).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    Label1(20).Caption = Format(veccos((Int(xcol / 7) + 1), 1), fg_Pict(6, 2))
    Label1(21).Caption = Format((veccos((Int(xcol / 7) + 1), 2)) + (veccos((Int(xcol / 7) + 1), 6)), fg_Pict(6, 2))
    Label1(22).Caption = Format(CCur(veccos((Int(xcol / 7) + 1), 1) + (veccos((Int(xcol / 7) + 1), 2)) + IIf(IsNull(veccos((Int(xcol / 7) + 1), 6)), 0, veccos((Int(xcol / 7) + 1), 6))), fg_Pict(6, 2))
    Label1(23).Caption = Format(veccos((Int(xcol / 7) + 1), 3), fg_Pict(6, 2))
    Label1(44).Caption = Format(nracre, fg_Pict(6, 2))
    If nracre > 0 Then
       
       Label1(45).Caption = Format(CCur((veccos((Int(xcol / 7) + 1), 1) + (veccos((Int(xcol / 7) + 1), 2)) + IIf(IsNull(veccos((Int(xcol / 7) + 1), 6)), 0, veccos((Int(xcol / 7) + 1), 6))) / nracre), fg_Pict(6, 2))
    
    Else
       
       Label1(45).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    Label1(46).Caption = Format(nracfo, fg_Pict(6, 2))
    If nracfo > 0 Then
       
       Label1(47).Caption = Format(CCur(veccos((Int(xcol / 7) + 1), 3) / nracfo), fg_Pict(6, 2))
    
    Else
       
       Label1(47).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    Label1(31).Caption = Format(toapla, fg_Pict(6, 2))
    Label1(32).Caption = Format((toaesf), fg_Pict(6, 2))
    Label1(33).Caption = Format(CCur(toapla + (toaesf) + TotLyD), fg_Pict(6, 2))
    Label1(34).Caption = Format(nracre, fg_Pict(6, 2))
    If nracre > 0 Then
       
       Label1(35).Caption = Format(CCur((toapla + toaesf + TotLyD) / nracre), fg_Pict(6, 2))
    
    Else
       
       Label1(35).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    Label1(36).Caption = Format(toafoo, fg_Pict(6, 2))
    Label1(37).Caption = Format(nracfo, fg_Pict(6, 2))
    If nracfo > 0 Then
       
       Label1(38).Caption = Format(CCur(toafoo / nracfo), fg_Pict(6, 2))
    
    Else
       
       Label1(38).Caption = Format(0, fg_Pict(6, 2))
    
    End If
    indcos = True
    fg_descarga

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
End Sub

Function TraerPosicionDia(Fecha As Long) As Long

On Error GoTo Man_Error

Dim i As Long
Dim j As Long

TraerPosicionDia = 0
For i = 2 To (vaSpread1.maxcols - 2) Step 7
    
    DoEvents
    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = i + 2
    
    If Trim(vaSpread1.text) <> "" Then
       
       If Fecha = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd") Then
           TraerPosicionDia = CLng((i / 7) + 1)
           Exit Function
       
       End If
    
    End If

Next i

Exit Function
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Function

Sub ExportarExcelMenuI()

On Error GoTo Ex_Error

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim RowSheet As Long
Dim oCol As String
Dim AoCol As String
Dim oColA As String
Dim MaxCol  As Long
Dim NumAsc As Long
Dim NombreServicio As String
Dim NombreServicioAux As String
Dim CodigoServicio As Long
Dim i As Long
Dim j As Long
Dim IndCol As Long
Dim IndColA As Long

fg_carga ""

'-------> Mover columnas
MaxCol = (vaSpread1.maxcols - 2) '- DateDiff("d", CDate(fg_Ctod1(Vg_FechaDesde)), CDate(fg_Ctod1(Vg_FechaHasta))) + 1

Dim VecDiaExcel() As Variant
oCol = ""
AoCol = ""
oColA = ""
IndCol = 1
IndColA = 65
oCol = Chr(IndCol + 64)
ReDim VecDiaExcel(MaxCol, 2)

For i = 1 To MaxCol
    
    '-------> Setear vector
    VecDiaExcel(i, 1) = Val(0) 'fecha
    VecDiaExcel(i, 2) = "" 'columna letra
    
    VecDiaExcel(i, 1) = 0 'Format(FecMin, "yyyymmdd") 'fecha
    VecDiaExcel(i, 2) = oCol 'columna letra
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1

Next i

'-------> Rutinas exportar excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

NumAsc = 66

CodigoServicio = M_MinSR1.fpLongInteger1(1).Value
NombreServicio = Trim(M_MinSR1.fpayuda(2).Caption)

Set oSheet = oBook.Worksheets.Add
NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
   
   NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)

End If
oSheet.Name = NombreServicioAux

'-------> Mover Ceco - Regimen
MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, "Contrato : " & M_MinSR1.fpText.text & " - " & Trim(M_MinSR1.fpayuda(0).Caption)
MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Regimen  : " & M_MinSR1.fpLongInteger1(0).Value & " - " & Trim(M_MinSR1.fpayuda(1).Caption)
MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, "Servicio : " & M_MinSR1.fpLongInteger1(1).Value & " - " & Trim(M_MinSR1.fpayuda(2).Caption)
MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, "Fecha    : " & M_MinSR1.FpFecDesde & " - " & M_MinSR1.FpFecHasta

'-------> Encabezado
For i = 1 To MaxCol
    
    vaSpread1.Row = SpreadHeader
    vaSpread1.Col = i
    
    oCol = VecDiaExcel(i, 2)
    DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
    
    If IsNumeric(vaSpread1.text) Then
       
       oCol = VecDiaExcel(i + 6, 2)
    
    End If
    
    
    If Not ValidarExisteDato(oExcel, oSheet, oCol, oCol, 5, 5, "") Then
       
       MoverDatosExcel oExcel, oSheet, oCol, oCol, 5, 5, vaSpread1.text
    
    End If
    DibujarLineas oExcel, oSheet, oCol, oCol, 5, 5
    
    vaSpread1.Row = SpreadHeader + 1
    vaSpread1.Col = i
    
    oCol = VecDiaExcel(i, 2)
    DibujarLineas oExcel, oSheet, oCol, oCol, 6, 6
    
    If IsNumeric(vaSpread1.text) Then
       
       oCol = VecDiaExcel(i + 6, 2)
    
    End If
    
    If Not ValidarExisteDato(oExcel, oSheet, oCol, oCol, 6, 6, "") Then
       
       MoverDatosExcel oExcel, oSheet, oCol, oCol, 6, 6, vaSpread1.text
    
    End If
    DibujarLineas oExcel, oSheet, oCol, oCol, 6, 6

    vaSpread1.Row = SpreadHeader + 2
    vaSpread1.Col = i
    oCol = VecDiaExcel(i, 2)
    DibujarLineas oExcel, oSheet, oCol, oCol, 7, 7
    
    If IsNumeric(vaSpread1.text) Then
       
       oCol = VecDiaExcel(i + 6, 2)
    
    End If
    
    DibujarLineas oExcel, oSheet, oCol, oCol, 7, 7
    If Not ValidarExisteDato(oExcel, oSheet, oCol, oCol, 7, 7, "") Then
       
       MoverDatosExcel oExcel, oSheet, oCol, oCol, 7, 7, vaSpread1.text
    
    End If

    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = i
    
    oCol = VecDiaExcel(i, 2)
    MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, vaSpread1.text
    DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
    
Next i

'-------> Detalle minuta bloque
RowSheet = 8
For i = 1 To MaxCol
    
    vaSpread1.Col = i
    oCol = VecDiaExcel(i, 2)
    
    For j = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = j
        MoverDatosExcel oExcel, oSheet, oCol, oCol, j + RowSheet, j + RowSheet, vaSpread1.text
    
    Next j
    
    DibujarLineas oExcel, oSheet, oCol, oCol, 1 + RowSheet, 8 + vaSpread1.MaxRows

Next i

'-------> Borrar columna código receta
'oCol = ""
'For i = 3 To MaxCol Step 7
'    oCol = oCol & VecDiaExcel(i + 5, 2) & ":" & VecDiaExcel(i + 5, 2) & ","
'Next i
''oExcel.Range("H:H,O:O").Select
'oExcel.Range(Mid(oCol, 1, Len(oCol) - 1)).Select
''Range("O1").Activate
'oExcel.Selection.Delete Shift:=xlToLeft

Dim aa As Variant

If Not IsEmpty(aa) Then
   
   oExcel.Cells.Replace What:="&0&;", Replacement:="", LookAt:=xlPart, SearchOrder _
         :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End If
If Not IsEmpty(aa) Then
   
   oExcel.Cells.Replace What:="&-1&;", Replacement:="", LookAt:=xlPart, SearchOrder _
         :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End If

aa = oExcel.Cells.Find(What:="&" & vg_codregimen & "&;", LookAt:=xlPart, SearchOrder _
      :=xlByRows, MatchCase:=False, SearchFormat:=False)

If Not IsEmpty(aa) Then
   
   oExcel.Cells.Replace What:="&" & vg_codregimen & "&;", Replacement:="", LookAt:=xlPart, SearchOrder _
         :=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False

End If

oExcel.Visible = True '------->Visualizar
Set oSheet = Nothing
Set oExcel = Nothing
Set oBook = Nothing

fg_descarga
'NashXl.Visible = True

Ex_Error:
    Resume Next

End Sub

Sub ExportarExcelMenuII()

On Error GoTo Ex_Error

Dim oExcel As Object
Dim oBook As Object
Dim oSheet As Object
Dim RowSheet As Long
Dim oCol As String
Dim AoCol As String
Dim oColA As String
Dim MaxCol  As Long
Dim NumAsc As Long
Dim NombreServicio As String
Dim NombreServicioAux As String
Dim CodigoServicio As Long
Dim i As Long
Dim j As Long
Dim IndCol As Long
Dim IndColA As Long

fg_carga ""

'-------> Mover columnas
MaxCol = (vaSpread1.maxcols - 2) '- DateDiff("d", CDate(fg_Ctod1(Vg_FechaDesde)), CDate(fg_Ctod1(Vg_FechaHasta))) + 1

Dim VecDiaExcel() As Variant
oCol = ""
AoCol = ""
oColA = ""
IndCol = 1
IndColA = 65
oCol = Chr(IndCol + 64)
ReDim VecDiaExcel(MaxCol, 2)

For i = 1 To MaxCol
    
    '-------> Setear vector
    VecDiaExcel(i, 1) = Val(0) 'fecha
    VecDiaExcel(i, 2) = "" 'columna letra
    
    VecDiaExcel(i, 1) = 0 'Format(FecMin, "yyyymmdd") 'fecha
    VecDiaExcel(i, 2) = oCol 'columna letra
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1

Next i

'-------> Rutinas exportar excel
Set oExcel = CreateObject("Excel.Application")
Set oBook = oExcel.Workbooks.Add

NumAsc = 66

CodigoServicio = M_MinSR1.fpLongInteger1(1).Value
NombreServicio = Trim(M_MinSR1.fpayuda(2).Caption)

Set oSheet = oBook.Worksheets.Add
NombreServicioAux = Mid(Trim(LimpiaDatoExcel(NombreServicio)), 1, 31)
If ValidarNombreHoja(oExcel, oSheet, NombreServicioAux) Then
   
   NombreServicioAux = Mid(CodigoServicio & NombreServicioAux, 1, 31)

End If
oSheet.Name = NombreServicioAux

'-------> Mover Ceco - Regimen
MoverDatosExcel oExcel, oSheet, "A", "A", 1, 1, "Contrato : " & M_MinSR1.fpText.text & " - " & Trim(M_MinSR1.fpayuda(0).Caption)
MoverDatosExcel oExcel, oSheet, "A", "A", 2, 2, "Regimen  : " & M_MinSR1.fpLongInteger1(0).Value & " - " & Trim(M_MinSR1.fpayuda(1).Caption)
MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, "Servicio : " & M_MinSR1.fpLongInteger1(1).Value & " - " & Trim(M_MinSR1.fpayuda(2).Caption)
MoverDatosExcel oExcel, oSheet, "A", "A", 3, 3, "Fecha    : " & M_MinSR1.FpFecDesde & " - " & M_MinSR1.FpFecHasta

'-------> Encabezado
oCol = ""
AoCol = ""
oColA = ""
IndCol = 1
IndColA = 65
oCol = Chr(IndCol + 64)

For i = 3 To MaxCol Step 7

    If i = 3 Then
       
       vaSpread1.Row = SpreadHeader + 3
       vaSpread1.Col = i - 2
    
       MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, vaSpread1.text
       DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
       
       If Chr(IndCol + 65) = "[" Then
          
          oColA = Chr(IndColA)
          IndColA = IndColA + 1
          IndCol = 0
       
       End If
       
       oCol = oColA & Chr(IndCol + 65)
       IndCol = IndCol + 1
    
    End If
    
    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = i
    
    MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, vaSpread1.text
    DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1
    
    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = i + 1
    
    MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, vaSpread1.text
    DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1
    
    vaSpread1.Row = SpreadHeader + 3
    vaSpread1.Col = i + 3
    
    MoverDatosExcel oExcel, oSheet, oCol, oCol, 8, 8, vaSpread1.text
    DibujarLineas oExcel, oSheet, oCol, oCol, 8, 8
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1
    
Next i

'-------> Detalle minuta bloque
RowSheet = 8
For j = 1 To vaSpread1.MaxRows
    
    vaSpread1.Row = j
    oCol = ""
    AoCol = ""
    oColA = ""
    IndCol = 1
    IndColA = 65
    oCol = Chr(IndCol + 64)

    For i = 3 To MaxCol Step 7
        
        vaSpread1.Col = i
        
        If i = 3 Then
           
           vaSpread1.Col = i - 2
           MoverDatosExcel oExcel, oSheet, "A", "A", j + RowSheet, j + RowSheet, vaSpread1.text

           If Chr(IndCol + 65) = "[" Then
              
              oColA = Chr(IndColA)
              IndColA = IndColA + 1
              IndCol = 0
           
           End If
           
           oCol = oColA & Chr(IndCol + 65)
           IndCol = IndCol + 1
        
        End If
           
           vaSpread1.Col = i
           MoverDatosExcel oExcel, oSheet, oCol, oCol, j + RowSheet, j + RowSheet, " "
           If Chr(IndCol + 65) = "[" Then
              
              oColA = Chr(IndColA)
              IndColA = IndColA + 1
              IndCol = 0
           
           End If
           oCol = oColA & Chr(IndCol + 65)
           IndCol = IndCol + 1

           vaSpread1.Col = i + 1
           MoverDatosExcel oExcel, oSheet, oCol, oCol, j + RowSheet, j + RowSheet, vaSpread1.text
           
           If Chr(IndCol + 65) = "[" Then
              
              oColA = Chr(IndColA)
              IndColA = IndColA + 1
              IndCol = 0
           
           End If
           oCol = oColA & Chr(IndCol + 65)
           IndCol = IndCol + 1

           vaSpread1.Col = i + 3
           MoverDatosExcel oExcel, oSheet, oCol, oCol, j + RowSheet, j + RowSheet, vaSpread1.text
           If Chr(IndCol + 65) = "[" Then
              
              oColA = Chr(IndColA)
              IndColA = IndColA + 1
              IndCol = 0
           
           End If
           oCol = oColA & Chr(IndCol + 65)
           IndCol = IndCol + 1

   Next i

Next j

'-------> Dibujar columnas
oCol = ""
AoCol = ""
oColA = ""
IndCol = 1
IndColA = 65
oCol = Chr(IndCol + 64)
RowSheet = 8
For i = 3 To MaxCol Step 7

    If i = 3 Then
       
       DibujarLineas oExcel, oSheet, oCol, oCol, RowSheet, RowSheet + vaSpread1.MaxRows
       
       If Chr(IndCol + 65) = "[" Then
          
          oColA = Chr(IndColA)
          IndColA = IndColA + 1
          IndCol = 0
       
       End If
       
       oCol = oColA & Chr(IndCol + 65)
       IndCol = IndCol + 1
    
    End If
    
    DibujarLineas oExcel, oSheet, oCol, oCol, RowSheet, RowSheet + vaSpread1.MaxRows
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1
    
    DibujarLineas oExcel, oSheet, oCol, oCol, RowSheet, RowSheet + vaSpread1.MaxRows
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1
    
    DibujarLineas oExcel, oSheet, oCol, oCol, RowSheet, RowSheet + vaSpread1.MaxRows
    
    If Chr(IndCol + 65) = "[" Then
       
       oColA = Chr(IndColA)
       IndColA = IndColA + 1
       IndCol = 0
    
    End If
    
    oCol = oColA & Chr(IndCol + 65)
    IndCol = IndCol + 1
    
Next i

'    DibujarLineas oExcel, oSheet, oCol, oCol, 1 + RowSheet, 8 + vaSpread1.MaxRows

oExcel.Visible = True '------->Visualizar
Set oSheet = Nothing
Set oExcel = Nothing
Set oBook = Nothing
fg_descarga

Ex_Error:
    Resume Next

End Sub

Sub Deshacer(StrRec As Variant)

'load in file
Dim ret As Integer
Screen.MousePointer = 11
ret = vaSpread1.LoadFromFile(LCase(App.Path) & "\" & StrRec)
If Dir(LCase(App.Path) & "\" & StrRec) <> "" Then Kill LCase(App.Path) & "\" & StrRec
ContadorDeshacer = ContadorDeshacer - 1
Screen.MousePointer = 0

End Sub

Sub GrabarCambios(ifil As Long, icol As Long, estado As String)

Dim ret
ContadorDeshacer = ContadorDeshacer + 1
ret = vaSpread1.SaveToFile(LCase(App.Path) & "\" & "spreadMBloque" & vg_NUsr & ContadorDeshacer & ".ss6", False)
Toolbar1.Buttons(34).Visible = True
Toolbar1.Buttons(34).Enabled = True

End Sub

Function ValidarBloqueoMinuta() As Boolean

On Error GoTo Man_Error

ValidarBloqueoMinuta = False
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 2 'vaSpread1.ActiveCol '3

If vaSpread1.BackColor = Shape1(1).FillColor Then
   
   ValidarBloqueoMinuta = True
   Call MsgBox("Minuta esta Bloqueada", vbCritical + vbOKOnly, MsgTitulo)
 
End If


Exit Function
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"

End Function

Function ValidarBloqueoMinutaDetalle(ByVal Row As Long, ByVal Col As Long) As Boolean

ValidarBloqueoMinutaDetalle = False
vaSpread1.Row = Row
vaSpread1.Col = Col

If vaSpread1.BackColor = Shape1(1).FillColor Then
   
   ValidarBloqueoMinutaDetalle = True
   Call MsgBox("Día minuta esta Bloqueada", vbCritical + vbOKOnly, MsgTitulo)
 
 End If

End Function

Sub MoverPorcentaje(Porcentaje, CodGrupoEstructura)

Dim i As Long
Dim CodGrupoEstructuraAux As Long

For i = iblockrow2 + 1 To vaSpread1.MaxRows - 1
    
    vaSpread1.Row = i
    vaSpread1.Col = vaSpread1.maxcols
    CodGrupoEstructuraAux = IIf(Trim(vaSpread1.text) = "", -1, vaSpread1.text)
    
    If CodGrupoEstructura = CodGrupoEstructuraAux Then
       
       vaSpread1.Col = 2
       vaSpread1.CellType = CellTypePercent
       vaSpread1.TypeHAlign = 1
       vaSpread1.TypePercentDecPlaces = 0
       vaSpread1.ForeColor = &HFF0000
       vaSpread1.text = Porcentaje
       vaSpread1.TypeNegRed = True
       Exit Sub
    
    End If

Next i

End Sub

Function PrimeraColumna(xcolini1 As Variant, maxcols As Long) As Integer

Dim i As Long
PrimeraColumna = 0

For i = 1 To maxcols
    
    If (vectorcol(i) - 2) = xcolini1 Or vectorcol(i) = xcolini1 Then PrimeraColumna = (vectorcol(i) - 2): Exit For

Next i

End Function

Function FinalColumna(xcolfin1 As Variant, maxcols As Long) As Integer

Dim i As Long
FinalColumna = 0

For i = 1 To maxcols
    
    If (vectorcol(i) - 2) = xcolfin1 Then FinalColumna = ((vectorcol(i) + 3)): Exit For
    If vectorcol(i) = xcolfin1 Then FinalColumna = (vectorcol(i) + 3): Exit For

Next i

End Function

Sub ResivarFilasTengaAsigEstGrupo()

On Error GoTo Man_Error

Dim i As Long
Dim j As Long
Dim codest As Long
Dim CodGrupoEst As Long
Dim Estreceta As Boolean

For i = 1 To vaSpread1.MaxRows - 1
    
    vaSpread1.Row = i
    vaSpread1.Col = vaSpread1.maxcols - 1
    
    If Trim(vaSpread1.text) <> "" Then
       
       codest = Val(vaSpread1.text)
    
    End If
    
    vaSpread1.Col = vaSpread1.maxcols
    
    If Trim(vaSpread1.text) <> "" Then
       
       CodGrupoEst = Val(vaSpread1.text)
    
    End If
    
    Estreceta = False
    For j = 3 To vaSpread1.maxcols - 2 Step 7
        
        vaSpread1.Col = j + 1
        
        If Trim(vaSpread1.text) <> "" Then
           
           '-------> Mover codigo estructura
           vaSpread1.Col = vaSpread1.maxcols - 1
           vaSpread1.text = codest
           '------> Mover codigo grupo estructura
           vaSpread1.Col = vaSpread1.maxcols
           vaSpread1.text = CodGrupoEst
           Estreceta = True
           Exit For
        
        End If
    
    Next j
    
    vaSpread1.Col = 1
    
    If Trim(vaSpread1.text) = "" And Not Estreceta Then
       '-------> Mover codigo estructura en blanco
       vaSpread1.Col = vaSpread1.maxcols - 1
       vaSpread1.text = ""
       '------> Mover codigo grupo estructura en blanco
       vaSpread1.Col = vaSpread1.maxcols
       vaSpread1.text = ""
    End If

Next i

Exit Sub
Man_Error:
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub Calcular_LyD()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim SeleccionOpt As Integer
Dim i As Long
Dim j As Long
Dim Fecha As Long
Dim FecMinAux As Long
Dim CostoAlim As Double
Dim CostoLyD As Double
Dim CostoRecetaLyD As Double
Dim CostoServicioAlim As Double
Dim CostoServicioLyD As Double
Dim PorAlim As Double

fg_carga ""

SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))

RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS = vg_db.Execute("sgpadm_Sel_CostoAlimLydMinutaBloque_V02 '" & vg_codcasino & "',  " & Vg_FechaDesde & ", " & Vg_FechaHasta & ", " & vg_codservicio & ", " & SeleccionOpt & "")

If Not RS.EOF Then

    Bar1(0).Min = 0
    Bar1(0).Value = 0
    Bar1(0).Max = RS.RecordCount
    Frame2(4).Visible = True
    Bar1(0).Visible = True
   
   Do While Not RS.EOF

      DoEvents
      Bar1(0).Value = Bar1(0).Value + 1
      If FecMinAux <> RS!min_fecmin Then
      
         If FecMinAux > 0 Then
            
            For i = 3 To (vaSpread1.maxcols - 2) Step 7
                
                vaSpread1.Row = SpreadHeader + 3
                vaSpread1.Col = i + 1
                
                If CStr(FecMinAux) = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd") Then
                   
                   Fecha = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")
                   '-------> Mover costo alimentación
                   CostoServicioAlim = veccos(TraerPosicionDia(Fecha), 1)
                   '-------> Mover costo receta LYD
                   CostoServicioLyD = veccos(TraerPosicionDia(Fecha), 2)
                   '-------> Sumar costo alimentación de los otros servicios
                   CostoAlim = (CostoAlim + CostoServicioAlim)
                   PorAlim = 0
                   '-------> Sacar % alimentación
                   If CostoAlim > 0 Then
                      
                      PorAlim = Round((CostoServicioAlim / CostoAlim) * 100, 2)
                   
                   End If
                   '-------> Sacar % LYD
'                   CostoLyD = 0
                   
                   If PorAlim > 0 Then
                      
                      CostoLyD = (CostoLyD * PorAlim) / 100
                   
                   End If
                   '-------> Mover costo LYD vector
                   veccos(TraerPosicionDia(Fecha), 6) = 0
                   veccos(TraerPosicionDia(Fecha), 6) = IIf(CostoServicioAlim = 0, 0, CostoLyD)
                   Exit For
                
                End If
            
            Next i
         
         End If
         
         FecMinAux = RS!min_fecmin
         
         '-------> Mover cero las variable
         CostoAlim = 0
         CostoLyD = 0
         CostoRecetaLyD = 0
         
      End If
      
      If Not RS!rec_LYD And Not RS!Ser_LYD Then ' Costo alimentación
         
         CostoAlim = CostoAlim + RS!Costo
      
      ElseIf RS!rec_LYD And RS!Ser_LYD Then   ' Costo limpieza y desechable
         
         CostoLyD = CostoLyD + RS!Costo
      
      End If

      RS.MoveNext
   
   Loop
   
   For i = 3 To (vaSpread1.maxcols - 2) Step 7
       
       vaSpread1.Row = SpreadHeader + 3
       vaSpread1.Col = i + 1
       
       If CStr(FecMinAux) = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd") Then
          
          Fecha = Format(CDate(Mid(Trim(vaSpread1.text), 5, Len(Trim(vaSpread1.text)))), "yyyymmdd")
          '-------> Mover costo alimentación
          CostoServicioAlim = veccos(TraerPosicionDia(Fecha), 1)
          '-------> Mover costo receta LYD
          CostoServicioLyD = veccos(TraerPosicionDia(Fecha), 2)
          '-------> Sumar costo alimentación de los otros servicios
          CostoAlim = (CostoAlim + CostoServicioAlim)
          '-------> Sacar % alimentación
          PorAlim = 0
          
          If CostoAlim > 0 Then
             
             PorAlim = Round((CostoServicioAlim / CostoAlim) * 100, 2)
          
          End If
          
          '-------> Sacar % LYD
          If PorAlim > 0 Then
             
             CostoLyD = (CostoLyD * PorAlim) / 100
          
          End If
          '-------> Mover costo LYD vector
          veccos(TraerPosicionDia(Fecha), 6) = 0
          veccos(TraerPosicionDia(Fecha), 6) = IIf(CostoServicioAlim = 0, 0, CostoLyD) 'CostoLyD
          Exit For
       
       End If
   
   Next i
   
End If
RS.Close
Set RS = Nothing
    
Frame2(4).Visible = False
Bar1(0).Visible = False
Call fg_descarga

Exit Sub
Man_Error:
    Frame2(4).Visible = False
    Bar1(0).Visible = False
    Call fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

