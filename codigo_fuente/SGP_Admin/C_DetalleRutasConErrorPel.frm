VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form C_DetalleRutasConErrorPel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Rutas con Error PEL"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14805
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar y Aceptar"
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
      Left            =   13320
      TabIndex        =   14
      Top             =   7680
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14535
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   5
         Left            =   11400
         TabIndex        =   12
         Top             =   6600
         Width           =   1260
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   13
            Top             =   135
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   4
         Left            =   10080
         TabIndex        =   10
         Top             =   6600
         Width           =   1260
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   11
            Top             =   135
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   7080
         TabIndex        =   8
         Top             =   6600
         Width           =   2940
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   9
            Top             =   135
            Width           =   2835
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   5640
         TabIndex        =   6
         Top             =   6600
         Width           =   1380
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   7
            Top             =   135
            Width           =   1275
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   2400
         TabIndex        =   4
         Top             =   6600
         Width           =   3180
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   0
            Left            =   45
            TabIndex        =   5
            Top             =   135
            Width           =   3075
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   6600
         Width           =   780
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   1
            Left            =   45
            TabIndex        =   3
            Top             =   135
            Width           =   675
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   14175
         _Version        =   393216
         _ExtentX        =   25003
         _ExtentY        =   11033
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
         MaxCols         =   9
         SpreadDesigner  =   "C_DetalleRutasConErrorPel.frx":0000
      End
   End
End
Attribute VB_Name = "C_DetalleRutasConErrorPel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
