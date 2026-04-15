VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form B_RecMBi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Receta Minuta Bloque"
   ClientHeight    =   7005
   ClientLeft      =   4215
   ClientTop       =   2310
   ClientWidth     =   10605
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   10605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   480
      TabIndex        =   0
      Top             =   0
      Width           =   9195
      Begin VB.TextBox FptNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   315
         Left            =   2370
         LinkTimeout     =   0
         MaxLength       =   80
         TabIndex        =   1
         Top             =   915
         Width           =   3195
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   2370
         TabIndex        =   7
         Top             =   570
         Width           =   5685
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   1905
         Picture         =   "B_RecMbl.frx":0000
         Top             =   480
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   1905
         Picture         =   "B_RecMbl.frx":030A
         Top             =   135
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
         Left            =   915
         TabIndex        =   6
         Top             =   645
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
         Left            =   930
         TabIndex        =   5
         Top             =   300
         Width           =   1020
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
         Left            =   915
         TabIndex        =   4
         Top             =   1035
         Width           =   1140
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registro 0"
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
         Left            =   5580
         TabIndex        =   3
         Top             =   1035
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2370
         TabIndex        =   2
         Top             =   240
         Width           =   5685
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   6
         Left            =   2415
         TabIndex        =   8
         Top             =   285
         Width           =   5685
      End
      Begin VB.Label lblSOMBRA 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   2415
         TabIndex        =   9
         Top             =   615
         Width           =   5685
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   5220
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   9705
      _Version        =   393216
      _ExtentX        =   17119
      _ExtentY        =   9208
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
      MaxCols         =   7
      MaxRows         =   20
      OperationMode   =   2
      RestrictRows    =   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_RecMbl.frx":0614
      VisibleCols     =   3
      VisibleRows     =   20
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   7005
      Left            =   9975
      TabIndex        =   11
      Top             =   0
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   12356
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
End
Attribute VB_Name = "B_RecMBi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
