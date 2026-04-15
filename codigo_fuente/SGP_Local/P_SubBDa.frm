VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form P_SubBDa 
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   8145
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2970
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   8250
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de productos"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de recetas"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Top             =   645
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   1530
         Width           =   7215
      End
      Begin VB.TextBox Text1 
         Height          =   585
         Index           =   1
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1515
         Visible         =   0   'False
         Width           =   7665
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de planificación"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   3300
      End
      Begin ACTIVEZIPLib.ActiveZip AZ 
         Left            =   7800
         Top             =   2280
         _Version        =   65536
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   2505
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de Origen"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   1335
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   7515
         Picture         =   "P_SubBDa.frx":0000
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Archivo en Proceso"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   2190
         Visible         =   0   'False
         Width           =   7650
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   495
      Top             =   2790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8145
      _ExtentX        =   14367
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "P_SubBDa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
