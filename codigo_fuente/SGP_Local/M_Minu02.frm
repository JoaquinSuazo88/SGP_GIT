VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{48932A52-981F-101B-A7FB-4A79242FD97B}#3.1#0"; "Tab32x30.ocx"
Begin VB.Form M_Minu02 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planificación Minutas"
   ClientHeight    =   6105
   ClientLeft      =   465
   ClientTop       =   1440
   ClientWidth     =   10965
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6105
   ScaleWidth      =   10965
   Begin TabproLib.vaTabPro vaTabPro1 
      Height          =   5145
      Left            =   0
      TabIndex        =   4
      Top             =   960
      Width           =   10935
      _Version        =   196609
      _ExtentX        =   19288
      _ExtentY        =   9075
      _StockProps     =   100
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabHeight       =   250
      TabsPerRow      =   0
      Tab             =   3
      AlignTextH      =   1
      AlignTextV      =   1
      Orientation     =   2
      ThreeD          =   -1  'True
      ShowFocusRect   =   0   'False
      TabShape        =   1
      MarginLeft      =   150
      MarginRight     =   150
      ApplyTo         =   2
      ActiveTabBold   =   -1  'True
      TabSeparator    =   -12
      OffsetFromClientTop=   -1  'True
      ShowEarMark     =   -1  'True
      BookShowMetalSpine=   -1  'True
      BookRingShowHole=   -1  'True
      DataFormat      =   ""
      BookCornerGuardWidth=   105
      BookCornerGuardLength=   390
      DrawFocusRect   =   1
      DataField       =   ""
      TabCaption      =   "M_Minu02.frx":0000
      PageEarMarkPictureNext=   "M_Minu02.frx":0329
      PageEarMarkPicturePrev=   "M_Minu02.frx":0345
      EarMarkPictureNext=   "M_Minu02.frx":0361
      EarMarkPicturePrev=   "M_Minu02.frx":037D
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Enabled         =   0   'False
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
         Left            =   -24615
         ScaleHeight     =   1035
         ScaleWidth      =   6675
         TabIndex        =   8
         Top             =   -17655
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
            Appearance      =   1
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
            Appearance      =   1
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
            TabIndex        =   12
            Top             =   0
            Width           =   195
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
            TabIndex        =   11
            Top             =   480
            Width           =   405
         End
      End
      Begin FPSpread.vaSpread vaSpread6 
         Height          =   135
         Left            =   -25095
         TabIndex        =   5
         Top             =   -20175
         Visible         =   0   'False
         Width           =   375
         _Version        =   393216
         _ExtentX        =   661
         _ExtentY        =   238
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         MaxRows         =   100
         SpreadDesigner  =   "M_Minu02.frx":0399
      End
      Begin FPSpread.vaSpread vaSpread4 
         Height          =   255
         Left            =   -22695
         TabIndex        =   6
         Top             =   -20175
         Visible         =   0   'False
         Width           =   735
         _Version        =   393216
         _ExtentX        =   1296
         _ExtentY        =   450
         _StockProps     =   64
         Enabled         =   0   'False
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
         MaxRows         =   100
         SpreadDesigner  =   "M_Minu02.frx":05CA
      End
      Begin FPSpread.vaSpread vaSpread3 
         Height          =   4860
         Left            =   -25935
         TabIndex        =   7
         Top             =   -19860
         Width           =   10935
         _Version        =   393216
         _ExtentX        =   19288
         _ExtentY        =   8573
         _StockProps     =   64
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   8
         MaxRows         =   100
         SpreadDesigner  =   "M_Minu02.frx":07FB
      End
      Begin FPSpread.vaSpread vaSpread1 
         DragIcon        =   "M_Minu02.frx":0FC4
         Height          =   4860
         Left            =   -25935
         TabIndex        =   13
         Top             =   -19860
         Width           =   10935
         _Version        =   393216
         _ExtentX        =   19288
         _ExtentY        =   8573
         _StockProps     =   64
         Enabled         =   0   'False
         ColsFrozen      =   1
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   249
         MaxRows         =   100
         RestrictRows    =   -1  'True
         SpreadDesigner  =   "M_Minu02.frx":1406
         UserResize      =   0
         VisibleCols     =   1
         VisibleRows     =   100
      End
      Begin FPSpread.vaSpread vaSpread5 
         DragIcon        =   "M_Minu02.frx":31A4
         Height          =   4860
         Left            =   -25935
         TabIndex        =   14
         Top             =   -19860
         Width           =   10935
         _Version        =   393216
         _ExtentX        =   19288
         _ExtentY        =   8573
         _StockProps     =   64
         Enabled         =   0   'False
         ColsFrozen      =   1
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         MaxRows         =   100
         RestrictRows    =   -1  'True
         SpreadDesigner  =   "M_Minu02.frx":35E6
         UserResize      =   0
         VisibleCols     =   1
         VisibleRows     =   100
      End
      Begin FPSpread.vaSpread vaSpread7 
         DragIcon        =   "M_Minu02.frx":3E9B
         Height          =   4860
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   10935
         _Version        =   393216
         _ExtentX        =   19288
         _ExtentY        =   8573
         _StockProps     =   64
         ColsFrozen      =   1
         EditEnterAction =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   15
         MaxRows         =   100
         RestrictRows    =   -1  'True
         SpreadDesigner  =   "M_Minu02.frx":42DD
         UserResize      =   0
         VisibleCols     =   1
         VisibleRows     =   100
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   840
      Width           =   10905
      Begin VB.Label Label1 
         Alignment       =   2  'Center
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
         Left            =   4875
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   5040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   23
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":4B92
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":4EAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":51C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":54E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":57FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":5B14
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":5E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":6148
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":6462
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":677C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":6A96
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":6DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":70CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":7466
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":7802
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":7B9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":7F3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":8254
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":8570
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":888A
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":9166
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":95BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Minu02.frx":9A16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   29
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Description     =   "Grabar Semana"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Description     =   "Grabar Semana"
            Object.ToolTipText     =   "Grabar Semana"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Deshacer"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Insertar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Subir"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bajar"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Receta"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Semana Anterior"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Semana Siguiente"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Semana Siguiente"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Minutas"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Borrar Días"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Copiar Plantilla Menu"
            ImageIndex      =   23
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Bloquear Minutas"
            ImageIndex      =   21
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cerrar"
            ImageIndex      =   17
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin FPSpread.vaSpread vaSpread8 
      Height          =   375
      Left            =   9720
      TabIndex        =   16
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
      _Version        =   393216
      _ExtentX        =   1931
      _ExtentY        =   661
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
      MaxCols         =   15
      MaxRows         =   100
      SpreadDesigner  =   "M_Minu02.frx":9E76
   End
   Begin FPSpread.vaSpread vaSpread2 
      Height          =   375
      Left            =   8280
      TabIndex        =   17
      Top             =   6120
      Visible         =   0   'False
      Width           =   1095
      _Version        =   393216
      _ExtentX        =   1931
      _ExtentY        =   661
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
      MaxCols         =   56
      MaxRows         =   100
      SpreadDesigner  =   "M_Minu02.frx":A64C
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Planificación Minutas"
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
      TabIndex        =   3
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
         Caption         =   "Grabar &Nuevo Ciclo Menú"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   2
         Visible         =   0   'False
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Copiar Plantilla Menu"
         Index           =   3
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu Plantilla 
         Caption         =   "C&opiar Minutas"
         Index           =   5
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Borrar Día(s)"
         Index           =   6
         Visible         =   0   'False
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   7
      End
      Begin VB.Menu Plantilla 
         Caption         =   "Ver &Receta"
         Index           =   8
      End
      Begin VB.Menu Plantilla 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu Plantilla 
         Caption         =   "&Cerrar"
         Index           =   10
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
      End
      Begin VB.Menu Plato 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu Plato 
         Caption         =   "Cambiar Plato &Menú"
         Index           =   2
      End
      Begin VB.Menu Plato 
         Caption         =   "Come&ntario"
         Index           =   3
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
   End
   Begin VB.Menu Main 
      Caption         =   "&Ver"
      Index           =   2
      NegotiatePosition=   2  'Middle
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
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Cambiar Plato &Menú"
         Index           =   2
      End
      Begin VB.Menu OpGrilla 
         Caption         =   "Come&ntario"
         Index           =   3
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
   End
End
Attribute VB_Name = "M_Minu02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConSql As ADODB.Recordset
Dim textocomentario As String, tipodatadj As String
Dim newcol3 As Long, newrow As Long, newcol1 As Long, newrow1 As Long, newcol2 As Long
Dim indcortarpegar As Long, wsmaxfilas As Long, SwCol As Long, XCol2 As Long, fecha As Long
Dim maxcolumna As Long, inddia As Long
Dim iblockrow As Integer, iblockrow2 As Integer, iblockcol As Integer, iblockcol2 As Integer, SwSalir As Integer
Dim aiblockrow As Integer, aiblockrow2 As Integer, aiblockcol As Integer, aiblockcol2 As Integer, IOpcion As Integer
Dim auxcol As Integer, auxcol2 As Integer, auxrow As Integer, auxrow2 As Integer
Dim iauxcol As Integer, iauxcol2 As Integer, iauxrow As Integer, iauxrow2 As Integer, indactivo As Integer, estado As Integer
Dim swgrabadoplan As Integer, swgrabadoest As Integer, swgrabadoadj As Integer
Dim grilla1 As Object, grilla2 As Object
Private Sub Form_Activate()
fg_descarga
End Sub
Private Sub Form_Load()
Me.Height = 6765
Me.Width = 11055
fg_centra Me
fg_carga (ss)
vaSpread1.MaxRows = 100: wsmaxfilas = 0: SwSalir = 0: maxcolumna = 0: tipodatadj = "1"
IndGrabadoTitulo = 0: IndGrabadoDetalle = 0: indactivo = 0: inddia = 0
iblockrow = 0: iblockrow2 = 0: estado = 0: swgrabadoplan = 0: swgrabadoest = 0: swgrabadoadj = 0
Label4.Caption = M_Minu01.fpayuda(2).Text & " - " & M_Minu01.fpayuda(3).Text
Label1(1).Caption = M_Minu01.fpayuda(1).Text & "(" & M_Minu01.fpText.Text & ")"
Set grilla1 = vaSpread1
Set Program = M_Minu02.vaSpread1

MoverPlantillaMinuta
MoverEstructuraFija

' *** Mover Salad Bar ***
maxcolumna = 2
Set grilla1 = vaSpread5
Set grilla2 = vaSpread6
tipodatadj = "1"
MoverDatosAdjunto

' *** Mover Postres ***
maxcolumna = 2
Set grilla1 = vaSpread7
Set grilla2 = vaSpread8
tipodatadj = "2"
MoverDatosAdjunto

Set grilla1 = vaSpread1
Set grilla2 = vaSpread2
Set Program = M_Minu02.vaSpread1
vaTabPro1.Tab = 0

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Cargando Detalle Planificación Minuta"
End Sub
Private Sub Form_Resize()
'If M_Minu02.WindowState <> 1 Then Label4.Move 0, 360, ScaleWidth, 435 'ScaleHeight - Toolbar1.Height
'If M_Minu02.WindowState <> 1 Then Frame1.Move 0, 840, ScaleWidth, 445
'If M_Minu02.WindowState <> 1 Then vaSpread1.Move 0, 1380, ScaleWidth, ScaleHeight - 1380
'Toolbar1.Refresh
End Sub
Private Sub Form_Unload(Cancel As Integer)
'Dim msg, Response   ' Declara variables.
Dim delrow As Long
If SwSalir = 0 Then
   If IndGrabadoTitulo = 1 Or IndGrabadoDetalle = 1 Then
      TITLE = "Mantención Planificación Minutas"
      msg = " Esta Seguro De Actualizar ?"
      Style = vbYesNoCancel + vbQuestion + vbDefaultButton2
      Help = "DEMO.HLP"
      Ctxt = 1000
      ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
      Select Case ws_respuesta
        Case Is = vbYes
          If vaSpread1.MaxRows > 100 Then
             delrow = vaSpread1.MaxRows - 100
             vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
          End If
          If swgrabadoplan = 1 Then inddia = 1: GrabarDatosPlantillaMinutas
          If swgrabadoest = 1 Then inddia = 1: GrabarDatosEstructuraFija
          If swgrabadoadj = 1 Then tipodatadj = "1": Set grilla1 = vaSpread5: Set grilla2 = vaSpread6: inddia = 1: GrabarDatosAdjuntos
          If swgrabadoadj = 1 Then tipodatadj = "2": Set grilla1 = vaSpread7: Set grilla2 = vaSpread8: inddia = 3: GrabarDatosAdjuntos
          IndGrabadoTitulo = 0: IndGrabadoDetalle = 0: swgrabadoplan = 0: swgrabadoest = 0: swgrabadoadj = 0
          Plantilla(0).Enabled = False
          Toolbar1.Buttons(1).Visible = True
          Toolbar1.Buttons(2).Visible = False
          SwSalir = 1
          Me.Hide
          Unload Me
        Case Is = vbCancel
          Cancel = -1
      End Select
   End If
End If
End Sub

Private Sub Plantilla_Click(Index As Integer)
Dim SubCReceta As String
Dim delrow As Long
Select Case Index
  Case 0
    If vaSpread1.MaxRows > 100 Then
       delrow = vaSpread1.MaxRows - 100
       vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
    End If
    TITLE = "Mantención Planificación Minutas"
    msg = " Esta Seguro De Actualizar ?"
    Style = vbYesNoCancel + vbQuestion + vbDefaultButton2
    Help = "DEMO.HLP"
    Ctxt = 1000
    ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
    Select Case ws_respuesta
      Case Is = vbYes
        If swgrabadoplan = 1 Then inddia = 1: GrabarDatosPlantillaMinutas
        If swgrabadoest = 1 Then inddia = 1: GrabarDatosEstructuraFija
        If swgrabadoadj = 1 Then tipodatadj = "1": Set grilla1 = vaSpread5: Set grilla2 = vaSpread6: inddia = 1: GrabarDatosAdjuntos
        If swgrabadoadj = 1 Then tipodatadj = "2": Set grilla1 = vaSpread7: Set grilla2 = vaSpread8: inddia = 3: GrabarDatosAdjuntos
        IndGrabadoTitulo = 0: IndGrabadoDetalle = 0: swgrabadoplan = 0: swgrabadoest = 0: swgrabadoadj = 0
        Plantilla(0).Enabled = False
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(2).Visible = False
      Case Is = vbCancel
        Exit Sub
    End Select
  Case 3
    ' *** Copiar Plantilla Menu *** '
    vg_opplanmenu = 0
    vaSpread1.Row = 2
    vaSpread1.Col = vaSpread1.ActiveCol
'    If vaSpread1.Col = 1 Then Exit Sub
    Select Case vaSpread1.Col
      Case 2, 3
        vaSpread1.Col = 9
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 10, 11
        vaSpread1.Col = 17
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 18, 19
        vaSpread1.Col = 25
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 26, 27
        vaSpread1.Col = 33
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 34, 35
        vaSpread1.Col = 41
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 42, 43
        vaSpread1.Col = 49
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 50, 51
        vaSpread1.Col = 57
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 58, 59
        vaSpread1.Col = 65
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 66, 67
        vaSpread1.Col = 73
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 74, 75
        vaSpread1.Col = 81
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 82, 83
        vaSpread1.Col = 89
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 90, 91
        vaSpread1.Col = 97
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 98, 99
        vaSpread1.Col = 105
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 106, 107
        vaSpread1.Col = 113
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 114, 115
        vaSpread1.Col = 121
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 122, 123
        vaSpread1.Col = 129
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 130, 131
        vaSpread1.Col = 137
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 138, 139
        vaSpread1.Col = 145
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 146, 147
        vaSpread1.Col = 153
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 154, 155
        vaSpread1.Col = 161
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 162, 163
        vaSpread1.Col = 169
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 170, 171
        vaSpread1.Col = 177
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 178, 179
        vaSpread1.Col = 185
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 186, 187
        vaSpread1.Col = 193
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 194, 195
        vaSpread1.Col = 201
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 202, 203
        vaSpread1.Col = 209
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 210, 211
        vaSpread1.Col = 217
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 218, 219
        vaSpread1.Col = 225
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 226, 227
        vaSpread1.Col = 233
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 234, 235
        vaSpread1.Col = 241
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 242, 243
        vaSpread1.Col = 249
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 250, 251
        vaSpread1.Col = 257
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 258, 259
        vaSpread1.Col = 265
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 266, 267
        vaSpread1.Col = 273
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 274, 275
        vaSpread1.Col = 281
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 282, 283
        vaSpread1.Col = 289
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 290, 291
        vaSpread1.Col = 297
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 298, 299
        vaSpread1.Col = 305
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 306, 307
        vaSpread1.Col = 313
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 314, 315
        vaSpread1.Col = 321
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 322, 323
        vaSpread1.Col = 329
        vg_fechainimenu = Val(vaSpread1.Text)
      Case 330, 331
        vaSpread1.Col = 337
        vg_fechainimenu = Val(vaSpread1.Text)
    
    End Select
    M_CpoPmi.Show 1
    If vg_opplanmenu = 1 Then
       MoverVecDia
       MoverPlantillaMinuta
    End If
'    M_Minu03.Show 1
  Case 5
    M_Minu05.Show 1
  Case 6
    M_Minu04.Show 1
  Case 8
    grilla1.Row = grilla1.ActiveRow
    grilla1.Col = grilla1.ActiveCol
    Select Case vaSpread1.Col
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        newcol3 = grilla1.Col
        grilla1.Col = newcol3 + 5 '7
        If Val(grilla1.Value) = 0 Then Exit Sub
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        newcol3 = grilla1.Col - 1
        grilla1.Col = newcol3 + 5 '7
        If Val(grilla1.Value) = 0 Then Exit Sub
    End Select
    grilla1.Col = newcol3 + 4
    If Val(grilla1.Value) = 1 Then
       grilla1.Col = newcol3 + 5
       vg_vercodreceta = Val(grilla1.Text)
       V_Recetas.Show 1
'       SubCReceta = vaSpread1.Value
'       vRet = Shell(dir_trabajo & "\Subrecet.exe " & vg_NUsr & "," & vg_Pass & "," & CStr(SubCReceta) & "," & CStr(WsCodPVenta) & ",", 1)
    Else
       MsgBox "Esto No Es Una Receta Para Vizualizar", vbCritical + vbOKOnly, "Mantención Planificación Minutas": Exit Sub
    End If
  Case 10
    If IndGrabadoTitulo = 1 Or IndGrabadoDetalle = 1 Then
       TITLE = "Mantención Planificación Minutas"
       msg = " Esta Seguro De Actualizar ?"
       Style = vbYesNoCancel + vbQuestion + vbDefaultButton2
       Help = "DEMO.HLP"
       Ctxt = 1000
       ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
       Select Case ws_respuesta
         Case Is = vbYes
           If vaSpread1.MaxRows > 100 Then
              delrow = vaSpread1.MaxRows - 100
              vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
           End If
           If swgrabadoplan = 1 Then inddia = 1: GrabarDatosPlantillaMinutas
           If swgrabadoest = 1 Then inddia = 1: GrabarDatosEstructuraFija
           If swgrabadoadj = 1 Then tipodatadj = "1": Set grilla1 = vaSpread5: Set grilla2 = vaSpread6: inddia = 1: GrabarDatosAdjuntos
           If swgrabadoadj = 1 Then tipodatadj = "2": Set grilla1 = vaSpread7: Set grilla2 = vaSpread8: inddia = 3: GrabarDatosAdjuntos
           swgrabadoplan = 0: swgrabadoest = 0: swgrabadoadj = 0
'           GrabarDatosPlantillaMinutas
         Case Is = vbCancel
           Exit Sub
       End Select
    End If
    SwSalir = 1
    Me.Hide
    Unload Me
End Select
End Sub
Private Sub Plato_Click(Index As Integer)
Dim auxp1 As Integer, auxp2 As Integer, auxp3 As Integer, auxp4 As Integer, auxp5 As Integer, auxp6 As Integer, i As Integer
Dim Del_Row As Integer, indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer

If estado = 1 Then Exit Sub
Select Case Index
 Case 0
   Plato(0).Enabled = False
   OpGrilla(0).Enabled = False
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   Toolbar1.Buttons(9).Visible = True
   Toolbar1.Buttons(10).Visible = False
   Select Case IOpcion
     Case 2
       IOpcion = 0
       auxcol = auxcol
       grilla1.Col = auxcol
       grilla1.Row = 101
       grilla1.Col2 = auxcol2 - 1
       grilla1.Row2 = 101
       grilla1.DestCol = auxcol
       grilla1.DestRow = auxrow
       grilla1.Action = 20
     Case 5
'Devolver Datos Insertados
       IOpcion = 0
       For i = auxcol To auxcol2
           grilla1.Col = auxcol
           grilla1.Row = auxrow2 + 1
           grilla1.Col2 = auxcol + 6
           grilla1.Row2 = 100
           grilla1.DestCol = auxcol
           grilla1.DestRow = auxrow
           grilla1.Action = 20
           i = auxcol + 8
           auxcol = auxcol + 8
       Next i
     Case 6
' Devolver datos eliminados
       IOpcion = 0
       newcol3 = 0
       For i = auxcol To auxcol2
           If i = 1 Then
              newcol3 = i: i = 1
           ElseIf i = 2 Or i = 10 Or i = 18 Or i = 26 Or i = 34 Or i = 42 Or i = 50 Or i = 58 Or i = 66 Or i = 74 Or i = 82 Or i = 90 Or i = 98 Or i = 106 Or i = 114 Or i = 122 Or i = 130 Or i = 138 Or i = 146 Or i = 154 Or i = 162 Or i = 170 Or i = 178 Or i = 186 Or i = 194 Or i = 202 Or i = 210 Or i = 218 Or i = 226 Or i = 234 Or i = 242 Or i = 250 Or i = 258 Or i = 266 Or i = 274 Or i = 282 Or i = 290 Or i = 298 Or i = 306 Or i = 314 Or i = 322 Or i = 330 Then
              newcol3 = i: i = newcol3 + 7
           ElseIf i = 3 Or i = 11 Or i = 19 Or i = 27 Or i = 35 Or i = 43 Or i = 51 Or i = 59 Or i = 67 Or i = 75 Or i = 83 Or i = 91 Or i = 99 Or i = 107 Or i = 115 Or i = 123 Or i = 131 Or i = 139 Or i = 147 Or i = 155 Or i = 163 Or i = 171 Or i = 179 Or i = 187 Or i = 195 Or i = 203 Or i = 211 Or i = 219 Or i = 227 Or i = 235 Or i = 243 Or i = 251 Or i = 259 Or i = 267 Or i = 275 Or i = 283 Or i = 291 Or i = 299 Or i = 307 Or i = 315 Or i = 323 Or i = 331 Then
              newcol3 = i - 1: i = newcol3 + 7
           End If
           grilla1.Col = newcol3
           grilla1.Row = auxrow
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = newcol3 + 6
           End If
           grilla1.Row2 = 100 - ((iauxrow2 - iauxrow) + 1)
           grilla1.DestCol = newcol3
           grilla1.DestRow = iauxrow2 + 1
           grilla1.Action = 20
            
           auxp5 = (iauxrow2 - iauxrow) + 1
           grilla1.Col = newcol3
           grilla1.Row = 1 + 100
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = newcol3 + 6
           End If
           grilla1.Row2 = auxp5 + 100
           grilla1.DestCol = newcol3
           grilla1.DestRow = auxrow
           grilla1.Action = 20
           If i = 1 Then
              auxcol = 2
           Else
              auxcol = auxcol + 8
           End If
       Next i
     Case 9
       For i = auxcol To auxcol2
           If grilla1.MaxRows <= 100 Then
              grilla1.MaxRows = (grilla1.MaxRows + (auxrow2 - auxrow)) + 1
              For auxp1 = 101 To grilla1.MaxRows
                  grilla1.Row = auxp1
                  grilla1.RowHidden = True
              Next auxp1
           End If
           grilla1.Col = auxcol
           grilla1.Row = auxrow
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = auxcol + 6
           End If
           grilla1.Row2 = auxrow2
           grilla1.DestCol = auxcol
           grilla1.DestRow = 101
           grilla1.Action = 20

'' ***      Copiar Datos a la fila Seleccionada *** '
           auxp6 = (auxrow2 - auxrow) + 1
           grilla1.Col = auxcol
           grilla1.Row = auxrow + auxp6
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = auxcol + 6
           End If
           grilla1.Row2 = auxrow2 + auxp6
           grilla1.DestCol = auxcol
           grilla1.DestRow = auxrow
           grilla1.Action = 20

' ***      Devolver Datos a la fila y restar ultima fila *** '
           grilla1.Col = auxcol
           grilla1.Row = 101
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = auxcol + 6
           End If
           grilla1.Row2 = 100 + ((auxrow2 - auxrow) + 1)
           grilla1.DestCol = auxcol
           grilla1.DestRow = auxrow + auxp6
           grilla1.Action = 20
           If i = 1 Then
              i = 2
              auxcol = 2
           Else
              i = auxcol + 7
              auxcol = auxcol + 8
           End If
       Next i
     Case 8
       For i = auxcol To auxcol2
           If grilla1.MaxRows <= 100 Then
              grilla1.MaxRows = (grilla1.MaxRows + (auxrow2 - auxrow)) + 1
              For auxp1 = 101 To grilla1.MaxRows
                  grilla1.Row = auxp1
                  grilla1.RowHidden = True
              Next auxp1
           End If
           grilla1.Col = auxcol
           grilla1.Row = auxrow
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = auxcol + 6
           End If
           grilla1.Row2 = auxrow2
           grilla1.DestCol = auxcol
           grilla1.DestRow = 101
           grilla1.Action = 20
        
' ***      Copiar Datos a la fila Seleccionada *** '
           auxp6 = (auxrow2 - auxrow) + 1
           grilla1.Col = auxcol
           grilla1.Row = auxrow - auxp6
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = auxcol + 6
           End If
           grilla1.Row2 = auxrow2 - auxp6
           grilla1.DestCol = auxcol
           grilla1.DestRow = auxrow
           grilla1.Action = 20
       
' ***      Devolver Datos a la fila y restar ultima fila *** '
           grilla1.Col = auxcol
           grilla1.Row = 101
           If i = 1 Then
              grilla1.Col2 = 1
           Else
              grilla1.Col2 = auxcol + 6
           End If
           grilla1.Row2 = 100 + ((auxrow - auxrow) + 1)
           grilla1.DestCol = auxcol
           grilla1.DestRow = auxrow - auxp6
           grilla1.Action = 20
           If i = 1 Then
              i = 2
              auxcol = 2
           Else
              i = auxcol + 7
              auxcol = auxcol + 8
           End If
       Next i
     Case 13
' destinacion de copiar y pegar datos
       IOpcion = 0
       newcol3 = 0
       If auxcol2 < iauxcol2 Then auxcol2 = iauxcol2
       For i = auxcol To auxcol2
           If i = 2 Or i = 10 Or i = 18 Or i = 26 Or i = 34 Or i = 42 Or i = 50 Or i = 58 Or i = 66 Or i = 74 Or i = 82 Or i = 90 Or i = 98 Or i = 106 Or i = 114 Or i = 122 Or i = 130 Or i = 138 Or i = 146 Or i = 154 Or i = 162 Or i = 170 Or i = 178 Or i = 186 Or i = 194 Or i = 202 Or i = 210 Or i = 218 Or i = 226 Or i = 234 Or i = 242 Or i = 250 Or i = 258 Or i = 266 Or i = 274 Or i = 282 Or i = 290 Or i = 298 Or i = 306 Or i = 314 Or i = 322 Or i = 330 Then
              newcol3 = i: i = newcol3 + 7
           ElseIf i = 3 Or i = 11 Or i = 19 Or i = 27 Or i = 35 Or i = 43 Or i = 51 Or i = 59 Or i = 67 Or i = 75 Or i = 83 Or i = 91 Or i = 99 Or i = 107 Or i = 115 Or i = 123 Or i = 131 Or i = 139 Or i = 147 Or i = 155 Or i = 163 Or i = 171 Or i = 179 Or i = 187 Or i = 195 Or i = 203 Or i = 211 Or i = 219 Or i = 227 Or i = 235 Or i = 243 Or i = 251 Or i = 259 Or i = 267 Or i = 275 Or i = 283 Or i = 291 Or i = 299 Or i = 307 Or i = 315 Or i = 323 Or i = 331 Then
              newcol3 = i - 1: i = newcol3 + 7
           End If
           If grilla1.MaxRows <= 100 Then grilla1.MaxRows = (grilla1.MaxRows + (iauxrow2 - iauxrow)) + 1
           auxp6 = iauxrow2 - iauxrow + auxrow
           grilla1.Col = newcol3
           grilla1.Row = auxrow
           grilla1.Col2 = newcol3 + 6
           grilla1.Row2 = auxp6
           grilla1.DestCol = iauxcol '26
           grilla1.DestRow = iauxrow
           grilla1.Action = 20
            
           auxp5 = (iauxrow2 - iauxrow) + 1
           grilla1.Col = newcol3
           grilla1.Row = 1 + 100
           grilla1.Col2 = newcol3 + 6
           grilla1.Row2 = auxp5 + 100
           grilla1.DestCol = newcol3
           grilla1.DestRow = auxrow
           grilla1.Action = 20
           auxcol = auxcol + 8
           iauxcol = iauxcol + 8
       Next i
     Case 14
       IOpcion = 0
       newcol3 = 0
       For i = auxcol To auxcol2
           If i = 2 Or i = 10 Or i = 18 Or i = 26 Or i = 34 Or i = 42 Or i = 50 Or i = 58 Or i = 66 Or i = 74 Or i = 82 Or i = 90 Or i = 98 Or i = 106 Or i = 114 Or i = 122 Or i = 130 Or i = 138 Or i = 146 Or i = 154 Or i = 162 Or i = 170 Or i = 178 Or i = 186 Or i = 194 Or i = 202 Or i = 210 Or i = 218 Or i = 226 Or i = 234 Or i = 242 Or i = 250 Or i = 258 Or i = 266 Or i = 274 Or i = 282 Or i = 290 Or i = 298 Or i = 306 Or i = 314 Or i = 322 Or i = 330 Then
              newcol3 = i: i = newcol3 + 7
           ElseIf i = 3 Or i = 11 Or i = 19 Or i = 27 Or i = 35 Or i = 43 Or i = 51 Or i = 59 Or i = 67 Or i = 75 Or i = 83 Or i = 91 Or i = 99 Or i = 107 Or i = 115 Or i = 123 Or i = 131 Or i = 139 Or i = 147 Or i = 155 Or i = 163 Or i = 171 Or i = 179 Or i = 187 Or i = 195 Or i = 203 Or i = 211 Or i = 219 Or i = 227 Or i = 235 Or i = 243 Or i = 251 Or i = 259 Or i = 267 Or i = 275 Or i = 283 Or i = 291 Or i = 299 Or i = 307 Or i = 315 Or i = 323 Or i = 331 Then
              newcol3 = i - 1: i = newcol3 + 7
           End If
           auxp5 = (iauxrow2 - iauxrow) + 1
           grilla1.Col = newcol3
           grilla1.Row = 1 + 100
           grilla1.Col2 = newcol3 + 6
           grilla1.Row2 = auxp5 + 100
           grilla1.DestCol = newcol3
           grilla1.DestRow = auxrow
           grilla1.Action = 20
           auxcol = auxcol + 8
       Next i
       
   End Select
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   grilla1.SetFocus
 Case 2
' Ingresar Recetas
   WsTipoMenu = 5
   iblockcol = grilla1.ActiveCol: aiblockcol = grilla1.ActiveCol
   iblockcol2 = grilla1.ActiveCol: aiblockcol2 = grilla1.ActiveCol
   iblockrow = grilla1.ActiveRow: aiblockrow = grilla1.ActiveRow
   iblockrow2 = grilla1.ActiveRow: aiblockrow2 = grilla1.ActiveRow
   grilla1.Row = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   Select Case grilla1.Col
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol3 = grilla1.Col
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       newcol3 = grilla1.Col - 1
   End Select
   IndPlantilla = 1
   ICGrilla = 0
   IndColumna = newcol3
   Select Case iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol = iblockcol
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol = iblockcol - 1
   End Select
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol2 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol2 = ((iblockcol2 + 7) - 1)
   End Select
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   grilla1.MaxRows = grilla1.MaxRows + 1
   grilla1.Row = grilla1.MaxRows
   grilla1.RowHidden = True
   grilla1.Col = iblockcol
   grilla1.Row = iblockrow
   grilla1.Col2 = iblockcol2 - 1
   grilla1.Row2 = iblockrow
' Copiar Columna
   grilla1.DestCol = iblockcol
   grilla1.DestRow = 101
   grilla1.Action = 19
   M_Minu03.Show 1
   If vaTabPro1.Tab = 0 Then
     swgrabadoplan = 1
     vaTabPro1.Tab = 0
   ElseIf vaTabPro1.Tab = 1 Then
     swgrabadoest = 1
     vaTabPro1.Tab = 1
   ElseIf vaTabPro1.Tab = 2 Then
     swgrabadoadj = 1
     vaTabPro1.Tab = 2
   ElseIf vaTabPro1.Tab = 3 Then
     swgrabadoadj = 1
     vaTabPro1.Tab = 3
   End If
   vaTabPro1.Refresh
   If ICGrilla = 0 Then
      If grilla1.MaxRows > 100 Then
         Del_Row = grilla1.MaxRows - 100
         grilla1.MaxRows = grilla1.MaxRows - Del_Row
      End If
      Exit Sub
   End If
   IOpcion = 2
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   IndPlantilla = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
 Case 3
    If vaTabPro1.Tab = 0 Then
       swgrabadoplan = 1
    ElseIf vaTabPro1.Tab = 1 Then
       swgrabadoest = 1
    ElseIf vaTabPro1.Tab = 2 Then
       swgrabadoadj = 1
    ElseIf vaTabPro1.Tab = 3 Then
       swgrabadoadj = 1
    End If
   iblockcol = grilla1.ActiveCol: aiblockcol = grilla1.ActiveCol
   iblockcol2 = grilla1.ActiveCol: aiblockcol2 = grilla1.ActiveCol
   iblockrow = grilla1.ActiveRow: aiblockrow = grilla1.ActiveRow
   iblockrow2 = grilla1.ActiveRow: aiblockrow2 = grilla1.ActiveRow
   grilla1.Row = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   If grilla1.Col = 1 Then Exit Sub
   Select Case grilla1.Col
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol3 = grilla1.Col
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       newcol3 = grilla1.Col - 1
   End Select
   grilla1.Col = newcol3
   If grilla1.Text <> "" Then
      grilla1.Col = newcol3 + 1
      textocomentario = ""
   Else
      grilla1.Col = newcol3 + 1
      textocomentario = grilla1.Text
   End If
   Select Case iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol = iblockcol
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol = iblockcol - 1
   End Select
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol2 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol2 = ((iblockcol2 + 7) - 1)
   End Select
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   grilla1.MaxRows = grilla1.MaxRows + 1
   grilla1.Row = grilla1.MaxRows
   grilla1.RowHidden = True
   grilla1.Col = iblockcol
   grilla1.Row = iblockrow
   grilla1.Col2 = iblockcol2 - 1
   grilla1.Row2 = iblockrow
'  Copiar Columna
   grilla1.DestCol = iblockcol
   grilla1.DestRow = 101
   grilla1.Action = 19
   M_Minu11.fpText1.Text = textocomentario
   ICGrilla = 0
   M_Minu11.Show 1
   vaTabPro1.Tab = 0
   vaTabPro1.Refresh
   If ICGrilla = 0 Or M_Minu11.fpText1.Text = "" Then
      If grilla1.MaxRows > 100 Then
         Del_Row = grilla1.MaxRows - 100
         grilla1.MaxRows = grilla1.MaxRows - Del_Row
      End If
      Exit Sub
   End If
   grilla1.Row = grilla1.ActiveRow
   grilla1.Col = newcol3
   grilla1.Row2 = grilla1.ActiveRow
   grilla1.Col2 = newcol3 + 6
   grilla1.BlockMode = True
' Limpiar Datos y Formato Celda
   grilla1.Action = 3
  ' Retorna Modo de la columna
   grilla1.BlockMode = False
        
   grilla1.Col = newcol3 + 1
   grilla1.Text = M_Minu11.fpText1.Text
   grilla1.Font.Bold = True
   grilla1.Font.Size = 9
   grilla1.Col = newcol3 + 4
   grilla1.Text = 6
   Unload M_Minu11
   IOpcion = 2
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
 Case 5
   If vaTabPro1.Tab = 0 Then
      swgrabadoplan = 1
   ElseIf vaTabPro1.Tab = 1 Then
      swgrabadoest = 1
   ElseIf vaTabPro1.Tab = 2 Then
      swgrabadoadj = 1
   ElseIf vaTabPro1.Tab = 3 Then
      swgrabadoadj = 1
   End If
   grilla1.Row = grilla1.ActiveRow
   indfila = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   IndColumna = grilla1.Col
   Select Case grilla1.Col
    Case 1
       newcol3 = grilla1.Col
    Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol3 = grilla1.Col
    Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       newcol3 = grilla1.Col - 1
   End Select
   For i = 1 To grilla1.MaxRows
       grilla1.Row = i: grilla1.Col = newcol3 + 1
       If grilla1.Text <> "" Then wsmaxfilas = grilla1.Row
   Next i
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   wsmaxfilas = (wsmaxfilas + (iblockrow2 - iblockrow) + 1)
   If wsmaxfilas > grilla1.MaxRows Then Exit Sub
   IOpcion = 5
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   grilla1.Row = indfila
   grilla1.Col = IndColumna
   If iblockcol < 0 Then iblockcol = 1: iblockcol2 = grilla1.MaxCols '255
   Select Case iblockcol
     Case 1
       iblockcol = iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol = iblockcol
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol = iblockcol - 1
   End Select
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol2 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol2 = ((iblockcol2 + 7) - 1)
   End Select
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   indcol = iblockcol
   For i = iblockcol To iblockcol2
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = 100 - ((iblockrow2 - iblockrow) + 1) ' 99
' Insertar Columna
       If i = 1 Then
          grilla1.DestCol = iblockcol
          grilla1.DestRow = iblockrow2 + 1
       Else
          grilla1.DestCol = iblockcol
          grilla1.DestRow = iblockrow2 + 1
       End If
       If grilla1.DestRow < 100 Then
          grilla1.Action = 20
          If i = 1 Then grilla1.CellBorderType = 16: grilla1.CellBorderStyle = 1: grilla1.Action = 16: grilla1.BackColor = &H8000000F
       End If
       If i = 1 Then
          i = iblockcol + 1
          iblockcol = iblockcol + 1
       Else
          i = iblockcol + 7
          iblockcol = iblockcol + 8
       End If

   Next i
   iblockcol = indcol
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
 Case 6
   If vaTabPro1.Tab = 0 Then
      swgrabadoplan = 1
   ElseIf vaTabPro1.Tab = 1 Then
      swgrabadoest = 1
   ElseIf vaTabPro1.Tab = 2 Then
      swgrabadoadj = 1
   ElseIf vaTabPro1.Tab = 3 Then
      swgrabadoadj = 1
   End If
   IOpcion = 6
   grilla1.Row = grilla1.ActiveRow
   indfila = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   aiblockcol = iblockcol
   aiblockrow = iblockrow
   aiblockcol2 = iblockcol2
   aiblockrow2 = iblockrow2
   If iblockcol < 0 Then iblockcol = 1: iblockcol2 = grilla1.MaxCols '255 '56
   Select Case iblockcol
     Case 1
       iblockcol = iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol = iblockcol
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol = iblockcol - 1
   End Select
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol2 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol2 = ((iblockcol2 + 7) - 1)
   End Select
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   indcol = iblockcol
   For i = iblockcol To iblockcol2
' Calcular y Mover Datos Ultima linea
       If grilla1.MaxRows <= 100 Then
          grilla1.MaxRows = (grilla1.MaxRows + (aiblockrow2 - aiblockrow)) + 1
          For auxp1 = 101 To grilla1.MaxRows
              grilla1.Row = auxp1
              grilla1.RowHidden = True
          Next auxp1
       End If
       auxp6 = aiblockrow2 - aiblockrow + iblockrow
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = auxp6
       grilla1.DestCol = iblockcol
       grilla1.DestRow = 100 + 1
       grilla1.Action = 20
'fin mover datos ultimas lineas
      
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = iblockrow2
       grilla1.BlockMode = True
' Limpiar Datos y Formato Celda
       grilla1.Action = 3
' Retorna Modo de la columna
       auxp6 = aiblockrow2 - aiblockrow + iblockrow + 1
       grilla1.BlockMode = False
       grilla1.Col = iblockcol
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row = auxp6
       grilla1.Row2 = 100
' Insertar Columna
       grilla1.DestCol = iblockcol
       If (iblockrow2 - auxp6) < 1 Then
          grilla1.DestRow = iblockrow
       Else
          grilla1.DestRow = auxp6 ' - IBlockRow2
       End If
       If grilla1.DestRow < 101 Then
          grilla1.Action = 20
'          vaSpread2.SetActiveCell
'          If i = 1 Then grilla1.CellBorderType = 16: grilla1.CellBorderStyle = 1: grilla1.Action = 16: grilla1.BackColor = &H8000000F
       End If
       If i = 1 Then
          grilla1.Col = 1
          grilla1.Row = 1
          grilla1.Col2 = 1
          grilla1.Row2 = 100
          grilla1.BlockMode = True
          grilla1.Lock = True
          grilla1.BlockMode = False
          grilla1.LockBackColor = &HC0C0C0
       End If
       
       If i = 1 Then
          i = 2
          iblockcol = 2
       Else
          i = iblockcol + 7
          iblockcol = iblockcol + 8
       End If
   Next i
   iblockcol = indcol
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
 Case 8
   If vaTabPro1.Tab = 0 Then
      swgrabadoplan = 1
   ElseIf vaTabPro1.Tab = 1 Then
      swgrabadoest = 1
   ElseIf vaTabPro1.Tab = 2 Then
      swgrabadoadj = 1
   ElseIf vaTabPro1.Tab = 3 Then
      swgrabadoadj = 1
   End If
   grilla1.Row = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   indfila = grilla1.ActiveRow
   Select Case grilla1.Col
     Case 1
       newcol1 = grilla1.Col
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol1 = grilla1.Col
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       newcol1 = grilla1.Col - 1
   End Select
   If grilla1.Row = 1 Then Exit Sub
   If (iblockrow - ((iblockrow2 - iblockrow) + 1)) < 1 Then
      MsgBox "Imposible subir la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
      Exit Sub
   End If
   IOpcion = 8
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   If iblockcol < 0 Then iblockcol = 1: iblockcol2 = grilla1.MaxCols '56
   Select Case iblockcol
     Case 1
       iblockcol = iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol = iblockcol
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol = iblockcol - 1
   End Select
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol2 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol2 = ((iblockcol2 + 7) - 1)
   End Select
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   indcol = iblockcol
   For i = iblockcol To iblockcol2
       If grilla1.MaxRows <= 100 Then
          grilla1.MaxRows = (grilla1.MaxRows + (iblockrow2 - iblockrow)) + 1
          For auxp1 = 101 To grilla1.MaxRows
              grilla1.Row = auxp1
              grilla1.RowHidden = True
          Next auxp1
       End If
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = iblockrow2
       grilla1.DestCol = iblockcol
       grilla1.DestRow = 101
       grilla1.Action = 20
         
' ***      Copiar Datos a la fila Seleccionada *** '
       auxp6 = (iblockrow2 - iblockrow) + 1
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow - auxp6
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = iblockrow2 - auxp6
       grilla1.DestCol = iblockcol
       grilla1.DestRow = iblockrow
       grilla1.Action = 20
        
' ***      Devolver Datos a la fila y restar ultima fila *** '
       grilla1.Col = iblockcol
       grilla1.Row = 101
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = 100 + ((iblockrow2 - iblockrow) + 1)
       grilla1.DestCol = iblockcol
       grilla1.DestRow = iblockrow - auxp6
       grilla1.Action = 20
       If i = 1 Then
          i = 2
          iblockcol = 2
       Else
          i = iblockcol + 7
          iblockcol = iblockcol + 8
       End If
   Next i
   iblockcol = indcol
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   grilla1.Action = 0
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
 Case 9
   If vaTabPro1.Tab = 0 Then
      swgrabadoplan = 1
   ElseIf vaTabPro1.Tab = 1 Then
      swgrabadoest = 1
   ElseIf vaTabPro1.Tab = 2 Then
      swgrabadoadj = 1
   ElseIf vaTabPro1.Tab = 3 Then
      swgrabadoadj = 1
   End If
   grilla1.Row = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   indfila = grilla1.ActiveRow
   Select Case grilla1.Col
     Case 1
       newcol1 = 1
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol1 = grilla1.Col
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       newcol1 = grilla1.Col - 1
   End Select
   If grilla1.Row = 100 Then Exit Sub
   If (iblockrow2 + ((iblockrow2 - iblockrow) + 1)) > 100 Then
      MsgBox "Imposible bajar la infomación ya que el bloque es mayor al bloque destino", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
      Exit Sub
   End If
   IOpcion = 9
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   If iblockcol < 0 Then iblockcol = 1: iblockcol2 = grilla1.MaxCols '56
   Select Case iblockcol
     Case 1
       iblockcol = iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol = iblockcol
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol = iblockcol - 1
   End Select
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       iblockcol2 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       iblockcol2 = ((iblockcol2 + 7) - 1)
   End Select
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   indcol = iblockcol
   For i = iblockcol To iblockcol2
       If grilla1.MaxRows <= 100 Then
          grilla1.MaxRows = (grilla1.MaxRows + (iblockrow2 - iblockrow)) + 1
          For auxp1 = 101 To grilla1.MaxRows
              grilla1.Row = auxp1
              grilla1.RowHidden = True
          Next auxp1
       End If
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = iblockrow2
       grilla1.DestCol = iblockcol
       grilla1.DestRow = 101
       grilla1.Action = 20
        
' ***      Copiar Datos a la fila Seleccionada *** '
       auxp6 = (iblockrow2 - iblockrow) + 1
       grilla1.Col = iblockcol
       grilla1.Row = iblockrow + auxp6
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = iblockrow2 + auxp6
       grilla1.DestCol = iblockcol
       grilla1.DestRow = iblockrow
       grilla1.Action = 20
        
' ***      Devolver Datos a la fila y restar ultima fila *** '
       grilla1.Col = iblockcol
       grilla1.Row = 101
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = iblockcol + 6
       End If
       grilla1.Row2 = 100 + ((iblockrow2 - iblockrow) + 1)
       grilla1.DestCol = iblockcol
       grilla1.DestRow = iblockrow + auxp6
       grilla1.Action = 20
       If i = 1 Then
          i = 2
          iblockcol = 2
       Else
          i = iblockcol + 7
          iblockcol = iblockcol + 8
       End If
   Next i
   iblockcol = indcol
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   grilla1.Action = 0
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
   Toolbar1.Buttons(6).Visible = True
   Toolbar1.Buttons(7).Visible = False
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
 Case 11, 12
   If vaTabPro1.Tab = 0 Then
      swgrabadoplan = 1
   ElseIf vaTabPro1.Tab = 1 Then
      swgrabadoest = 1
   ElseIf vaTabPro1.Tab = 2 Then
      swgrabadoadj = 1
   ElseIf vaTabPro1.Tab = 3 Then
      swgrabadoadj = 1
   End If
   grilla1.Row = grilla1.ActiveRow
   grilla1.Col = grilla1.ActiveCol
   aiblockrow = iblockrow
   aiblockrow2 = iblockrow2
   aiblockcol = iblockcol
   aiblockcol2 = iblockcol2
   Select Case grilla1.Col
     Case 1
       newcol3 = grilla1.Col
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol3 = grilla1.Col
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       aiblockcol = aiblockcol - 1
       newcol3 = grilla1.Col - 1
   End Select
  
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If
   Plato(13).Enabled = True
   OpGrilla(13).Enabled = True
   Toolbar1.Buttons(6).Visible = False
   Toolbar1.Buttons(7).Visible = True

   If iblockcol < 1 Then aiblockcol = 1: aiblockcol2 = grilla1.MaxCols '56 '
   If Index = 11 Then
      indcortarpegar = 0
   Else
      indcortarpegar = 1
   End If
 Case 13
   If vaTabPro1.Tab = 0 Then
      swgrabadoplan = 1
   ElseIf vaTabPro1.Tab = 1 Then
      swgrabadoest = 1
   ElseIf vaTabPro1.Tab = 2 Then
      swgrabadoadj = 1
   ElseIf vaTabPro1.Tab = 3 Then
      swgrabadoadj = 1
   End If
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
      IOpcion = 13
   Else
      If (iblockrow2 - iblockrow) > (aiblockrow2 - aiblockrow) Then
         MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
         Exit Sub
      End If
      If aiblockcol <> iblockcol2 And aiblockcol = 1 Then
         MsgBox "Imposible Pegar la infomación ya que el área de Copiar y el área de Pegado" & VgLinea & "tienen formas distintas, intente lo siguiente :" & VgLinea & "* Haga clic en una única misma celda y luego eliga Pegar." & VgLinea & "* Seleccione un rectángulo con el mismo tamaño y forma y luego eliga Pegar.", vbInformation + vbOKOnly, "Detalle Planificación Minutas"
         Exit Sub
      End If
      IOpcion = 14
   End If
   grilla1.Col = grilla1.ActiveCol
   Select Case grilla1.Col
     Case 1
       newcol3 = grilla1.Col
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       newcol3 = grilla1.Col
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       newcol3 = grilla1.Col - 1
   End Select
   If indcortarpegar = 0 Then OpGrilla(13).Enabled = False: Toolbar1.Buttons(6).Visible = True: Toolbar1.Buttons(7).Visible = False
   Plato(0).Enabled = True
   OpGrilla(0).Enabled = True
   Toolbar1.Buttons(9).Visible = False
   Toolbar1.Buttons(10).Visible = True
'   IOpcion = 13
   If grilla1.MaxRows > 100 Then
      Del_Row = grilla1.MaxRows - 100
      grilla1.MaxRows = grilla1.MaxRows - Del_Row
   End If

   ' destinacion de copiar y pegar datos
   If iblockcol < 1 Then iblockcol = 1: iblockcol2 = grilla1.MaxCols '56
   If aiblockcol2 = grilla1.MaxCols Then aiblockcol2 = grilla1.MaxCols - 1 '56 Then aiblockcol2 = 49
   auxp1 = 0: auxp2 = 0: auxp3 = 0: auxp4 = 0: auxp5 = 0: auxp6 = 0: newcol3 = 0
   Select Case aiblockcol
     Case 1
       auxp1 = aiblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       auxp1 = aiblockcol + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       auxp1 = (aiblockcol + 7) - 1
   End Select
        
   Select Case aiblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       auxp2 = aiblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       auxp2 = (aiblockcol2 + 7) - 1
     Case 336
       auxp2 = (aiblockcol2 + 7) - 6
   End Select
   If (auxp2 - auxp1) < 1 Then
      auxp5 = 1
   Else
      auxp5 = (auxp2 / 8) + 1
   End If
   Select Case iblockcol
     Case 1
       auxp3 = iblockcol
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       auxp3 = iblockcol + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       auxp3 = (iblockcol + 7) - 1
   
   End Select
        
   Select Case iblockcol2
     Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
       auxp4 = iblockcol2 + 7
     Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
       auxp4 = (iblockcol2 + 7) - 1
   End Select
        
   If auxp3 = auxp4 And auxp1 = auxp2 Then
      Select Case iblockcol2
        Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
          iblockcol2 = (iblockcol2 + 7)
        Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
          iblockcol2 = ((iblockcol2 + 7) - 1)
      End Select
   ElseIf auxp1 = auxp3 And auxp2 = auxp4 Then
      Select Case iblockcol2
        Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
          iblockcol2 = (iblockcol2 + 7)
        Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
          iblockcol2 = ((iblockcol2 + 7) - 1)
      End Select
   ElseIf auxp1 <> auxp2 Then
      Select Case iblockcol2
        Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
          iblockcol2 = ((iblockcol2 + 7) + auxp2 - auxp1)
        Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
          iblockcol2 = ((iblockcol2 + 7 - 1) + auxp2 - auxp1)
      End Select
   End If
   auxp1 = 1: auxp3 = 1
   auxcol = iblockcol: auxcol2 = iblockcol2
   auxrow = iblockrow: auxrow2 = iblockrow2
   iauxcol = aiblockcol: iauxcol2 = aiblockcol2
   iauxrow = aiblockrow: iauxrow2 = aiblockrow2
   indcol = aiblockcol: indcol2 = iblockcol2
   indrow = aiblockrow: indrow2 = aiblockrow2
   For i = iblockcol To iblockcol2
       If i = 1 Then
          newcol3 = i: i = 1
       ElseIf i = 2 Or i = 10 Or i = 18 Or i = 26 Or i = 34 Or i = 42 Or i = 50 Or i = 58 Or i = 66 Or i = 74 Or i = 82 Or i = 90 Or i = 98 Or i = 106 Or i = 114 Or i = 122 Or i = 130 Or i = 138 Or i = 146 Or i = 154 Or i = 162 Or i = 170 Or i = 178 Or i = 186 Or i = 194 Or i = 202 Or i = 210 Or i = 218 Or i = 226 Or i = 234 Or i = 242 Or i = 250 Or i = 258 Or i = 266 Or i = 274 Or i = 282 Or i = 290 Or i = 298 Or i = 306 Or i = 314 Or i = 322 Or i = 330 Then
          newcol3 = i: i = newcol3 + 7
       ElseIf i = 3 Or i = 11 Or i = 19 Or i = 27 Or i = 35 Or i = 43 Or i = 51 Or i = 59 Or i = 67 Or i = 75 Or i = 83 Or i = 91 Or i = 99 Or i = 107 Or i = 115 Or i = 123 Or i = 131 Or i = 139 Or i = 147 Or i = 155 Or i = 163 Or i = 171 Or i = 179 Or i = 187 Or i = 195 Or i = 203 Or i = 211 Or i = 219 Or i = 227 Or i = 235 Or i = 243 Or i = 251 Or i = 259 Or i = 267 Or i = 275 Or i = 283 Or i = 291 Or i = 299 Or i = 307 Or i = 315 Or i = 323 Or i = 331 Then
          newcol3 = i - 1: i = newcol3 + 7
       End If
       
' Calcular y Mover Datos Ultima linea
       If grilla1.MaxRows <= 100 Then
          grilla1.MaxRows = (grilla1.MaxRows + (aiblockrow2 - aiblockrow)) + 1
          For auxp2 = 101 To grilla1.MaxRows
              grilla1.Row = auxp2
              grilla1.RowHidden = True
          Next auxp2
       End If
       auxp6 = aiblockrow2 - aiblockrow + iblockrow
       grilla1.Col = newcol3
       grilla1.Row = iblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = newcol3 + 6
       End If
       grilla1.Row2 = auxp6
       grilla1.DestCol = newcol3
       grilla1.DestRow = 100 + 1
       grilla1.Action = 19 '20
'fin mover datos ultimas lineas
            
       grilla1.Col = aiblockcol
       grilla1.Row = aiblockrow
       If i = 1 Then
          grilla1.Col2 = 1
       Else
          grilla1.Col2 = aiblockcol + 6
       End If
       grilla1.Row2 = aiblockrow2
       If i = 1 Then
          aiblockcol = 2
       Else
          aiblockcol = aiblockcol + 8
       End If
       grilla1.DestCol = newcol3
       grilla1.DestRow = iblockrow
       If indcortarpegar = 1 Then
          grilla1.Action = 19
       Else
          grilla1.Action = 20
          Plato(13).Enabled = False
          OpGrilla(13).Enabled = False
       End If
       If auxp5 = auxp1 Then
          If i = 1 Then
             aiblockcol = 2: auxp1 = 0
          Else
             aiblockcol = indcol: auxp1 = 0
          End If
       End If
       auxp1 = auxp1 + 1
   Next i
   aiblockcol = indcol: iblockcol2 = indcol2
   aiblockrow = indrow: aiblockrow2 = indrow2
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
End Select
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If estado = 1 And Button.Index <> 29 Then Exit Sub
Select Case Button.Index
  Case 2
    Plantilla_Click (0)
  Case 4
    Plato_Click (11)
  Case 5
     Plato_Click (12)
  Case 7
     Plato_Click (13)
  Case 10
     Plato_Click (0)
  Case 12
     Plato_Click (5)
  Case 13
     Plato_Click (6)
  Case 15
     Plato_Click (8)
  Case 16
     Plato_Click (9)
  Case 19
     Ver_Click (3)
  Case 18
     Plantilla_Click (8)
  Case 21
     Ver_Click (2)
  Case 23
    Plantilla_Click (5)
  Case 24
    Plantilla_Click (6)
  Case 25
    Plantilla_Click (3)
  Case 27
    BloquearMinuta
  Case 29
    Plantilla_Click (10)
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
If Row > 0 Then
    If Col <> 0 Then SwCol = 0
    If Col = 0 Then SwCol = 1
    indactivo = 1
    iblockrow = vaSpread1.ActiveRow
    iblockrow2 = vaSpread1.ActiveRow
    iblockcol = vaSpread1.ActiveCol
    iblockcol2 = vaSpread1.ActiveCol
    vaSpread1.Row = vaSpread1.ActiveRow
    indfila = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    IndColumna = vaSpread1.ActiveCol
End If
End Sub
Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row > 0 And estado = 0 Then
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = vaSpread1.ActiveCol
   vaSpread1.Col = Col
   If vaSpread1.Col = 1 Then Exit Sub
   Select Case vaSpread1.Col
    Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
     newcol2 = vaSpread1.Col + 1
     vaSpread1.Col = newcol2
    Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
     newcol2 = vaSpread1.Col
   End Select
   If newcol2 < 1 Then
      newcol2 = 2
   ElseIf newcol2 > 330 Then
      newcol2 = 229
   End If
'   textocomentario = vaSpread1.Text
'   newcol2 = newcol2 - 1
'   If newcol2 < 1 Then newcol2 = 2
'   vaSpread1.Col = newcol2
'   If vaSpread1.Value = "" And textocomentario <> "" Then
'      Plato_Click (3)
'   Else
      Plato_Click (2)
'   End If
End If
End Sub
Private Sub vaSpread1_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread1.ChangeMade = False Then Exit Sub
If Col = 1 And estado = 0 Then
   swgrabadoplan = 1
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
'   Toolbar1.Buttons(6).Visible = True
'   Toolbar1.Buttons(7).Visible = False
'   Toolbar1.Buttons(9).Visible = False
'   Toolbar1.Buttons(10).Visible = True
End If
End Sub
Private Sub vaSpread1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IBorrado As Long, delrow As Long, i As Long
Dim auxp1 As Integer, auxp2 As Integer, auxp3 As Integer, auxp4 As Integer, auxp5 As Integer, auxp6 As Integer
Dim indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer

IBorrado = 0
Select Case KeyCode
  Case 86 And estado <> 0
    Exit Sub
  Case 46 And estado = 0
    If vaTabPro1.Tab = 0 Then
       swgrabadoplan = 1
    ElseIf vaTabPro1.Tab = 1 Then
       swgrabadoest = 1
    ElseIf vaTabPro1.Tab = 2 Then
       swgrabadoadj = 1
    ElseIf vaTabPro1.Tab = 3 Then
       swgrabadoadj = 1
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = vaSpread1.ActiveCol
    If vaSpread1.Col = 1 Then Exit Sub
    Select Case vaSpread1.Col
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        newcol3 = vaSpread1.Col
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        newcol3 = vaSpread1.Col - 1
    End Select
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    IOpcion = 14
    If vaSpread1.MaxRows > 100 Then
       delrow = vaSpread1.MaxRows - 100
       vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
    End If
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
    Select Case iblockcol
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol = iblockcol
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol = iblockcol - 1
    End Select
    Select Case iblockcol2
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol2 = iblockcol2 + 7
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol2 = ((iblockcol2 + 7) - 1)
    End Select
    auxcol = iblockcol: auxcol2 = iblockcol2
    auxrow = iblockrow: auxrow2 = iblockrow2
    iauxcol = aiblockcol: iauxcol2 = aiblockcol2
    iauxrow = aiblockrow: iauxrow2 = aiblockrow2
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    For i = iblockcol To iblockcol2
'' Calcular y Mover Datos Ultima linea
        If vaSpread1.MaxRows <= 100 Then
           vaSpread1.MaxRows = (vaSpread1.MaxRows + (aiblockrow2 - aiblockrow)) + 1
           For auxp1 = 101 To vaSpread1.MaxRows
               vaSpread1.Row = auxp1
               vaSpread1.RowHidden = True
           Next auxp1
        End If
        auxp6 = aiblockrow2 - aiblockrow + iblockrow
        vaSpread1.Col = iblockcol
        vaSpread1.Row = iblockrow
        vaSpread1.Col2 = iblockcol + 6
        vaSpread1.Row2 = auxp6
        vaSpread1.DestCol = iblockcol
        vaSpread1.DestRow = 100 + 1
        vaSpread1.Action = 20
''fin mover datos ultimas lineas
      
        vaSpread1.Col = iblockcol
        vaSpread1.Row = iblockrow
        vaSpread1.Col2 = iblockcol + 6
        vaSpread1.Row2 = iblockrow2
        vaSpread1.BlockMode = True
'' Limpiar Datos y Formato Celda
        vaSpread1.Action = 3
        i = iblockcol + 7
        iblockcol = iblockcol + 8
    Next i
    iblockcol = auxcol
    vaSpread1.BlockMode = False
    IndGrabadoDetalle = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(9).Visible = False
    Toolbar1.Buttons(10).Visible = True
    indactivo = 0
End Select
End Sub
Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal newrow As Long, Cancel As Boolean)
vaSpread1.Row = Row
vaSpread1.Col = Col
iblockrow = vaSpread1.ActiveRow
iblockrow2 = vaSpread1.ActiveRow
iblockcol = vaSpread1.ActiveCol
iblockcol2 = vaSpread1.ActiveCol
Select Case Col
  Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259
    vaSpread1.Col = Col
    newcol1 = vaSpread1.Col
    newrow1 = vaSpread1.ActiveRow
    WsNumPlanificacion = Val(vaSpread1.Value)
    vaSpread1.Col = newcol1 + 1
    If Val(vaSpread1.Value) <> WsNumPlanificacion Then
       vaSpread1.EditModeReplace = True
       vaSpread1.OperationMode = 0
       vaSpread1.Col = newcol1
       vaSpread1.Row = newrow1
       vaSpread1.Col = newcol1 + 1
       vaSpread1.Value = WsNumPlanificacion
       IndGrabadoDetalle = 1
       Plantilla(0).Enabled = True
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = True
       vaSpread1.Row = newrow1
    End If
End Select
End Sub
Private Sub vaspread1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
  If vaSpread1.Visible = True And estado = 0 Then
     Indgrilla1 = 0
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
End Select
End Sub
Private Sub GrabarDatosPlantillaMinutas()

Dim grilladescripcion1 As String, grilladescripcion2 As String
Dim textogrilla As String, textoplantilla As String
Dim grillacodreceta1 As Long, grillatiporeceta1 As Long, i As Long, j As Long, wsfecha As Long
Dim grillacodreceta2 As Long, grillatiporeceta2 As Long, mes As Long, planificado As Long
Dim cantplanificada As Long, correlativo As Long
Dim tiporecplant As Long, codrecplant As Long, contregdetalle As Long, indiceminutas As Long
Dim valreceta As Double

On Error GoTo Man_Error

mes = 42 * 8 + 1
TotalRegistro = 42: ContRegistro = 0: contregdetalle = 0
gauge1.Value = 0: gauge.Value = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh

fg_carga (ss)

' *** Grabar Estructura Servicio *** '

vg_db.BeginTrans
  For i = 1 To 1
      planificado = 0
      correlativo = 0
      For j = 1 To 100
          vaSpread1.Row = j
          vaSpread1.Col = i + 8
          wsfecha = Val(Mid(vaSpread1.Value, 1, 4) & Mid(vaSpread1.Value, 5, 2) & "00")
          vaSpread1.Col = i
          If vaSpread1.Text <> "" Then planificado = 6: Exit For
      Next j
      indiceminutas = 0
      Set ConSql = vg_db.Execute("select ind_minuta " & _
                   "From  Sdx_EncMinutas " & _
                   "where cod_casino='" & vg_codcasino & "' " & _
                   "and   cod_regimen=" & vg_codpventa & " " & _
                   "and   cod_servicio=" & vg_codservicio & " " & _
                   "and   dia_minuta=0 " & _
                   "and   ind_borrado=0", , adCmdText)
      If Not ConSql.EOF Then
         indiceminutas = ConSql!ind_minuta
         correlativo = ConSql!ind_minuta
         If ConSql!ind_minuta > 0 And planificado = 0 Then
            vg_db.Execute "Delete Sdx_EncMinutas from Sdx_EncMinutas " & _
                          "where cod_casino='" & vg_codcasino & "' " & _
                          "and   cod_regimen=" & vg_codpventa & " " & _
                          "and   cod_servicio=" & vg_codservicio & " " & _
                          "and   ind_minuta=" & indiceminutas & " " & _
                          "and   dia_minuta=0"
    
            vg_db.Execute "Delete Sdx_DetMinutas from Sdx_DetMinutas " & _
                          "where ind_minuta=" & indiceminutas & " " & _
                          "and   num_dia=0"
         End If
         ConSql.Close: Set ConSql = Nothing
      Else
         ConSql.Close: Set ConSql = Nothing
         If planificado > 0 Then
            Set ConSql = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=41", , adCmdText)
            If Not ConSql.EOF Then
               vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                             "Where Parametro_Num=41"
            Else
               vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                             "values (41, 'Parametro Nueva Planificación Minutas', 1)"
            End If
            ConSql.Close: Set ConSql = Nothing
   
            Set ConSql = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                         "Where Parametro_Num=41", , adCmdText)
            If Not ConSql.EOF Then
               indiceminutas = ConSql!Parametro_Val
               correlativo = ConSql!Parametro_Val
            End If
            ConSql.Close: Set ConSql = Nothing
             
            vg_db.Execute "insert into Sdx_EncMinutas (cod_casino, cod_regimen, cod_servicio, " & _
                          "ind_minuta, dia_minuta, fecha_minuta, op_minuta, ind_borrado) values " & _
                          "('" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & ", " & _
                          "" & indiceminutas & ", 0, " & Format(Date, "yyyymm") & ", '0', 0)"
         End If
      End If
'    Set ConSql = vg_db.Execute("sod_id_encplanificacionminuta '" & vg_codcasino & "', " & _
'                 "" & vg_codpventa & ", " & vg_codservicio & ", 0, " & planificado & "", , adCmdStoredProc)
'    If Not ConSql.EOF Then correlativo = ConSql!indminuta
'    ConSql.Close: Set ConSql = Nothing
      If planificado > 0 Then
         planificado = 0
         For j = 1 To 100
             grilladescripcion1 = "": grilladescripcion2 = ""
             vaSpread1.Row = j: vaSpread2.Row = j
             vaSpread1.Col = i: vaSpread2.Col = i
             grilladescripcion1 = vaSpread1.Value
             grilladescripcion2 = vaSpread2.Value
             If grilladescripcion1 <> grilladescripcion2 Then
                vaSpread2.Col = i: vaSpread1.Col = i
                vaSpread2.Value = vaSpread1.Value
                vaSpread1.Col = i
                If vaSpread1.Value <> "" Then
                   If planificado = 0 Then planificado = 5
                   textogrilla = vaSpread1.Value
                   textocomentario = vaSpread1.Value
                   tiporecplant = 6: codrecplant = 0: valreceta = 0
                   textocomentario = "txtAltComment"
                   Set ConSql = vg_db.Execute("select * " & _
                                "from Sdx_DetMinutas " & _
                                "where ind_minuta=" & correlativo & " " & _
                                "and   num_linea=" & j & " " & _
                                "and num_dia=0", , adCmdText)
                   If Not ConSql.EOF Then
                      vg_db.Execute "update Sdx_DetMinutas " & _
                                    "Set tipo_minuta=" & tiporecplant & ", " & _
                                    "cod_item=" & codrecplant & ", " & _
                                    "descripcion='" & LTrim(textogrilla) & "', " & _
                                    "ind_borrado=0 " & _
                                    "where ind_minuta=" & correlativo & " " & _
                                    "and   num_linea=" & j & " " & _
                                    "and   num_dia=0 " & _
                                    "and  (ltrim(descripcion)<>'" & LTrim(textogrilla) & "' " & _
                                    "or   tipo_minuta<>" & tiporecplant & " " & _
                                    "or   cod_item<>" & codrecplant & ")"
                   Else
                      vg_db.Execute "insert into Sdx_DetMinutas (ind_minuta, num_linea, " & _
                                     "num_dia, tipo_minuta, cod_item, descripcion, " & _
                                    "ind_borrado) values (" & correlativo & ", " & j & ", 0, " & _
                                    "" & tiporecplant & ", " & codrecplant & ", '" & textogrilla & "', 0)"
                   End If
                   ConSql.Close: Set ConSql = Nothing
'                  vg_db.Execute "sod_iu_detplanificacionminuta " & correlativo & ", " & j & ", 0, " & tiporecplant & ", " & _
'                                 "" & codrecplant & ", '" & textogrilla & "'"
                Else
                   vg_db.Execute "Delete Sdx_DetMinutas from Sdx_DetMinutas " & _
                                 "where ind_minuta=" & correlativo & " " & _
                                 "and   num_linea=" & j & " " & _
                                 "and   num_dia=0"
'                     vg_db.Execute "sod_d_detplanificacionminuta " & correlativo & ", " & j & ", 0"
                End If
             Else
                vaSpread1.Col = i
                If vaSpread1.Value <> "" Then If planificado = 0 Then planificado = 5
             End If
             vaSpread2.Col = i
             vaSpread1.Col = i
             vaSpread2.Value = vaSpread1.Value
         Next j
      Else
         For j = 1 To 100
             vaSpread1.Row = j
             vaSpread2.Row = j
             vaSpread2.Col = isodexhoch
             vaSpread1.Col = i
             vaSpread2.Value = vaSpread1.Value
         Next j
      End If
  Next i

' Fin Grabar Estructura Servicio *** '

  inddia = 1
  For i = 2 To mes Step 8
      ContRegistro = ContRegistro + 1
      gauge1.Value = Val((ContRegistro / TotalRegistro) * 100)
      Label3.Caption = ""
      Label3.Caption = "Día : " & inddia
      planificado = 0
      correlativo = 0
      For j = 1 To 100
          vaSpread1.Row = j
          vaSpread1.Col = i + 7
          wsfecha = Val(vaSpread1.Value)
          vaSpread1.Col = i + 1
          If vaSpread1.Text <> "" Then planificado = 6: Exit For
      Next j
      indiceminutas = 0
      Set ConSql = vg_db.Execute("select ind_minuta " & _
                   "From  Sdx_EncMinutas " & _
                   "where cod_casino='" & vg_codcasino & "' " & _
                   "and   cod_regimen=" & vg_codpventa & " " & _
                   "and   cod_servicio=" & vg_codservicio & " " & _
                   "and   dia_minuta=" & inddia & " " & _
                   "and   ind_borrado=0", , adCmdText)
      If Not ConSql.EOF Then
         indiceminutas = ConSql!ind_minuta
         correlativo = ConSql!ind_minuta
         If ConSql!ind_minuta > 0 And planificado = 0 Then
            vg_db.Execute "Delete Sdx_EncMinutas from Sdx_EncMinutas " & _
                          "where cod_casino='" & vg_codcasino & "' " & _
                          "and   cod_regimen=" & vg_codpventa & " " & _
                          "and   cod_servicio=" & vg_codservicio & " " & _
                          "and   ind_minuta=" & indiceminutas & " " & _
                          "and   dia_minuta=" & inddia & ""
    
            vg_db.Execute "Delete Sdx_DetMinutas from Sdx_DetMinutas " & _
                          "where ind_minuta=" & indiceminutas & " " & _
                          "and   num_dia=" & inddia & ""
         End If
         ConSql.Close: Set ConSql = Nothing
      Else
         ConSql.Close: Set ConSql = Nothing
         If planificado > 0 Then
            Set ConSql = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=41", , adCmdText)
            If Not ConSql.EOF Then
               vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                             "Where Parametro_Num=41"
            Else
               vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                             "values (41, 'Parametro Nueva Planificación Minutas', 1)"
            End If
            ConSql.Close: Set ConSql = Nothing
   
            Set ConSql = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                         "Where Parametro_Num=41", , adCmdText)
            If Not ConSql.EOF Then
               indiceminutas = ConSql!Parametro_Val
               correlativo = ConSql!Parametro_Val
            End If
            ConSql.Close: Set ConSql = Nothing
             
            vg_db.Execute "insert into Sdx_EncMinutas (cod_casino, cod_regimen, cod_servicio, " & _
                          "ind_minuta, dia_minuta, fecha_minuta, op_minuta, ind_borrado) values " & _
                          "('" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & ", " & _
                          "" & indiceminutas & ", " & inddia & ", " & Format(Date, "yyyymm") & ", '0', 0)"
         End If
      End If
'      Set ConSql = vg_db.Execute("sod_id_encplanificacionminuta '" & vg_codcasino & "', " & _
'                   "" & vg_codpventa & ", " & vg_codservicio & ", " & inddia & ", " & planificado & "", , adCmdStoredProc)
'      If Not ConSql.EOF Then correlativo = ConSql!indminuta
'      ConSql.Close: Set ConSql = Nothing
      gauge.Value = 0: contregdetalle = 0
      If planificado > 0 Then
         planificado = 0
         For j = 1 To 100
             contregdetalle = contregdetalle + 1
             gauge.Value = Val((contregdetalle / 100) * 100)
             grilladescripcion1 = "": grillacodreceta1 = 0: grillatiporeceta1 = 0
             grilladescripcion2 = "": grillacodreceta2 = 0: grillatiporeceta2 = 0
             vaSpread1.Row = j: vaSpread2.Row = j
             vaSpread1.Col = i + 1: vaSpread2.Col = i + 1
             grilladescripcion1 = vaSpread1.Value
             grilladescripcion2 = vaSpread2.Value
             vaSpread1.Col = i + 4: vaSpread2.Col = i + 4
             grillacodreceta1 = Val(vaSpread1.Value)
             grillacodreceta2 = Val(vaSpread2.Value)
             vaSpread1.Col = i + 5: vaSpread2.Col = i + 5
             grillatiporeceta1 = Val(vaSpread1.Value)
             grillatiporeceta2 = Val(vaSpread2.Value)
             If grilladescripcion1 <> grilladescripcion2 Or grillacodreceta1 <> grillacodreceta2 Or grillatiporeceta1 <> grillatiporeceta2 Then
                vaSpread2.Col = i + 1: vaSpread1.Col = i + 1
                vaSpread2.Value = vaSpread1.Value
                vaSpread2.Col = i + 2: vaSpread1.Col = i + 2
                vaSpread2.Value = vaSpread1.Value
                vaSpread2.Col = i + 4: vaSpread1.Col = i + 4
                vaSpread2.Value = vaSpread1.Value
                vaSpread2.Col = i + 5: vaSpread1.Col = i + 5
                vaSpread2.Value = vaSpread1.Value
                vaSpread2.Col = i + 6: vaSpread1.Col = i + 6
                vaSpread2.Value = vaSpread1.Value
                vaSpread1.Col = i + 1
                If vaSpread1.Value <> "" Then
                   If planificado = 0 Then planificado = 5
                   textogrilla = vaSpread1.Value
                   textocomentario = vaSpread1.Value
                   vaSpread1.Col = i + 2
                   cantplanificada = Val(vaSpread1.Value)
                   If cantplanificada > 0 Then planificado = 3
                   vaSpread1.Col = i + 4
                   tiporecplant = Val(vaSpread1.Value)
                   If tiporecplant = 6 Then textocomentario = "txtAltComment"
                   vaSpread1.Col = i + 5
                   codrecplant = Val(vaSpread1.Value)
                   vaSpread1.Col = i + 6
                   valreceta = Val(vaSpread1.Value)
                   Set ConSql = vg_db.Execute("select * " & _
                                "from Sdx_DetMinutas " & _
                                "where ind_minuta=" & correlativo & " " & _
                                "and   num_linea=" & j & " " & _
                                "and num_dia=" & inddia & "", , adCmdText)
                   If Not ConSql.EOF Then
                      vg_db.Execute "update Sdx_DetMinutas " & _
                                    "Set tipo_minuta=" & tiporecplant & ", " & _
                                    "cod_item=" & codrecplant & ", " & _
                                    "descripcion='" & LTrim(textogrilla) & "', " & _
                                    "ind_borrado=0 " & _
                                    "where ind_minuta=" & correlativo & " " & _
                                    "and   num_linea=" & j & " " & _
                                    "and   num_dia=" & inddia & " " & _
                                    "and  (ltrim(descripcion)<>'" & LTrim(textogrilla) & "' " & _
                                    "or   tipo_minuta<>" & tiporecplant & " " & _
                                    "or   cod_item<>" & codrecplant & ")"
                   Else
                      vg_db.Execute "insert into Sdx_DetMinutas (ind_minuta, num_linea, " & _
                                     "num_dia, tipo_minuta, cod_item, descripcion, " & _
                                    "ind_borrado) values (" & correlativo & ", " & j & ", " & inddia & ", " & _
                                    "" & tiporecplant & ", " & codrecplant & ", '" & textogrilla & "', 0)"
                   End If
                   ConSql.Close: Set ConSql = Nothing
'                   vg_db.Execute "sod_iu_detplanificacionminuta " & correlativo & ", " & j & ", " & inddia & ", " & tiporecplant & ", " & _
'                                 "" & codrecplant & ", '" & textogrilla & "'"
                Else
                   vg_db.Execute "Delete Sdx_DetMinutas from Sdx_DetMinutas " & _
                                 "where ind_minuta=" & correlativo & " " & _
                                 "and   num_linea=" & j & " " & _
                                 "and   num_dia=" & inddia & ""
'                   vg_db.Execute "sod_d_detplanificacionminuta " & correlativo & ", " & j & ", " & inddia & ""
                End If
             Else
                vaSpread1.Col = i + 1
                If vaSpread1.Value <> "" Then If planificado = 0 Then planificado = 5
             End If
             vaSpread2.Col = i + 1: vaSpread1.Col = i + 1
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 2: vaSpread1.Col = i + 2
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 4: vaSpread1.Col = i + 4
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 5: vaSpread1.Col = i + 5
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 6: vaSpread1.Col = i + 6
             vaSpread2.Value = vaSpread1.Value
       
         Next j
      Else
         For j = 1 To 100
             vaSpread1.Row = j: vaSpread2.Row = j
             vaSpread2.Col = i + 1: vaSpread1.Col = i + 1
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 2: vaSpread1.Col = i + 2
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 4: vaSpread1.Col = i + 4
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 5: vaSpread1.Col = i + 5
             vaSpread2.Value = vaSpread1.Value
             vaSpread2.Col = i + 6: vaSpread1.Col = i + 6
             vaSpread2.Value = vaSpread1.Value
         Next j
      End If
      inddia = inddia + 1
  Next i
vg_db.CommitTrans
Picture1.Visible = False: gauge.Visible = False
vaSpread1.Refresh
fg_descarga

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub GrabarDatosEstructuraFija()
Dim grilladescripcion1 As String, grilladescripcion2 As String
Dim textogrilla As String, textoplantilla As String
Dim grillacodreceta1 As Long, grillatiporeceta1 As Long, i As Long, j As Long, wsfecha As Long
Dim grillacodreceta2 As Long, grillatiporeceta2 As Long, mes As Long, planificado As Long
Dim cantplanificada As Long, correlativo As Long
Dim tiporecplant As Long, codrecplant As Long, contregdetalle As Long, indiceestfija As Long
Dim valreceta As Double

On Error GoTo Man_Error

mes = 7 * 8 + 1
TotalRegistro = 7: ContRegistro = 0: contregdetalle = 0
gauge1.Value = 0: gauge.Value = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh

fg_carga (ss)

' *** Grabar Estructura Servicio *** '

vg_db.BeginTrans
  For i = 1 To 1
      planificado = 0
      correlativo = 0
      For j = 1 To 100
          vaSpread3.Row = j
          vaSpread3.Col = i + 8
          wsfecha = Val(Mid(vaSpread3.Value, 1, 4) & Mid(vaSpread3.Value, 5, 2) & "00")
          vaSpread3.Col = i
          If vaSpread3.Text <> "" Then planificado = 6: Exit For
      Next j
      indiceestfija = 0
      Set ConSql = vg_db.Execute("select ind_estfija " & _
                   "From  Sdx_EncEstructuraFija " & _
                   "where cod_casino='" & vg_codcasino & "' " & _
                   "and   cod_regimen=" & vg_codpventa & " " & _
                   "and   cod_servicio=" & vg_codservicio & " " & _
                   "and   dia_estfija=0 " & _
                   "and   ind_borrado=0", , adCmdText)
      If Not ConSql.EOF Then
         indiceestfija = ConSql!ind_estfija
         correlativo = ConSql!ind_estfija
         If ConSql!ind_estfija > 0 And planificado = 0 Then
            vg_db.Execute "Delete Sdx_EncEstructuraFija from Sdx_EncEstructuraFija " & _
                          "where cod_casino='" & vg_codcasino & "' " & _
                          "and   cod_regimen=" & vg_codpventa & " " & _
                          "and   cod_servicio=" & vg_codservicio & " " & _
                          "and   ind_estfija=" & indiceestfija & " " & _
                          "and   dia_estfija=0"
    
            vg_db.Execute "Delete Sdx_DetEstructuraFija from Sdx_DetEstructuraFija " & _
                          "where ind_estfija=" & indiceestfija & " " & _
                          "and   num_dia=0"
         End If
         ConSql.Close: Set ConSql = Nothing
      Else
         ConSql.Close: Set ConSql = Nothing
         If planificado > 0 Then
            Set ConSql = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=42", , adCmdText)
            If Not ConSql.EOF Then
               vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                             "Where Parametro_Num=42"
            Else
               vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                             "values (42, 'Parametro Estructura Fijas', 1)"
            End If
            ConSql.Close: Set ConSql = Nothing
   
            Set ConSql = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                         "Where Parametro_Num=42", , adCmdText)
            If Not ConSql.EOF Then
               indiceestfija = ConSql!Parametro_Val
               correlativo = ConSql!Parametro_Val
            End If
            ConSql.Close: Set ConSql = Nothing
             
            vg_db.Execute "insert into Sdx_EncEstructuraFija (cod_casino, cod_regimen, cod_servicio, " & _
                          "ind_estfija, dia_estfija, fecha_estfija, op_estfija, ind_borrado) values " & _
                          "('" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & ", " & _
                          "" & indiceestfija & ", 0, " & Format(Date, "yyyymm") & ", '0', 0)"
         End If
      End If
'    Set ConSql = vg_db.Execute("sod_id_encplanificacionminuta '" & vg_codcasino & "', " & _
'                 "" & vg_codpventa & ", " & vg_codservicio & ", 0, " & planificado & "", , adCmdStoredProc)
'    If Not ConSql.EOF Then correlativo = ConSql!indminuta
'    ConSql.Close: Set ConSql = Nothing
      If planificado > 0 Then
         planificado = 0
         For j = 1 To 100
             grilladescripcion1 = "": grilladescripcion2 = ""
             vaSpread3.Row = j: vaSpread4.Row = j
             vaSpread3.Col = i: vaSpread4.Col = i
             grilladescripcion1 = vaSpread3.Value
             grilladescripcion2 = vaSpread4.Value
             If grilladescripcion1 <> grilladescripcion2 Then
                vaSpread4.Col = i: vaSpread3.Col = i
                vaSpread4.Value = vaSpread3.Value
                vaSpread3.Col = i
                If vaSpread3.Value <> "" Then
                   If planificado = 0 Then planificado = 5
                   textogrilla = vaSpread3.Value
                   textocomentario = vaSpread3.Value
                   tiporecplant = 6: codrecplant = 0: valreceta = 0
                   textocomentario = "txtAltComment"
                   Set ConSql = vg_db.Execute("select * " & _
                                "from Sdx_DetEstructuraFija " & _
                                "where ind_estfija=" & correlativo & " " & _
                                "and   num_linea=" & j & " " & _
                                "and num_dia=0", , adCmdText)
                   If Not ConSql.EOF Then
                      vg_db.Execute "update Sdx_DetEstructuraFija " & _
                                    "Set tipo_estfija=" & tiporecplant & ", " & _
                                    "cod_item=" & codrecplant & ", " & _
                                    "descripcion='" & LTrim(textogrilla) & "', " & _
                                    "ind_borrado=0 " & _
                                    "where ind_estfija=" & correlativo & " " & _
                                    "and   num_linea=" & j & " " & _
                                    "and   num_dia=0 " & _
                                    "and  (ltrim(descripcion)<>'" & LTrim(textogrilla) & "' " & _
                                    "or   tipo_estfija<>" & tiporecplant & " " & _
                                    "or   cod_item<>" & codrecplant & ")"
                   Else
                      vg_db.Execute "insert into Sdx_DetEstructuraFija (ind_estfija, num_linea, " & _
                                     "num_dia, tipo_estfija, cod_item, descripcion, " & _
                                    "ind_borrado) values (" & correlativo & ", " & j & ", 0, " & _
                                    "" & tiporecplant & ", " & codrecplant & ", '" & textogrilla & "', 0)"
                   End If
                   ConSql.Close: Set ConSql = Nothing
'                  vg_db.Execute "sod_iu_detplanificacionminuta " & correlativo & ", " & j & ", 0, " & tiporecplant & ", " & _
'                                 "" & codrecplant & ", '" & textogrilla & "'"
                Else
                   vg_db.Execute "Delete Sdx_DetEstructuraFija from Sdx_DetEstructuraFija " & _
                                 "where ind_estfija=" & correlativo & " " & _
                                 "and   num_linea=" & j & " " & _
                                 "and   num_dia=0"
'                     vg_db.Execute "sod_d_detplanificacionminuta " & correlativo & ", " & j & ", 0"
                End If
             Else
                vaSpread3.Col = i
                If vaSpread3.Value <> "" Then If planificado = 0 Then planificado = 5
             End If
             vaSpread4.Col = i: vaSpread3.Col = i
             vaSpread4.Value = vaSpread3.Value
         Next j
      Else
         For j = 1 To 100
             vaSpread3.Row = j: vaSpread4.Row = j
             vaSpread4.Col = isodexhoch
             vaSpread3.Col = i
             vaSpread4.Value = vaSpread3.Value
         Next j
      End If
  Next i

' *** Grabar Estructura Servicio *** '

  inddia = 1
  For i = 2 To mes Step 8
      ContRegistro = ContRegistro + 1
      gauge1.Value = Val((ContRegistro / TotalRegistro) * 100)
      Label3.Caption = ""
      Label3.Caption = "Día : " & inddia
      planificado = 0
      correlativo = 0
      For j = 1 To 100
          vaSpread3.Row = j
          vaSpread3.Col = i + 7
          wsfecha = Val(vaSpread3.Value)
          vaSpread3.Col = i + 1
          If vaSpread3.Text <> "" Then planificado = 6: Exit For
      Next j
      indiceestfija = 0
      Set ConSql = vg_db.Execute("select ind_estfija " & _
                   "From  Sdx_EncEstructuraFija " & _
                   "where cod_casino='" & vg_codcasino & "' " & _
                   "and   cod_regimen=" & vg_codpventa & " " & _
                   "and   cod_servicio=" & vg_codservicio & " " & _
                   "and   dia_estfija=" & inddia & " " & _
                   "and   ind_borrado=0", , adCmdText)
      If Not ConSql.EOF Then
         indiceestfija = ConSql!ind_estfija
         correlativo = ConSql!ind_estfija
         If ConSql!ind_estfija > 0 And planificado = 0 Then
            
            vg_db.Execute "Delete Sdx_EncEstructuraFija from Sdx_EncEstructuraFija " & _
                          "where cod_casino='" & vg_codcasino & "' " & _
                          "and   cod_regimen=" & vg_codpventa & " " & _
                          "and   cod_servicio=" & vg_codservicio & " " & _
                          "and   ind_estfija=" & indiceestfija & " " & _
                          "and   dia_estfija=" & inddia & ""
    
            vg_db.Execute "Delete Sdx_DetEstructuraFija from Sdx_DetEstructuraFija " & _
                          "where ind_estfija=" & indiceestfija & " " & _
                          "and   num_dia=" & inddia & ""
         End If
         ConSql.Close: Set ConSql = Nothing
      Else
         ConSql.Close: Set ConSql = Nothing
         If planificado > 0 Then
            Set ConSql = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=42", , adCmdText)
            If Not ConSql.EOF Then
               vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                             "Where Parametro_Num=42"
            Else
               vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                             "values (42, 'Parametro Estructura Fijas', 1)"
            End If
            ConSql.Close: Set ConSql = Nothing
   
            Set ConSql = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                         "Where Parametro_Num=42", , adCmdText)
            If Not ConSql.EOF Then
               indiceestfija = ConSql!Parametro_Val
               correlativo = ConSql!Parametro_Val
            End If
            ConSql.Close: Set ConSql = Nothing
             
            vg_db.Execute "insert into Sdx_EncEstructuraFija (cod_casino, cod_regimen, cod_servicio, " & _
                          "ind_estfija, dia_estfija, fecha_estfija, op_estfija, ind_borrado) values " & _
                          "('" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & ", " & _
                          "" & indiceestfija & ", " & inddia & ", " & Format(Date, "yyyymm") & ", '0', 0)"
         End If
      End If
'      Set ConSql = vg_db.Execute("sod_id_encplanificacionminuta '" & vg_codcasino & "', " & _
'                   "" & vg_codpventa & ", " & vg_codservicio & ", " & inddia & ", " & planificado & "", , adCmdStoredProc)
'      If Not ConSql.EOF Then correlativo = ConSql!indestfija
'      ConSql.Close: Set ConSql = Nothing
      gauge.Value = 0: contregdetalle = 0
      If planificado > 0 Then
         planificado = 0
         For j = 1 To 100
             contregdetalle = contregdetalle + 1
             gauge.Value = Val((contregdetalle / 100) * 100)
             grilladescripcion1 = "": grillacodreceta1 = 0: grillatiporeceta1 = 0
             grilladescripcion2 = "": grillacodreceta2 = 0: grillatiporeceta2 = 0
             vaSpread3.Row = j: vaSpread4.Row = j
             vaSpread3.Col = i + 1: vaSpread4.Col = i + 1
             grilladescripcion1 = vaSpread3.Value
             grilladescripcion2 = vaSpread4.Value
             vaSpread3.Col = i + 4: vaSpread4.Col = i + 4
             grillacodreceta1 = Val(vaSpread3.Value)
             grillacodreceta2 = Val(vaSpread4.Value)
             vaSpread3.Col = i + 5: vaSpread4.Col = i + 5
             grillatiporeceta1 = Val(vaSpread3.Value)
             grillatiporeceta2 = Val(vaSpread4.Value)
             If grilladescripcion1 <> grilladescripcion2 Or grillacodreceta1 <> grillacodreceta2 Or grillatiporeceta1 <> grillatiporeceta2 Then
                vaSpread4.Col = i + 1: vaSpread3.Col = i + 1
                vaSpread4.Value = vaSpread3.Value
                vaSpread4.Col = i + 2: vaSpread3.Col = i + 2
                vaSpread4.Value = vaSpread3.Value
                vaSpread4.Col = i + 4: vaSpread3.Col = i + 4
                vaSpread4.Value = vaSpread3.Value
                vaSpread4.Col = i + 5: vaSpread3.Col = i + 5
                vaSpread4.Value = vaSpread3.Value
                vaSpread4.Col = i + 6: vaSpread3.Col = i + 6
                vaSpread4.Value = vaSpread3.Value
                vaSpread3.Col = i + 1
                If vaSpread3.Value <> "" Then
                   If planificado = 0 Then planificado = 5
                   textogrilla = vaSpread3.Value
                   textocomentario = vaSpread3.Value
                   vaSpread3.Col = i + 2
                   cantplanificada = Val(vaSpread3.Value)
                   If cantplanificada > 0 Then planificado = 3
                   vaSpread3.Col = i + 4
                   tiporecplant = Val(vaSpread3.Value)
                   If tiporecplant = 6 Then textocomentario = "txtAltComment"
                   vaSpread3.Col = i + 5
                   codrecplant = Val(vaSpread3.Value)
                   vaSpread3.Col = i + 6
                   valreceta = Val(vaSpread3.Value)
                   Set ConSql = vg_db.Execute("select * " & _
                                "from Sdx_DetEstructuraFija " & _
                                "where ind_estfija=" & correlativo & " " & _
                                "and   num_linea=" & j & " " & _
                                "and num_dia=" & inddia & "", , adCmdText)
                   If Not ConSql.EOF Then
                      vg_db.Execute "update Sdx_DetEstructuraFija " & _
                                    "Set tipo_estfija=" & tiporecplant & ", " & _
                                    "cod_item=" & codrecplant & ", " & _
                                    "descripcion='" & LTrim(textogrilla) & "', " & _
                                    "ind_borrado=0 " & _
                                    "where ind_estfija=" & correlativo & " " & _
                                    "and   num_linea=" & j & " " & _
                                    "and   num_dia=" & inddia & " " & _
                                    "and  (ltrim(descripcion)<>'" & LTrim(textogrilla) & "' " & _
                                    "or   tipo_estfija<>" & tiporecplant & " " & _
                                    "or   cod_item<>" & codrecplant & ")"
                   Else
                      vg_db.Execute "insert into Sdx_DetEstructuraFija (ind_estfija, num_linea, " & _
                                     "num_dia, tipo_estfija, cod_item, descripcion, " & _
                                    "ind_borrado) values (" & correlativo & ", " & j & ", " & inddia & ", " & _
                                    "" & tiporecplant & ", " & codrecplant & ", '" & textogrilla & "', 0)"
                   End If
                   ConSql.Close: Set ConSql = Nothing
'                   vg_db.Execute "sod_iu_detplanificacionminuta " & correlativo & ", " & j & ", " & inddia & ", " & tiporecplant & ", " & _
'                                 "" & codrecplant & ", '" & textogrilla & "'"
                Else
                   vg_db.Execute "Delete Sdx_DetEstructuraFija from Sdx_DetEstructuraFija " & _
                                 "where ind_estfija=" & correlativo & " " & _
                                 "and   num_linea=" & j & " " & _
                                 "and   num_dia=" & inddia & ""
'                   vg_db.Execute "sod_d_detplanificacionminuta " & correlativo & ", " & j & ", " & inddia & ""
                End If
             Else
                vaSpread3.Col = i + 1
                If vaSpread3.Value <> "" Then If planificado = 0 Then planificado = 5
             End If
             vaSpread4.Col = i + 1: vaSpread3.Col = i + 1
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 2: vaSpread3.Col = i + 2
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 4: vaSpread3.Col = i + 4
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 5: vaSpread3.Col = i + 5
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 6: vaSpread3.Col = i + 6
             vaSpread4.Value = vaSpread3.Value
       
         Next j
      Else
         For j = 1 To 100
             vaSpread3.Row = j: vaSpread4.Row = j
             vaSpread4.Col = i + 1: vaSpread3.Col = i + 1
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 2: vaSpread3.Col = i + 2
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 4: vaSpread3.Col = i + 4
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 5: vaSpread3.Col = i + 5
             vaSpread4.Value = vaSpread3.Value
             vaSpread4.Col = i + 6: vaSpread3.Col = i + 6
             vaSpread4.Value = vaSpread3.Value
         Next j
      End If
      inddia = inddia + 1
  Next i
vg_db.CommitTrans
Picture1.Visible = False: gauge.Visible = False
vaSpread3.Refresh
fg_descarga

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub GrabarDatosAdjuntos()

Dim grilladescripcion1 As String, grilladescripcion2 As String
Dim textogrilla As String, textoplantilla As String
Dim grillacodreceta1 As Long, grillatiporeceta1 As Long, i As Long, j As Long, wsfecha As Long
Dim grillacodreceta2 As Long, grillatiporeceta2 As Long, mes As Long, planificado As Long
Dim cantplanificada As Long, correlativo As Long
Dim tiporecplant As Long, codrecplant As Long, contregdetalle As Long, indicedatadj As Long
Dim valreceta As Double

On Error GoTo Man_Error

mes = 7 * 8 + 1
TotalRegistro = 7: ContRegistro = 0: contregdetalle = 0
gauge1.Value = 0: gauge.Value = 0
Picture1.Visible = True: Label3.Visible = True: gauge.Visible = True
Picture1.Refresh: Label3.Refresh: gauge.Refresh: gauge1.Refresh

fg_carga (ss)

' *** Grabar Datos Adjuntos *** '

vg_db.BeginTrans
  For i = 2 To mes Step 8
      ContRegistro = ContRegistro + 1
      gauge1.Value = Val((ContRegistro / TotalRegistro) * 100)
      Label3.Caption = ""
      Label3.Caption = "Día : " & inddia
      planificado = 0
      correlativo = 0
      For j = 1 To 100
          grilla1.Row = j
          grilla1.Col = i + 7
          wsfecha = Val(grilla1.Value)
          grilla1.Col = i + 1
          If grilla1.Text <> "" Then planificado = 6: Exit For
      Next j
      indicedatadj = 0
      Set ConSql = vg_db.Execute("select ind_datadj " & _
                   "From  Sdx_EncDatosAdjuntos " & _
                   "where cod_casino='" & vg_codcasino & "' " & _
                   "and   cod_regimen=" & vg_codpventa & " " & _
                   "and   cod_servicio=" & vg_codservicio & " " & _
                   "and   dia_datadj=" & inddia & " " & _
                   "and   tipo_datadj='" & tipodatadj & "' " & _
                   "and   ind_borrado=0", , adCmdText)
      If Not ConSql.EOF Then
         indicedatadj = ConSql!ind_datadj
         correlativo = ConSql!ind_datadj
         If ConSql!ind_datadj > 0 And planificado = 0 Then
            
            vg_db.Execute "Delete Sdx_EncDatosAdjuntos from Sdx_EncDatosAdjuntos " & _
                          "where cod_casino='" & vg_codcasino & "' " & _
                          "and   cod_regimen=" & vg_codpventa & " " & _
                          "and   cod_servicio=" & vg_codservicio & " " & _
                          "and   ind_datadj=" & indicedatadj & " " & _
                          "and   tipo_datadj='" & tipodatadj & "' " & _
                          "and   dia_datadj=" & inddia & ""
    
            vg_db.Execute "Delete Sdx_DetDatosAdjuntos from Sdx_DetDatosAdjuntos " & _
                          "where  ind_datadj=" & indicedatadj & " " & _
                          "and    num_dia=" & inddia & ""
         End If
         ConSql.Close: Set ConSql = Nothing
      Else
         ConSql.Close: Set ConSql = Nothing
         If planificado > 0 Then
            Set ConSql = vg_db.Execute("select * from Sdx_Parametro holdlock where Parametro_Num=43", , adCmdText)
            If Not ConSql.EOF Then
               vg_db.Execute "Update Sdx_Parametro Set Parametro_Val = Parametro_Val + 1 " & _
                             "Where Parametro_Num=43"
            Else
               vg_db.Execute "insert into Sdx_Parametro (Parametro_Num, Parametro_Desc, Parametro_Val) " & _
                             "values (43, 'Parametro Datos Adjuntos', 1)"
            End If
            ConSql.Close: Set ConSql = Nothing
   
            Set ConSql = vg_db.Execute("select Parametro_Val From Sdx_Parametro " & _
                         "Where Parametro_Num=43", , adCmdText)
            If Not ConSql.EOF Then
               indicedatadj = ConSql!Parametro_Val
               correlativo = ConSql!Parametro_Val
            End If
            ConSql.Close: Set ConSql = Nothing
             
            vg_db.Execute "insert into Sdx_EncDatosAdjuntos (cod_casino, cod_regimen, cod_servicio, " & _
                          "ind_datadj, dia_datadj, tipo_datadj, fecha_datadj, op_datadj, ind_borrado) values " & _
                          "('" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & ", " & _
                          "" & indicedatadj & ", " & inddia & ", '" & tipodatadj & "', " & Format(Date, "yyyymm") & ", '0', 0)"
         End If
      End If
'      Set ConSql = vg_db.Execute("sod_id_encplanificacionminuta '" & vg_codcasino & "', " & _
'                   "" & vg_codpventa & ", " & vg_codservicio & ", " & inddia & ", " & planificado & "", , adCmdStoredProc)
'      If Not ConSql.EOF Then correlativo = ConSql!inddatadj
'      ConSql.Close: Set ConSql = Nothing
      gauge.Value = 0: contregdetalle = 0
      If planificado > 0 Then
         planificado = 0
         For j = 1 To 100
             contregdetalle = contregdetalle + 1
             gauge.Value = Val((contregdetalle / 100) * 100)
             grilladescripcion1 = "": grillacodreceta1 = 0: grillatiporeceta1 = 0
             grilladescripcion2 = "": grillacodreceta2 = 0: grillatiporeceta2 = 0
             grilla1.Row = j: grilla2.Row = j
             grilla1.Col = i + 1: grilla2.Col = i + 1
             grilladescripcion1 = grilla1.Value
             grilladescripcion2 = grilla2.Value
             grilla1.Col = i + 4: grilla2.Col = i + 4
             grillacodreceta1 = Val(grilla1.Value)
             grillacodreceta2 = Val(grilla2.Value)
             grilla1.Col = i + 5: grilla2.Col = i + 5
             grillatiporeceta1 = Val(grilla1.Value)
             grillatiporeceta2 = Val(grilla2.Value)
             If grilladescripcion1 <> grilladescripcion2 Or grillacodreceta1 <> grillacodreceta2 Or grillatiporeceta1 <> grillatiporeceta2 Then
                grilla2.Col = i + 1: grilla1.Col = i + 1
                grilla2.Value = grilla1.Value
                grilla2.Col = i + 2: grilla1.Col = i + 2
                grilla2.Value = grilla1.Value
                grilla2.Col = i + 4: grilla1.Col = i + 4
                grilla2.Value = grilla1.Value
                grilla2.Col = i + 5: grilla1.Col = i + 5
                grilla2.Value = grilla1.Value
                grilla2.Col = i + 6: grilla1.Col = i + 6
                grilla2.Value = grilla1.Value
                grilla1.Col = i + 1
                If grilla1.Value <> "" Then
                   If planificado = 0 Then planificado = 5
                   textogrilla = grilla1.Value
                   textocomentario = grilla1.Value
                   grilla1.Col = i + 2
                   cantplanificada = Val(grilla1.Value)
                   If cantplanificada > 0 Then planificado = 3
                   grilla1.Col = i + 4
                   tiporecplant = Val(grilla1.Value)
                   If tiporecplant = 6 Then textocomentario = "txtAltComment"
                   grilla1.Col = i + 5
                   codrecplant = Val(grilla1.Value)
                   grilla1.Col = i + 6
                   valreceta = Val(grilla1.Value)
                   Set ConSql = vg_db.Execute("select * " & _
                                "from Sdx_DetDatosAdjuntos " & _
                                "where ind_datadj=" & correlativo & " " & _
                                "and   num_linea=" & j & " " & _
                                "and num_dia=" & inddia & "", , adCmdText)
                   If Not ConSql.EOF Then
                      vg_db.Execute "update Sdx_DetDatosAdjuntos " & _
                                    "Set tipo_datadj=" & tiporecplant & ", " & _
                                    "cod_item=" & codrecplant & ", " & _
                                    "descripcion='" & LTrim(textogrilla) & "', " & _
                                    "ind_borrado=0 " & _
                                    "where ind_datadj=" & correlativo & " " & _
                                    "and   num_linea=" & j & " " & _
                                    "and   num_dia=" & inddia & " " & _
                                    "and  (ltrim(descripcion)<>'" & LTrim(textogrilla) & "' " & _
                                    "or   tipo_datadj<>" & tiporecplant & " " & _
                                    "or   cod_item<>" & codrecplant & ")"
                   Else
                      vg_db.Execute "insert into Sdx_DetDatosAdjuntos (ind_datadj, num_linea, " & _
                                     "num_dia, tipo_datadj, cod_item, descripcion, " & _
                                    "ind_borrado) values (" & correlativo & ", " & j & ", " & inddia & ", " & _
                                    "" & tiporecplant & ", " & codrecplant & ", '" & textogrilla & "', 0)"
                   End If
                   ConSql.Close: Set ConSql = Nothing
'                   vg_db.Execute "sod_iu_detplanificacionminuta " & correlativo & ", " & j & ", " & inddia & ", " & tiporecplant & ", " & _
'                                 "" & codrecplant & ", '" & textogrilla & "'"
                Else
                   vg_db.Execute "Delete Sdx_DetDatosAdjuntos from Sdx_DetDatosAdjuntos " & _
                                 "where ind_datadj=" & correlativo & " " & _
                                 "and   num_linea=" & j & " " & _
                                 "and   num_dia=" & inddia & ""
'                   vg_db.Execute "sod_d_detplanificacionminuta " & correlativo & ", " & j & ", " & inddia & ""
                End If
             Else
                grilla1.Col = i + 1
                If grilla1.Value <> "" Then If planificado = 0 Then planificado = 5
             End If
             grilla2.Col = i + 1: grilla1.Col = i + 1
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 2: grilla1.Col = i + 2
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 4: grilla1.Col = i + 4
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 5: grilla1.Col = i + 5
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 6: grilla1.Col = i + 6
             grilla2.Value = grilla1.Value
       
         Next j
      Else
         For j = 1 To 100
             grilla1.Row = j: grilla2.Row = j
             grilla2.Col = i + 1: grilla1.Col = i + 1
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 2: grilla1.Col = i + 2
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 4: grilla1.Col = i + 4
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 5: grilla1.Col = i + 5
             grilla2.Value = grilla1.Value
             grilla2.Col = i + 6: grilla1.Col = i + 6
             grilla2.Value = grilla1.Value
         Next j
      End If
      inddia = inddia + 1
  Next i
vg_db.CommitTrans
Picture1.Visible = False: gauge.Visible = False
grilla1.Refresh
               
fg_descarga

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
vg_db.Rollback
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
Private Sub MoverVecDia()
Dim X As Long, i As Long, j As Long, indi As Long, indj As Long
Dim fecha As String
Dim mes As Long
grilla1.MaxCols = 0
grilla1.MaxCols = 8 * maxcolumna + 1
grilla1.Row = 0
grilla2.MaxCols = 0
grilla2.MaxCols = 8 * maxcolumna + 1
j = 3: indi = 1: indj = 1
For i = 1 To maxcolumna
    
    grilla1.Col = j - 2
    grilla1.ColsFrozen = 1
    grilla1.VisibleCols = 1
    If maxcolumna = 42 Then
       grilla1.ColWidth(j - 2) = 15
    ElseIf maxcolumna = 1 Then
       grilla1.ColWidth(j - 2) = 20
    ElseIf maxcolumna = 2 Then
       grilla1.ColHidden = True
    End If
    grilla1.Text = "Estructura Servicio"
    
    grilla1.Col = j - 1
    grilla1.ColWidth(j - 1) = 2
    grilla1.Text = " "
    If indi > 7 Then
       indi = 1: indj = indj + 1
    End If
    
    grilla1.Col = j
    If maxcolumna = 42 Then
       grilla1.ColWidth(j) = 21
       grilla1.Text = "Semana " & indj & " (Día " & i & ")"
    ElseIf maxcolumna = 1 Then
       grilla1.ColWidth(j) = 31
       grilla1.Text = " "
    ElseIf maxcolumna = 2 Then
       grilla1.ColWidth(j) = 21
       grilla1.Text = " "
    End If
    indi = indi + 1
    
    grilla1.Col = j + 4 + 1
    grilla1.ColHidden = True
'    vaSpread1.ColWidth(j + 4 + 1) = 4.75
'    vaSpread1.Text = "Costo Plato"
    
    grilla1.Col = j + 1
    grilla1.ColHidden = True
    grilla1.Col = j + 2
    grilla1.ColHidden = True
    grilla1.Col = j + 3
    grilla1.ColHidden = True
    grilla1.Col = j + 4
    grilla1.ColHidden = True
    grilla1.Col = j + 6
    grilla1.ColHidden = True
    
    X = 1
    For X = 1 To 100
        grilla1.Row = X
        grilla1.Col = j + 6
        If i < 10 Then
           grilla1.Text = "0" & i
        Else
           grilla1.Text = i
        End If
    Next X
    grilla1.Row = 0
    j = j + 8
Next i
For i = 1 To 100
    grilla1.Row = i
    grilla1.Col = 1
    If estado = 0 Then
       grilla1.CellType = 1
    Else
       grilla1.CellType = 5
    End If
    grilla1.CellBorderType = 1 + 2 + 4 + 8
    grilla1.CellBorderStyle = 1
    grilla1.Action = 16
    grilla1.TypeHAlign = 0
    grilla1.Font.Bold = True
    grilla1.Font.Size = 9
    grilla1.Value = ""
    grilla1.BackColor = &H80000004
Next i
End Sub
Sub MoverPlantillaMinuta()
Dim j As Long
fg_carga ""
estado = 0
Set ConSql = vg_db.Execute("select * " & _
             "From Sdx_BloqueoMinutas " & _
             "where codigo_casino='" & vg_codcasino & "' " & _
             "and   codigo_segmento=0 " & _
             "and   codigo_pventa=" & vg_codpventa & " " & _
             "and   codigo_servicio=" & vg_codservicio & " " & _
             "and   fecha=0", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_bloqueominutas 1, '" & vg_codcasino & "', 0, " & vg_codpventa & ", " & vg_codservicio & ", 0, '" & "" & "', '" & "" & "'", , adCmdStoredProc)
If Not ConSql.EOF Then
   If ConSql!estado <> 0 Then
      estado = ConSql!estado
      Toolbar1.Buttons(27).Image = 22
      Toolbar1.Buttons(27).ToolTipText = "Minuta Bloqueada"
   End If
End If
ConSql.Close: Set ConSql = Nothing

Set grilla1 = vaSpread1
Set grilla2 = vaSpread2
maxcolumna = 42
MoverVecDia

j = 0

 Set ConSql = vg_db.Execute("SELECT Sdx_DetMinutas.ind_minuta, " & _
              "Sdx_DetMinutas.num_linea, Sdx_DetMinutas.num_dia, " & _
              "Sdx_DetMinutas.tipo_minuta, Sdx_DetMinutas.cod_item, " & _
              "Sdx_DetMinutas.descripcion, Sdx_DetMinutas.ind_borrado, " & _
              "Sdx_EncMinutas.cod_casino, Sdx_EncMinutas.cod_regimen, " & _
              "Sdx_EncMinutas.cod_servicio, PB00078.Rcpe_Desc " & _
              "FROM PB00078 INNER JOIN (Sdx_EncMinutas INNER JOIN Sdx_DetMinutas " & _
              "ON Sdx_EncMinutas.ind_minuta = Sdx_DetMinutas.ind_minuta) " & _
              "ON PB00078.Rcpe_No = Sdx_DetMinutas.ind_minuta " & _
              "Where (((Sdx_DetMinutas.ind_borrado)=0) " & _
              "And    ((Sdx_EncMinutas.cod_casino)='" & vg_codcasino & "') " & _
              "And    ((Sdx_EncMinutas.cod_regimen)=" & vg_codpventa & ") " & _
              "And    ((Sdx_EncMinutas.cod_servicio)=" & vg_codservicio & ") " & _
              "And    ((Sdx_EncMinutas.ind_borrado)=0) " & _
              "And    ((Sdx_DetMinutas.ind_borrado)=0)) " & _
              "ORDER BY Sdx_DetMinutas.num_linea, Sdx_DetMinutas.num_dia", , adCmdText)
'  Set ConSql = vg_db.Execute("sod_s_planificacionminuta 1, '" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & "", , adCmdStoredProc)
  If Not ConSql.EOF Then
     Do While Not ConSql.EOF
        If ConSql!num_dia = 0 Then
           j = 1
        Else
           j = (((ConSql!num_dia * 8) - 8) + 1) + 1 ' modificación suma 1 más
        End If
        vaSpread1.Row = ConSql!num_linea
        vaSpread2.Row = ConSql!num_linea
      
        Select Case ConSql!tipo_minuta
          Case 1
          
            vaSpread1.Col = j
            vaSpread1.CellType = 5
            vaSpread1.TypeHAlign = 2
            vaSpread1.Value = "R"
            vaSpread1.ForeColor = &HFF&
            vaSpread1.BackColor = &H80FF80
            
            vaSpread1.Col = j + 1
            vaSpread1.CellType = 5
            vaSpread1.TypeHAlign = 0
            vaSpread1.Value = Trim(ConSql!descripcion)
          
            vaSpread2.Col = j + 1
            vaSpread2.CellType = 5
            vaSpread2.TypeHAlign = 0
            vaSpread2.Value = Trim(ConSql!descripcion)
                          
            vaSpread1.Col = j + 2
            vaSpread1.CellType = 3
            vaSpread1.TypeIntegerMin = 1
            vaSpread1.TypeIntegerMax = 9999999
            vaSpread1.TypeHAlign = 1
            vaSpread1.TypeSpin = False
            vaSpread1.TypeIntegerSpinInc = 1
            vaSpread1.TypeIntegerSpinWrap = False
            vaSpread1.Value = 0
            vaSpread1.ForeColor = &HFF0000
                       
            vaSpread1.Col = j + 3
            vaSpread1.Value = 0
                       
            vaSpread1.Col = j + 4
            vaSpread1.Value = ConSql!tipo_minuta
                          
            vaSpread2.Col = j + 4
            vaSpread2.Value = ConSql!tipo_minuta
          
            vaSpread1.Col = j + 5
            vaSpread1.Value = ConSql!cod_item
                          
            vaSpread2.Col = j + 5
            vaSpread2.Value = ConSql!cod_item
          
            vaSpread1.Col = j + 6
            vaSpread1.TypeHAlign = 1
            vaSpread1.Value = Format(0, fg_Pict(6, 2))
            vaSpread1.ForeColor = &HFF0000
          Case 6
            If j = 1 Then
               vaSpread1.Col = j
               If estado = 0 Then
                  vaSpread1.CellType = 1
               Else
                  vaSpread1.CellType = 5
               End If
               vaSpread1.TypeHAlign = 0
               vaSpread1.Font.Bold = True
               vaSpread1.Font.Size = 9
               vaSpread1.Value = Trim(ConSql!descripcion)
           
               vaSpread2.Col = j
               vaSpread2.CellType = 1
               vaSpread2.TypeHAlign = 0
               vaSpread2.Font.Bold = True
               vaSpread2.Font.Size = 9
               vaSpread2.Value = Trim(ConSql!descripcion)
            Else
               vaSpread1.Col = j + 1
               vaSpread1.CellType = 5
               vaSpread1.TypeHAlign = 0
               vaSpread1.Font.Bold = True
               vaSpread1.Font.Size = 9
               vaSpread1.Value = Trim(ConSql!descripcion)
                       
               vaSpread2.Col = j + 1
               vaSpread2.CellType = 5
               vaSpread2.TypeHAlign = 0
               vaSpread2.Font.Bold = True
               vaSpread2.Font.Size = 9
               vaSpread2.Value = Trim(ConSql!descripcion)
                     
               vaSpread1.Col = j + 3
               vaSpread1.Value = 0
                     
               vaSpread1.Col = j + 4
               vaSpread1.Value = 6
                      
               vaSpread2.Col = j + 4
               vaSpread2.Value = 6
          
               vaSpread1.Col = j + 5
               vaSpread1.Value = 0
                     
               vaSpread2.Col = j + 5
               vaSpread2.Value = 0
          
               vaSpread1.Col = j + 6
               vaSpread1.Value = ""
          
            End If
        End Select
        ConSql.MoveNext
     Loop
  End If
  ConSql.Close: Set ConSql = Nothing: fg_descarga

vaSpread1.Row = 1: vaSpread1.Col = 1
iblockrow = vaSpread1.Row: aiblockrow = vaSpread1.Row
iblockrow2 = vaSpread1.Row: aiblockrow2 = vaSpread1.Row
iblockcol = vaSpread1.Col: aiblockcol = vaSpread1.Col
iblockcol2 = vaSpread1.Col: aiblockcol2 = vaSpread1.Col
End Sub
Private Sub vaSpread3_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
indactivo = 1
iblockrow = BlockRow
iblockrow2 = BlockRow2
iblockcol = BlockCol
iblockcol2 = BlockCol2
If BlockRow < 0 Then iblockrow = 1
If BlockRow2 < 0 Then iblockrow2 = 100
If BlockRow2 > 100 Then iblockrow2 = 100
End Sub
Private Sub vaSpread3_Click(ByVal Col As Long, ByVal Row As Long)
If Row > 0 Then
    If Col <> 0 Then SwCol = 0
    If Col = 0 Then SwCol = 1
    indactivo = 1
    iblockrow = vaSpread3.ActiveRow
    iblockrow2 = vaSpread3.ActiveRow
    iblockcol = vaSpread3.ActiveCol
    iblockcol2 = vaSpread3.ActiveCol
    vaSpread1.Row = vaSpread3.ActiveRow
    indfila = vaSpread3.ActiveRow
    vaSpread1.Col = vaSpread3.ActiveCol
    IndColumna = vaSpread3.ActiveCol
End If
End Sub
Private Sub vaSpread3_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row > 0 And estado = 0 Then
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   vaSpread3.Row = vaSpread3.ActiveRow
   vaSpread3.Col = vaSpread3.ActiveCol
   vaSpread3.Col = Col
   If vaSpread3.Col = 1 Then Exit Sub
   Select Case vaSpread3.Col
    Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
     newcol2 = vaSpread3.Col + 1
     vaSpread3.Col = newcol2
    Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
     newcol2 = vaSpread3.Col
   End Select
   If newcol2 < 1 Then
      newcol2 = 2
   ElseIf newcol2 > 330 Then
      newcol2 = 229
   End If
'   textocomentario = vaSpread1.Text
'   newcol2 = newcol2 - 1
'   If newcol2 < 1 Then newcol2 = 2
'   vaSpread1.Col = newcol2
'   If vaSpread1.Value = "" And textocomentario <> "" Then
'      Plato_Click (3)
'   Else
      Plato_Click (2)
'   End If
End If
End Sub
Private Sub vaSpread3_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread3.ChangeMade = False Then Exit Sub
If Col = 1 And estado = 0 Then
   IndGrabadoDetalle = 1
   swgrabadoest = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
'   Toolbar1.Buttons(6).Visible = True
'   Toolbar1.Buttons(7).Visible = False
'   Toolbar1.Buttons(9).Visible = False
'   Toolbar1.Buttons(10).Visible = True
End If
End Sub
Private Sub vaSpread3_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IBorrado As Long, delrow As Long, i As Long
Dim auxp1 As Integer, auxp2 As Integer, auxp3 As Integer, auxp4 As Integer, auxp5 As Integer, auxp6 As Integer
Dim indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer

IBorrado = 0
Select Case KeyCode
  Case 86 And estado <> 0
    Exit Sub
  Case 46 And estado = 0
    If vaTabPro1.Tab = 0 Then
       swgrabadoplan = 1
    ElseIf vaTabPro1.Tab = 1 Then
       swgrabadoest = 1
    ElseIf vaTabPro1.Tab = 2 Then
       swgrabadoadj = 1
    ElseIf vaTabPro1.Tab = 3 Then
       swgrabadoadj = 1
    End If
    vaSpread3.Row = vaSpread3.ActiveRow
    vaSpread3.Col = vaSpread3.ActiveCol
    If vaSpread3.Col = 1 Then Exit Sub
    Select Case vaSpread3.Col
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        newcol3 = vaSpread3.Col
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        newcol3 = vaSpread3.Col - 1
    End Select
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    IOpcion = 14
    If vaSpread3.MaxRows > 100 Then
       delrow = vaSpread3.MaxRows - 100
       vaSpread3.MaxRows = vaSpread3.MaxRows - delrow
    End If
    If indactivo = 0 Then iblockcol = vaSpread3.ActiveCol: iblockcol2 = vaSpread3.ActiveCol: iblockrow = vaSpread3.ActiveRow: iblockrow2 = vaSpread3.ActiveRow
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread3.MaxCols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    Select Case iblockcol
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol = iblockcol
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol = iblockcol - 1
    End Select
    Select Case iblockcol2
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol2 = iblockcol2 + 7
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol2 = ((iblockcol2 + 7) - 1)
    End Select
    auxcol = iblockcol: auxcol2 = iblockcol2
    auxrow = iblockrow: auxrow2 = iblockrow2
    iauxcol = aiblockcol: iauxcol2 = aiblockcol2
    iauxrow = aiblockrow: iauxrow2 = aiblockrow2
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    For i = iblockcol To iblockcol2
'' Calcular y Mover Datos Ultima linea
        If vaSpread3.MaxRows <= 100 Then
           vaSpread3.MaxRows = (vaSpread3.MaxRows + (aiblockrow2 - aiblockrow)) + 1
           For auxp1 = 101 To vaSpread3.MaxRows
               vaSpread3.Row = auxp1
               vaSpread3.RowHidden = True
           Next auxp1
        End If
        auxp6 = aiblockrow2 - aiblockrow + iblockrow
        vaSpread3.Col = iblockcol
        vaSpread3.Row = iblockrow
        vaSpread3.Col2 = iblockcol + 6
        vaSpread3.Row2 = auxp6
        vaSpread3.DestCol = iblockcol
        vaSpread3.DestRow = 100 + 1
        vaSpread3.Action = 20
''fin mover datos ultimas lineas
      
        vaSpread3.Col = iblockcol
        vaSpread3.Row = iblockrow
        vaSpread3.Col2 = iblockcol + 6
        vaSpread3.Row2 = iblockrow2
        vaSpread3.BlockMode = True
'' Limpiar Datos y Formato Celda
        vaSpread3.Action = 3
        i = iblockcol + 7
        iblockcol = iblockcol + 8
    Next i
    iblockcol = auxcol
    vaSpread3.BlockMode = False
    IndGrabadoDetalle = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(9).Visible = False
    Toolbar1.Buttons(10).Visible = True
    indactivo = 0
End Select
End Sub
Private Sub vaSpread3_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal newrow As Long, Cancel As Boolean)
vaSpread3.Row = Row
vaSpread3.Col = Col
iblockrow = vaSpread3.ActiveRow
iblockrow2 = vaSpread3.ActiveRow
iblockcol = vaSpread3.ActiveCol
iblockcol2 = vaSpread3.ActiveCol
Select Case Col
  Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259
    vaSpread3.Col = Col
    newcol1 = vaSpread3.Col
    newrow1 = vaSpread3.ActiveRow
    WsNumPlanificacion = Val(vaSpread3.Value)
    vaSpread3.Col = newcol1 + 1
    If Val(vaSpread3.Value) <> WsNumPlanificacion Then
       vaSpread3.EditModeReplace = True
       vaSpread3.OperationMode = 0
       vaSpread3.Col = newcol1
       vaSpread3.Row = newrow1
       vaSpread3.Col = newcol1 + 1
       vaSpread3.Value = WsNumPlanificacion
       IndGrabadoDetalle = 1
       Plantilla(0).Enabled = True
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = True
       vaSpread3.Row = newrow1
    End If
End Select
End Sub
Private Sub vaspread3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
  If vaSpread3.Visible = True And estado = 0 Then
     Indgrilla1 = 0
     PopupMenu MenuDetalle
  End If
End Select
End Sub
Private Sub vaspread5_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
indactivo = 1
iblockrow = BlockRow
iblockrow2 = BlockRow2
iblockcol = BlockCol
iblockcol2 = BlockCol2
If BlockRow < 0 Then iblockrow = 1
If BlockRow2 < 0 Then iblockrow2 = 100
If BlockRow2 > 100 Then iblockrow2 = 100
End Sub
Private Sub vaspread5_Click(ByVal Col As Long, ByVal Row As Long)
If Row > 0 Then
    If Col <> 0 Then SwCol = 0
    If Col = 0 Then SwCol = 1
    indactivo = 1
    iblockrow = vaSpread5.ActiveRow
    iblockrow2 = vaSpread5.ActiveRow
    iblockcol = vaSpread5.ActiveCol
    iblockcol2 = vaSpread5.ActiveCol
    vaSpread1.Row = vaSpread5.ActiveRow
    indfila = vaSpread5.ActiveRow
    vaSpread1.Col = vaSpread5.ActiveCol
    IndColumna = vaSpread5.ActiveCol
End If
End Sub
Private Sub vaspread5_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row > 0 And estado = 0 Then
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   vaSpread5.Row = vaSpread5.ActiveRow
   vaSpread5.Col = vaSpread5.ActiveCol
   vaSpread5.Col = Col
   If vaSpread5.Col = 1 Then Exit Sub
   Select Case vaSpread5.Col
    Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
     newcol2 = vaSpread5.Col + 1
     vaSpread5.Col = newcol2
    Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
     newcol2 = vaSpread5.Col
   End Select
   If newcol2 < 1 Then
      newcol2 = 2
   ElseIf newcol2 > 330 Then
      newcol2 = 229
   End If
'   textocomentario = vaSpread1.Text
'   newcol2 = newcol2 - 1
'   If newcol2 < 1 Then newcol2 = 2
'   vaSpread1.Col = newcol2
'   If vaSpread1.Value = "" And textocomentario <> "" Then
'      Plato_Click (3)
'   Else
      Plato_Click (2)
'   End If
End If
End Sub
Private Sub vaspread5_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread5.ChangeMade = False Then Exit Sub
If Col = 1 And estado = 0 Then
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
'   Toolbar1.Buttons(6).Visible = True
'   Toolbar1.Buttons(7).Visible = False
'   Toolbar1.Buttons(9).Visible = False
'   Toolbar1.Buttons(10).Visible = True
End If
End Sub
Private Sub vaspread5_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IBorrado As Long, delrow As Long, i As Long
Dim auxp1 As Integer, auxp2 As Integer, auxp3 As Integer, auxp4 As Integer, auxp5 As Integer, auxp6 As Integer
Dim indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer

IBorrado = 0
Select Case KeyCode
  Case 86 And estado <> 0
    Exit Sub
  Case 46 And estado = 0
    If vaTabPro1.Tab = 0 Then
       swgrabadoplan = 1
    ElseIf vaTabPro1.Tab = 1 Then
       swgrabadoest = 1
    ElseIf vaTabPro1.Tab = 2 Then
       swgrabadoadj = 1
    ElseIf vaTabPro1.Tab = 3 Then
       swgrabadoadj = 1
    End If
    vaSpread5.Row = vaSpread5.ActiveRow
    vaSpread5.Col = vaSpread5.ActiveCol
    If vaSpread5.Col = 1 Then Exit Sub
    Select Case vaSpread5.Col
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        newcol3 = vaSpread5.Col
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        newcol3 = vaSpread5.Col - 1
    End Select
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    IOpcion = 14
    If vaSpread5.MaxRows > 100 Then
       delrow = vaSpread5.MaxRows - 100
       vaSpread5.MaxRows = vaSpread5.MaxRows - delrow
    End If
    If indactivo = 0 Then iblockcol = vaSpread5.ActiveCol: iblockcol2 = vaSpread5.ActiveCol: iblockrow = vaSpread5.ActiveRow: iblockrow2 = vaSpread5.ActiveRow
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread5.MaxCols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    Select Case iblockcol
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol = iblockcol
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol = iblockcol - 1
    End Select
    Select Case iblockcol2
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol2 = iblockcol2 + 7
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol2 = ((iblockcol2 + 7) - 1)
    End Select
    auxcol = iblockcol: auxcol2 = iblockcol2
    auxrow = iblockrow: auxrow2 = iblockrow2
    iauxcol = aiblockcol: iauxcol2 = aiblockcol2
    iauxrow = aiblockrow: iauxrow2 = aiblockrow2
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    For i = iblockcol To iblockcol2
'' Calcular y Mover Datos Ultima linea
        If vaSpread5.MaxRows <= 100 Then
           vaSpread5.MaxRows = (vaSpread5.MaxRows + (aiblockrow2 - aiblockrow)) + 1
           For auxp1 = 101 To vaSpread5.MaxRows
               vaSpread5.Row = auxp1
               vaSpread5.RowHidden = True
           Next auxp1
        End If
        auxp6 = aiblockrow2 - aiblockrow + iblockrow
        vaSpread5.Col = iblockcol
        vaSpread5.Row = iblockrow
        vaSpread5.Col2 = iblockcol + 6
        vaSpread5.Row2 = auxp6
        vaSpread5.DestCol = iblockcol
        vaSpread5.DestRow = 100 + 1
        vaSpread5.Action = 20
''fin mover datos ultimas lineas
      
        vaSpread5.Col = iblockcol
        vaSpread5.Row = iblockrow
        vaSpread5.Col2 = iblockcol + 6
        vaSpread5.Row2 = iblockrow2
        vaSpread5.BlockMode = True
'' Limpiar Datos y Formato Celda
        vaSpread5.Action = 3
        i = iblockcol + 7
        iblockcol = iblockcol + 8
    Next i
    iblockcol = auxcol
    vaSpread5.BlockMode = False
    IndGrabadoDetalle = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(9).Visible = False
    Toolbar1.Buttons(10).Visible = True
    indactivo = 0
End Select
End Sub
Private Sub vaspread5_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal newrow As Long, Cancel As Boolean)
vaSpread5.Row = Row
vaSpread5.Col = Col
iblockrow = vaSpread5.ActiveRow
iblockrow2 = vaSpread5.ActiveRow
iblockcol = vaSpread5.ActiveCol
iblockcol2 = vaSpread5.ActiveCol
Select Case Col
  Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259
    vaSpread5.Col = Col
    newcol1 = vaSpread5.Col
    newrow1 = vaSpread5.ActiveRow
    WsNumPlanificacion = Val(vaSpread5.Value)
    vaSpread5.Col = newcol1 + 1
    If Val(vaSpread5.Value) <> WsNumPlanificacion Then
       vaSpread5.EditModeReplace = True
       vaSpread5.OperationMode = 0
       vaSpread5.Col = newcol1
       vaSpread5.Row = newrow1
       vaSpread5.Col = newcol1 + 1
       vaSpread5.Value = WsNumPlanificacion
       IndGrabadoDetalle = 1
       Plantilla(0).Enabled = True
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = True
       vaSpread5.Row = newrow1
    End If
End Select
End Sub
Private Sub vaspread5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
  If vaSpread5.Visible = True And estado = 0 Then
     Indgrilla1 = 0
     PopupMenu MenuDetalle
  End If
End Select
End Sub
Private Sub vaspread7_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)
indactivo = 1
iblockrow = BlockRow
iblockrow2 = BlockRow2
iblockcol = BlockCol
iblockcol2 = BlockCol2
If BlockRow < 0 Then iblockrow = 1
If BlockRow2 < 0 Then iblockrow2 = 100
If BlockRow2 > 100 Then iblockrow2 = 100
End Sub
Private Sub vaspread7_Click(ByVal Col As Long, ByVal Row As Long)
If Row > 0 Then
    If Col <> 0 Then SwCol = 0
    If Col = 0 Then SwCol = 1
    indactivo = 1
    iblockrow = vaSpread7.ActiveRow
    iblockrow2 = vaSpread7.ActiveRow
    iblockcol = vaSpread7.ActiveCol
    iblockcol2 = vaSpread7.ActiveCol
    vaSpread1.Row = vaSpread7.ActiveRow
    indfila = vaSpread7.ActiveRow
    vaSpread1.Col = vaSpread7.ActiveCol
    IndColumna = vaSpread7.ActiveCol
End If
End Sub
Private Sub vaspread7_DblClick(ByVal Col As Long, ByVal Row As Long)
If Row > 0 And estado = 0 Then
   Plato(13).Enabled = False
   OpGrilla(13).Enabled = False
   vaSpread7.Row = vaSpread7.ActiveRow
   vaSpread7.Col = vaSpread7.ActiveCol
   vaSpread7.Col = Col
   If vaSpread7.Col = 1 Then Exit Sub
   Select Case vaSpread7.Col
    Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
     newcol2 = vaSpread7.Col + 1
     vaSpread7.Col = newcol2
    Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
     newcol2 = vaSpread7.Col
   End Select
   If newcol2 < 1 Then
      newcol2 = 2
   ElseIf newcol2 > 330 Then
      newcol2 = 229
   End If
'   textocomentario = vaSpread1.Text
'   newcol2 = newcol2 - 1
'   If newcol2 < 1 Then newcol2 = 2
'   vaSpread1.Col = newcol2
'   If vaSpread1.Value = "" And textocomentario <> "" Then
'      Plato_Click (3)
'   Else
      Plato_Click (2)
'   End If
End If
End Sub
Private Sub vaspread7_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)
If vaSpread7.ChangeMade = False Then Exit Sub
If Col = 1 And estado = 0 Then
   IndGrabadoDetalle = 1
   Plantilla(0).Enabled = True
   Toolbar1.Buttons(1).Visible = False
   Toolbar1.Buttons(2).Visible = True
'   Toolbar1.Buttons(6).Visible = True
'   Toolbar1.Buttons(7).Visible = False
'   Toolbar1.Buttons(9).Visible = False
'   Toolbar1.Buttons(10).Visible = True
End If
End Sub
Private Sub vaspread7_KeyDown(KeyCode As Integer, Shift As Integer)
Dim IBorrado As Long, delrow As Long, i As Long
Dim auxp1 As Integer, auxp2 As Integer, auxp3 As Integer, auxp4 As Integer, auxp5 As Integer, auxp6 As Integer
Dim indcol As Integer, indrow As Integer, indcol2 As Integer, indrow2 As Integer

IBorrado = 0
Select Case KeyCode
  Case 86 And estado <> 0
    Exit Sub
  Case 46 And estado = 0
    If vaTabPro1.Tab = 0 Then
       swgrabadoplan = 1
    ElseIf vaTabPro1.Tab = 1 Then
       swgrabadoest = 1
    ElseIf vaTabPro1.Tab = 2 Then
       swgrabadoadj = 1
    ElseIf vaTabPro1.Tab = 3 Then
       swgrabadoadj = 1
    End If
    vaSpread7.Row = vaSpread7.ActiveRow
    vaSpread7.Col = vaSpread7.ActiveCol
    If vaSpread7.Col = 1 Then Exit Sub
    Select Case vaSpread7.Col
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        newcol3 = vaSpread7.Col
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        newcol3 = vaSpread7.Col - 1
    End Select
    Plato(0).Enabled = True
    OpGrilla(0).Enabled = True
    Plato(13).Enabled = False
    OpGrilla(13).Enabled = False
    IOpcion = 14
    If vaSpread7.MaxRows > 100 Then
       delrow = vaSpread7.MaxRows - 100
       vaSpread7.MaxRows = vaSpread7.MaxRows - delrow
    End If
    If indactivo = 0 Then iblockcol = vaSpread7.ActiveCol: iblockcol2 = vaSpread7.ActiveCol: iblockrow = vaSpread7.ActiveRow: iblockrow2 = vaSpread7.ActiveRow
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    If iblockcol < 0 Then iblockcol = 2: iblockcol2 = vaSpread7.MaxCols
    aiblockcol = iblockcol
    aiblockrow = iblockrow
    aiblockcol2 = iblockcol2
    aiblockrow2 = iblockrow2
    Select Case iblockcol
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol = iblockcol
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol = iblockcol - 1
    End Select
    Select Case iblockcol2
      Case 2, 10, 18, 26, 34, 42, 50, 58, 66, 74, 82, 90, 98, 106, 114, 122, 130, 138, 146, 154, 162, 170, 178, 186, 194, 202, 210, 218, 226, 234, 242, 250, 258, 266, 274, 282, 290, 298, 306, 314, 322, 330
        iblockcol2 = iblockcol2 + 7
      Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259, 267, 275, 283, 291, 299, 307, 315, 323, 331
        iblockcol2 = ((iblockcol2 + 7) - 1)
    End Select
    auxcol = iblockcol: auxcol2 = iblockcol2
    auxrow = iblockrow: auxrow2 = iblockrow2
    iauxcol = aiblockcol: iauxcol2 = aiblockcol2
    iauxrow = aiblockrow: iauxrow2 = aiblockrow2
    indcol = aiblockcol: indcol2 = iblockcol2
    indrow = aiblockrow: indrow2 = aiblockrow2
    For i = iblockcol To iblockcol2
'' Calcular y Mover Datos Ultima linea
        If vaSpread7.MaxRows <= 100 Then
           vaSpread7.MaxRows = (vaSpread7.MaxRows + (aiblockrow2 - aiblockrow)) + 1
           For auxp1 = 101 To vaSpread7.MaxRows
               vaSpread7.Row = auxp1
               vaSpread7.RowHidden = True
           Next auxp1
        End If
        auxp6 = aiblockrow2 - aiblockrow + iblockrow
        vaSpread7.Col = iblockcol
        vaSpread7.Row = iblockrow
        vaSpread7.Col2 = iblockcol + 6
        vaSpread7.Row2 = auxp6
        vaSpread7.DestCol = iblockcol
        vaSpread7.DestRow = 100 + 1
        vaSpread7.Action = 20
''fin mover datos ultimas lineas
      
        vaSpread7.Col = iblockcol
        vaSpread7.Row = iblockrow
        vaSpread7.Col2 = iblockcol + 6
        vaSpread7.Row2 = iblockrow2
        vaSpread7.BlockMode = True
'' Limpiar Datos y Formato Celda
        vaSpread7.Action = 3
        i = iblockcol + 7
        iblockcol = iblockcol + 8
    Next i
    iblockcol = auxcol
    vaSpread7.BlockMode = False
    IndGrabadoDetalle = 1
    Plantilla(0).Enabled = True
    Toolbar1.Buttons(1).Visible = False
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(6).Visible = True
    Toolbar1.Buttons(7).Visible = False
    Toolbar1.Buttons(9).Visible = False
    Toolbar1.Buttons(10).Visible = True
    indactivo = 0
End Select
End Sub
Private Sub vaspread7_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal newrow As Long, Cancel As Boolean)
vaSpread7.Row = Row
vaSpread7.Col = Col
iblockrow = vaSpread7.ActiveRow
iblockrow2 = vaSpread7.ActiveRow
iblockcol = vaSpread7.ActiveCol
iblockcol2 = vaSpread7.ActiveCol
Select Case Col
  Case 3, 11, 19, 27, 35, 43, 51, 59, 67, 75, 83, 91, 99, 107, 115, 123, 131, 139, 147, 155, 163, 171, 179, 187, 195, 203, 211, 219, 227, 235, 243, 251, 259
    vaSpread7.Col = Col
    newcol1 = vaSpread7.Col
    newrow1 = vaSpread7.ActiveRow
    WsNumPlanificacion = Val(vaSpread7.Value)
    vaSpread7.Col = newcol1 + 1
    If Val(vaSpread7.Value) <> WsNumPlanificacion Then
       vaSpread7.EditModeReplace = True
       vaSpread7.OperationMode = 0
       vaSpread7.Col = newcol1
       vaSpread7.Row = newrow1
       vaSpread7.Col = newcol1 + 1
       vaSpread7.Value = WsNumPlanificacion
       IndGrabadoDetalle = 1
       Plantilla(0).Enabled = True
       Toolbar1.Buttons(1).Visible = False
       Toolbar1.Buttons(2).Visible = True
       vaSpread7.Row = newrow1
    End If
End Select
End Sub
Private Sub vaspread7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Button
Case 2
  If vaSpread7.Visible = True And estado = 0 Then
     Indgrilla1 = 0
     PopupMenu MenuDetalle
  End If
End Select
End Sub
Private Sub vaTabPro1_TabActivate(TabToActivate As Integer)
Select Case TabToActivate
  Case 0
    Set Program = M_Minu02.vaSpread1
    Set grilla1 = vaSpread1
    vaTabPro1.Tab = 0
  Case 1
    Set grilla1 = vaSpread3
    Set Program = M_Minu02.vaSpread3
    vaTabPro1.Tab = 1
  Case 2
    Set grilla1 = vaSpread5
    Set Program = M_Minu02.vaSpread5
    vaTabPro1.Tab = 2
  Case 3
    Set grilla1 = vaSpread7
    Set Program = M_Minu02.vaSpread7
    vaTabPro1.Tab = 3
End Select
End Sub
Private Sub Ver_Click(Index As Integer)
Dim delrow As Long
Select Case Index
  Case 5
'    M_MINU21.Show 1
  Case 6
    If vaSpread1.Col = 1 Then Exit Sub
    If IndGrabadoDetalle = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, "Planificación Minutas": Exit Sub
    M_Minu06.Show 1
  Case 7
    If IndGrabadoDetalle = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, "Planificación Minutas": Exit Sub
    M_Minu07.Show 1
  Case 8
    If IndGrabadoDetalle = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, "Planificación Minutas": Exit Sub
    M_Minu08.Show 1
  Case 9
    If IndGrabadoDetalle = 1 Then MsgBox "Actualice Datos, para ver Información", vbInformation + vbOKOnly, "Planificación Minutas": Exit Sub
    M_Minu09.Show 1
End Select
If vaSpread1.MaxRows > 100 Then
   delrow = vaSpread1.MaxRows - 100
   vaSpread1.MaxRows = vaSpread1.MaxRows - delrow
End If
End Sub
Sub BloquearMinuta()
' *** bloquear minuta *** '
TITLE = "Planificación Minutas"
msg = " Esta Seguro Bloquear Minuta ?"
Style = vbYesNo + vbQuestion + vbDefaultButton2
Help = "DEMO.HLP"
Ctxt = 1000
ws_respuesta = MsgBox(msg, Style, TITLE, Help, Ctxt)
Select Case ws_respuesta
  Case Is = vbYes
    
    ' *** Leer Bloque Minutas *** '
    
    fecha = Val(vg_ano & vg_mes)
    Set ConSql = vg_db.Execute("select  * " & _
                 "From Sdx_BloqueoMinutas " & _
                 "where codigo_casino='" & vg_codcasino & "' " & _
                 "and   codigo_segmento=0 " & _
                 "and   codigo_pventa=" & vg_codpventa & " " & _
                 "and   codigo_servicio=" & vg_codservicio & " " & _
                 "and   fecha=0", , adCmdText)
'    Set ConSql = vg_db.Execute("sod_s_bloqueominutas 1, '" & vg_codcasino & "', 0, " & vg_codpventa & ", " & vg_codservicio & ", 0, '', ''", , adCmdStoredProc)
    If Not ConSql.EOF Then
       If ConSql!estado = 0 Then
          ConSql.Close: Set ConSql = Nothing
          vg_db.Execute "update Sdx_BloqueoMinutas set estado=" & 1 & ", " & _
                        "usuario='" & vg_Nuser & "' where " & _
                        "codigo_casino='" & vg_codcasino & "' " & _
                        "and codigo_pventa=" & vg_codpventa & " " & _
                        "and codigo_servicio=" & vg_codservicio & ""
          estado = 1
          Toolbar1.Buttons(27).Image = 22
          Toolbar1.Buttons(27).ToolTipText = "Minuta Bloqueada"
       End If
    Else
       ConSql.Close: Set ConSql = Nothing
       vg_db.Execute "insert into Sdx_BloqueoMinutas(codigo_casino, codigo_segmento, codigo_pventa, " & _
                     "codigo_servicio, fecha, fecha_proceso, " & _
                     "estado, usuario) values ('" & vg_codcasino & "', 0, " & _
                     "" & vg_codpventa & ", " & vg_codservicio & ", " & _
                     "0, " & _
                     "" & Val(Format(Date, "yyyymmdd")) & ", " & 1 & ", '" & vg_NUsr & "')"
       estado = 1
       Toolbar1.Buttons(27).Image = 22
       Toolbar1.Buttons(27).ToolTipText = "Minuta Bloqueado"
    End If
  
  End Select

End Sub
Sub MoverEstructuraFija()
Dim j As Long
fg_carga ""
vaSpread3.MaxRows = 100
vaSpread4.MaxRows = 100
maxcolumna = 1
Set grilla1 = vaSpread3
Set grilla2 = vaSpread4
MoverVecDia

j = 0
Set ConSql = vg_db.Execute("SELECT Sdx_DetEstructuraFija.ind_estfija, Sdx_DetEstructuraFija.num_linea, " & _
             "Sdx_DetEstructuraFija.num_dia, Sdx_DetEstructuraFija.tipo_estfija, " & _
             "Sdx_DetEstructuraFija.cod_item, Sdx_DetEstructuraFija.descripcion, " & _
             "Sdx_DetEstructuraFija.ind_borrado, PB00078.Rcpe_Desc, Sdx_EncEstructuraFija.cod_casino, " & _
             "Sdx_EncEstructuraFija.cod_regimen, Sdx_EncEstructuraFija.cod_servicio, " & _
             "Sdx_EncEstructuraFija.ind_borrado, Sdx_DetEstructuraFija.ind_borrado, PB00078.Del_Ind " & _
             "from (Sdx_EncEstructuraFija INNER JOIN Sdx_DetEstructuraFija ON " & _
             "Sdx_EncEstructuraFija.ind_estfija = Sdx_DetEstructuraFija.ind_estfija) " & _
             "LEFT JOIN PB00078 ON Sdx_DetEstructuraFija.cod_item = PB00078.Rcpe_No " & _
             "WHERE Sdx_EncEstructuraFija.cod_casino='" & vg_codcasino & "' " & _
             "and   Sdx_EncEstructuraFija.cod_regimen=" & vg_codpventa & " " & _
             "and   Sdx_EncEstructuraFija.cod_servicio=" & vg_codservicio & " " & _
             "and   Sdx_EncEstructuraFija.ind_borrado=0 " & _
             "and   Sdx_DetEstructuraFija.ind_borrado=0 " & _
             "and   PB00078.Del_Ind=0 " & _
             "order by Sdx_DetEstructuraFija.num_linea, Sdx_DetEstructuraFija.num_dia", , adCmdText)
'Set ConSql = vg_db.Execute("sod_s_estructurafija 1, '" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & "", , adCmdStoredProc)
  If Not ConSql.EOF Then
     Do While Not ConSql.EOF
        If ConSql!num_dia = 0 Then
           j = 1
        Else
           j = (((ConSql!num_dia * 8) - 8) + 1) + 1 ' modificación suma 1 más
        End If
        vaSpread3.Row = ConSql!num_linea
        vaSpread4.Row = ConSql!num_linea
      
        Select Case ConSql!tipo_estfija
          Case 1
          
            vaSpread3.Col = j
            vaSpread3.CellType = 5
            vaSpread3.TypeHAlign = 2
            vaSpread3.Value = "R"
            vaSpread3.ForeColor = &HFF&
            vaSpread3.BackColor = &H80FF80
            
            vaSpread3.Col = j + 1
            vaSpread3.CellType = 5
            vaSpread3.TypeHAlign = 0
            vaSpread3.Value = Trim(ConSql!descripcion)
          
            vaSpread4.Col = j + 1
            vaSpread4.CellType = 5
            vaSpread4.TypeHAlign = 0
            vaSpread4.Value = Trim(ConSql!descripcion)
                          
            vaSpread3.Col = j + 2
            vaSpread3.CellType = 3
            vaSpread3.TypeIntegerMin = 1
            vaSpread3.TypeIntegerMax = 9999999
            vaSpread3.TypeHAlign = 1
            vaSpread3.TypeSpin = False
            vaSpread3.TypeIntegerSpinInc = 1
            vaSpread3.TypeIntegerSpinWrap = False
            vaSpread3.Value = 0
            vaSpread3.ForeColor = &HFF0000
                       
            vaSpread3.Col = j + 3
            vaSpread3.Value = 0
                       
            vaSpread3.Col = j + 4
            vaSpread3.Value = ConSql!tipo_estfija
                          
            vaSpread4.Col = j + 4
            vaSpread4.Value = ConSql!tipo_estfija
          
            vaSpread3.Col = j + 5
            vaSpread3.Value = ConSql!cod_item
                          
            vaSpread4.Col = j + 5
            vaSpread4.Value = ConSql!cod_item
          
            vaSpread3.Col = j + 6
            vaSpread3.TypeHAlign = 1
            vaSpread3.Value = Format(0, fg_Pict(6, 2))
            vaSpread3.ForeColor = &HFF0000
          Case 6
            If j = 1 Then
               vaSpread3.Col = j
               If estado = 0 Then
                  vaSpread3.CellType = 1
               Else
                  vaSpread3.CellType = 5
               End If
               vaSpread3.TypeHAlign = 0
               vaSpread3.Font.Bold = True
               vaSpread3.Font.Size = 9
               vaSpread3.Value = Trim(ConSql!descripcion)
           
               vaSpread4.Col = j
               vaSpread4.CellType = 1
               vaSpread4.TypeHAlign = 0
               vaSpread4.Font.Bold = True
               vaSpread4.Font.Size = 9
               vaSpread4.Value = Trim(ConSql!descripcion)
            Else
               vaSpread3.Col = j + 1
               vaSpread3.CellType = 5
               vaSpread3.TypeHAlign = 0
               vaSpread3.Font.Bold = True
               vaSpread3.Font.Size = 9
               vaSpread3.Value = Trim(ConSql!descripcion)
                       
               vaSpread4.Col = j + 1
               vaSpread4.CellType = 5
               vaSpread4.TypeHAlign = 0
               vaSpread4.Font.Bold = True
               vaSpread4.Font.Size = 9
               vaSpread4.Value = Trim(ConSql!descripcion)
                     
               vaSpread3.Col = j + 3
               vaSpread3.Value = 0
                     
               vaSpread3.Col = j + 4
               vaSpread3.Value = 6
                      
               vaSpread4.Col = j + 4
               vaSpread4.Value = 6
          
               vaSpread3.Col = j + 5
               vaSpread3.Value = 0
                     
               vaSpread4.Col = j + 5
               vaSpread4.Value = 0
          
               vaSpread3.Col = j + 6
               vaSpread3.Value = ""
          
            End If
        End Select
        ConSql.MoveNext
     Loop
  End If
  ConSql.Close: Set ConSql = Nothing: fg_descarga

vaSpread3.Row = 1: vaSpread3.Col = 1
iblockrow = vaSpread3.Row: aiblockrow = vaSpread3.Row
iblockrow2 = vaSpread3.Row: aiblockrow2 = vaSpread3.Row
iblockcol = vaSpread3.Col: aiblockcol = vaSpread3.Col
iblockcol2 = vaSpread3.Col: aiblockcol2 = vaSpread3.Col

End Sub
Sub MoverDatosAdjunto()
Dim j As Long
fg_carga ""
MoverVecDia
j = 0

  Set ConSql = vg_db.Execute("SELECT Sdx_DetDatosAdjuntos.ind_datadj, Sdx_DetDatosAdjuntos.num_linea, " & _
               "Sdx_DetDatosAdjuntos.num_dia, Sdx_DetDatosAdjuntos.tipo_datadj, " & _
               "Sdx_DetDatosAdjuntos.cod_item, Sdx_DetDatosAdjuntos.descripcion, " & _
               "Sdx_DetDatosAdjuntos.ind_borrado, PB00078.Rcpe_Desc " & _
               "from (Sdx_EncDatosAdjuntos INNER JOIN Sdx_DetDatosAdjuntos ON " & _
               "Sdx_EncDatosAdjuntos.ind_datadj = Sdx_DetDatosAdjuntos.ind_datadj) " & _
               "LEFT JOIN PB00078 ON Sdx_DetDatosAdjuntos.cod_item = PB00078.Rcpe_No " & _
               "where Sdx_EncDatosAdjuntos.cod_casino='" & vg_codcasino & "' " & _
               "and   Sdx_EncDatosAdjuntos.cod_regimen=" & vg_codpventa & " " & _
               "and   Sdx_EncDatosAdjuntos.cod_servicio=" & vg_codservicio & " " & _
               "and   Sdx_EncDatosAdjuntos.tipo_datadj='" & tipodatadj & "' " & _
               "and   Sdx_EncDatosAdjuntos.ind_borrado=0 " & _
               "and   Sdx_DetDatosAdjuntos.ind_borrado=0 " & _
               "order by Sdx_DetDatosAdjuntos.num_linea, Sdx_DetDatosAdjuntos.num_dia", , adCmdText)
'  Set ConSql = vg_db.Execute("sod_s_datosadjuntos 1, '" & vg_codcasino & "', " & vg_codpventa & ", " & vg_codservicio & ", '" & tipodatadj & "'", , adCmdStoredProc)
  If Not ConSql.EOF Then
     Do While Not ConSql.EOF
        If ConSql!num_dia = 0 Then
           j = 1
        Else
           If ConSql!num_dia >= 3 And ConSql!num_dia <= 4 Then
              j = (((((ConSql!num_dia - 2) * 8) - 8) + 1) + 1) ' modificación suma 1 más
           Else
              j = (((ConSql!num_dia * 8) - 8) + 1) + 1 ' modificación suma 1 más
           End If
        End If
        grilla1.Row = ConSql!num_linea
        grilla2.Row = ConSql!num_linea
      
        Select Case ConSql!tipo_datadj
          Case 1
          
            grilla1.Col = j
            grilla1.CellType = 5
            grilla1.TypeHAlign = 2
            grilla1.Value = "R"
            grilla1.ForeColor = &HFF&
            grilla1.BackColor = &H80FF80
            
            grilla1.Col = j + 1
            grilla1.CellType = 5
            grilla1.TypeHAlign = 0
            grilla1.Value = Trim(ConSql!descripcion)
          
            grilla2.Col = j + 1
            grilla2.CellType = 5
            grilla2.TypeHAlign = 0
            grilla2.Value = Trim(ConSql!descripcion)
                          
            grilla1.Col = j + 2
            grilla1.CellType = 3
            grilla1.TypeIntegerMin = 1
            grilla1.TypeIntegerMax = 9999999
            grilla1.TypeHAlign = 1
            grilla1.TypeSpin = False
            grilla1.TypeIntegerSpinInc = 1
            grilla1.TypeIntegerSpinWrap = False
            grilla1.Value = 0
            grilla1.ForeColor = &HFF0000
                       
            grilla1.Col = j + 3
            grilla1.Value = 0
                       
            grilla1.Col = j + 4
            grilla1.Value = ConSql!tipo_datadj
                          
            grilla2.Col = j + 4
            grilla2.Value = ConSql!tipo_datadj
          
            grilla1.Col = j + 5
            grilla1.Value = ConSql!cod_item
                          
            grilla2.Col = j + 5
            grilla2.Value = ConSql!cod_item
          
            grilla1.Col = j + 6
            grilla1.TypeHAlign = 1
            grilla1.Value = Format(0, fg_Pict(6, 2))
            grilla1.ForeColor = &HFF0000
          Case 6
            If j = 1 Then
               grilla1.Col = j
               If estado = 0 Then
                  grilla1.CellType = 1
               Else
                  grilla1.CellType = 5
               End If
               grilla1.TypeHAlign = 0
               grilla1.Font.Bold = True
               grilla1.Font.Size = 9
               grilla1.Value = Trim(ConSql!descripcion)
           
               grilla2.Col = j
               grilla2.CellType = 1
               grilla2.TypeHAlign = 0
               grilla2.Font.Bold = True
               grilla2.Font.Size = 9
               grilla2.Value = Trim(ConSql!descripcion)
            Else
               grilla1.Col = j + 1
               grilla1.CellType = 5
               grilla1.TypeHAlign = 0
               grilla1.Font.Bold = True
               grilla1.Font.Size = 9
               grilla1.Value = Trim(ConSql!descripcion)
                       
               grilla2.Col = j + 1
               grilla2.CellType = 5
               grilla2.TypeHAlign = 0
               grilla2.Font.Bold = True
               grilla2.Font.Size = 9
               grilla2.Value = Trim(ConSql!descripcion)
                     
               grilla1.Col = j + 3
               grilla1.Value = 0
                     
               grilla1.Col = j + 4
               grilla1.Value = 6
                      
               grilla2.Col = j + 4
               grilla2.Value = 6
          
               grilla1.Col = j + 5
               grilla1.Value = 0
                     
               grilla2.Col = j + 5
               grilla2.Value = 0
          
               grilla1.Col = j + 6
               grilla1.Value = ""
          
            End If
        End Select
        ConSql.MoveNext
     Loop
  End If
  ConSql.Close: Set ConSql = Nothing: fg_descarga

grilla1.Row = 1: grilla1.Col = 1
End Sub
