VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form M_Receta 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receta"
   ClientHeight    =   10605
   ClientLeft      =   2085
   ClientTop       =   2295
   ClientWidth     =   19755
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
   MaxButton       =   0   'False
   ScaleHeight     =   10605
   ScaleWidth      =   19755
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":031A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":094E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":0F82
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":18D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":1BEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":1F04
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":221E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2538
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2852
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2B6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":2E86
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":31A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":34BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":37D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":3AF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":3E0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4124
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":443E
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4758
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4A72
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":4D8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":50A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":53C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":56DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":59F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "M_Receta.frx":5D0E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9975
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   19575
      _ExtentX        =   34528
      _ExtentY        =   17595
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   6
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Receta"
      TabPicture(0)   =   "M_Receta.frx":6028
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Receta.frx":6044
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1(3)"
      Tab(1).Control(1)=   "Frame3(0)"
      Tab(1).Control(2)=   "Frame1(2)"
      Tab(1).Control(3)=   "ImageList4"
      Tab(1).Control(4)=   "Toolbar2"
      Tab(1).Control(5)=   "vaSpread1(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Metodos Preparación"
      TabPicture(2)   =   "M_Receta.frx":6060
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame4"
      Tab(2).Control(1)=   "Frame5(0)"
      Tab(2).Control(2)=   "Label3(2)"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Grupo Vulnerable"
      TabPicture(3)   =   "M_Receta.frx":607C
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label3(1)"
      Tab(3).Control(1)=   "Frame5(1)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "Hipersensibilidad Alimentaria"
      TabPicture(4)   =   "M_Receta.frx":6098
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5(2)"
      Tab(4).Control(1)=   "Label3(3)"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Index           =   2
         Left            =   -74640
         TabIndex        =   56
         Top             =   1080
         Width           =   11655
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   4335
            Index           =   2
            Left            =   420
            TabIndex        =   57
            Top             =   405
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   7646
            _Version        =   393217
            BackColor       =   -2147483624
            BorderStyle     =   0
            HideSelection   =   0   'False
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"M_Receta.frx":60B4
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   -74505
         TabIndex        =   36
         Top             =   495
         Width           =   10095
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            Left            =   675
            TabIndex        =   38
            Text            =   "Combo2"
            Top             =   270
            Width           =   2055
         End
         Begin VB.ComboBox Combo2 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            Left            =   2790
            TabIndex        =   37
            Text            =   "Combo2"
            Top             =   270
            Width           =   1005
         End
         Begin MSComctlLib.Toolbar Toolbar3 
            Height          =   360
            Left            =   3915
            TabIndex        =   39
            Top             =   225
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   635
            ButtonWidth     =   609
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Appearance      =   1
            Style           =   1
            ImageList       =   "ImageList4"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   10
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Negrilla"
                  ImageIndex      =   8
                  Style           =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Italica"
                  ImageIndex      =   9
                  Style           =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Subrayado"
                  ImageIndex      =   10
                  Style           =   1
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Style           =   3
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Alinear a la Izquierda"
                  ImageIndex      =   11
                  Style           =   1
                  Value           =   1
               EndProperty
               BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Centrar"
                  ImageIndex      =   12
                  Style           =   1
               EndProperty
               BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Alinear a la Derecha"
                  ImageIndex      =   13
                  Style           =   1
               EndProperty
               BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Justificar"
                  ImageIndex      =   14
                  Style           =   1
               EndProperty
               BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
               BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Object.ToolTipText     =   "Vriñeta"
                  ImageIndex      =   15
                  Style           =   1
               EndProperty
            EndProperty
            BorderStyle     =   1
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fonts"
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
            Left            =   135
            TabIndex        =   40
            Top             =   315
            Width           =   480
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1095
         Index           =   1
         Left            =   2760
         TabIndex        =   29
         Top             =   480
         Width           =   13575
         Begin VB.ComboBox Combo3 
            Height          =   315
            Index           =   0
            Left            =   10320
            Style           =   2  'Dropdown List
            TabIndex        =   81
            Top             =   360
            Width           =   1455
         End
         Begin VB.CheckBox Check2 
            Caption         =   "No Mostrar Recetas No Vigentes"
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
            Left            =   10320
            TabIndex        =   62
            Top             =   720
            Value           =   1  'Checked
            Width           =   3015
         End
         Begin EditLib.fpText fpTnombre 
            Height          =   315
            Left            =   2640
            TabIndex        =   30
            Top             =   330
            Visible         =   0   'False
            Width           =   2505
            _Version        =   196608
            _ExtentX        =   4410
            _ExtentY        =   870
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
            ThreeDOutsideHighlightColor=   -2147483628
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   -1  'True
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
            AutoBeep        =   0   'False
            AutoCase        =   0
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   2
            MarginTop       =   2
            MarginRight     =   2
            MarginBottom    =   2
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   -1  'True
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime FpFecDesde 
            Height          =   315
            Left            =   12000
            TabIndex        =   83
            Top             =   360
            Width           =   1305
            _Version        =   196608
            _ExtentX        =   2302
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
            ButtonStyle     =   2
            ButtonWidth     =   0
            ButtonWrap      =   -1  'True
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
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
            Text            =   "01/09/2013"
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
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Org. Compras"
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
            Left            =   9120
            TabIndex        =   82
            Top             =   420
            Width           =   1155
         End
         Begin VB.Label fpayuda1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   2520
            TabIndex        =   52
            Top             =   720
            Visible         =   0   'False
            Width           =   6015
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Lista Precio Asociada"
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
            Index           =   7
            Left            =   480
            TabIndex        =   51
            Top             =   780
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   " Buscar Texto"
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
            Left            =   1350
            TabIndex        =   32
            Top             =   400
            Visible         =   0   'False
            Width           =   1200
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Label2"
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
            Index           =   10
            Left            =   5280
            TabIndex        =   31
            Top             =   400
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   2550
            TabIndex        =   53
            Top             =   765
            Visible         =   0   'False
            Width           =   6135
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6255
         Left            =   1200
         TabIndex        =   23
         Top             =   1620
         Width           =   16905
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   8
            Left            =   14880
            TabIndex        =   79
            Top             =   5520
            Width           =   1740
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   11
               Left            =   45
               TabIndex        =   80
               Top             =   135
               Width           =   1635
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   7
            Left            =   12120
            TabIndex        =   77
            Top             =   5520
            Width           =   2700
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   10
               Left            =   45
               TabIndex        =   78
               Top             =   135
               Width           =   2595
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   6
            Left            =   10920
            TabIndex        =   75
            Top             =   5520
            Width           =   1140
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   8
               Left            =   45
               TabIndex        =   76
               Top             =   135
               Width           =   1035
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   5
            Left            =   8760
            TabIndex        =   73
            Top             =   5520
            Width           =   900
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   5
               Left            =   45
               TabIndex        =   74
               Top             =   135
               Width           =   795
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   4
            Left            =   6240
            TabIndex        =   71
            Top             =   5520
            Width           =   2460
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   4
               Left            =   45
               TabIndex        =   72
               Top             =   135
               Width           =   2355
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   3
            Left            =   3960
            TabIndex        =   69
            Top             =   5520
            Width           =   2220
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   3
               Left            =   45
               TabIndex        =   70
               Top             =   135
               Width           =   2115
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   2
            Left            =   1080
            TabIndex        =   67
            Top             =   5520
            Width           =   2820
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   2
               Left            =   45
               TabIndex        =   68
               Top             =   135
               Width           =   2715
            End
         End
         Begin VB.Frame Frame3 
            Height          =   435
            Index           =   1
            Left            =   240
            TabIndex        =   65
            Top             =   5520
            Width           =   780
            Begin VB.TextBox TextDet2 
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Index           =   1
               Left            =   45
               TabIndex        =   66
               Top             =   135
               Width           =   675
            End
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4695
            Index           =   0
            Left            =   240
            TabIndex        =   24
            Top             =   720
            Width           =   16575
            _Version        =   393216
            _ExtentX        =   29236
            _ExtentY        =   8281
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   12
            MaxRows         =   20
            SpreadDesigner  =   "M_Receta.frx":612F
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H80000018&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   14745
            Top             =   6030
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Recetas Vigentes"
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
            Left            =   15105
            TabIndex        =   44
            Top             =   6000
            Width           =   1515
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00D9D9FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   12480
            Top             =   6030
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Recetas No Vigentes"
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
            Left            =   12840
            TabIndex        =   43
            Top             =   6000
            Width           =   1815
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
            Index           =   11
            Left            =   2235
            TabIndex        =   28
            Top             =   435
            Width           =   885
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Categoria Dietetica"
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
            Index           =   12
            Left            =   2235
            TabIndex        =   27
            Top             =   180
            Width           =   1650
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   8
            Left            =   3960
            TabIndex        =   26
            Top             =   180
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   9
            Left            =   3960
            TabIndex        =   25
            Top             =   420
            Width           =   450
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5355
         Index           =   3
         Left            =   -74880
         TabIndex        =   14
         Top             =   360
         Width           =   19425
         Begin VB.CheckBox ChAMD 
            Caption         =   "Integra Receta AMD"
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
            Left            =   6000
            TabIndex        =   141
            Top             =   1680
            Width           =   2055
         End
         Begin VB.ComboBox ECoccion 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   5925
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   4920
            Width           =   2205
         End
         Begin VB.ComboBox PSalsa 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   4920
            Width           =   2205
         End
         Begin VB.ComboBox EtiquetadoSello 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   4480
            Width           =   2205
         End
         Begin VB.Frame Frame14 
            Caption         =   "Par. Adicional 2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   15480
            TabIndex        =   123
            Top             =   3000
            Width           =   1575
            Begin VB.ListBox ParAdi2 
               Height          =   1860
               Index           =   0
               ItemData        =   "M_Receta.frx":25B5D
               Left            =   120
               List            =   "M_Receta.frx":25B64
               Style           =   1  'Checkbox
               TabIndex        =   124
               Top             =   240
               Width           =   1350
            End
         End
         Begin VB.Frame Frame13 
            Caption         =   "Par. Adicional 1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   12960
            TabIndex        =   121
            Top             =   3000
            Width           =   2415
            Begin VB.ListBox ParAdi1 
               Height          =   1860
               Index           =   0
               ItemData        =   "M_Receta.frx":25B71
               Left            =   120
               List            =   "M_Receta.frx":25B78
               Style           =   1  'Checkbox
               TabIndex        =   122
               Top             =   240
               Width           =   2190
            End
         End
         Begin VB.Frame Frame12 
            Caption         =   "Estilo Alimentación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   10440
            TabIndex        =   119
            Top             =   3000
            Width           =   2415
            Begin VB.ListBox EstAli 
               Height          =   1860
               Index           =   0
               ItemData        =   "M_Receta.frx":25B85
               Left            =   120
               List            =   "M_Receta.frx":25B8C
               Style           =   1  'Checkbox
               TabIndex        =   120
               Top             =   240
               Width           =   2190
            End
         End
         Begin VB.Frame Frame11 
            Caption         =   "Alergeno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2295
            Left            =   8160
            TabIndex        =   117
            Top             =   3000
            Width           =   2175
            Begin VB.ListBox Alergeno 
               Height          =   1860
               Index           =   0
               ItemData        =   "M_Receta.frx":25B98
               Left            =   120
               List            =   "M_Receta.frx":25B9F
               Style           =   1  'Checkbox
               TabIndex        =   118
               Top             =   240
               Width           =   1950
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Intolerancia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2535
            Left            =   17160
            TabIndex        =   113
            Top             =   120
            Width           =   2055
            Begin VB.ListBox Intolerancia 
               Height          =   2085
               Index           =   0
               ItemData        =   "M_Receta.frx":25BAD
               Left            =   120
               List            =   "M_Receta.frx":25BB4
               Style           =   1  'Checkbox
               TabIndex        =   114
               Top             =   240
               Width           =   1695
            End
         End
         Begin VB.Frame Frame1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   2595
            Index           =   0
            Left            =   17160
            TabIndex        =   100
            Top             =   2640
            Width           =   2040
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   2
               Left            =   1080
               TabIndex        =   101
               Top             =   915
               Width           =   855
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   1323
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "900000"
               MinValue        =   "-900000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   1
               Left            =   1095
               TabIndex        =   102
               Top             =   1905
               Width           =   855
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   1323
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               DecimalPlaces   =   2
               DecimalPoint    =   ""
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "-9000000000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   0
               Left            =   1080
               TabIndex        =   103
               Top             =   270
               Width           =   855
               _Version        =   196608
               _ExtentX        =   1508
               _ExtentY        =   556
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               Text            =   "1"
               DecimalPlaces   =   0
               DecimalPoint    =   ""
               FixedPoint      =   0   'False
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "-9000000000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ""
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   3
               Left            =   1095
               TabIndex        =   104
               Top             =   1575
               Width           =   855
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   1323
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               DecimalPlaces   =   2
               DecimalPoint    =   ""
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "-9000000000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   4
               Left            =   1095
               TabIndex        =   105
               Top             =   1245
               Width           =   855
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   1323
               Enabled         =   -1  'True
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "900000"
               MinValue        =   "-900000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   5
               Left            =   1080
               TabIndex        =   106
               Top             =   600
               Width           =   855
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   1323
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               DecimalPlaces   =   2
               DecimalPoint    =   "."
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "900000"
               MinValue        =   "-900000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin EditLib.fpDoubleSingle fpDouble1 
               Height          =   315
               Index           =   6
               Left            =   1095
               TabIndex        =   143
               Top             =   2240
               Width           =   855
               _Version        =   196608
               _ExtentX        =   2646
               _ExtentY        =   1323
               Enabled         =   0   'False
               MousePointer    =   0
               Object.TabStop         =   -1  'True
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BackColor       =   -2147483624
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
               ButtonDefaultAction=   -1  'True
               ThreeDText      =   0
               ThreeDTextHighlightColor=   -2147483637
               ThreeDTextShadowColor=   -2147483632
               ThreeDTextOffset=   1
               AlignTextH      =   2
               AlignTextV      =   0
               AllowNull       =   -1  'True
               NoSpecialKeys   =   0
               AutoAdvance     =   0   'False
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
               DecimalPlaces   =   2
               DecimalPoint    =   ""
               FixedPoint      =   -1  'True
               LeadZero        =   0
               MaxValue        =   "9000000000"
               MinValue        =   "-9000000000"
               NegFormat       =   1
               NegToggle       =   0   'False
               Separator       =   ","
               UseSeparator    =   0   'False
               IncInt          =   1
               IncDec          =   1
               BorderGrayAreaColor=   -2147483637
               ThreeDOnFocusInvert=   0   'False
               ThreeDFrameColor=   -2147483637
               Appearance      =   0
               BorderDropShadow=   1
               BorderDropShadowColor=   -2147483632
               BorderDropShadowWidth=   3
               ButtonColor     =   -2147483637
               AutoMenu        =   0   'False
               ButtonAlign     =   0
               OLEDropMode     =   0
               OLEDragMode     =   0
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "H.Carbono"
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
               TabIndex        =   142
               Top             =   2280
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Raciones"
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
               Left            =   90
               TabIndex        =   112
               Top             =   375
               Width           =   765
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C. Servida"
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
               Left            =   90
               TabIndex        =   111
               Top             =   1990
               Width           =   840
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "G. Neto Nut."
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
               Left            =   90
               TabIndex        =   110
               Top             =   1000
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C. Bruta"
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
               Left            =   90
               TabIndex        =   109
               Top             =   1680
               Width           =   780
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "C. Serv. M."
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
               Index           =   13
               Left            =   90
               TabIndex        =   108
               Top             =   1305
               Width           =   870
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "G. Neto"
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
               Index           =   14
               Left            =   90
               TabIndex        =   107
               Top             =   710
               Width           =   600
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Zona"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   15480
            TabIndex        =   98
            Top             =   120
            Width           =   1575
            Begin VB.CommandButton Command1 
               Caption         =   "Todos"
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
               Left            =   360
               TabIndex        =   125
               Top             =   240
               Width           =   855
            End
            Begin VB.ListBox Zona 
               Height          =   2085
               Index           =   0
               ItemData        =   "M_Receta.frx":25BC6
               Left            =   120
               List            =   "M_Receta.frx":25BCD
               Style           =   1  'Checkbox
               TabIndex        =   99
               Top             =   480
               Width           =   1335
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Tipo Negocio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   12960
            TabIndex        =   96
            Top             =   120
            Width           =   2415
            Begin VB.ListBox TipoNegocio 
               Height          =   2310
               Index           =   0
               ItemData        =   "M_Receta.frx":25BD7
               Left            =   120
               List            =   "M_Receta.frx":25BDE
               Style           =   1  'Checkbox
               TabIndex        =   97
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.ComboBox Sellos 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   3660
            Width           =   2205
         End
         Begin VB.ComboBox EfectoMeteorizante 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   5925
            Style           =   2  'Dropdown List
            TabIndex        =   93
            Top             =   4060
            Width           =   2205
         End
         Begin VB.ComboBox CatCompleja 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   5925
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   3660
            Width           =   2205
         End
         Begin VB.ComboBox Costo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   0
            ItemData        =   "M_Receta.frx":25BEF
            Left            =   1800
            List            =   "M_Receta.frx":25BF1
            Style           =   2  'Dropdown List
            TabIndex        =   86
            Top             =   2420
            Width           =   2205
         End
         Begin VB.Frame Frame7 
            Caption         =   "Estacionalidad"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   10440
            TabIndex        =   63
            Top             =   120
            Width           =   2415
            Begin VB.ListBox List1 
               Height          =   2310
               Index           =   0
               ItemData        =   "M_Receta.frx":25BF3
               Left            =   120
               List            =   "M_Receta.frx":25BF5
               Style           =   1  'Checkbox
               TabIndex        =   64
               Top             =   240
               Width           =   2175
            End
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   315
            Index           =   0
            ItemData        =   "M_Receta.frx":25BF7
            Left            =   1800
            List            =   "M_Receta.frx":25BF9
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   2040
            Width           =   2205
         End
         Begin VB.Frame Frame6 
            Caption         =   " Ofertas"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   8160
            TabIndex        =   59
            Top             =   120
            Width           =   2175
            Begin VB.ListBox List1 
               Height          =   2310
               Index           =   1
               ItemData        =   "M_Receta.frx":25BFB
               Left            =   120
               List            =   "M_Receta.frx":25BFD
               Style           =   1  'Checkbox
               TabIndex        =   60
               Top             =   240
               Width           =   1950
            End
         End
         Begin VB.ComboBox Combo2 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   5925
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2040
            Width           =   2205
         End
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Caption         =   "Fecha Vencimiento"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   41
            Top             =   1710
            Width           =   1950
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   2040
            TabIndex        =   15
            Top             =   240
            Width           =   6045
            _Version        =   196608
            _ExtentX        =   10663
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483624
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            MaxLength       =   80
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   16
            Top             =   600
            Width           =   6045
            _Version        =   196608
            _ExtentX        =   10663
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483624
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            MaxLength       =   80
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   3
            Left            =   2460
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1305
            Width           =   5625
            _Version        =   196608
            _ExtentX        =   9922
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483638
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   2
            Left            =   2460
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   990
            Width           =   5625
            _Version        =   196608
            _ExtentX        =   9922
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483638
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Index           =   0
            Left            =   2340
            TabIndex        =   42
            Top             =   1655
            Width           =   1500
            _Version        =   196608
            _ExtentX        =   2646
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483624
            ForeColor       =   -2147483640
            ThreeDInsideStyle=   0
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   0
            ThreeDOutsideHighlightColor=   -2147483628
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
            ThreeDTextHighlightColor=   -2147483633
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   0
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   2
            MarginTop       =   2
            MarginRight     =   2
            MarginBottom    =   2
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "12/10/2004"
            DateCalcMethod  =   0
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   0
            Left            =   5925
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   2880
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   1
            Left            =   5925
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   3240
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpLongInteger fpSIP1 
            Height          =   315
            Index           =   0
            Left            =   4800
            TabIndex        =   135
            Top             =   3240
            Width           =   705
            _Version        =   196608
            _ExtentX        =   1244
            _ExtentY        =   556
            Enabled         =   -1  'True
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   4
            Left            =   1800
            TabIndex        =   136
            TabStop         =   0   'False
            Top             =   2840
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   5
            Left            =   1800
            TabIndex        =   137
            TabStop         =   0   'False
            Top             =   3240
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   6
            Left            =   1800
            TabIndex        =   138
            TabStop         =   0   'False
            Top             =   4080
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   7
            Left            =   5925
            TabIndex        =   139
            TabStop         =   0   'False
            Top             =   2520
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fpayuda 
            Height          =   315
            Index           =   8
            Left            =   5925
            TabIndex        =   140
            TabStop         =   0   'False
            Top             =   4480
            Width           =   2145
            _Version        =   196608
            _ExtentX        =   3784
            _ExtentY        =   556
            Enabled         =   0   'False
            MousePointer    =   0
            Object.TabStop         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483648
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
            ButtonDefaultAction=   -1  'True
            ThreeDText      =   0
            ThreeDTextHighlightColor=   -2147483637
            ThreeDTextShadowColor=   -2147483632
            ThreeDTextOffset=   1
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   0
            AutoAdvance     =   0   'False
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
            ControlType     =   3
            Text            =   ""
            CharValidationText=   ""
            MaxLength       =   255
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
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   8
            Left            =   5400
            Picture         =   "M_Receta.frx":25BFF
            Top             =   4440
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   7
            Left            =   5400
            Picture         =   "M_Receta.frx":25F09
            Top             =   2400
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   6
            Left            =   1200
            Picture         =   "M_Receta.frx":26213
            Top             =   3990
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   5
            Left            =   1200
            Picture         =   "M_Receta.frx":2651D
            Top             =   3120
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   4
            Left            =   1200
            Picture         =   "M_Receta.frx":26827
            Top             =   2760
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   1
            Left            =   5400
            Picture         =   "M_Receta.frx":26B31
            Top             =   3120
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   5400
            Picture         =   "M_Receta.frx":26E3B
            Top             =   2760
            Width           =   480
         End
         Begin VB.Label Label16 
            Caption         =   "S. I. P."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4095
            TabIndex        =   132
            Top             =   3315
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "E. Cocción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   3
            Left            =   4095
            TabIndex        =   130
            Top             =   4990
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "P. Salsa"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   128
            Top             =   4990
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Etiq. Sello"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   126
            Top             =   4570
            Width           =   975
         End
         Begin VB.Label Label15 
            Caption         =   "T. HH"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   116
            Top             =   4140
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "T. Cocción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   4095
            TabIndex        =   115
            Top             =   4570
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Sellos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   60
            TabIndex        =   94
            Top             =   3750
            Width           =   615
         End
         Begin VB.Label Label12 
            Caption         =   "Efecto Meteorizante"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4095
            TabIndex        =   92
            Top             =   4180
            Width           =   1815
         End
         Begin VB.Label Label11 
            Caption         =   "Ing. C. Gar."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   91
            Top             =   3315
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Cat. Compleja"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4095
            TabIndex        =   89
            Top             =   3750
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "M. Cocción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   88
            Top             =   2940
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "T. Ing. P."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   4095
            TabIndex        =   87
            Top             =   2940
            Width           =   855
         End
         Begin VB.Label Label7 
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
            Height          =   180
            Left            =   60
            TabIndex        =   85
            Top             =   2530
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Color"
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
            Left            =   4095
            TabIndex        =   84
            Top             =   2530
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "Unidad Receta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   60
            TabIndex        =   61
            Top             =   2145
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Receta"
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
            Index           =   23
            Left            =   4095
            TabIndex        =   55
            Top             =   2145
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Receta"
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
            Left            =   60
            TabIndex        =   22
            Top             =   330
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre Fantasia"
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
            Left            =   60
            TabIndex        =   21
            Top             =   705
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Categoria Dietetica"
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
            Index           =   8
            Left            =   60
            TabIndex        =   20
            Top             =   1065
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo Plato"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   60
            TabIndex        =   19
            Top             =   1365
            Width           =   1695
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   2
            Left            =   1920
            Picture         =   "M_Receta.frx":27145
            Top             =   900
            Width           =   480
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   3
            Left            =   1920
            Picture         =   "M_Receta.frx":2744F
            Top             =   1215
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Aporte Nutricionales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4110
         Index           =   0
         Left            =   -58440
         TabIndex        =   12
         Top             =   5760
         Width           =   2885
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   3660
            Left            =   165
            TabIndex        =   13
            Top             =   360
            Width           =   2550
            _Version        =   393216
            _ExtentX        =   4498
            _ExtentY        =   6456
            _StockProps     =   64
            DisplayRowHeaders=   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   3
            MaxRows         =   15
            SpreadDesigner  =   "M_Receta.frx":27759
            ScrollBarTrack  =   3
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Index           =   2
         Left            =   -66720
         TabIndex        =   3
         Top             =   9240
         Width           =   7890
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   5
            Left            =   3255
            TabIndex        =   11
            Top             =   190
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. : "
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
            Left            =   2295
            TabIndex        =   10
            Top             =   195
            Width           =   870
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Costo : "
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
            Left            =   5940
            TabIndex        =   9
            Top             =   195
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Gr.Net.Verd. : "
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
            Left            =   120
            TabIndex        =   8
            Top             =   195
            Width           =   1230
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   3
            Left            =   6720
            TabIndex        =   7
            Top             =   195
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   4
            Left            =   1320
            TabIndex        =   6
            Top             =   190
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "P.A.V.B. % : "
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
            Left            =   4140
            TabIndex        =   5
            Top             =   195
            Width           =   1080
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "0.00"
            Height          =   315
            Index           =   7
            Left            =   5200
            TabIndex        =   4
            Top             =   190
            Width           =   495
         End
      End
      Begin MSComctlLib.ImageList ImageList4 
         Left            =   -66510
         Top             =   5790
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":27BC3
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":27EDD
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":281F7
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":28511
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":2882B
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":28B45
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":28E5F
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":29179
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":29493
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":297AD
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":29AC7
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":29C21
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":29D7B
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":29ED5
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_Receta.frx":2A02F
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   390
         Left            =   -75000
         TabIndex        =   33
         Top             =   9360
         Width           =   8115
         _ExtentX        =   14314
         _ExtentY        =   688
         ButtonWidth     =   2858
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         TextAlignment   =   1
         ImageList       =   "ImageList4"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Ing.    "
               Description     =   "Agregar Ingrediente"
               Object.ToolTipText     =   "Agregar Ingrediente"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Agrega Linea    "
               Description     =   "Insertar Linea"
               Object.ToolTipText     =   "Insertar Linea"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Borra Linea    "
               Description     =   "Borrar Ingrediente"
               Object.ToolTipText     =   "Borrar Ingrediente"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Up    "
               Description     =   "Mueve Up"
               Object.ToolTipText     =   "Mueve Up"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Mueve Dn    "
               Description     =   "Mueve Dn"
               Object.ToolTipText     =   "Mueve Dn"
               ImageIndex      =   5
            EndProperty
         EndProperty
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   3420
         Index           =   1
         Left            =   -74880
         TabIndex        =   34
         Top             =   5940
         Width           =   15255
         _Version        =   393216
         _ExtentX        =   26908
         _ExtentY        =   6033
         _StockProps     =   64
         ColsFrozen      =   2
         DisplayRowHeaders=   0   'False
         EditEnterAction =   5
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
         GridShowHoriz   =   0   'False
         MaxCols         =   15
         MaxRows         =   200
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Receta.frx":2A189
         VisibleCols     =   10
         VisibleRows     =   40
         ScrollBarTrack  =   3
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4695
         Index           =   0
         Left            =   -74640
         TabIndex        =   45
         Top             =   1320
         Width           =   10455
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   4455
            Index           =   0
            Left            =   180
            TabIndex        =   46
            Top             =   165
            Width           =   10095
            _ExtentX        =   17806
            _ExtentY        =   7858
            _Version        =   393217
            BackColor       =   -2147483624
            BorderStyle     =   0
            HideSelection   =   0   'False
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"M_Receta.frx":2BDFF
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Index           =   1
         Left            =   -74640
         TabIndex        =   47
         Top             =   960
         Width           =   11655
         Begin RichTextLib.RichTextBox RichTextBox1 
            Height          =   4335
            Index           =   1
            Left            =   420
            TabIndex        =   48
            Top             =   405
            Width           =   10815
            _ExtentX        =   19076
            _ExtentY        =   7646
            _Version        =   393217
            BackColor       =   -2147483624
            BorderStyle     =   0
            HideSelection   =   0   'False
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            AutoVerbMenu    =   -1  'True
            TextRTF         =   $"M_Receta.frx":2BE7C
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
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
         Left            =   -74640
         TabIndex        =   58
         Top             =   840
         Width           =   555
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
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
         Left            =   -74640
         TabIndex        =   50
         Top             =   6120
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
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
         Left            =   -74640
         TabIndex        =   49
         Top             =   720
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
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
         Left            =   -74520
         TabIndex        =   35
         Top             =   6060
         Width           =   585
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   19755
      _ExtentX        =   34846
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "M_Receta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private RS                              As New ADODB.Recordset
Private RS1                             As New ADODB.Recordset
Private itab                            As Integer
Private i                               As Long
Private itexto                          As Integer
Private indlec                          As Integer
Private ibusca                          As Long
Private codigo                          As Long
Private CodCatDie                       As Long
Private codtipplato                     As Long
Private Est                             As Boolean
Private EstOrg                          As Boolean
Private vecdatos(1)                     As String
Private vecdatos1(5)
Private rec_indppr                      As String
Private MsgTitulo                       As String
Private codpro1                         As String
Private codpro2                         As String
Private modo                            As String
Private metodoreceta                    As String
Private nombusca                        As String
Private grupovulnerable                 As String
Private HipersensabilidadAlimentaria    As String
Private canpro1                         As Double
Private canpro2                         As Double
Private pctapr1                         As Double
Private pctapr2                         As Double
Private pctcoc1                         As Double
Private pctcoc2                         As Double
Private pctnut1                         As Double
Private pctnut2                         As Double
Private cospro                          As Double
Private candiet                         As Double
Private cosrec                          As Double
Private Indmetpre                       As Integer
Private Indgruvul                       As Integer
Private ComboValOrig                    As String
Private Index                           As Long
Private TipoIngPrincipal                As Long
Private MetodoCoccion                   As Long
Private IngCruceGarnitura               As Long
Private TiempoHH                        As Long
Private Color                           As Long
Private TiempoCoccion                   As Long

Private CantRecordCosto                 As Long
Private CantRecordTiPrincipal           As Long
Private CantRecordSIPrincipal           As Long
Private CantRecordECoccion              As Long
Private CantRecordPSalsa                As Long
Private CantRecordMetodoCoccion         As Long
Private CantRecordCatCompleja           As Long
Private CantRecordIngCruceGar           As Long
Private CantRecordEfectoMeteorizante    As Long
Private CantRecordSellos                As Long
Private CantRecordEtiquetadoSello       As Long
Private CantRecordTCoccion              As Long
Private CantRecordTHH                   As Long
Private CantRecordColor                 As Long
Private CantRecordOferta                As Long
Private CantRecordEstacionalidad        As Long
Private CantRecordTipoNegocio           As Long
Private CantRecordZona                  As Long
Private CantRecordIntolerancia          As Long

Private Sub ConfiControlesReceta(Index As Integer, habilita As Boolean)
  
  Select Case Index
    
    Case 1
        
        Frame1(3).Enabled = IIf(habilita = True, True, False)
        fpText1(0).Enabled = IIf(habilita = True, True, False)
        fpText1(1).Enabled = IIf(habilita = True, True, False)
        Image1(2).Enabled = IIf(habilita = True, True, False)
        Image1(3).Enabled = IIf(habilita = True, True, False)
        Check1(0).Enabled = IIf(habilita = True, True, False)
        fpDateTime1(0).Enabled = IIf(habilita = True, True, False)
        Combo2(2).Enabled = IIf(habilita = True, True, False)
        
        If VarSitioRemoto = False Then
            
            Let Toolbar2.Enabled = IIf(habilita = True, True, False)
        
        Else
            
            Let Toolbar2.Enabled = False
        
        End If
        
        Frame4.Enabled = IIf(habilita = True, True, False)
        Frame5(0).Enabled = IIf(habilita = True, True, False)
'        Frame5(1).Enabled = IIf(habilita = True, True, False)
        RichTextBox1(0).Locked = IIf(habilita = True, False, True)
'        RichTextBox1(0).Enabled = IIf(habilita = True, True, False)
        Combo2(0).Enabled = IIf(habilita = True, True, False)
        Combo2(1).Enabled = IIf(habilita = True, True, False)
        Toolbar3.Enabled = IIf(habilita = True, True, False)
'        RichTextBox1(1).Enabled = IIf(habilita = True, True, False)
'        RichTextBox1(1).Locked = IIf(habilita = True, False, True)
'        RichTextBox1(2).Locked = IIf(habilita = True, False, True)
        fpDouble1(4).Enabled = IIf(habilita = True, True, False)
  
  End Select

End Sub

Private Sub Alergeno_Click(Index As Integer)

On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub CatCompleja_Click(Index As Integer)
    
On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo

If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then

   Hab_Des 0

Else

   Exit Sub
 
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub ChAMD_Click()
   
   If itexto = 0 And modo = "M" Then
      
      Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
      If Check1(0).Value = 1 Then fpDateTime1(0).Enabled = True: fpDateTime1(0).text = Format(Date, "dd/mm/yyyy") Else fpDateTime1(0).Enabled = False: fpDateTime1(0).text = "  /  /    "
   
   End If

End Sub

Private Sub Check1_Click(Index As Integer)

Select Case Index

Case 0
   
   If itexto = 0 And modo = "M" Then
      
      Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
      If Check1(0).Value = 1 Then fpDateTime1(0).Enabled = True: fpDateTime1(0).text = Format(Date, "dd/mm/yyyy") Else fpDateTime1(0).Enabled = False: fpDateTime1(0).text = "  /  /    "
   
   End If

End Select

End Sub

Private Sub Check2_Click()

On Error GoTo Man_Error

Mover_ListaReceta

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Combo1_Click(Index As Integer)
    
On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0 Else Exit Sub
 
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
 
End Sub

Private Sub Combo2_Click(Index As Integer)

On Error GoTo Man_Error

If Est Then Exit Sub

Select Case Index

Case 0
    
    RichTextBox1(0).SelFontName = Combo2(0).text

Case 1
    
    RichTextBox1(0).SelFontSize = Combo2(1).text

Case 2
      
      If modo = "" Then modo = "M"
      
      If vg_Indppr = 3 Or vg_Indppr = 1 Then
        
        If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo:  If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
      
      End If
      
      If vg_Indppr = 2 And ComboValOrig = vg_Indppr Then
        
        If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
          
          Gl_Ac_Botones Me, 3, 0, modo:
          Combo2(2).ListIndex = fg_buscacbo(Combo2, 2, 1, (ComboValOrig))
          Combo2(2).Enabled = False
          
          If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
        
        End If
      
      End If

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
 
End Sub

Private Sub Combo3_Click(Index As Integer)

On Error GoTo Man_Error

If EstOrg Then Exit Sub

Mover_ListaReceta

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim i As Long

For i = 0 To Zona(0).ListCount - 1
   
    Zona(0).Selected(i) = True
        
Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
 
End Sub

Private Sub Costo_Click(Index As Integer)
    
On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo

If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then

   Hab_Des 0
   
Else
   
   Exit Sub

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub ECoccion_Click(Index As Integer)

On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo

If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then

   Hab_Des 0

Else

   Exit Sub

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub EfectoMeteorizante_Click(Index As Integer)
    
On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo

If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then

   Hab_Des 0

Else

   Exit Sub

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub EstAli_Click(Index As Integer)

On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub EtiquetadoSello_Click(Index As Integer)

On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo

If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then

   Hab_Des 0

Else
   
   Exit Sub
 
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
 
End Sub

Private Sub Form_Activate()

fg_descarga

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Me.Height = 11040
Me.Width = 19845
fg_centra Me
Me.HelpContextID = vg_OpcM
MsgTitulo = "Recetas"
Est = True
EstOrg = True

CantRecordCosto = 0
CantRecordTiPrincipal = 0
CantRecordSIPrincipal = 0
CantRecordECoccion = 0
CantRecordPSalsa = 0
CantRecordMetodoCoccion = 0
CantRecordCatCompleja = 0
CantRecordIngCruceGar = 0
CantRecordEfectoMeteorizante = 0
CantRecordSellos = 0
CantRecordEtiquetadoSello = 0
CantRecordTCoccion = 0
CantRecordTHH = 0
CantRecordColor = 0
CantRecordOferta = 0
CantRecordEstacionalidad = 0
CantRecordTipoNegocio = 0
CantRecordZona = 0
CantRecordIntolerancia = 0

'-------> Llenar palabra
With Combo2(0)
    
    .AddItem Screen.Fonts(0)
    For i = 1 To Screen.FontCount - 1
        
        .AddItem Screen.Fonts(i)
    
    Next
    .ListIndex = 0

End With

With Combo2(2)
    
    .Clear
    .AddItem "Real" & Space(150) & "(1)"
    .AddItem "Propuesta" & Space(150) & "(2)"

End With

'-------> Ini : Cargar Org. Compras
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_OrgCompras_V02 ''")
Combo3(0).Clear
Do While Not RS.EOF
   
   Combo3(0).AddItem Trim(RS!ID_Orgcompra) & Space(150) & "(" & RS!ID_Orgcompra & ")"
   RS.MoveNext

Loop
RS.Close
Set RS = Nothing
Combo3(0).ListIndex = -1
'-------> Fin : Cargar Org. Compras

'-------> Traer Parametro
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_Parametros 'parcosrec'")
If Not RS.EOF Then Combo3(0).ListIndex = fg_buscacbostring(Combo3, 0, 4, (RS!par_valor))
RS.Close
Set RS = Nothing

'-------> Llenar tamaño
With Combo2(1)
    
    .AddItem 8
    For i = 9 To 72
        
        .AddItem i
    
    Next i
    .ListIndex = 0

End With

Est = False
'-------> Mover nutrientes
With vaSpread2
    
    Dim IndApo As Long
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_s_nutriente 1, 0, ''")
    IndApo = 1
    If Not RS.EOF Then
       
       .MaxRows = 0
       .MaxRows = RS.RecordCount
       Do While Not RS.EOF
          
          DoEvents
'          .MaxRows = .MaxRows + 1
          .Row = IndApo '.MaxRows
                   
          .Col = 1
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!nut_codigo)
                   
          .Col = 2
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!nut_nombre)
                   
          .Col = 3
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignRight
          .text = Format(0, fg_Pict(6, 2))
          .ForeColor = &HFF0000
          
          IndApo = IndApo + 1
          RS.MoveNext
       
       Loop
    
    End If
    RS.Close
    Set RS = Nothing

End With

itexto = 1
vg_dbndecimal = 2
modo = ""
Gl_Mo_Botones Me, 3
Gl_Ac_Botones Me, 3, 3, modo
Hab_Des 3
vg_filcatdie = 0
vg_filtippla = 0

If vg_newcodrec > 0 Then
   
   Me.HelpContextID = "1030000"
   If Not vg_modreceta Then vg_modreceta = IIf(Mid(ValidarUsuario(Me), 2, 1) = 0, True, False)
    
    modo = "M"
    Hab_Des 0
    SSTab1.Tab = 1
    SSTab1.TabEnabled(2) = False
    Gl_Ac_Botones Me, 3, 3, modo
    MoverDetalleDatos
    
    If vg_newestrec = True Then
        
        Frame1(0).Enabled = False
        Frame1(3).Enabled = False
        Let Toolbar2.Enabled = False
        modo = "M"
    
    End If
    
    SSTab1.TabEnabled(2) = True
    SSTab1.TabEnabled(3) = True
    Exit Sub

End If

'------> Mover parametro categoria receta
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_parametro 2, 'catdefecto-" & Trim(vg_NUsr) & "' , ''")

If Not RS.EOF Then vg_filcatdie = RS!par_valor: Label2(8).Caption = fg_BuscaenArbol(RS!par_valor, "a_recetacatdie", "car_codigo")
RS.Close
Set RS = Nothing

'-------> Mover parametro lista precio
Vg_FechaDesde = 0
vg_codlpr = 0
fpayuda1(3).Caption = ""

'If RS.State = 1 Then RS.Close
'RS.CursorLocation = adUseClient
'vg_db.CursorLocation = adUseClient
'Set RS = vg_db.Execute("sgpadm_s_listaprecio 8, 0, 0, '" & vg_NUsr & "'")
'If Not RS.EOF Then
'
'   fpayuda1(3).Caption = "(" & RS!lpr_codigo & ") " & Trim(RS!lpr_nombre) & " " & Mid(RS!dlp_anomes, 5, 2) & "/" & Mid(RS!dlp_anomes, 1, 4)
'   vg_codlpr = RS!lpr_codigo
'   Vg_FechaDesde = RS!dlp_anomes
'
'End If
'RS.Close
'Set RS = Nothing
FpFecDesde = Format(Date, "dd/mm/yyyy")
Mover_ListaReceta

With SSTab1
    
    .Tab = 0
    '-------> Activar carpetar Grupo Vulnerable, Hipersensibilidad y boton buscar Ingrediente
    Me.HelpContextID = 1091000
    .TabVisible(3) = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    RichTextBox1(1).Locked = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", True, False)
    Me.HelpContextID = 1092000
    .TabVisible(4) = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    RichTextBox1(2).Locked = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", True, False)
    Me.HelpContextID = 1093000
'    Toolbar1.Buttons(19).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Toolbar1.Buttons(19).ButtonMenus(1).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = 1094000
    Toolbar1.Buttons(19).ButtonMenus(2).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)

End With

Me.HelpContextID = vg_OpcM

EstOrg = False

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpDateTime1_Change(Index As Integer)

On Error GoTo Man_Error

If IsDate(fpDateTime1(Index).text) = False Then Exit Sub

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpDouble1_Change(Index As Integer)

On Error GoTo Man_Error

If Index = 6 Then

   Exit Sub
   
End If

If Val(fpDouble1(Index).Value) <> Val(vecdatos1(Index)) And itexto = 0 And modo = "M" Then
   
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0

End If

If Index = 4 Then
 
 If vg_Indppr = 1 Or vg_Indppr = 3 Then
   
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
     
     Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
   
   End If
 
 ElseIf vg_Indppr = 2 And ComboValOrig = vg_Indppr Then
   
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
     
     Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
   
   End If
 
 End If

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpDouble1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub FpFecDesde_Change()

On Error GoTo Man_Error

If IsDate(FpFecDesde.text) = False Then Exit Sub

If EstOrg Then Exit Sub

Mover_ListaReceta

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub FpFecDesde_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error
    
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub fpSIP1_Change(Index As Integer)

On Error GoTo Man_Error

   Dim RS As New ADODB.Recordset
  
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
  
   Set RS = vg_db.Execute("sgpadm_Sel_CodigoTipoIngPrincipalRecetaActivo " & Val(fpSIP1(0).Value) & "")
   fpayuda(1).text = ""
   If Not RS.EOF Then
   
      fpayuda(1).text = Trim(RS!NombreTipoIngPrincipal)
   
   End If
   
   RS.Close
   Set RS = Nothing
   
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
     
     Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
   
   End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    

End Sub

Private Sub fpSIP1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpText1_Change(Index As Integer)

On Error GoTo Man_Error

If fpText1(Index).text <> vecdatos(Index) And itexto = 0 And modo = "M" Then
  
  If vg_Indppr = 1 Or vg_Indppr = 3 Then
   
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
     
     Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
   
   End If
  
  ElseIf vg_Indppr = 2 And vg_Indppr = ComboValOrig Then
   
   If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
     
     Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
   
   End If
  
  End If

End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Sub

Private Sub fpTnombre_Change()

Dim estact          As Variant
Dim FindString      As String
Dim SwActiva        As Long
Dim IRow            As Long
Dim SourceString    As String
Dim indactivo       As Long
With vaSpread1(0)
    
    If .MaxRows < 1 Then Exit Sub
    FindString = Trim(fpTnombre.text)
    If fpTnombre.text = "" Then
        
        .Visible = False
        SwActiva = 0
        For i = 1 To .MaxRows
            .Row = i
            .RowHidden = False
            If SwActiva = 0 Then
                .Col = 1
                codigo = Val(.text)
                SwActiva = 1
            End If
        Next i
    
    Else
        
        SwActiva = 0
        .Visible = False
        IRow = 0
        For i = 1 To .MaxRows
            .Row = i
            .Col = 2
            SourceString = Trim(.Value)
            indactivo = UCase(Trim(SourceString)) Like "*" & UCase(FindString) & "*"
            If indactivo = -1 Then
                If SwActiva = 0 Then
                    .Col = 1
                    codigo = Val(.text)
                    SwActiva = 1
                End If
                If .RowHidden = True Then .RowHidden = False
                IRow = IRow + 1
            Else
                If .RowHidden = False Then .RowHidden = True
            End If
        Next i
    
    End If
    
    .Visible = True
    .SetActiveCell 1, 1
    
    With SSTab1
        
        If SwActiva = 1 Then
            
            Label1(10).Caption = Format(IRow, fg_Pict(7, 0)) & " Reg. Encontrados"
            .TabEnabled(1) = True: .TabEnabled(2) = True: .TabEnabled(3) = True
        
        Else
            
            Label1(10).Caption = Format(0, fg_Pict(7, 0)) & " Reg. Encontrados"
            codigo = 0: .TabEnabled(1) = False: .TabEnabled(2) = False:: .TabEnabled(3) = False: Exit Sub
        
        End If
    
    End With

End With

'vaSpread1(0).Visible = False
'Set RS = vg_db.Execute("sgpadm_s_receta_V06 2, 0, '%" & LimpiaDato(UCase(fptnombre.text)) & "%', " & vg_filcatdie & ", " & vg_filtippla & ", 0, '" & vg_NUsr & "'")
'If RS.EOF Or RS!nreg = 0 Then RS.Close: Set RS = Nothing: ibusca = 0: vaSpread1(0).Visible = True: vaSpread1(0).MaxRows = 0: Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Reg. Encontrados": SSTab1.TabEnabled(1) = False: SSTab1.TabEnabled(2) = False: Exit Sub
'If ibusca <> RS!nreg Then ibusca = RS!nreg: vaSpread1(0).MaxRows = RS!nreg
'RS.Close: Set RS = Nothing
'
'Set RS1 = vg_db.Execute("sgpadm_s_receta_V06 23, " & vg_codlpr & ", '%" & LimpiaDato(UCase(fptnombre.text)) & "%',  " & vg_filcatdie & ", " & vg_filtippla & ", " & Val(vg_fechaDESDE) & ", '" & vg_NUsr & "'")
'i = 1
'vaSpread1(0).Row = -1: vaSpread1(0).Col = -1
'vaSpread1(0).BackColor = Shape1(0).FillColor
'If Not RS1.EOF Then
'   Do While Not RS1.EOF
'     DoEvents
'      vaSpread1(0).Row = i
'
'      vaSpread1(0).Col = 1
''     vaSpread1(0).CellType = 5
'      vaSpread1(0).TypeHAlign = TypeHAlignLeft
'      vaSpread1(0).text = RS1!rec_codigo
'
'      vaSpread1(0).Col = -1
'      If RS1!rec_fecvig <= Val(Format(Date, "yyyymmdd")) And RS1!rec_fecvig > 0 Then vaSpread1(0).BackColor = Shape1(1).FillColor
'
'      vaSpread1(0).Col = 2
'      vaSpread1(0).CellType = CellTypeStaticText
'      vaSpread1(0).TypeHAlign = TypeHAlignLeft
'      vaSpread1(0).text = Trim(RS1!rec_nombre)
'
'      vaSpread1(0).Col = 3
'      vaSpread1(0).CellType = CellTypeStaticText
'      vaSpread1(0).TypeHAlign = TypeHAlignRight
'      vaSpread1(0).text = IIf(IsNull(RS1!cosrec) = True, "", Format(CCur(RS1!cosrec), fg_Pict(6, 2)))
'
'      vaSpread1(0).Col = 4
'      vaSpread1(0).CellType = CellTypeStaticText
'      vaSpread1(0).TypeHAlign = TypeHAlignCenter
'      vaSpread1(0).ForeColor = &HFF00&
'      vaSpread1(0).FontBold = True
'      vaSpread1(0).text = IIf(Trim(RS1!rec_metpre) = "1", "X", "")
'
'
'      vaSpread1(0).Col = 5
'      vaSpread1(0).CellType = CellTypeStaticText
'      vaSpread1(0).TypeHAlign = TypeHAlignCenter
'      vaSpread1(0).ForeColor = &HFF00&
'      vaSpread1(0).FontBold = True
'      vaSpread1(0).text = IIf(Trim(RS1!rec_gruvul) = "1", "X", "")
'
'      vaSpread1(0).Col = 6
'      vaSpread1(0).CellType = CellTypeStaticText
'      vaSpread1(0).TypeHAlign = TypeHAlignLeft
'      vaSpread1(0).text = IIf(IsNull(RS1!rec_Indppr) Or Trim(RS1!rec_Indppr) = "", "", IIf(RS1!rec_Indppr = "1", "Real", "Propuesta"))
'
'      i = i + 1
'      RS1.MoveNext
'   Loop
'   SSTab1.TabEnabled(1) = True
'   SSTab1.TabEnabled(2) = True
'End If
'RS1.Close: Set RS1 = Nothing
'
'vaSpread1(0).Visible = True
'vaSpread1(0).SetActiveCell 1, 1
'Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Reg. Encontrados"
End Sub

Private Sub Intolerancia_Click(Index As Integer)
     
On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub List1_Click(Index As Integer)
     
On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub ParAdi1_Click(Index As Integer)

On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub ParAdi2_Click(Index As Integer)

On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub PSalsa_Click(Index As Integer)

On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub RichTextBox1_Change(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    If RichTextBox1(0).TextRTF <> metodoreceta And itexto = 0 And modo = "M" Then
       
       If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
         
         If vg_Indppr = 3 Or vg_Indppr = 1 Then
          
            Gl_Ac_Botones Me, 3, 0, modo
          
            If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
             
                Hab_Des 0
                SSTab1.TabEnabled(4) = False
          
            End If
         
         ElseIf vg_Indppr = 2 And vg_Indppr = ComboValOrig Then
          
            Gl_Ac_Botones Me, 3, 0, modo
            If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
         
         End If
         
       End If
       
    End If

Case 1
    
    If RichTextBox1(1).TextRTF <> grupovulnerable And itexto = 0 And modo = "M" Then
       
       If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
         
         If vg_Indppr = 3 Or vg_Indppr = 1 Then
            
            Me.HelpContextID = 1091000
            Gl_Ac_Botones Me, 3, 0, modo
           If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 4
           SSTab1.TabEnabled(3) = True
           SSTab1.TabEnabled(4) = False
           Me.HelpContextID = vg_OpcM
         
         ElseIf vg_Indppr = 2 And vg_Indppr = ComboValOrig Then
           
           Gl_Ac_Botones Me, 3, 0, modo
           If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 4
         
         End If
       
       End If
    
    End If

Case 2
    
    If RichTextBox1(2).TextRTF <> HipersensabilidadAlimentaria And itexto = 0 And modo = "M" Then
       
       If Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
         
         If vg_Indppr = 3 Or vg_Indppr = 1 Then
            
            Me.HelpContextID = 1092000
            Gl_Ac_Botones Me, 3, 0, modo
            If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 4
            SSTab1.TabEnabled(3) = False
            SSTab1.TabEnabled(4) = True
            Me.HelpContextID = vg_OpcM
         
         ElseIf vg_Indppr = 2 And vg_Indppr = ComboValOrig Then
           
           Gl_Ac_Botones Me, 3, 0, modo
           If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 4
         
         End If
       
       End If
    
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub RichTextBox1_Click(Index As Integer)

On Error GoTo Man_Error

'------- Bold
If RichTextBox1(Index).SelBold = False Or IsNull(RichTextBox1(Index).SelBold) Then
   
   Toolbar3.Buttons(1).Value = 1: Toolbar3.Buttons(1).Value = 0

ElseIf RichTextBox1(Index).SelBold = True Then
   
   Toolbar3.Buttons(1).Value = 0: Toolbar3.Buttons(1).Value = 1

End If

'------- Italic
If RichTextBox1(Index).SelItalic = False Or IsNull(RichTextBox1(Index).SelItalic) Then
   
   Toolbar3.Buttons(2).Value = 1: Toolbar3.Buttons(2).Value = 0

ElseIf RichTextBox1(Index).SelItalic = True Then
   
   Toolbar3.Buttons(2).Value = 0: Toolbar3.Buttons(2).Value = 1

End If

'------- Subrayado
If RichTextBox1(Index).SelUnderline = False Or IsNull(RichTextBox1(Index).SelUnderline) Then
   
   Toolbar3.Buttons(3).Value = 1: Toolbar3.Buttons(3).Value = 0

ElseIf RichTextBox1(Index).SelUnderline = True Then
   
   Toolbar3.Buttons(3).Value = 0: Toolbar3.Buttons(3).Value = 1

End If

'------- Viñetas
If RichTextBox1(Index).SelBullet = False Or IsNull(RichTextBox1(Index).SelBullet) Then
   
   Toolbar3.Buttons(10).Value = 1: Toolbar3.Buttons(10).Value = 0

ElseIf RichTextBox1(Index).SelBullet = True Then
   
   Toolbar3.Buttons(10).Value = 0: Toolbar3.Buttons(10).Value = 1

End If

If RichTextBox1(Index).SelAlignment = 0 Then
   
   Toolbar3.Buttons(5).Value = 0
   Toolbar3.Buttons(5).Value = 1
   Toolbar3.Buttons(6).Value = 1
   Toolbar3.Buttons(6).Value = 0
   Toolbar3.Buttons(7).Value = 1
   Toolbar3.Buttons(7).Value = 0

End If

If RichTextBox1(Index).SelAlignment = 2 Then
   
   Toolbar3.Buttons(5).Value = 1
   Toolbar3.Buttons(5).Value = 0
   Toolbar3.Buttons(6).Value = 0
   Toolbar3.Buttons(6).Value = 1
   Toolbar3.Buttons(7).Value = 1
   Toolbar3.Buttons(7).Value = 0

End If

If RichTextBox1(Index).SelAlignment = 1 Then
   
   Toolbar3.Buttons(5).Value = 1
   Toolbar3.Buttons(5).Value = 0
   Toolbar3.Buttons(6).Value = 1
   Toolbar3.Buttons(6).Value = 0
   Toolbar3.Buttons(7).Value = 0
   Toolbar3.Buttons(7).Value = 1

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub RichTextBox1_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

On Error GoTo Man_Error

'------- Bold
If RichTextBox1(Index).SelBold = False Or IsNull(RichTextBox1(Index).SelBold) Then
   
   Toolbar3.Buttons(1).Value = 1: Toolbar3.Buttons(1).Value = 0

ElseIf RichTextBox1(Index).SelBold = True Then
   
   Toolbar3.Buttons(1).Value = 0: Toolbar3.Buttons(1).Value = 1

End If

'------- Italic
If RichTextBox1(Index).SelItalic = False Or IsNull(RichTextBox1(Index).SelItalic) Then
   
   Toolbar3.Buttons(2).Value = 1: Toolbar3.Buttons(2).Value = 0

ElseIf RichTextBox1(Index).SelItalic = True Then
   
   Toolbar3.Buttons(2).Value = 0: Toolbar3.Buttons(2).Value = 1

End If

'------- Subrayado
If RichTextBox1(Index).SelUnderline = False Or IsNull(RichTextBox1(Index).SelUnderline) Then
   
   Toolbar3.Buttons(3).Value = 1: Toolbar3.Buttons(3).Value = 0

ElseIf RichTextBox1(Index).SelUnderline = True Then
   
   Toolbar3.Buttons(3).Value = 0: Toolbar3.Buttons(3).Value = 1

End If

'------- Viñetas
If RichTextBox1(Index).SelBullet = False Or IsNull(RichTextBox1(Index).SelBullet) Then
   
   Toolbar3.Buttons(10).Value = 1: Toolbar3.Buttons(10).Value = 0

ElseIf RichTextBox1(Index).SelBullet = True Then
   
   Toolbar3.Buttons(10).Value = 0: Toolbar3.Buttons(10).Value = 1

End If

If RichTextBox1(Index).SelAlignment = 0 Then
   
   Toolbar3.Buttons(5).Value = 0
   Toolbar3.Buttons(5).Value = 1
   Toolbar3.Buttons(6).Value = 1
   Toolbar3.Buttons(6).Value = 0
   Toolbar3.Buttons(7).Value = 1
   Toolbar3.Buttons(7).Value = 0
   
End If

If RichTextBox1(Index).SelAlignment = 2 Then

   Toolbar3.Buttons(5).Value = 1
   Toolbar3.Buttons(5).Value = 0
   Toolbar3.Buttons(6).Value = 0
   Toolbar3.Buttons(6).Value = 1
   Toolbar3.Buttons(7).Value = 1
   Toolbar3.Buttons(7).Value = 0
   
End If

If RichTextBox1(Index).SelAlignment = 1 Then

   Toolbar3.Buttons(5).Value = 1
   Toolbar3.Buttons(5).Value = 0
   Toolbar3.Buttons(6).Value = 1
   Toolbar3.Buttons(6).Value = 0
   Toolbar3.Buttons(7).Value = 0
   Toolbar3.Buttons(7).Value = 1

End If

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

vg_codigo = 0

Select Case Index

Case 0

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_tipoingredienteprincipalreceta", "sec_", "Ingrediente Principal", "TipIngPri"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    TipoIngPrincipal = vg_codigo
    fpayuda(0).text = vg_nombre

Case 1

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_tipoingredienteprincipalreceta", "sec_", "Segundo Ingrediente Principal", "TipIngPri"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    fpSIP1(0).Value = vg_codigo
    fpayuda(1).text = vg_nombre

Case 2
    
    vg_nombre = ""
    vg_codigo = 0
    vg_left = fpayuda(2).Left + 550
    B_ArbEst.MoverDatosTvwDir "a_recetacatdie", "car_", "Categoria Dietetica"
    B_ArbEst.Show 1
    Me.Refresh
    If vg_codigo = 0 Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    CodCatDie = vg_codigo
    fpayuda(2).text = vg_nombre

Case 3
    
    vg_nombre = ""
    vg_codigo = 0
    vg_left = fpayuda(3).Left + 550
    B_ArbEst.MoverDatosTvwDir "a_recetatippla", "tip_", "Tipo Plato"
    B_ArbEst.Show 1
    Me.Refresh
    If vg_codigo = 0 Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
    codtipplato = vg_codigo
    fpayuda(3).text = vg_nombre

Case 4

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_metodococcionreceta", "sec_", "Metodo Cocción", "MetCocc"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    MetodoCoccion = vg_codigo
    fpayuda(4).text = vg_nombre

Case 5

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_ingredientecrucegarniturareceta", "sec_", "Ingrediente Cruce Garnitura", "IngCruGar"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    IngCruceGarnitura = vg_codigo
    fpayuda(5).text = vg_nombre

Case 6

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_tiempohhreceta", "sec_", "Tiempo HH", "TiempoHH"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    TiempoHH = vg_codigo
    fpayuda(6).text = vg_nombre

Case 7

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_color", "sec_", "Color", "Color"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    Color = vg_codigo
    fpayuda(7).text = vg_nombre

Case 8

    vg_left = Command1.Left + 3801
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_tiempococcionreceta", "sec_", "Tiempo Cocción ", "TiempoCoccion"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
    
       Gl_Ac_Botones Me, 3, 0, modo
       If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
       
          Hab_Des 0
          
      End If
      
    End If
    
    TiempoCoccion = vg_codigo
    fpayuda(8).text = vg_nombre

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Sellos_Click(Index As Integer)
    
On Error GoTo Man_Error

Gl_Ac_Botones Me, 3, 0, modo

If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then

   Hab_Des 0
   
Else

   Exit Sub
 
End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
 
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error GoTo Man_Error

Dim Izquerda As Long

Toolbar1.Buttons(24).Enabled = False
Select Case SSTab1.Tab

Case 0

Case 1
    
    If vaSpread1(0).MaxRows > 0 And modo = "M" Then
       
       modo = "M"
       If vg_newcodrec < 1 Then
          
          Frame1(3).Enabled = False
          MoverDetalleDatos
          Frame1(3).Enabled = True
          
      End If
    
    ElseIf vaSpread1(0).MaxRows < 1 And modo = "M" Then
       
       SSTab1.Tab = 0
       Exit Sub
    
    End If

Case 2
    
    MoverDatosPropuesta
    If vaSpread1(0).MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Izquerda = 0
    CargaMetodoReceta

Case 3
    
    MoverDatosPropuesta
    If vaSpread1(0).MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Izquerda = 0
    CargaGrupoVulnerable

Case 4
    
    MoverDatosPropuesta
    If vaSpread1(0).MaxRows < 1 Then SSTab1.Tab = 0: Exit Sub
    Izquerda = 0
    CargaHipersensabilidadAlimentaria

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub LlenaCombo(Index, CodReceta As Long)

On Error GoTo Man_Error

Dim RS  As New ADODB.Recordset
Dim Sql As String

'INI ARI
'Carga la el combox
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Combo1(0).Clear
Sql = " sgpadm_Sel_UnidadRecetaActiva "
Set RS = vg_db.Execute(Sql)

'-------> Inicio LLenar grilla
Do While Not RS.EOF

   Combo1(0).AddItem RS("Descripcion") & Space(150) & RS("Codigo_unidad")
   RS.MoveNext

Loop

RS.Close
Set RS = Nothing
Combo1(0).ListIndex = -1

'If IsNull(unidadReceta) Or Trim(unidadReceta) = 0 Then
'
'   Combo1(0).ListIndex = -1
'
'Else
'
'   Combo1(0).ListIndex = fg_buscacboNuevo(Combo1, 0, 10, (unidadReceta))
'
'End If

'FIN ARI

With Combo2(2)
  
  Select Case vg_Indppr
  
  Case 2
    
    .Clear
    .AddItem "Propuesta" & Space(150) & "(2)"
  
  Case 1, 3
    
    .Clear
    .AddItem "Real" & Space(150) & "(1)"
    .AddItem "Propuesta" & Space(150) & "(2)"
  
  End Select
  
  If Index = 1 Then
    
    .Clear
    .AddItem "Real" & Space(150) & "(1)"
    .AddItem "Propuesta" & Space(150) & "(2)"
  
  End If

End With

'INI ARI
List1(0).Clear
'Dim Sql As String

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaEstacionalidad " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
Dim contador As Long
contador = 0
         
Do While Not RS.EOF
         
   List1(0).AddItem RS("Glosa") & Space(150) & RS("IdEstacionalidad")
       
   If RS("selected") = 1 Then List1(0).Selected(contador) = True
          
   RS.MoveNext
   contador = contador + 1
    
Loop
    
RS.Close
Set RS = Nothing

'Carga la LisBOX
List1(1).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_OfertasReceta " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   List1(1).AddItem RS("Descripcion") & Space(150) & RS("codigo_oferta")
   If RS("selected") = 1 Then List1(1).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing

'Ini : Carga Tipo Negocio
TipoNegocio(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaTipoNegocio " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   TipoNegocio(0).AddItem RS("NombreTipoNegocio") & Space(150) & RS("IdTipoNegocio")
   If RS("selected") = 1 Then TipoNegocio(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Tipo Negocio

'Ini : Carga Org. Compras(Zona)
Zona(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaZona " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   Zona(0).AddItem RS("NombreOrgCompra") & Space(150) & RS("IdOrgCompra")
   If RS("selected") = 1 Then Zona(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Org. Compras(Zona)

'Ini : Carga Intolerancia
Intolerancia(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaIntolerancia " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   Intolerancia(0).AddItem RS("NombreIntolerancia") & Space(150) & RS("IdIntolerancia")
   If RS("selected") = 1 Then Intolerancia(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Intolerancia

'Ini : Carga Alergeno
Alergeno(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaAlergeno " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   Alergeno(0).AddItem RS("NombreAlergeno") & Space(150) & RS("IdAlergeno")
   If RS("selected") = 1 Then Alergeno(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Alergeno

'Ini : Carga Estilo Alimentación
EstAli(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaEstiloAlimentacion " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   EstAli(0).AddItem RS("NombreEstiloAlimentacion") & Space(150) & RS("IdEstiloAlimentacion")
   If RS("selected") = 1 Then EstAli(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Estilo Alimentación

'Ini : Carga Parametro Adicional N°1
ParAdi1(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaParametroadicional1 " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   ParAdi1(0).AddItem RS("NombreParametroAdicional1") & Space(150) & RS("IdParametroadicional1")
   If RS("selected") = 1 Then ParAdi1(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Parametro Adicional N°1

'Ini : Carga Parametro Adicional N°2
ParAdi2(0).Clear

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Sql = " sgpadm_Sel_RecetaParametroadicional2 " & CodReceta & ""
Set RS = vg_db.Execute(Sql)
contador = 0
         
Do While Not RS.EOF
         
   ParAdi2(0).AddItem RS("NombreParametroAdicional2") & Space(150) & RS("IdParametroadicional2")
   If RS("selected") = 1 Then ParAdi2(0).Selected(contador) = True
          
    RS.MoveNext
    contador = contador + 1
    
Loop
RS.Close
Set RS = Nothing
'Fin : Carga Parametro Adicional N°2

'-------> Ini : Costo Receta
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_CostoRecetaActivo")
If RS.RecordCount > CantRecordCosto Then
   
   CantRecordCosto = RS.RecordCount
   Costo(0).Clear

   Do While Not RS.EOF
   
      Costo(0).AddItem RS("NombreCosto") & Space(150) & RS("IdCosto")
      RS.MoveNext

   Loop

End If

RS.Close
Set RS = Nothing
Costo(0).ListIndex = -1
'-------> Fin : Cargar Costo Receta

'-------> Ini : Categoria Compleja Receta
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_CategorizacionComplejaRecetaActivo")
If RS.RecordCount > CantRecordCatCompleja Then
   
   CantRecordCatCompleja = RS.RecordCount

   CatCompleja(0).Clear
 
   Do While Not RS.EOF
   
      CatCompleja(0).AddItem RS("NombreCategorizacionCompleja") & Space(150) & RS("IdCategorizacionCompleja")
      RS.MoveNext

   Loop

End If

RS.Close
Set RS = Nothing
CatCompleja(0).ListIndex = -1
'-------> Fin : Categoria Compleja Receta

'-------> Ini : Efecto Meteorizante Receta
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EfectoMeteorizanteRecetaActivo")
If RS.RecordCount > CantRecordEfectoMeteorizante Then
   
   CantRecordEfectoMeteorizante = RS.RecordCount

   EfectoMeteorizante(0).Clear
   
   Do While Not RS.EOF
   
      EfectoMeteorizante(0).AddItem RS("NombreEfectoMeteorizante") & Space(150) & RS("IdEfectoMeteorizante")
      RS.MoveNext

   Loop
   
End If

RS.Close
Set RS = Nothing
EfectoMeteorizante(0).ListIndex = -1
'-------> Fin : Efecto Meteorizante Receta

'-------> Ini : Sellos Receta
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_SellosRecetaActivo")
If RS.RecordCount > CantRecordSellos Then
   
   CantRecordSellos = RS.RecordCount

   Sellos(0).Clear
   Do While Not RS.EOF
   
      Sellos(0).AddItem RS("NombreSellos") & Space(150) & RS("IdSellos")
      RS.MoveNext

   Loop
   
End If
'-------> Fin : Sellos Receta

RS.Close
Set RS = Nothing
Sellos(0).ListIndex = -1
'-------> Fin : Sellos Receta

'-------> Ini : Etiquetado Sello Receta
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EtiquetadoSelloRecetaActivo")
If RS.RecordCount > CantRecordEtiquetadoSello Then
   
   CantRecordEtiquetadoSello = RS.RecordCount

   EtiquetadoSello(0).Clear
   Do While Not RS.EOF
   
      EtiquetadoSello(0).AddItem RS("NombreEtiquetadoSello") & Space(150) & RS("IdEtiquetadoSello")
      RS.MoveNext

   Loop
   
End If

RS.Close
Set RS = Nothing
EtiquetadoSello(0).ListIndex = -1
'-------> Fin : Etiquetado Sello Receta

'-------> Ini : Equipamiento Cocción
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_EquipamientoCoccionActivoReceta")
If RS.RecordCount > CantRecordECoccion Then
   
   CantRecordECoccion = RS.RecordCount

   ECoccion(0).Clear
   Do While Not RS.EOF
   
      ECoccion(0).AddItem RS("Nombre") & Space(150) & RS("IdEquipamientoCoccion")
      RS.MoveNext

   Loop
   
End If

RS.Close
Set RS = Nothing
ECoccion(0).ListIndex = -1
'-------> Fin : Equipamiento Cocción

'-------> Ini : Parametro Salsa
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_ParametroSalsaActivoReceta")
If RS.RecordCount > CantRecordPSalsa Then
   
   CantRecordPSalsa = RS.RecordCount

   PSalsa(0).Clear
   Do While Not RS.EOF
   
      PSalsa(0).AddItem RS("Nombre") & Space(150) & RS("IdParametroSalsa")
      RS.MoveNext

   Loop
   
End If

RS.Close
Set RS = Nothing
PSalsa(0).ListIndex = -1
'-------> Fin : Parametro Salsa

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    
End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 1 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""
   TextDet2(10).text = ""
   TextDet2(11).text = ""

ElseIf Index = 2 Then
   
   TextDet2(1).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""
   TextDet2(10).text = ""
   TextDet2(11).text = ""

ElseIf Index = 3 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""
   TextDet2(10).text = ""
   TextDet2(11).text = ""

ElseIf Index = 4 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""
   TextDet2(10).text = ""
   TextDet2(11).text = ""

ElseIf Index = 5 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(8).text = ""
   TextDet2(10).text = ""
   TextDet2(11).text = ""


ElseIf Index = 8 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(10).text = ""
   TextDet2(11).text = ""

ElseIf Index = 10 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""
   TextDet2(11).text = ""

ElseIf Index = 11 Then
   
   TextDet2(1).text = ""
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""
   TextDet2(10).text = ""

End If

For i = 1 To vaSpread1(0).MaxRows
           
    vaSpread1(0).Row = i
    vaSpread1(0).Col = 12
    vaSpread1(0).text = 0
    
Next

Select Case Index

Case 1, 2, 3, 4, 5, 8, 10, 11
    
    vaSpread1(0).Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1(0).MaxRows
           
           vaSpread1(0).Row = i
           vaSpread1(0).Col = Index
           indactivo = UCase(Trim(vaSpread1(0).Value)) Like IIf(Index = 1, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1(0).Col = Index
           
           If indactivo = -1 And Trim(vaSpread1(0).text) <> "" Then
              
              vaSpread1(0).Col = 12
              
              If Val(vaSpread1(0).Value) <> 1 Then
                              
                 vaSpread1(0).Col = 1
              
                 If vaSpread1(0).RowHidden = True Then
                 
                    vaSpread1(0).RowHidden = False
                    vaSpread1(0).Col = 12
                    vaSpread1(0).text = 1
                 
                 Else
                 
                    vaSpread1(0).Col = 12
                    vaSpread1(0).text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1(0).Col = 12
              EstBuq = vaSpread1(0).Value
              vaSpread1(0).Col = 2
              
              If vaSpread1(0).RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1(0).RowHidden = True
                 
                 vaSpread1(0).Col = 12
                 vaSpread1(0).text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1(0).SetActiveCell Index + 1, 1
        vaSpread1(0).ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1(0).ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1(0).SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1(0).SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1(0).Sort -1, -1, vaSpread1(0).maxcols, vaSpread1(0).MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1(0).MaxRows
           
           vaSpread1(0).Row = i
           If vaSpread1(0).RowHidden = True Then vaSpread1(0).RowHidden = False
           
           vaSpread1(0).Col = 12
           vaSpread1(0).text = 0
       
       Next
       
       vaSpread1(0).SetActiveCell Index, vaSpread1(0).SearchCol(Index, 0, vaSpread1(0).MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1(0).SetActiveCell Index, 1
    
    End If
    
    vaSpread1(0).Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TipoNegocio_Click(Index As Integer)
     
On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim indice      As Long
Dim tiprec      As String
Dim CodRec      As String
Dim Esteli      As Boolean
Dim Cod1        As Integer
Dim Cod2        As Integer
Dim ExistTabla  As Boolean
Dim TippRec     As String
Dim Sql         As String
Dim RS          As New ADODB.Recordset

Select Case Button.Index

Case 1 '-------> Agregar
    
    modo = "A"
    itexto = 1
    LimpiarVariable
    Hab_Des 0: SSTab1.Tab = 1
    SSTab1.TabEnabled(1) = True
    SSTab1.TabEnabled(2) = False
    SSTab1.TabEnabled(3) = False
    LlenaCombo 0, 0
    Gl_Ac_Botones Me, 3, 0, modo: ConfiControlesReceta 1, True
    itexto = 0

Case 3 '-------> Modificar
    
    modo = "M"
    If vaSpread1(0).MaxRows < 1 Then
    
       Exit Sub
       
    End If
    
    Hab_Des 0
    Combo2(2).ListIndex = fg_buscacbo(Combo2, 2, 1, (ComboValOrig))
    Combo2(2).Enabled = True
    Gl_Ac_Botones Me, 3, 0, modo
    
    If SSTab1.Tab = 1 Or SSTab1.Tab = 0 Then
       
       SSTab1.TabEnabled(2) = False
       SSTab1.Tab = 1
    
    ElseIf SSTab1.Tab = 2 Then
       
       SSTab1.TabEnabled(1) = False
       CargaMetodoReceta
    
    End If

Case 5 '-------> Borrar
    
    If vaSpread1(0).MaxRows < 1 Then
    
       Exit Sub
    
    End If
    
    Esteli = False
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    
    'Borrar receta sgp
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_Del_Receta_V02 " & codigo & "")
    If Not RS.EOF Then
       
       If RS(0) > 0 Then
          
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          RS.Close
          Set RS = Nothing
          Exit Sub
       
       End If
    
    End If
    RS.Close
    Set RS = Nothing
    
    'Borrar tabla gramaje
    If Esteli = False Then
       
       Esteli = True
       MsgBox "Atención, ingrediente sera borrado de la tabla gramaje, ya que fue eliminado en receta", vbCritical + vbOKOnly, MsgTitulo
    
    End If
    
    vg_db.Execute "sgpadm_Del_TablaGramajeReceta " & codigo & ""
    'Borrar receta tecfood
    
    CodRec = ""
    vaSpread1(0).Row = vaSpread1(0).ActiveRow
    vaSpread1(0).DeleteRows vaSpread1(0).Row, 1
    vaSpread1(0).MaxRows = vaSpread1(0).MaxRows - 1
    vaSpread1(0).Row = vaSpread1(0).MaxRows
    Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Registros"
    
    If vaSpread1(0).MaxRows < 1 Then
       
       Label1(1).Visible = False
       fpTnombre.Visible = False
       Label1(10).Visible = False
       SSTab1.TabEnabled(1) = False
       SSTab1.TabEnabled(2) = False
       SSTab1.Tab = 0
       Gl_Ac_Botones Me, 3, 2, modo
    
    Else
       
       SSTab1.TabEnabled(1) = True
       SSTab1.Tab = 0
       fpTnombre.SetFocus
    
    End If

Case 7 '-------> Actualizar Lista
    
    TextDet2(1).text = ""
    TextDet2(2).text = ""
    TextDet2(3).text = ""
    TextDet2(4).text = ""
    TextDet2(5).text = ""
    TextDet2(8).text = ""
    TextDet2(10).text = ""
    TextDet2(11).text = ""
    Mover_ListaReceta
    SSTab1.Tab = 0

Case 10 '-------> Cancelar
    
    If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
'    LlenaCombo 1, 0
    Combo2(2).ListIndex = fg_buscacbo(Combo2, 2, 1, (ComboValOrig))
    Combo2(2).Enabled = True
    '---->
    If SSTab1.Tab = 1 Then
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("sgpadm_s_receta_V07 2, 0, '%" & LimpiaDato(UCase(fpTnombre.text)) & "%', " & vg_filcatdie & "," & vg_filtippla & ", 0, '" & vg_NUsr & "'")
       If RS.EOF Or RS!nReg = 0 Then
       
          RS.Close
          Set RS = Nothing
          Hab_Des 2
          Gl_Ac_Botones Me, 3, 2, modo
          SSTab1.Tab = 0
          Exit Sub
          
       End If
       
       modo = "M"
       RS.Close
       Set RS = Nothing
       
       If modo = "A" Then
          
          SSTab1.Tab = 0
       
       ElseIf modo = "M" Then
          
          LlenaCombo 1, vg_newcodrec
          MoverDetalleDatos
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
          SSTab1.TabEnabled(4) = True
       
       End If
       
       If vg_newcodrec = 0 Then
          
          Gl_Ac_Botones Me, 3, 1, modo: Hab_Des 1
       
       ElseIf vg_newcodrec > 0 Then
          
          Gl_Ac_Botones Me, 3, 3, modo: Hab_Des 0
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
          SSTab1.TabEnabled(4) = True
       
       End If
    
    ElseIf SSTab1.Tab = 2 Then
       
       vaSpread1(0).Row = vaSpread1(0).ActiveRow
       vaSpread1(0).Col = 1
       codigo = Val(vaSpread1(0).text)
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("sgpadm_s_receta_V07 4, " & codigo & ", '', 0, 0, 0, '" & vg_NUsr & "'")
       If RS.EOF Then
       
          RS.Close
          Set RS = Nothing
          Hab_Des 2
          Gl_Ac_Botones Me, 3, 2, modo
          SSTab1.Tab = 0
          Exit Sub
          
       End If
       
       modo = "M"
       RS.Close
       Set RS = Nothing
       
       If modo = "A" Then
          
          SSTab1.Tab = 0
       
       ElseIf modo = "M" Then
          
          CargaMetodoReceta
          SSTab1.TabEnabled(1) = True
       
       End If
       
       If vg_newcodrec = 0 Then
          
          Gl_Ac_Botones Me, 3, 1, modo
          Hab_Des 1
          SSTab1.TabEnabled(4) = True
       
       ElseIf vg_newcodrec > 0 Then
          
          Gl_Ac_Botones Me, 3, 3, modo
          Hab_Des 0
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
       
       End If
    
    ElseIf SSTab1.Tab = 3 Then
       
       vaSpread1(0).Row = vaSpread1(0).ActiveRow
       vaSpread1(0).Col = 1
       codigo = Val(vaSpread1(0).text)
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("sgpadm_s_receta_V07 10, " & codigo & ", '', 0, 0, 0, '" & vg_NUsr & "'")
       If RS.EOF Then
       
          RS.Close
          Set RS = Nothing
          Hab_Des 2
          Gl_Ac_Botones Me, 3, 2, modo
          SSTab1.Tab = 0
          Exit Sub
          
       End If
       
       modo = "M"
       RS.Close
       Set RS = Nothing
       
       If modo = "A" Then
          
          SSTab1.Tab = 0
       
       ElseIf modo = "M" Then
          
          CargaGrupoVulnerable
          SSTab1.TabEnabled(1) = True
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
          SSTab1.TabEnabled(4) = True
       
       End If
       
       If vg_newcodrec = 0 Then
          
          Gl_Ac_Botones Me, 3, 1, modo
          Hab_Des 1
       
       ElseIf vg_newcodrec > 0 Then
          
          Gl_Ac_Botones Me, 3, 3, modo
          Hab_Des 0
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
       
       End If
    
    ElseIf SSTab1.Tab = 4 Then
       
       vaSpread1(0).Row = vaSpread1(0).ActiveRow
       vaSpread1(0).Col = 1
       codigo = Val(vaSpread1(0).text)
       
       If RS.State = 1 Then RS.Close
       RS.CursorLocation = adUseClient
       vg_db.CursorLocation = adUseClient
       
       Set RS = vg_db.Execute("sgpadm_s_receta_V07 10, " & codigo & ", '', 0, 0, 0, '" & vg_NUsr & "'")
       If RS.EOF Then
       
          RS.Close
          Set RS = Nothing
          Hab_Des 2
          Gl_Ac_Botones Me, 3, 2, modo
          SSTab1.Tab = 0
          Exit Sub
          
       End If
       
       modo = "M"
       RS.Close
       Set RS = Nothing
       
       If modo = "A" Then
          
          SSTab1.Tab = 0
       
       ElseIf modo = "M" Then
          
          CargaHipersensabilidadAlimentaria
          SSTab1.TabEnabled(1) = True
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
       
       End If
       
       If vg_newcodrec = 0 Then
          
          Gl_Ac_Botones Me, 3, 1, modo
          Hab_Des 1
       
       ElseIf vg_newcodrec > 0 Then
          
          Gl_Ac_Botones Me, 3, 3, modo
          Hab_Des 0
          SSTab1.TabEnabled(2) = True
          SSTab1.TabEnabled(3) = True
       
       End If
    
    End If
    Me.HelpContextID = 1093000
'    Toolbar1.Buttons(19).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Toolbar1.Buttons(19).ButtonMenus(1).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = vg_OpcM
    Me.HelpContextID = 1094000
    Toolbar1.Buttons(19).ButtonMenus(2).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = vg_OpcM

Case 12 '------> Confirmar
    
    Dim coddi1                 As Long
    Dim coddi2                 As Long
    Dim codti1                 As Long
    Dim codti2                 As Long
    Dim codti3                 As Long
    Dim fecvig                 As Long
    Dim IndIngSumaTablaGramaje As String
    Dim StrFamb                As String
    Dim StrFam                 As String
    Dim Var_Sellos             As Long
    Dim Var_EtiquetadoSello    As Long
    Dim Var_Costo              As Long
    Dim Var_TiPrincipal        As Long
    Dim Var_MetodoCoccion      As Long
    Dim Var_CatCompleja        As Long
    Dim Var_IngCruceGar        As Long
    Dim Var_EfectoMeteorizante As Long
    Dim Var_TCoccion           As Long
    Dim Var_THH                As Long
    Dim Var_Color              As Long
    Dim Var_ECoccion           As Long
    Dim Var_SIPrincipal        As Long
    Dim Var_PSalsa             As Long
    Dim Var_IntegraAMD         As String
    
    If SSTab1.Tab = 1 Then
       
       If (LimpiaDato(Trim(fpText1(0).text)) = "" Or LimpiaDato(Trim(fpText1(1).text)) = "" Or fpayuda(2).text = "" Or fpayuda(3).text = "" Or Val(fpDouble1(0).Value) = 0) Then
    
          MsgBox "Falta información...", vbExclamation + vbOKOnly, MsgTitulo
          Exit Sub
    
       End If
       
       If Costo(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Costo Receta", 16
          Costo(0).SetFocus
          Exit Sub
      
       End If
       
       If Trim(fpayuda(0).text) = "" Then
         
          MsgBox "Debe Seleccionar ítem Tipo Ingrediente Principal", 16
          Exit Sub
      
       End If
       
       If Trim(fpayuda(4).text) = "" Then
         
          MsgBox "Debe Seleccionar ítem Método Cocción", 16
          Exit Sub
      
       End If
       
       If CatCompleja(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Categorización Compleja", 16
          CatCompleja(0).SetFocus
          Exit Sub
      
       End If
       
       If Trim(fpayuda(5).text) = "" Then
         
          MsgBox "Debe Seleccionar ítem Ingrediente Cruce Garnitura", 16
          Exit Sub
      
       End If
       
       If EfectoMeteorizante(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Efecto Meteorizante", 16
          EfectoMeteorizante(0).SetFocus
          Exit Sub
      
       End If
       
       If Sellos(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Sellos", 16
          Sellos(0).SetFocus
          Exit Sub
      
       End If
       
       If EtiquetadoSello(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Eiquetado Sello", 16
          EtiquetadoSello(0).SetFocus
          Exit Sub
      
       End If
       
       If Trim(fpayuda(8).text) = "" Then
         
          MsgBox "Debe Seleccionar ítem Tiempo Cocción", 16
          Exit Sub
      
       End If
        
       If Trim(fpayuda(6).text) = "" Then
         
          MsgBox "Debe Seleccionar ítem Tiempo HH", 16
          Exit Sub
      
       End If
        
       If Trim(fpayuda(7).text) = "" Then
         
          MsgBox "Debe Seleccionar ítem Color", 16
          Exit Sub
      
       End If
                           
       If ECoccion(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Equipamiento Cocción", 16
          ECoccion(0).SetFocus
          Exit Sub
      
       End If
              
       
       If PSalsa(0) = "" Then
         
          MsgBox "Debe Seleccionar ítem Parametro Salsa", 16
          PSalsa(0).SetFocus
          Exit Sub
      
       End If
       
       Var_Sellos = IIf(Sellos(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(Sellos(0), 0, 10, "")))
       Var_EtiquetadoSello = IIf(EtiquetadoSello(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(EtiquetadoSello(0), 0, 10, "")))
       Var_Costo = IIf(Costo(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(Costo(0), 0, 10, "")))
       Var_TiPrincipal = TipoIngPrincipal 'IIf(TiPrincipal(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(TiPrincipal(0), 0, 10, "")))
       Var_MetodoCoccion = MetodoCoccion 'IIf(MetodoCoccion(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(MetodoCoccion(0), 0, 10, "")))
       Var_CatCompleja = IIf(CatCompleja(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(CatCompleja(0), 0, 10, "")))
       Var_IngCruceGar = IngCruceGarnitura 'IIf(IngCruceGar(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(IngCruceGar(0), 0, 10, "")))
       Var_EfectoMeteorizante = IIf(EfectoMeteorizante(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(EfectoMeteorizante(0), 0, 10, "")))
       Var_TCoccion = TiempoCoccion 'IIf(TCoccion(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(TCoccion(0), 0, 10, "")))
       Var_THH = TiempoHH 'IIf(THH(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(THH(0), 0, 10, "")))
       Var_Color = Color 'IIf(Color(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(Color(0), 0, 10, "")))
       Var_SIPrincipal = fpSIP1(0).Value 'IIf(SIPrincipal(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(SIPrincipal(0), 0, 10, "")))
       Var_ECoccion = IIf(ECoccion(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(ECoccion(0), 0, 10, "")))
       Var_PSalsa = IIf(PSalsa(0).ListIndex = -1, -1, Val(fg_codigolistaNuevo(PSalsa(0), 0, 10, "")))
       Var_IntegraAMD = IIf(ChAMD.Value = 1, "1", "2")
    
       'Validar que no se repita el primer ingrediente principal con Segundo ingrediente
       If Var_TiPrincipal = Var_SIPrincipal Then
       
          MsgBox "Código ing. principal, debe ser distinto segundo código ingrediente principal", 16
          Exit Sub
       
       
       End If
       
       If Var_SIPrincipal > 0 And Trim(fpayuda(1).text) = "" Then
       
          MsgBox "Código segundo ing. principal no existe maestra ing. principal o bien esta desactivado, si no lo va ocupar debe dejar con valor cero o bien cambielo por otro codigo vigente", 16
          Exit Sub
       
       
       End If
       
    End If
    
    If vaSpread1(1).MaxRows > 0 Then
       
       For i = 1 To vaSpread1(1).MaxRows
           
           vaSpread1(1).Row = i
           vaSpread1(1).Col = 3
           If Val(vaSpread1(1).Value) <= 0 And vaSpread1(1).CellType = CellTypeCurrency Then
           
              MsgBox "Columna cantidad bruta, existe valor cero", vbExclamation + vbOKOnly, MsgTitulo
              Exit Sub
              
           End If
           
           vaSpread1(1).Col = 5
           If Val(vaSpread1(1).Value) <= 0 And vaSpread1(1).CellType = CellTypeCurrency Then
           
              MsgBox "Columna %aprovechamiento con valor cero", vbExclamation + vbOKOnly, MsgTitulo
              Exit Sub
           
           End If
           
           vaSpread1(1).Col = 7 'actualizado col
           If Val(vaSpread1(1).Value) <= 0 And vaSpread1(1).CellType = CellTypeCurrency Then
           
              MsgBox "Columna %a.cocción con valor cero", vbExclamation + vbOKOnly, MsgTitulo
              Exit Sub
              
           End If
           
           vaSpread1(1).Col = 9 'actualizado col
           If Val(vaSpread1(1).Value) <= 0 And vaSpread1(1).CellType = CellTypeCurrency Then
           
              MsgBox "Columna %p.nutricional con valor cero", vbExclamation + vbOKOnly, MsgTitulo
              Exit Sub
              
           End If
       
       Next i
    
    End If
    
    If IsDate(fpDateTime1(0).text) = True Then
       
       fecvig = IIf(Check1(0).Value = 0, 0, Mid(fpDateTime1(0).text, 7, 4) & Mid(fpDateTime1(0).text, 4, 2) & Mid(fpDateTime1(0).text, 1, 2))
    
    End If
    
    'Verifica si existen ingredientes con tabla de gramaje
      ExistTabla = False
      For i = 1 To vaSpread1(1).MaxRows
                    
          vaSpread1(1).Row = i
          vaSpread1(1).Col = 1
          Cod1 = Val(vaSpread1(1).Value)
          
          vaSpread1(1).Col = 12 'actualizado col
          Cod2 = Val(vaSpread1(1).Value)
                    
          If Cod1 <> Cod2 Then
             
             If RS.State = 1 Then RS.Close
             RS.CursorLocation = adUseClient
             vg_db.CursorLocation = adUseClient
             
             Set RS = vg_db.Execute("sgpadm_Sel_TablaGramajeCecoReceta " & codigo & "," & Cod2 & "")
             If Not (RS.EOF And RS.BOF) Then
                
                ExistTabla = True
                Exit For
             
             End If
          
          End If
      
      Next i
      
      If ExistTabla = True Then
       
         If MsgBox("Existen ingredientes asociados a una tabla de gramaje que serán modificados" & VgLinea & "               Cancela proceso...", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
          
            Exit Sub
             
         End If
          
      End If
     '******************************************************
     
    '------- Fin validar tipo de plato
    If vg_newcodrec > 0 Then

       TippRec = "0"
    
    Else
       
       TippRec = "0"
    
    End If
    
    If modo = "A" Or modo = "M" Then
       
       indice = 0
       
       If modo = "A" Then
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_iu_receta_V05 'A', 0, " & CodCatDie & ", " & codtipplato & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & LimpiaDato(Trim(fpText1(1).text)) & "', '', '', " & _
                                   "' ', 1, '" & TippRec & "', " & fecvig & " , '', " & Val(fg_codigocbo(Combo2, 2, 1, "")) & ", " & Val(fpDouble1(4).Value) & ", '', " & Var_Sellos & ", " & Var_Costo & ", " & Var_TiPrincipal & ", " & Var_MetodoCoccion & ", " & Var_CatCompleja & ", " & Var_IngCruceGar & ", " & Var_EfectoMeteorizante & ", " & Var_TCoccion & ", " & Var_THH & ", " & Var_Color & ", " & Var_EtiquetadoSello & ", " & Var_SIPrincipal & ", " & Var_ECoccion & ", " & Var_PSalsa & ", '" & Var_IntegraAMD & "'")
            
            If Not RS.EOF Then
               
               indice = RS!indice
            
            End If
            RS.Close
            Set RS = Nothing
            '-------> Grabar detalle recetas
            '-------> Actualiza grilla Receta
            
            With vaSpread1(0)
                
                .MaxRows = .MaxRows + 1: .Row = .MaxRows
                .Col = 1
                .Lock = True
                .text = Trim(indice)
                
                .Col = 2
                .Lock = True
                .text = LimpiaDato(Trim(fpText1(0).text))
                
                .Col = 3
                .Lock = True
                .text = LimpiaDato(Trim(fpayuda(2).text))
                
                .Col = 4
                .Lock = True
                .text = LimpiaDato(Trim(fpayuda(3).text))
                
                .Col = 5
                .Lock = True
                .TypeHAlign = TypeHAlignRight
                .text = Label2(3).Caption
                
                .Col = 6
                .Lock = True
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .text = "0"
                
                .Col = 7
                .Lock = True
                .CellType = CellTypeCheckBox
                .TypeHAlign = TypeHAlignCenter
                .text = "0"
                
                .Col = 8
                .Lock = True
                .text = IIf(Val(fg_codigocbo(Combo2, 2, 1, "")) = "1", "Real", "Propuesta")
                
                .Col = 11
                .text = ""
                
                .Col = 12
                .text = 0
                
                For i = 0 To List1(0).ListCount - 1
   
                    If List1(0).Selected(i) = True Then
                            
                       If Trim(.text) = "" Then
                                  
                          .text = .text & UCase(Mid(List1(0).List(i), 1, 3))
                                
                       Else
                                   
                          .text = .text & "-" & UCase(Mid(List1(0).List(i), 1, 3))
                                
                       End If
                            
                    End If
       
                Next i
                    
                .Col = 10
                .text = ""
                For i = 0 To List1(1).ListCount - 1
   
                    If List1(1).Selected(i) = True Then
                            
                       If Trim(.text) = "" Then
                                   
                          .text = .text & UCase(Mid(List1(1).List(i), 1, 3))
                                
                       Else
                                   
                          .text = .text & "-" & UCase(Mid(List1(1).List(i), 1, 3))
                                
                       End If
                            
                    End If
       
                Next i
                
                .SetActiveCell .ActiveCol, .ActiveRow
            
            End With
            
            With vaSpread1(1)
                
                For i = 1 To .MaxRows
                    
                    .Row = i
                    .Col = 1
                    
                    If .text <> "" Then
                       
                       codpro1 = 0
                       canpro1 = 0
                       pctapr1 = 0
                       pctcoc1 = 0
                       pctnut1 = 0
                       
                       .Col = 1
                       codpro1 = .text
                       
                       .Col = 3
                       canpro1 = .text
                       
                       .Col = 5
                       pctapr1 = .text
                       
                       .Col = 7 'actualizado col
                       pctcoc1 = .text
                       
                       .Col = 9 'actualizado col
                       pctnut1 = .text
                       
                       .Col = 11 'actualizado col
                       cospro = .text
                       
                       .Col = 13 'actualizado col
                       IndIngSumaTablaGramaje = .text
                       
                       vg_db.Execute "sgpadm_iu_recetadet 'A', " & indice & ", " & i & ", '" & codpro1 & "', " & canpro1 & ", " & cospro & ", " & pctapr1 & ", " & pctcoc1 & ", " & pctnut1 & ", '', '" & IndIngSumaTablaGramaje & "'"
                    
                    End If
                
                Next i
            
            End With
          
          codigo = indice
          If vg_newcodrec > 0 Then vg_newcodrec = indice: vg_newnomrec = LimpiaDato(Trim(fpText1(0).text)): Exit Sub
             
          If fecvig <= Val(Format(Date, "yyyymmdd")) And fecvig > 0 Then vaSpread1(0).BackColor = Shape1(1).FillColor Else vaSpread1(0).BackColor = Shape1(0).FillColor
       
          Else
            
            If SSTab1.Tab = 1 Then
               
               indlec = 0
               Esteli = False
               If fpDouble1(1).Value = "" Then fpDouble1(1).Value = 0
               If fpDouble1(2).Value = "" Then fpDouble1(2).Value = 0
               If fpDouble1(3).Value = "" Then fpDouble1(3).Value = 0
               If fpDouble1(4).Value = "" Then fpDouble1(4).Value = 0
               If fpDouble1(5).Value = "" Then fpDouble1(5).Value = 0
               If fpDouble1(6).Value = "" Then fpDouble1(6).Value = 0
               
               indlec = 0
               If VarSitioRemoto = False Then
                   
                   '-------> Actualizar encabezado recetas
                   vg_db.Execute "sgpadm_iu_receta_V05 'M1', " & codigo & ", " & CodCatDie & ", " & codtipplato & ", '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & LimpiaDato(Trim(fpText1(1).text)) & "', '', '', '', " & _
                                 "1, '', " & fecvig & ", '', " & Val(fg_codigocbo(Combo2, 2, 1, "")) & ", " & Val(fpDouble1(4).Value) & ", '', " & Var_Sellos & ", " & Var_Costo & ", " & Var_TiPrincipal & ", " & Var_MetodoCoccion & ", " & Var_CatCompleja & ", " & Var_IngCruceGar & ", " & Var_EfectoMeteorizante & ", " & Var_TCoccion & ", " & Var_THH & ", " & Var_Color & ", " & Var_EtiquetadoSello & ", " & Var_SIPrincipal & ", " & Var_ECoccion & ", " & Var_PSalsa & ", '" & Var_IntegraAMD & "'"
                   
                   '-------> Grabar detalle recetas
                   With vaSpread1(0)
                        
                        .Row = vaSpread1(0).ActiveRow
                        .Col = 1
                        .text = Trim(codigo)
                        
                        .Col = 2
                        .text = fpText1(0).text
                        
                        .Col = 3
                        .text = fpayuda(2).text
                        
                        .Col = 4
                        .text = fpayuda(3).text
                        
                        .Col = 5
                        .TypeHAlign = TypeHAlignRight
                        .text = Label2(3).Caption
                        
                        .Col = 8
                        .text = IIf(Val(fg_codigocbo(Combo2, 2, 1, "")) = "1", "Real", "Propuesta")
                    
                        .Col = 11
                        .text = ""
                        
                        For i = 0 To List1(0).ListCount - 1
   
                            If List1(0).Selected(i) = True Then
                            
                                If Trim(.text) = "" Then
                                   
                                   .text = .text & UCase(Mid(List1(0).List(i), 1, 3))
                                
                                Else
                                   
                                   .text = .text & "-" & UCase(Mid(List1(0).List(i), 1, 3))
                                
                                End If
                            
                            End If
       
                        Next i
                    
                        .Col = 10
                        .text = ""
                        
                        For i = 0 To List1(1).ListCount - 1
   
                            If List1(1).Selected(i) = True Then
                            
                                If Trim(.text) = "" Then
                                   
                                   .text = .text & UCase(Mid(List1(1).List(i), 1, 3))
                                
                                Else
                                   
                                   .text = .text & "-" & UCase(Mid(List1(1).List(i), 1, 3))
                                
                                End If
                            
                            End If
       
                        Next i
                    
                    End With
                   
                    '-------> Actualiza tabla de gramaje ceco
                    With vaSpread1(1)
                        
                        For i = 1 To .MaxRows
                            
                            .Row = i
                            
                            .Col = 1
                            Cod1 = Val(.Value)
                            
                            .Col = 12 'actualizado col
                            Cod2 = Val(.Value)
                            
                            If Cod1 <> Cod2 Then
                               
                               vg_db.Execute "sgpadm_Upd_TablaGramajeReceta " & codigo & ", '" & Cod2 & "', '" & Cod1 & "'"
                            
                            End If
                            
                        Next i
                        
                       indlec = 0
                       vg_db.Execute "DELETE b_recetadet FROM b_recetadet WHERE red_codigo = " & codigo & ""
                       
                       For i = 1 To .MaxRows
                           
                           .Row = i
                           .Col = 1
                           
                           If .text <> "" Then
                              
                              codpro1 = 0
                              canpro1 = 0
                              pctapr1 = 0
                              pctcoc1 = 0
                              pctnut1 = 0
                              cospro = 0
                              
                              .Col = 1
                              codpro1 = .text
                              
                              .Col = 3
                              canpro1 = .text
                              
                              .Col = 5
                              pctapr1 = .text
                              
                              .Col = 7 'actualizado col
                              pctcoc1 = .text
                              
                              .Col = 9 'actualizado col
                              pctnut1 = .text
                              
                              .Col = 11 'actualizado col
                              cospro = .text
                              
                              .Col = 13 'actualizado col
                              IndIngSumaTablaGramaje = .text
                              
                              vg_db.Execute "sgpadm_iu_recetadet 'A', " & codigo & ", " & i & ", '" & codpro1 & "', " & canpro1 & ", " & cospro & ", " & pctapr1 & ", " & pctcoc1 & ", " & pctnut1 & ", '', '" & IndIngSumaTablaGramaje & "'"
                           
                           End If
                       
                       Next i
                    
                    End With
                   
                   ' ACTUALIZA GRILLA
                   With vaSpread1(0)
                        
                        .Row = .ActiveRow
                        .Col = -1
                        If fecvig <= Val(Format(Date, "yyyymmdd")) And fecvig > 0 Then
                           
                           .BackColor = Shape1(1).FillColor
                        
                        Else
                           
                           .BackColor = Shape1(0).FillColor
                        
                        End If
                   
                   End With
               
               Else '-------> grabar receta sitio remotos
                   
                   If RS.State = 1 Then RS.Close
                   RS.CursorLocation = adUseClient
                   vg_db.CursorLocation = adUseClient
                   Set RS = vg_db.Execute("SELECT DISTINCT rec_cecori FROM cas_b_receta WHERE rec_cecori = '" & vg_codcasino & "' AND rec_codigo = " & vg_newcodrec & "") ', vg_db, adOpenStatic
                   
                   If RS.EOF Then
                      
                      '-------> Actualizar encabezado recetas sitios remotos
                      vg_db.Execute "sgpadm_i_recetasitrem '" & vg_codcasino & "', " & vg_newcodrec & ""
                   
                   End If
                   RS.Close
                   Set RS = Nothing
                   
                   With vaSpread1(1)
                       
                       tiprec = IIf(vg_codregimen > 9999 And vg_codservicio > 9999, vg_codregimen, -1)
                       indlec = 0
                       vg_db.Execute "DELETE cas_b_recetadet FROM cas_b_recetadet WHERE red_cecori = '" & vg_codcasino & "' AND red_tiprec = " & tiprec & " AND red_codigo = " & codigo & ""
                       
                       For i = 1 To .MaxRows
                           
                           .Row = i
                           .Col = 1
                           
                           If .text <> "" Then
                              
                              codpro1 = 0
                              canpro1 = 0
                              pctapr1 = 0
                              pctcoc1 = 0
                              pctnut1 = 0
                              cospro = 0
                              
                              .Col = 1
                              codpro1 = .text
                              
                              .Col = 3
                              canpro1 = .text
                              
                              .Col = 5
                              pctapr1 = .text
                              
                              .Col = 7 'actualizado col
                              pctcoc1 = .text
                              
                              .Col = 9 'actualizado col
                               pctnut1 = .text
                              
                              .Col = 11 'actualizado col
                              cospro = .text
                              
                              vg_db.Execute "sgpadm_i_recetadetsitrem '" & vg_codcasino & "', " & codigo & ", " & i & ", '" & codpro1 & "', " & canpro1 & ", " & cospro & ", " & pctapr1 & ", " & pctcoc1 & ", " & pctnut1 & ", " & tiprec & ", '" & vg_codcasino & "'"
                           
                           End If
                       
                       Next i
                    
                    End With
               
               End If
            
            ElseIf SSTab1.Tab = 2 Then
               
               vg_db.Execute "sgpadm_iu_receta_V05 'M2', " & codigo & ", 0, 0, '', '', '" & IIf(Trim(RichTextBox1(0).text) = "", "", Trim(LimpiaDato(RichTextBox1(0).TextRTF))) & "', '', '', 0, '', 0, '', '', 0, '', 0,0,0,0,0,0,0,0,0,0,0,0,0,0,''"
               vaSpread1(0).Col = 5
               vaSpread1(0).TypeHAlign = TypeHAlignCenter
               vaSpread1(0).Value = IIf(IsNull(RichTextBox1(0).text) Or RichTextBox1(0).text = "", "0", "1")
            
            ElseIf SSTab1.Tab = 3 Then
                
               vg_db.Execute "sgpadm_iu_receta_V05 'M3', " & codigo & ", 0, 0, '', '', '', '','', 0, '', '', '" & IIf(Trim(RichTextBox1(1).text) = "", "", Trim(LimpiaDato(RichTextBox1(1).TextRTF))) & "', '', 0, '',0,0,0,0,0,0,0,0,0,0,0,0,0,0,''"
               vaSpread1(0).Col = 6
               vaSpread1(0).TypeHAlign = TypeHAlignCenter
               vaSpread1(0).Value = IIf(IsNull(RichTextBox1(1).text) Or RichTextBox1(1).text = "", "0", "1")
            
            ElseIf SSTab1.Tab = 4 Then
                
                vg_db.Execute "sgpadm_iu_receta_V05 'M5', " & codigo & ", 0, 0, '', '', '', '','', 0, '', '', '', '', 0, '" & IIf(Trim(RichTextBox1(2).text) = "", "", Trim(LimpiaDato(RichTextBox1(2).TextRTF))) & "',0,0,0,0,0,0,0,0,0,0,0,0,0,0,''"
            
            End If
       
       End If
       
       itexto = 1
       modo = "M"
       Gl_Ac_Botones Me, 3, IIf(vg_newcodrec > 0, 3, 1), modo
       Hab_Des IIf(vg_newcodrec > 0, 0, 1)
       Label1(10).Caption = Format(vaSpread1(0).MaxRows, fg_Pict(7, 0)) & " Registros"
       LimpiarVariableAux
       itexto = 0
       With SSTab1
            
            If vg_newcodrec > 0 Then
               
               .TabEnabled(1) = True
               .TabEnabled(2) = True
               .TabEnabled(3) = True
               .TabEnabled(4) = True
            
            Else
               
               .TabEnabled(1) = True
               .TabEnabled(2) = True
               .TabEnabled(3) = True
               .TabEnabled(4) = True
            
            End If
        
        End With
    
    End If

    Dim MyBufferOfertas        As String
    Dim MyBufferEstacionalidad As String
    Dim MyBufferTipoNegocio    As String
    Dim MyBufferZona           As String
    Dim MyBufferIntolerancia   As String
    Dim MyBufferAlergeno       As String
    Dim MyBufferEstiloAlim     As String
    Dim MyBufferParAdi1        As String
    Dim MyBufferParAdi2        As String
   
    Dim CodOferta              As Long
    
    Dim IdCodigoGral           As Long
    Dim IdZona                 As String
    Dim iselecc                As Long
    Dim RS1                    As New ADODB.Recordset
    Dim descripcion            As String
    Dim unidadReceta           As Long
       
    'crear xlm Ofertas
    Let MyBufferOfertas = ""
    Let MyBufferOfertas = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferOfertas = MyBufferOfertas & "<Oferta>"

    For i = 0 To List1(1).ListCount - 1

        If List1(1).Selected(i) = True Then

           iselecc = 1

        Else

           iselecc = 0

        End If
       
       descripcion = ""
       descripcion = List1(1).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))

       MyBufferOfertas = MyBufferOfertas & " <DetOferta"
       MyBufferOfertas = MyBufferOfertas & " COfe = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferOfertas = MyBufferOfertas & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferOfertas = MyBufferOfertas & "/>"

    Next i
    
    MyBufferOfertas = MyBufferOfertas & "</Oferta>"
    
    Dim Nombre As String
    Dim IdCodigo As Long
    iselecc = 0
    
    ' crear xml Estacionalidad
    
    Let MyBufferEstacionalidad = ""
    Let MyBufferEstacionalidad = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferEstacionalidad = MyBufferEstacionalidad & "<Estac>"
    
    For i = 0 To List1(0).ListCount - 1
   
        If List1(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = List1(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferEstacionalidad = MyBufferEstacionalidad & " <DetEstac"
       MyBufferEstacionalidad = MyBufferEstacionalidad & " Est = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferEstacionalidad = MyBufferEstacionalidad & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferEstacionalidad = MyBufferEstacionalidad & "/>"

    Next i
    
    MyBufferEstacionalidad = MyBufferEstacionalidad & "</Estac>"
  
    ' crear xml tipo negocio
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferTipoNegocio = ""
    Let MyBufferTipoNegocio = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferTipoNegocio = MyBufferTipoNegocio & "<TipNeg>"
    
    For i = 0 To TipoNegocio(0).ListCount - 1
   
        If TipoNegocio(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = TipoNegocio(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferTipoNegocio = MyBufferTipoNegocio & " <DetTipNeg"
       MyBufferTipoNegocio = MyBufferTipoNegocio & " TNeg = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferTipoNegocio = MyBufferTipoNegocio & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferTipoNegocio = MyBufferTipoNegocio & "/>"

    Next i
    
    MyBufferTipoNegocio = MyBufferTipoNegocio & "</TipNeg>"
  
    ' crear xml zona
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferZona = ""
    Let MyBufferZona = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferZona = MyBufferZona & "<Zona>"
    
    For i = 0 To Zona(0).ListCount - 1
   
        If Zona(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = Zona(0).List(i)
       IdZona = 0
       IdZona = fg_codigolistaNuevo(descripcion, 1, 10, "")
       
       MyBufferZona = MyBufferZona & " <DetZona"
       MyBufferZona = MyBufferZona & " Zon = " & Chr(34) & IdZona & Chr(34)
       MyBufferZona = MyBufferZona & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferZona = MyBufferZona & "/>"

    Next i
    
    MyBufferZona = MyBufferZona & "</Zona>"
  
    ' crear xml intolerancia
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferIntolerancia = ""
    Let MyBufferIntolerancia = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferIntolerancia = MyBufferIntolerancia & "<Intol>"
    
    For i = 0 To Intolerancia(0).ListCount - 1
   
        If Intolerancia(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = Intolerancia(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferIntolerancia = MyBufferIntolerancia & " <DetIntol"
       MyBufferIntolerancia = MyBufferIntolerancia & " Intol = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferIntolerancia = MyBufferIntolerancia & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferIntolerancia = MyBufferIntolerancia & "/>"

    Next i
    
    MyBufferIntolerancia = MyBufferIntolerancia & "</Intol>"
  
    ' crear xml alergeno
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferAlergeno = ""
    Let MyBufferAlergeno = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferAlergeno = MyBufferAlergeno & "<Alerg>"
    
    For i = 0 To Alergeno(0).ListCount - 1
   
        If Alergeno(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = Alergeno(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferAlergeno = MyBufferAlergeno & " <DetAlerg"
       MyBufferAlergeno = MyBufferAlergeno & " Alerg = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferAlergeno = MyBufferAlergeno & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferAlergeno = MyBufferAlergeno & "/>"

    Next i
    
    MyBufferAlergeno = MyBufferAlergeno & "</Alerg>"
    
    ' crear xml estilo alimentación
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferEstiloAlim = ""
    Let MyBufferEstiloAlim = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferEstiloAlim = MyBufferEstiloAlim & "<EstAli>"
    
    For i = 0 To EstAli(0).ListCount - 1
   
        If EstAli(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = EstAli(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferEstiloAlim = MyBufferEstiloAlim & " <DetEstAli"
       MyBufferEstiloAlim = MyBufferEstiloAlim & " EstAli = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferEstiloAlim = MyBufferEstiloAlim & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferEstiloAlim = MyBufferEstiloAlim & "/>"

    Next i
    
    MyBufferEstiloAlim = MyBufferEstiloAlim & "</EstAli>"
    
    ' crear xml parametro adicional n°1
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferParAdi1 = ""
    Let MyBufferParAdi1 = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferParAdi1 = MyBufferParAdi1 & "<ParAdi1>"
    
    For i = 0 To ParAdi1(0).ListCount - 1
   
        If ParAdi1(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = ParAdi1(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferParAdi1 = MyBufferParAdi1 & " <DetParAdi1"
       MyBufferParAdi1 = MyBufferParAdi1 & " ParAdi1 = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferParAdi1 = MyBufferParAdi1 & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferParAdi1 = MyBufferParAdi1 & "/>"

    Next i
    
    MyBufferParAdi1 = MyBufferParAdi1 & "</ParAdi1>"
    
    ' crear xml parametro adicional n°2
    iselecc = 0
    IdCodigoGral = 0
    
    Let MyBufferParAdi2 = ""
    Let MyBufferParAdi2 = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBufferParAdi2 = MyBufferParAdi2 & "<ParAdi2>"
    
    For i = 0 To ParAdi2(0).ListCount - 1
   
        If ParAdi2(0).Selected(i) = True Then
           
           iselecc = 1
        
        Else
           
           iselecc = 0
        
        End If
       
       descripcion = ""
       descripcion = ParAdi2(0).List(i)
       IdCodigoGral = 0
       IdCodigoGral = Val(fg_codigolistaNuevo(descripcion, 1, 10, ""))
       
       MyBufferParAdi2 = MyBufferParAdi2 & " <DetParAdi2"
       MyBufferParAdi2 = MyBufferParAdi2 & " ParAdi2 = " & Chr(34) & IdCodigoGral & Chr(34)
       MyBufferParAdi2 = MyBufferParAdi2 & " Sel = " & Chr(34) & iselecc & Chr(34)
       MyBufferParAdi2 = MyBufferParAdi2 & "/>"

    Next i
    
    MyBufferParAdi2 = MyBufferParAdi2 & "</ParAdi2>"
       
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS = vg_db.Execute("sgpadm_InsDelUpd_XmlOfeEstTNegZonaInt '" & MyBufferOfertas & "', '" & MyBufferEstacionalidad & "', '" & MyBufferTipoNegocio & "', '" & MyBufferZona & "', '" & MyBufferIntolerancia & "', '" & MyBufferAlergeno & "', '" & MyBufferEstiloAlim & "', '" & MyBufferParAdi1 & "', '" & MyBufferParAdi2 & "', " & codigo & ", '" & UCase(vg_NUsr) & "'")

    If Not RS.EOF Then
            
       If RS(0) > 0 Then
                   
          MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
          Exit Sub
      
'       Else
   
'          MsgBox "Proceso Termino Correctamente ", vbInformation + vbOKOnly, MsgTitulo
               
       End If
            
    End If
    RS.Close
    Set RS = Nothing
    
    If Combo1(0).ListIndex = -1 Then
       
       MsgBox "debe seleccionar una Undidad de Receta ", vbExclamation
       Exit Sub
    
    End If
    
    unidadReceta = IIf(Combo1(0).ListIndex = -1, "Null", Val(fg_codigolistaNuevo(Combo1(0), 0, 10, "")))
    Sql = " sgpadm_iu_codUnidadReceta "
    Sql = Sql & Trim(codigo) & ","
    Sql = Sql & unidadReceta
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Set RS1 = vg_db.Execute(Sql)
 
    MoverDetalleDatos
    
    Me.HelpContextID = 1093000
    Toolbar1.Buttons(19).ButtonMenus(1).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = vg_OpcM
    Me.HelpContextID = 1094000
    Toolbar1.Buttons(19).ButtonMenus(2).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = vg_OpcM

Case 17 '------> Filtrar
    
    SSTab1.Tab = 0
    B_DieTip.Show 1
    Label2(8).Caption = "Todos": Label2(9).Caption = "Todos"
    If vg_filnomtippla <> "" Then Label2(9).Caption = vg_filnomtippla
    If vg_filnomcatdie <> "" Then Label2(8).Caption = vg_filnomcatdie
    If vg_opcion = 2 Then Exit Sub
    Mover_ListaReceta
    Me.HelpContextID = 1093000
    Toolbar1.Buttons(19).ButtonMenus(1).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = vg_OpcM
    Me.HelpContextID = 1094000
    Toolbar1.Buttons(19).ButtonMenus(2).Enabled = IIf(Mid(ValidaPerfil(Me), 1, 1) = "0", False, True)
    Me.HelpContextID = vg_OpcM

'Case 19 '-------> Cambiar productos en recetas
'    If vaSpread1(0).MaxRows < 1 Then Exit Sub
'    SSTab1.Tab = 0
'    M_ReePro.Show 1
'    Me.HelpContextID = vg_OpcM
Case 21 '-------> Imprimir
    
    If vaSpread1(0).MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, "Receta": Exit Sub
    I_Receta.Label1(12).Caption = Label1(12).Caption
    I_Receta.Label1(11).Caption = Label1(11).Caption
    I_Receta.Label2(8).Caption = Label2(8).Caption
    I_Receta.Label2(9).Caption = Label2(9).Caption
    I_Receta.Show 1

Case 24 '-------> Ver vinculo Ingrediente
    
    With vaSpread1(1)
        
        If .MaxRows < 1 Then Exit Sub
        .Row = .ActiveRow
        .Col = 1
        If .MaxRows < 1 Or Trim(.text) = "" Then Exit Sub
        M_VinPro.LlenaDatos Trim(.text)
        M_VinPro.Show 1
        Me.Refresh
    
    End With

Case 26 '-------> Salir
    
    vg_filcatdie = 0
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If indlec = 1 Then RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: indlec = 0
If indlec = 2 Then RS.Close: Set RS = Nothing: indlec = 0
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Mover_ListaReceta()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""
SSTab1.Enabled = False
With vaSpread1(0)

    .Visible = False
    .Row = -1
    .Col = -1
    .BackColor = Shape1(0).FillColor
    .MaxRows = 0
    itab = 0
    Dim X As Boolean
    
    ' Control displays text tips aligned to pointer with focus
    .TextTip = 2
    .TextTipDelay = 250
    X = .SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    
    '------- Mover encabezado recetas
    Dim IndRec As Long
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Dim OrgCompras As String
    OrgCompras = fg_codigocbo(Combo3, 0, 4, "") 'fg_buscacbostring(Combo3, 0, 4, "")
'    Set RS = vg_db.Execute("sgpadm_s_receta_V06 19, " & vg_codlpr & ", '" & IIf(Check2.Value = 1, "x", "") & "', " & vg_filcatdie & ", " & vg_filtippla & ", " & Val(Vg_FechaDesde) & ", '" & vg_NUsr & "'")
    Set RS = vg_db.Execute("sgpadm_Sel_CostoResumenrecetaOrgCompras_V02 " & vg_filcatdie & ", " & vg_filtippla & ", '" & OrgCompras & "', '" & vg_NUsr & "', '" & IIf(Check2.Value = 1, "x", "") & "', 1, " & Format(FpFecDesde, "yyyymmdd") & "")
    IndRec = 1
    .MaxRows = RS.RecordCount
    
    If Not RS.EOF Then
       
       Do While Not RS.EOF
     
          DoEvents
          .Row = IndRec
         
          .Col = 1
          .TypeHAlign = TypeHAlignLeft
          .Lock = True
          .text = RS!rec_codigo
          
           .Col = -1
           If RS!rec_fecvig <= Val(Format(Date, "yyyymmdd")) And RS!rec_fecvig > 0 Then .BackColor = Shape1(1).FillColor
 
          .Col = 2
          .CellType = CellTypeStaticText
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!rec_nombre)
          
          .Col = 3
          .CellType = CellTypeStaticText
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(Mid(RS!rec_catdie, 1, Len(RS!rec_catdie) - 1))
          
          .Col = 4
          .CellType = CellTypeStaticText
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(Mid(RS!rec_tippla, 1, Len(RS!rec_tippla) - 1))
    
          .Col = 5
          .CellType = CellTypeStaticText
          .Lock = True
          .TypeHAlign = TypeHAlignRight
          .text = IIf(IsNull(RS!cosrec) = True Or Trim(RS!cosrec) = 0#, 0#, Format(CCur(RS!cosrec), fg_Pict(6, 2)))
          
          .Col = 6
          .CellType = CellTypeCheckBox
          .Lock = True
          .TypeHAlign = TypeHAlignCenter
          .FontBold = True
          .Value = IIf(Trim(RS!rec_metpre) = "1", "1", "0")
            
          .Col = 7
          .CellType = CellTypeCheckBox
          .Lock = True
          .TypeHAlign = TypeHAlignCenter
          .FontBold = True
          .Value = IIf(Trim(RS!rec_gruvul) = "1", "1", "0")
          
          .Col = 8
          .CellType = CellTypeStaticText
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!rec_indppr) Or Trim(RS!rec_indppr) = "", "", IIf(RS!rec_indppr = "1", "Real", "Propuesta"))
          
          .Col = 9
          .CellType = CellTypeStaticText
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = IIf(IsNull(RS!rec_tippla) = True, "", Trim(RS!rec_tippla))
             
          .Col = 10
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = RS!ofertas_asoc
  
          .Col = 11
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = RS!Estacionalidad_asoc
  
          .Col = 12
          .Lock = True
          .TypeHAlign = TypeHAlignLeft
          .text = 0
  
          IndRec = IndRec + 1
          RS.MoveNext
       
       Loop
       modo = "M"
       
       If .MaxRows > 1 Then
          
          .Row = 1
          .Col = 1
          codigo = .text
          .SetActiveCell 1, 1
       
       End If
       Label1(1).Visible = True: fpTnombre.Visible = True
       Label1(10).Visible = True: Gl_Ac_Botones Me, 3, 1, modo
       SSTab1.TabEnabled(1) = True: SSTab1.TabEnabled(2) = True: SSTab1.TabEnabled(3) = True
       
       '------- Grabar Categoría Dietética por Defecto Recetas
       If vg_filcatdie > 0 And (Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1") Then
          
          vg_db.Execute "sgpadm_iu_param 'UI', 'catdefecto-" & Trim(vg_NUsr) & "', 'Parametro Categoria Diatetica', 'N', '" & vg_filcatdie & "'"
       
       End If
    
    Else
       
       Label1(1).Visible = True: fpTnombre.Visible = True
       Label1(10).Visible = False: SSTab1.Tab = 0
       SSTab1.TabEnabled(1) = False: SSTab1.TabEnabled(2) = False: SSTab1.TabEnabled(3) = False
       Gl_Ac_Botones Me, 3, 2, modo
    
    End If
    RS.Close
    Set RS = Nothing
    
    SSTab1.Enabled = True
    Label1(10).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
    fpTnombre.text = ""
    .Visible = True
    fg_descarga
    modo = ""
    MoverDetalleDatos
    
End With

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, "Mantención Recetas"

End Sub

Private Sub MoverDatosPropuesta()

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT rec_indppr FROM b_receta with (nolock) WHERE rec_codigo = " & codigo & "")
If Not RS.EOF Then
  
  ComboValOrig = RS!rec_indppr

End If
RS.Close
Set RS = Nothing

End Sub

Private Sub MoverDetalleDatos()

On Error GoTo Man_Error

Dim grnetverd As Double
Dim canpavb As Double
Dim HuellaCarbono As Double
Dim IndentificadorIngSumaTablaGramaje  As String
Dim proveedor As String
Dim Material As String
Dim FIniConv As String
Dim FFinConv As String
Dim RS As New ADODB.Recordset

'-------> esta variables se definen para rescatar precio receta minuta bloque
Dim Fecha As Long
Dim xColIni As Variant, xRowIni As Variant, xcolfin As Variant, xRowFin As Variant
Dim SeleccionOpt As String
Dim unidadReceta As Long
Dim OrgCompras As String
'-------> Fin

Dim X As Boolean

ComboValOrig = "0"
fg_carga ""
Frame1(3).Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
Toolbar2.Enabled = Frame1(3).Enabled
itexto = 1
modo = ""
LimpiarVariable
If vg_newcodrec > 0 Then codigo = vg_newcodrec

Frame1(3).Caption = codigo

'-------> Lectura de encabezado recetas
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

LlenaCombo 0, codigo

Set RS = vg_db.Execute("sgpadm_s_receta_V07 8, " & codigo & ", '', 0, 0, 0, '" & vg_NUsr & "'")
If Not RS.EOF Then
   
   fpText1(0).text = Trim(RS!rec_nombre): vecdatos(0) = Trim(RS!rec_nombre)
   unidadReceta = IIf(IsNull(RS!cod_uniReceta), 0, RS!cod_uniReceta)
   
   fpText1(1).text = Trim(RS!rec_nomfan)
   vecdatos(1) = Trim(RS!rec_nomfan)
   fpayuda(2).text = ""
   fpayuda(2).text = Trim(RS!CatDietetica)
   fpayuda(3).text = ""
   fpayuda(3).text = Trim(RS!tippla)
   fpDouble1(0).Value = RS!rec_basrac
   vecdatos1(0) = RS!rec_basrac
   fpDouble1(4).Value = IIf(IsNull(RS!rec_canser) = True, 0, RS!rec_canser)
   CodCatDie = RS!rec_catdie
   codtipplato = RS!rec_tippla
   fpDateTime1(0).text = IIf(IsNull(RS!rec_fecvig) Or RS!rec_fecvig = 0, "  /  /    ", Mid(RS!rec_fecvig, 7, 2) & "/" & Mid(RS!rec_fecvig, 5, 2) & "/" & Mid(RS!rec_fecvig, 1, 4))
   Check1(0).Value = IIf(IsNull(RS!rec_fecvig) Or RS!rec_fecvig = 0, 0, 1)
   ChAMD.Value = IIf(IsNull(RS!IdIntegraAMD) Or RS!IdIntegraAMD = "1", 1, 0)
   ComboValOrig = ""
   
   'INI unidad medida - Costo Receta - Metodo Cocción - Categorización Compleja - Ing Cruce Garnitura - Efecto Meteorizante - Sellos - Tiempo Cocción - Tiempo HH - Color
        'Mover campos nuevos al combo
        Combo1(0).ListIndex = IIf(IsNull(unidadReceta), -1, fg_buscacboNuevo(Combo1, 0, 10, IIf(IsNull(unidadReceta), -1, unidadReceta)))
        
        If RS!IdCosto > 0 Then
           
           Costo(0).ListIndex = fg_buscacboNuevo(Costo, 0, 10, RS!IdCosto)
        
        End If
        
        fpayuda(0).text = ""
        
        If RS!IdTipoIngPrincipal > 0 Then
        
           TipoIngPrincipal = RS!IdTipoIngPrincipal
           fpayuda(0).text = RS!NombreTipoIngPrincipal
        
        End If
        
        If RS!IdMetodoCoccion > 0 Then
        
           MetodoCoccion = RS!IdMetodoCoccion
           fpayuda(4).text = RS!NombreMetodoCoccion
        
        End If
        
        If RS!IdCategorizacionCompleja > 0 Then
        
           CatCompleja(0).ListIndex = fg_buscacboNuevo(CatCompleja, 0, 10, RS!IdCategorizacionCompleja)
        
        End If
        
        If RS!IdIngCruceGarnitura > 0 Then
        
           IngCruceGarnitura = RS!IdIngCruceGarnitura
           fpayuda(5).text = RS!NombreIngCruceGarnitura
        
        End If
        
        If RS!IdEfectoMeteorizante > 0 Then
        
           EfectoMeteorizante(0).ListIndex = fg_buscacboNuevo(EfectoMeteorizante, 0, 10, RS!IdEfectoMeteorizante)
        
        End If
        
        If RS!IdSellos > 0 Then
        
           Sellos(0).ListIndex = fg_buscacboNuevo(Sellos, 0, 10, RS!IdSellos)
        
        End If
        
        If RS!IdEtiquetadoSello > 0 Then
        
           EtiquetadoSello(0).ListIndex = fg_buscacboNuevo(EtiquetadoSello, 0, 10, RS!IdEtiquetadoSello)
        
        End If
        
        If RS!IdTiempoCoccion > 0 Then
        
           TiempoCoccion = RS!IdTiempoCoccion
           fpayuda(8).text = RS!tiempococcionHora
        
        End If
        
        If RS!IdTiempoHh > 0 Then
        
           TiempoHH = RS!IdTiempoHh
           fpayuda(6).text = RS!tiempohhHora
        
        End If
        
        If RS!IdColor > 0 Then
        
           Color = RS!IdColor
           fpayuda(7).text = RS!NombreColor
           
        End If
   
        fpayuda(1).text = ""
        If RS!IdSegundoingprincipal > 0 Then
        
'           SIPrincipal(0).ListIndex = fg_buscacboNuevo(SIPrincipal, 0, 10, RS!IdSegundoingprincipal)
           fpSIP1(0).Value = RS!IdSegundoingprincipal
           fpayuda(1).text = RS!NombreSegtipoingprincipal
           
        End If
   
        If RS!IdEquipamientoCoccion > 0 Then
        
           ECoccion(0).ListIndex = fg_buscacboNuevo(ECoccion, 0, 10, RS!IdEquipamientoCoccion)
           
        End If
   
   
        If RS!IdParametroSalsa > 0 Then
        
           PSalsa(0).ListIndex = fg_buscacboNuevo(PSalsa, 0, 10, RS!IdParametroSalsa)
           
        End If
   
   'FIN unidad medida - Costo Receta - Metodo Cocción - Categorización Compleja - Ing Cruce Garnitura - Efecto Meteorizante - Sellos - Tiempo Cocción - Tiempo HH - Color
 
   
   If IsNull(RS!rec_indppr) Or Trim(RS!rec_indppr) = "" Then
     
     Combo2(2).ListIndex = -1
   
   Else
     
     Est = True
     ComboValOrig = RS!rec_indppr
     If vg_Indppr = 2 And Not vg_modreceta Then Gl_Ac_BotonesRealPropuesta Me, 1, 1, modo, vg_Indppr, ComboValOrig ': LlenaCombo
     Combo2(2).ListIndex = fg_buscacbo(Combo2, 2, 1, (RS!rec_indppr))
     Est = False
     
     '----->  Se agrega validación
     vg_modreceta = IIf(vg_Indppr <> "2" And vg_newcodrec < 1, False, True)
     If vg_Indppr = 2 And vg_Indppr <> ComboValOrig Then
        
        vg_modreceta = True
        ConfiControlesReceta 1, False
     
     Else
        
        If vg_newcodrec < 1 Then vg_modreceta = False
        Call ConfiControlesReceta(1, True)
     
     End If
     '----->
   
   End If

End If
RS.Close
Set RS = Nothing

' Control displays text tips aligned to pointer with focus
vaSpread1(1).TextTip = 2
vaSpread1(1).TextTipDelay = 2000
X = vaSpread1(1).SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)

'-------> Lectura detalle recetas
cosrec = 0
grnetverd = 0
canpavb = 0
IndentificadorIngSumaTablaGramaje = 0
OrgCompras = fg_codigocbo(Combo3, 0, 4, "")

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If vg_newestrec = False Then
    
    If vg_RecetaReal = 0 Then
        
        If VarSitioRemoto = False Then
            
'            Set RS = vg_db.Execute("sgpadm_s_receta_V06 18, " & codigo & ", '', " & vg_codlpr & ", 0, " & Val(Vg_FechaDesde) & ", '" & vg_NUsr & "'")
           Set RS = vg_db.Execute("sgpadm_Sel_CostoDetalladorecetaOrgCompras_V02 " & codigo & ", '" & OrgCompras & "', 1, " & Format(FpFecDesde, "yyyymmdd") & "")
            
        Else
           
           If vg_opcionmenubloque = "2" Then
           
              M_MinSR2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinSR2.vaSpread1.Col = xColIni
              M_MinSR2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinSR2.vaSpread1.text, 5, Len(M_MinSR2.vaSpread1.text)), "yyyymmdd"))
           
           ElseIf vg_opcionmenubloque = "1" Then
           
              M_MinBloqueADM2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinBloqueADM2.vaSpread1.Col = xColIni
              M_MinBloqueADM2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinBloqueADM2.vaSpread1.text, 5, Len(M_MinBloqueADM2.vaSpread1.text)), "yyyymmdd"))
           
           End If
           
           SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
           Set RS = vg_db.Execute("sgpadm_Sel_DetalleRecetaMinBloque_V03 '" & vg_codcasino & "', " & codigo & ", " & vg_auxtiprec & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Fecha & ", " & SeleccionOpt & "")
           vg_modreceta = True
           RichTextBox1(0).Enabled = False
           ConfiControlesReceta 1, False
        
        End If

    Else
        
        If VarSitioRemoto = False Then
            
            Set RS = vg_db.Execute("sgpadm_s_recetaReal  " & codigo & "," & vg_codsubseg & "," & vg_codregimen & "," & vg_codservicio & "," & vg_Zona & "," & vg_codlpr & "," & Val(Vg_FechaDesde) & " ")
        
        Else
           
           If vg_opcionmenubloque = "2" Then
           
              M_MinSR2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinSR2.vaSpread1.Col = xColIni
              M_MinSR2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinSR2.vaSpread1.text, 5, Len(M_MinSR2.vaSpread1.text)), "yyyymmdd"))
           
           ElseIf vg_opcionmenubloque = "1" Then
           
              M_MinBloqueADM2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinBloqueADM2.vaSpread1.Col = xColIni
              M_MinBloqueADM2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinBloqueADM2.vaSpread1.text, 5, Len(M_MinBloqueADM2.vaSpread1.text)), "yyyymmdd"))
           
           End If
           
           SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
           Set RS = vg_db.Execute("sgpadm_Sel_DetalleRecetaMinBloque_V03 '" & vg_codcasino & "', " & codigo & ", " & vg_auxtiprec & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Fecha & ", " & SeleccionOpt & "")
           vg_modreceta = True
           RichTextBox1(0).Enabled = False
        
        End If
        ConfiControlesReceta 1, False
   
   End If

Else
    If vg_RecetaReal = 0 Then
        
        If VarSitioRemoto = False Then
           
           Set RS = vg_db.Execute("sgpadm_s_receta_V07 18, " & codigo & ", '', " & vg_codlpr & ", 0, " & Val(Vg_FechaDesde) & ", '" & vg_NUsr & "'")
        
        Else
           
           If vg_opcionmenubloque = "2" Then
           
              M_MinSR2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinSR2.vaSpread1.Col = xColIni
              M_MinSR2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinSR2.vaSpread1.text, 5, Len(M_MinSR2.vaSpread1.text)), "yyyymmdd"))
           
           ElseIf vg_opcionmenubloque = "1" Then
           
              M_MinBloqueADM2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinBloqueADM2.vaSpread1.Col = xColIni
              M_MinBloqueADM2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinBloqueADM2.vaSpread1.text, 5, Len(M_MinBloqueADM2.vaSpread1.text)), "yyyymmdd"))
           
           End If
           
           SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
           Set RS = vg_db.Execute("sgpadm_Sel_DetalleRecetaMinBloque_V03 '" & vg_codcasino & "', " & codigo & ", " & vg_auxtiprec & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Fecha & ", " & SeleccionOpt & "")
           vg_modreceta = True
           RichTextBox1(0).Enabled = False
        
        End If
   
   Else
        
        If VarSitioRemoto = False Then
            
            Set RS = vg_db.Execute("sgpadm_s_recetaReal  " & codigo & "," & vg_codsubseg & "," & vg_codregimen & "," & vg_codservicio & "," & vg_Zona & "," & vg_codlpr & "," & Val(Vg_FechaDesde) & "")
        
        Else
           
           If vg_opcionmenubloque = "2" Then
           
              M_MinSR2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinSR2.vaSpread1.Col = xColIni
              M_MinSR2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinSR2.vaSpread1.text, 5, Len(M_MinSR2.vaSpread1.text)), "yyyymmdd"))
           
           ElseIf vg_opcionmenubloque = "1" Then
           
              M_MinBloqueADM2.vaSpread1.GetSelection 1, xColIni, xRowIni, xcolfin, xRowFin
              M_MinBloqueADM2.vaSpread1.Col = xColIni
              M_MinBloqueADM2.vaSpread1.Row = SpreadHeader + 3
              Fecha = CLng(Format(Mid(M_MinBloqueADM2.vaSpread1.text, 5, Len(M_MinBloqueADM2.vaSpread1.text)), "yyyymmdd"))
           
           End If
           
           SeleccionOpt = IIf(M_MinSR1.optPrecioConvenio.Value = True, 1, IIf(M_MinSR1.optPrecioGenerico = True, 2, 3))
           Set RS = vg_db.Execute("sgpadm_Sel_DetalleRecetaMinBloque_V03 '" & vg_codcasino & "', " & codigo & ", " & vg_auxtiprec & ", " & vg_codregimen & ", " & vg_codservicio & ", " & Fecha & ", " & SeleccionOpt & "")
           vg_modreceta = True
           RichTextBox1(0).Enabled = False
        
        End If
        
        ConfiControlesReceta 1, False
   
   End If

End If

vg_RecetaReal = 0
HuellaCarbono = 0

If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      vaSpread1(1).Row = RS!red_nroite
      If vg_RecetaReal = 0 And VarSitioRemoto = False Then
         
         IndentificadorIngSumaTablaGramaje = RS!red_IndentificadorIngSumaTablaGramaje
         proveedor = RS!proveedor
         Material = RS!Material
         FIniConv = Format(RS!FIniConvenios, "dd/mm/yyyy")
         FFinConv = Format(RS!FFinConvenios, "dd/mm/yyyy")
      
      Else
         
         vaSpread1(1).Col = -1
         vaSpread1(1).Row = -1
         vaSpread1(1).Lock = True
         
         IndentificadorIngSumaTablaGramaje = 0
         proveedor = ""
         Material = ""
         FIniConv = ""
         FFinConv = ""
      
      End If
      
      vaSpread1(1).Row = RS!red_nroite
      If IsNull(RS!ing_nombre) = False Then
         
         formatearcelda vaSpread1(1).Row, RS!red_codpro, RS!ing_nombre & " - (" & IIf(RS!ing_indppr = "1", "Real)", "Propuesta)"), RS!unm_nomcor, RS!red_canpro, IIf(IsNull(RS!red_pctapr), 0, RS!red_pctapr), IIf(IsNull(RS!red_pctcoc), 0, RS!red_pctcoc), IIf(IsNull(RS!red_pctnut), 0, RS!red_pctnut), RS!canservida, RS!canneta, RS!precos, True, IndentificadorIngSumaTablaGramaje, proveedor, Material, FIniConv, FFinConv, RS!HuellaCarbono
         If RS!ing_indgrv = 1 Then grnetverd = CCur(grnetverd + RS!cangvneta)
         cosrec = CCur(RS!precos + cosrec)
      
      End If
      
      RS.MoveNext
   
   Loop

End If
RS.Close
Set RS = Nothing

If vaSpread1(1).MaxRows > 0 Then
   
   vaSpread1(1).Row = 1

End If

'-------> Calcular Aporte de alto valor biologigo
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_s_calpavb " & codigo & "")
If Not RS.EOF Then canpavb = Format(RS!canpavb, fg_Pict(6, 2))
RS.Close
Set RS = Nothing

'-------> Calcular resumen aportes Nutricionales
With vaSpread2
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If VarSitioRemoto = False Then
       
       Set RS = vg_db.Execute("sgpadm_s_calresaponut " & codigo & ", " & vg_codsubseg & ", " & vg_codregimen & ", " & IIf(Trim(vg_Zona) = "", 0, vg_Zona) & "")
    
    Else
       
       Set RS = vg_db.Execute("sgpadm_Sel_DetAporteRecetaMinutaBloque_V02 " & codigo & ", " & vg_codregimen & ", '" & vg_codcasino & "'")
    
    End If

    .ClearRange 1, 1, 3, .MaxRows, True ' Limpia grilla aportes nutricionales
    '.DeleteRows 1, .MaxRows
    If Not RS.EOF Then
       
       i = 1
       Do While Not RS.EOF
          
          .Row = i
                   
          .Col = 1
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!nut_codigo)
                   
          .Col = 2
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!nut_nombre)
          
          .Col = 3
          .CellType = CellTypeStaticText
          .TypeHAlign = TypeHAlignRight
          .text = Format(RS!candiet, fg_Pict(6, 2))
          .ForeColor = &HFF0000
             
          If RS!nut_codigo = 3 Then candiet = CCur(candiet + RS!candiet)
          
          RS.MoveNext
          i = i + 1
       
       Loop
    
    End If
    RS.Close
    Set RS = Nothing
    
End With

Label2(4).Caption = Format(grnetverd, fg_Pict(6, 2))
Label2(5).Caption = Format(CCur(canpavb / fpDouble1(0).Value), fg_Pict(6, 2))

If candiet > 0 Then
   
   Label2(7).Caption = Format(CCur(((canpavb / fpDouble1(0).Value) / candiet) * 100), fg_Pict(6, 2))

Else
   
   Label2(7).Caption = Format(0, fg_Pict(6, 2))

End If

If cosrec > 0 Then
   
   Label2(3).Caption = Format(CCur(cosrec / fpDouble1(0).Value), fg_Pict(6, 2))

End If

Call calnetoservido

itexto = 0
modo = "M"
fg_descarga

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub LimpiarVariable()

On Error GoTo Man_Error

fpText1(0).text = ""
fpText1(1).text = ""
fpDouble1(1).Value = ""
fpDouble1(2).Value = ""
fpDouble1(3).Value = ""
If fpDouble1(3).Value = "" Then fpDouble1(4).Value = ""
fpDouble1(5).Value = ""
fpDouble1(6).Value = ""
ChAMD.Value = 0

'Ini : Tipo Ingrediente principal & Segundo & metodo coccion & Ing Cruce Garnitura & Tiempo HH & Color & Tiempo Coccion

fpayuda(0).text = ""
fpayuda(1).text = ""
fpayuda(4).text = ""
fpayuda(5).text = ""
fpayuda(6).text = ""
fpayuda(7).text = ""
fpayuda(8).text = ""

TipoIngPrincipal = 0
fpSIP1(0).Value = 0
MetodoCoccion = 0
IngCruceGarnitura = 0
TiempoHH = 0
Color = 0
TiempoCoccion = 0

'Fin : Tipo Ingrediente principal & Segundo & metodo coccion & Ing Cruce Garnitura & Tiempo HH & Color & Tiempo Coccion


fpayuda(2).text = ""
fpayuda(3).text = ""
vaSpread1(1).MaxRows = 0
vaSpread1(1).MaxRows = 200
Label2(3).Caption = Format(0, fg_Pict(6, 2))
Label2(4).Caption = Format(0, fg_Pict(6, 2))
Label2(5).Caption = Format(0, fg_Pict(6, 2))
Label2(7).Caption = Format(0, fg_Pict(6, 2))
fpDateTime1(0).text = "  /  /    "

For i = 0 To 1
    
    vecdatos(i) = ""

Next i

For i = 0 To 5
    
    vecdatos1(i) = ""

Next i

With vaSpread2
    
    For i = 1 To .MaxRows
       
       .Row = i
       .Col = 3
       .CellType = CellTypeStaticText
       .TypeHAlign = TypeHAlignRight
       .text = Format(0, fg_Pict(6, 2))
       .ForeColor = &HFF0000
    
    Next i

End With
candiet = 0
CodCatDie = 0
codtipplato = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub LimpiarVariableAux()

On Error GoTo Man_Error

For i = 0 To 1
    
    vecdatos(i) = ""

Next i

For i = 0 To 5
    
    vecdatos1(i) = ""

Next i

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Function Hab_Des(op As Integer)

With SSTab1

    Select Case op
    
    Case 0
        
        If modo = "A" Or modo = "M" Then
           
           .TabEnabled(0) = False
           
           If .Tab = 1 Or .Tab = 0 Then
              
              .TabEnabled(1) = True
              .TabEnabled(2) = False
              .TabEnabled(3) = False
              .TabEnabled(4) = False
           
           ElseIf .Tab = 2 Then
              
              .TabEnabled(1) = False
              .TabEnabled(3) = False
           
           ElseIf .Tab = 3 Then
              
              .TabEnabled(1) = False
              .TabEnabled(2) = False
           
           End If
           fpTnombre.Enabled = False
        
        End If
    
    Case 1
        
        .TabEnabled(0) = True
        If .Tab = 1 Or .Tab = 0 Then
           
           .TabEnabled(2) = True
           .TabEnabled(3) = True
        
        ElseIf SSTab1.Tab = 2 Then
           
           .TabEnabled(1) = True
           .TabEnabled(3) = True
        
        ElseIf .Tab = 3 Then
           
           .TabEnabled(1) = True
           .TabEnabled(2) = True
        
        End If
        fpTnombre.Enabled = True
    
    Case 2
        
        .TabEnabled(0) = True
        fpTnombre.Enabled = True
        .TabEnabled(1) = False
    
    Case 3
        
        .Tab = 0
        .TabEnabled(0) = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
    
    Case 4
        
        .TabEnabled(0) = False
        fpTnombre.Enabled = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
    
    Case 5
        
        .TabEnabled(0) = False
        fpTnombre.Enabled = True
        .TabEnabled(1) = False
        .TabEnabled(2) = False
        .TabEnabled(3) = False
    
    End Select

End With

End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Select Case ButtonMenu

Case "Copiar Recetas"
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    vg_swpegreceta = 0
    vg_codreceta = 0
    
    vaSpread1(0).Row = vaSpread1(0).ActiveRow
    vaSpread1(0).Col = 1
    vg_codreceta = Val(vaSpread1(0).Value)
    nombusca = ""
    vg_swpegreceta = 0
    M_CpoRec.Show 1
    Me.Refresh
    If vg_swpegreceta = 1 Then nombusca = fpTnombre.text: fpTnombre.text = "": fpTnombre.text = nombusca

Case "Pegar Recetas"
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    vg_swpegreceta = 0
    vg_codreceta = 0
    vaSpread1(0).Row = vaSpread1(0).ActiveRow
    vaSpread1(0).Col = 1
    vg_codreceta = Val(vaSpread1(0).Value)
    M_PegRec.Show 1
    If SSTab1.TabEnabled(1) = True And vg_swpegreceta = 1 Then MoverDetalleDatos

Case "Mover Recetas"
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    SSTab1.Tab = 0
    vg_swmovreceta = 0
    M_MovRec.LlenarRecetas Label2(9).Caption, vg_filcatdie, vg_filtippla
    M_MovRec.Show 1
    
    If vg_swmovreceta = 1 And fpTnombre.text <> "" Then
       
       nombusca = fpTnombre.text
       fpTnombre.text = ""
       fpTnombre.text = nombusca
    
    ElseIf vg_swmovreceta = 1 And fpTnombre.text = "" Then
       
       fpTnombre.text = " "
    
    End If

Case "Reemplazar Ingredientes"
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    SSTab1.Tab = 0
    M_ReePro.MoverDatosIniciales "ReeIng"
    M_ReePro.Show 1
    Me.HelpContextID = vg_OpcM

Case "Reemplazar % Ingredientes"
    
    If vaSpread1(0).MaxRows < 1 Then Exit Sub
    SSTab1.Tab = 0
    M_ReePro.MoverDatosIniciales "Ree%in"
    M_ReePro.Show 1
    Me.HelpContextID = vg_OpcM


Case "Bach - Input Receta", "Bach - Input Método Receta"
   
   'Abrimos el Commondialog con ShowOpen
    CD.DialogTitle = "Seleccione un archivo excel"
    CD.Filter = "Archivos xls|*.xls|Archivos xlsx|*.xlsx"
    CD.DefaultExt = "*.xls|*.xlsx"
    CD.FilterIndex = 2
    CD.Flags = cdlOFNFileMustExist
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.FileName = ""
    CD.ShowOpen

    'Si seleccionamos un archivo mostramos la ruta
    If CD.FileName <> "" Then

       Dim ObjExcel As excel.Application
       Dim ObjW As excel.Workbook
       Set ObjExcel = New excel.Application
       Set ObjW = ObjExcel.Workbooks.Open(CD.FileName)
       Dim count As Integer
       Dim i     As Integer
       Dim cSpi  As Long

       If MsgBox(IIf(ButtonMenu = "Bach - Input Receta", "Esta Seguro Subir Recetas?...", "Esta Seguro Subir Método Recetas?..."), vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
          
          ObjW.Application.DisplayAlerts = False
          ObjW.Close
          Set ObjExcel = Nothing
          Set ObjW = Nothing
          
          Exit Sub
       
       End If
               
       '-------> Buscar spid
       cSpi = 0
       Set RS = vg_db.Execute("SELECT @@spid spid")
       If Not RS.EOF Then cSpi = RS!spid
       RS.Close: Set RS = Nothing
         
       'Validar datos
       count = 0
       If ButtonMenu = "Bach - Input Receta" Then
       
          For i = 1 To ObjW.Sheets.count
 
              If Trim(ObjW.Sheets(i).Name) = "ENCABEZADO DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "DETALLE DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "OFERTA DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "ESTACIONALIDAD DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "TIPO NEGOCIO RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "ZONA" Or _
                 Trim(ObjW.Sheets(i).Name) = "INTOLERANCIA" Or _
                 Trim(ObjW.Sheets(i).Name) = "ALERGENO" Or _
                 Trim(ObjW.Sheets(i).Name) = "ESTILO ALIMENTACION" Or _
                 Trim(ObjW.Sheets(i).Name) = "PAR 1" Or _
                 Trim(ObjW.Sheets(i).Name) = "PAR 2" _
                 Then
                 
                 count = count + 1
                 If Not ValidarPlantillaExcel(CD.FileName, ObjW.Sheets(i).Name, count, cSpi, True) Then
                     
                    ObjW.Application.DisplayAlerts = False
                    ObjW.Close
                    Set ObjExcel = Nothing
                    Set ObjW = Nothing
                    Exit Sub
                     
                 End If
              
              End If
    
          Next
    
          'Inserta datos
          count = 0
          For i = 1 To ObjW.Sheets.count
 
              If Trim(ObjW.Sheets(i).Name) = "ENCABEZADO DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "DETALLE DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "OFERTA DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "ESTACIONALIDAD DE RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "TIPO NEGOCIO RECETA" Or _
                 Trim(ObjW.Sheets(i).Name) = "ZONA" Or _
                 Trim(ObjW.Sheets(i).Name) = "INTOLERANCIA" Or _
                 Trim(ObjW.Sheets(i).Name) = "ALERGENO" Or _
                 Trim(ObjW.Sheets(i).Name) = "ESTILO ALIMENTACION" Or _
                 Trim(ObjW.Sheets(i).Name) = "PAR 1" Or _
                 Trim(ObjW.Sheets(i).Name) = "PAR 2" _
                 Then
       
                 count = count + 1
                 If Not ValidarPlantillaExcel(CD.FileName, ObjW.Sheets(i).Name, count, cSpi, False) Then
                     
                    ObjW.Application.DisplayAlerts = False
                    ObjW.Close
                    Set ObjExcel = Nothing
                    Set ObjW = Nothing
                    Exit Sub
                     
                 End If
              
              End If
    
          Next
    
          ObjW.Application.DisplayAlerts = False
          ObjW.Close
          Set ObjExcel = Nothing
          Set ObjW = Nothing

          If Not ValidarPlantillaExcel(CD.FileName, "ENCABEZADO DE RECETA", 12, cSpi, True) Then
      
          End If
       
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_Excel"), Me.HelpContextID, "", "", "")
       
          MsgBox "Proceso finalizo correctamente", vbInformation, MsgTitulo
        
       ElseIf ButtonMenu = "Bach - Input Método Receta" Then
       
       
          For i = 1 To ObjW.Sheets.count
 
              If Trim(ObjW.Sheets(i).Name) = "ENCABEZADO DE RECETA" Then
       
                 count = count + 1
                 If Not ValidarPlantillaExcelMetodo(CD.FileName, ObjW.Sheets(i).Name, count, cSpi, True) Then
                     
                    ObjW.Application.DisplayAlerts = False
                    ObjW.Close
                    Set ObjExcel = Nothing
                    Set ObjW = Nothing
                    Exit Sub
                     
                 End If
              
              End If
    
          Next

          If Not ValidarPlantillaExcelMetodo(CD.FileName, "ENCABEZADO DE RECETA", 12, cSpi, True) Then
      
          End If
       
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_Excel_MerPre_ConChef_SugChef"), Me.HelpContextID, "", "", "")
       
          MsgBox "Proceso finalizo correctamente", vbInformation, MsgTitulo
       
       End If
       
    Else
        'Si no mostramos un texto de advertencia de que no se seleccionó _
        ninguno, ya que FileName devuelve una cadena vacía
        MsgBox "No seleccionó ningún archivo"

    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Public Function ValidarPlantillaExcel(NombreArchivo As String, NomSheet As String, opSheet As Integer, cSpi As Long, OpValIns As Boolean) As Boolean

On Error GoTo Man_Error

Dim i                      As Long
Dim PathXls                As String
Dim File_Ext               As String
Dim NomHoja                As String
Dim dbexcel                As Database
Dim cn                     As ADODB.Connection
Dim RS                     As New ADODB.Recordset
Dim RsExcel                As ADODB.Recordset
Dim MyBuffer               As String
Dim MyBuffer_Orig          As String
Dim NomArchivoExcel        As String

'Definición variables excel
Dim xlApp                  As Object
Dim xlWb                   As Object
Dim xlWs                   As Object

Dim CodReceta              As Long
Dim CodDiet                As Long
Dim CodPla                 As Long
Dim nomrec                 As String
Dim NomFanRec              As String
Dim MetPreparacion         As String
Dim ConDelChef             As String
Dim SugDelChef             As String
Dim unidad                 As Long
Dim TipoReceta             As Long
Dim Costo                  As Long
Dim Color                  As Long
Dim MetodoCoccion          As Long
Dim TipoIngPrincipal       As Long
Dim IngCruceGarnitura      As Long
Dim CategorizacionCompleja As Long
Dim Sellos                 As Long
Dim EfectoMeteorizante     As Long
Dim TiempoHH               As Long
Dim TiempoCoccion          As Long
Dim EtiquetadoSello        As Long
Dim SIPrincipal            As Long
Dim ECoccion               As Long
Dim PSalsa                 As Long

Dim NumLin                 As Long
Dim CodIng                 As String
Dim cantidad               As Double
Dim PorAprv                As Double
Dim PorCoc                 As Double
Dim PorNut                 As Double

Dim CodOfe                 As Long
Dim codest                 As Long
Dim CodTipoNegocio         As Long
Dim CodZona                As String
Dim CodIntolerancia        As Long
Dim CodAlergeno            As Long
Dim CodEstiloAlimentacion  As Long
Dim CodPar1                As Long
Dim CodPar2                As Long
Dim CodIntRecetaAMD        As String

Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

ValidarPlantillaExcel = True
PathXls = Trim(NombreArchivo)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))

With cn
     
     Select Case File_Ext
        
        ' Excel 97/2003
        Case "XLS"
          
          .Provider = "Microsoft.Jet.OLEDB.4.0"
          .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
          
        ' Excel 2010
        Case "XLSX"

          .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
          .ConnectionString = "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
     
     End Select
     
     .Open

End With

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<Rec>"
Let MyBuffer_Orig = MyBuffer & "</Rec>"

RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic

RsExcel.Open ("SELECT * FROM [" & NomSheet & "$]"), cn

If RsExcel.EOF Then Exit Function

RsExcel.MoveFirst

If RsExcel.Fields(0).Value = "*" Then
   
   ValidarPlantillaExcel = False
   MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
   
   Exit Function
   
End If

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Or IsNull(RsExcel.Fields(0).Value) Then Exit Do
           
   Select Case opSheet
   
    Case 1 'validar primera hoja
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód# Receta" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Categoría_Dietética" Or _
         Trim(RsExcel.Fields(2).Name) <> "Cód#_Tipo de Plato" Or Trim(RsExcel.Fields(3).Name) <> "Nombre_Receta" Or _
         Trim(RsExcel.Fields(4).Name) <> "Nombre_Receta_Fantasía" Or Trim(RsExcel.Fields(5).Name) <> "Método Preparación" Or _
         Trim(RsExcel.Fields(6).Name) <> "Consejo del Chef" Or Trim(RsExcel.Fields(7).Name) <> "Sugerencia del chef" Or _
         Trim(RsExcel.Fields(8).Name) <> "Unidad" Or _
         Trim(RsExcel.Fields(9).Name) <> "Costo" Or Trim(RsExcel.Fields(10).Name) <> "Color" Or _
         Trim(RsExcel.Fields(11).Name) <> "Metodo_Coccion" Or Trim(RsExcel.Fields(12).Name) <> "Tipo_Ingrediente_Principal" Or _
         Trim(RsExcel.Fields(13).Name) <> "Ing_Cruce_Garnitura" Or Trim(RsExcel.Fields(14).Name) <> "Categorizacion_Compleja" Or _
         Trim(RsExcel.Fields(15).Name) <> "Sellos" Or Trim(RsExcel.Fields(16).Name) <> "Efecto_Meteorizante" Or _
         Trim(RsExcel.Fields(17).Name) <> "Tiempo_HH" Or Trim(RsExcel.Fields(18).Name) <> "Tiempo_Coccion" Or Trim(RsExcel.Fields(19).Name) <> "Etiquetado_Sello" _
         Or Trim(RsExcel.Fields(20).Name) <> "Segundo_Ing_Principal" Or Trim(RsExcel.Fields(21).Name) <> "Equipamiento_Cocción" Or Trim(RsExcel.Fields(22).Name) <> "Parametro_Salsa" Or Trim(RsExcel.Fields(23).Name) <> "Integra_Receta_AMD" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor codigo receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código categoria dietetica esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(2).Value) Or Trim(RsExcel.Fields(2).Value) = "" Or Not IsNumeric(RsExcel.Fields(2).Value) Then
   
         MsgBox "Valor código tipo plato esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(3).Value) Or Trim(RsExcel.Fields(3).Value) = "" Then
   
         MsgBox "Valor nombre receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(4).Value) Or Trim(RsExcel.Fields(4).Value) = "" Then
   
         MsgBox "Valor nombre fantasia receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(8).Value) Or Trim(RsExcel.Fields(8).Value) = "" Or Not IsNumeric(RsExcel.Fields(8).Value) Then
   
         MsgBox "Valor código unidad receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
          
      If IsNull(RsExcel.Fields(9).Value) Or Trim(RsExcel.Fields(9).Value) = "" Or Not IsNumeric(RsExcel.Fields(9).Value) Then
   
         MsgBox "Valor código costo receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(10).Value) Or Trim(RsExcel.Fields(10).Value) = "" Or Not IsNumeric(RsExcel.Fields(10).Value) Then
   
         MsgBox "Valor código color de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(11).Value) Or Trim(RsExcel.Fields(11).Value) = "" Or Not IsNumeric(RsExcel.Fields(11).Value) Then
   
         MsgBox "Valor código metodo cocción de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(12).Value) Or Trim(RsExcel.Fields(12).Value) = "" Or Not IsNumeric(RsExcel.Fields(12).Value) Then
   
         MsgBox "Valor código tipo ingrediente principal de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(13).Value) Or Trim(RsExcel.Fields(13).Value) = "" Or Not IsNumeric(RsExcel.Fields(13).Value) Then
   
         MsgBox "Valor código ingrediente cruce garnitura de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(14).Value) Or Trim(RsExcel.Fields(14).Value) = "" Or Not IsNumeric(RsExcel.Fields(14).Value) Then
   
         MsgBox "Valor código categorización Compleja de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(15).Value) Or Trim(RsExcel.Fields(15).Value) = "" Or Not IsNumeric(RsExcel.Fields(15).Value) Then
   
         MsgBox "Valor código sello de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(16).Value) Or Trim(RsExcel.Fields(16).Value) = "" Or Not IsNumeric(RsExcel.Fields(16).Value) Then
   
         MsgBox "Valor código efecto meteorizante de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(17).Value) Or Trim(RsExcel.Fields(17).Value) = "" Or Not IsNumeric(RsExcel.Fields(17).Value) Then
   
         MsgBox "Valor código tiempo HH de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(18).Value) Or Trim(RsExcel.Fields(18).Value) = "" Or Not IsNumeric(RsExcel.Fields(18).Value) Then
   
         MsgBox "Valor código tiempo cocción de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(19).Value) Or Trim(RsExcel.Fields(19).Value) = "" Or Not IsNumeric(RsExcel.Fields(19).Value) Then
   
         MsgBox "Valor código etiquetado sello de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(20).Value) Or Trim(RsExcel.Fields(20).Value) = "" Or Not IsNumeric(RsExcel.Fields(20).Value) Then
   
         MsgBox "Valor código etiquetado segundo ingrediente principal de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(21).Value) Or Trim(RsExcel.Fields(21).Value) = "" Or Not IsNumeric(RsExcel.Fields(21).Value) Then
   
         MsgBox "Valor código equipamiento cocción de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(22).Value) Or Trim(RsExcel.Fields(22).Value) = "" Or Not IsNumeric(RsExcel.Fields(22).Value) Then
   
         MsgBox "Valor código parametro salsa de la receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(23).Value) Or Trim(RsExcel.Fields(23).Value) = "" Or Not IsNumeric(RsExcel.Fields(23).Value) Then
   
         MsgBox "Valor código integra receta AMD esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      CodReceta = RsExcel.Fields(0).Value
      CodDiet = RsExcel.Fields(1).Value
      CodPla = RsExcel.Fields(2).Value
      nomrec = RsExcel.Fields(3).Value
      NomFanRec = RsExcel.Fields(4).Value
      unidad = RsExcel.Fields(8).Value
      Costo = RsExcel.Fields(9).Value
      Color = RsExcel.Fields(10).Value
      MetodoCoccion = RsExcel.Fields(11).Value
      TipoIngPrincipal = RsExcel.Fields(12).Value
      IngCruceGarnitura = RsExcel.Fields(13).Value
      CategorizacionCompleja = RsExcel.Fields(14).Value
      Sellos = RsExcel.Fields(15).Value
      EfectoMeteorizante = RsExcel.Fields(16).Value
      TiempoHH = RsExcel.Fields(17).Value
      TiempoCoccion = RsExcel.Fields(18).Value
      EtiquetadoSello = RsExcel.Fields(19).Value
      SIPrincipal = RsExcel.Fields(20).Value
      ECoccion = RsExcel.Fields(21).Value
      PSalsa = RsExcel.Fields(22).Value
      CodIntRecetaAMD = RsExcel.Fields(23).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
    
      CodDiet = Replace(Trim(CodDiet), Chr(34), "&quot;")
      CodDiet = Replace(Trim(CodDiet), Chr(38), "&amp;")
      CodDiet = Replace(Trim(CodDiet), Chr(39), "&apos;")
      CodDiet = Replace(Trim(CodDiet), Chr(60), "&lt;")
      CodDiet = Replace(Trim(CodDiet), Chr(62), "&gt;")
    
      CodPla = Replace(Trim(CodPla), Chr(34), "&quot;")
      CodPla = Replace(Trim(CodPla), Chr(38), "&amp;")
      CodPla = Replace(Trim(CodPla), Chr(39), "&apos;")
      CodPla = Replace(Trim(CodPla), Chr(60), "&lt;")
      CodPla = Replace(Trim(CodPla), Chr(62), "&gt;")
    
      nomrec = Replace(Trim(nomrec), Chr(34), "&quot;")
      nomrec = Replace(Trim(nomrec), Chr(38), "&amp;")
      nomrec = Replace(Trim(nomrec), Chr(39), "&apos;")
      nomrec = Replace(Trim(nomrec), Chr(60), "&lt;")
      nomrec = Replace(Trim(nomrec), Chr(62), "&gt;")
    
      NomFanRec = Replace(Trim(NomFanRec), Chr(34), "&quot;")
      NomFanRec = Replace(Trim(NomFanRec), Chr(38), "&amp;")
      NomFanRec = Replace(Trim(NomFanRec), Chr(39), "&apos;")
      NomFanRec = Replace(Trim(NomFanRec), Chr(60), "&lt;")
      NomFanRec = Replace(Trim(NomFanRec), Chr(62), "&gt;")
    
      unidad = Replace(Trim(unidad), Chr(34), "&quot;")
      unidad = Replace(Trim(unidad), Chr(38), "&amp;")
      unidad = Replace(Trim(unidad), Chr(39), "&apos;")
      unidad = Replace(Trim(unidad), Chr(60), "&lt;")
      unidad = Replace(Trim(unidad), Chr(62), "&gt;")
      
      Costo = Replace(Trim(Costo), Chr(34), "&quot;")
      Costo = Replace(Trim(Costo), Chr(38), "&amp;")
      Costo = Replace(Trim(Costo), Chr(39), "&apos;")
      Costo = Replace(Trim(Costo), Chr(60), "&lt;")
      Costo = Replace(Trim(Costo), Chr(62), "&gt;")
      
      Color = Replace(Trim(Color), Chr(34), "&quot;")
      Color = Replace(Trim(Color), Chr(38), "&amp;")
      Color = Replace(Trim(Color), Chr(39), "&apos;")
      Color = Replace(Trim(Color), Chr(60), "&lt;")
      Color = Replace(Trim(Color), Chr(62), "&gt;")
      
      MetodoCoccion = Replace(Trim(MetodoCoccion), Chr(34), "&quot;")
      MetodoCoccion = Replace(Trim(MetodoCoccion), Chr(38), "&amp;")
      MetodoCoccion = Replace(Trim(MetodoCoccion), Chr(39), "&apos;")
      MetodoCoccion = Replace(Trim(MetodoCoccion), Chr(60), "&lt;")
      MetodoCoccion = Replace(Trim(MetodoCoccion), Chr(62), "&gt;")
      
      TipoIngPrincipal = Replace(Trim(TipoIngPrincipal), Chr(34), "&quot;")
      TipoIngPrincipal = Replace(Trim(TipoIngPrincipal), Chr(38), "&amp;")
      TipoIngPrincipal = Replace(Trim(TipoIngPrincipal), Chr(39), "&apos;")
      TipoIngPrincipal = Replace(Trim(TipoIngPrincipal), Chr(60), "&lt;")
      TipoIngPrincipal = Replace(Trim(TipoIngPrincipal), Chr(62), "&gt;")
      
      IngCruceGarnitura = Replace(Trim(IngCruceGarnitura), Chr(34), "&quot;")
      IngCruceGarnitura = Replace(Trim(IngCruceGarnitura), Chr(38), "&amp;")
      IngCruceGarnitura = Replace(Trim(IngCruceGarnitura), Chr(39), "&apos;")
      IngCruceGarnitura = Replace(Trim(IngCruceGarnitura), Chr(60), "&lt;")
      IngCruceGarnitura = Replace(Trim(IngCruceGarnitura), Chr(62), "&gt;")
      
      CategorizacionCompleja = Replace(Trim(CategorizacionCompleja), Chr(34), "&quot;")
      CategorizacionCompleja = Replace(Trim(CategorizacionCompleja), Chr(38), "&amp;")
      CategorizacionCompleja = Replace(Trim(CategorizacionCompleja), Chr(39), "&apos;")
      CategorizacionCompleja = Replace(Trim(CategorizacionCompleja), Chr(60), "&lt;")
      CategorizacionCompleja = Replace(Trim(CategorizacionCompleja), Chr(62), "&gt;")
      
      Sellos = Replace(Trim(Sellos), Chr(34), "&quot;")
      Sellos = Replace(Trim(Sellos), Chr(38), "&amp;")
      Sellos = Replace(Trim(Sellos), Chr(39), "&apos;")
      Sellos = Replace(Trim(Sellos), Chr(60), "&lt;")
      Sellos = Replace(Trim(Sellos), Chr(62), "&gt;")
      
      EfectoMeteorizante = Replace(Trim(EfectoMeteorizante), Chr(34), "&quot;")
      EfectoMeteorizante = Replace(Trim(EfectoMeteorizante), Chr(38), "&amp;")
      EfectoMeteorizante = Replace(Trim(EfectoMeteorizante), Chr(39), "&apos;")
      EfectoMeteorizante = Replace(Trim(EfectoMeteorizante), Chr(60), "&lt;")
      EfectoMeteorizante = Replace(Trim(EfectoMeteorizante), Chr(62), "&gt;")
      
      TiempoHH = Replace(Trim(TiempoHH), Chr(34), "&quot;")
      TiempoHH = Replace(Trim(TiempoHH), Chr(38), "&amp;")
      TiempoHH = Replace(Trim(TiempoHH), Chr(39), "&apos;")
      TiempoHH = Replace(Trim(TiempoHH), Chr(60), "&lt;")
      TiempoHH = Replace(Trim(TiempoHH), Chr(62), "&gt;")
      
      TiempoCoccion = Replace(Trim(TiempoCoccion), Chr(34), "&quot;")
      TiempoCoccion = Replace(Trim(TiempoCoccion), Chr(38), "&amp;")
      TiempoCoccion = Replace(Trim(TiempoCoccion), Chr(39), "&apos;")
      TiempoCoccion = Replace(Trim(TiempoCoccion), Chr(60), "&lt;")
      TiempoCoccion = Replace(Trim(TiempoCoccion), Chr(62), "&gt;")
      
      EtiquetadoSello = Replace(Trim(EtiquetadoSello), Chr(34), "&quot;")
      EtiquetadoSello = Replace(Trim(EtiquetadoSello), Chr(38), "&amp;")
      EtiquetadoSello = Replace(Trim(EtiquetadoSello), Chr(39), "&apos;")
      EtiquetadoSello = Replace(Trim(EtiquetadoSello), Chr(60), "&lt;")
      EtiquetadoSello = Replace(Trim(EtiquetadoSello), Chr(62), "&gt;")
      
      SIPrincipal = Replace(Trim(SIPrincipal), Chr(34), "&quot;")
      SIPrincipal = Replace(Trim(SIPrincipal), Chr(38), "&amp;")
      SIPrincipal = Replace(Trim(SIPrincipal), Chr(39), "&apos;")
      SIPrincipal = Replace(Trim(SIPrincipal), Chr(60), "&lt;")
      SIPrincipal = Replace(Trim(SIPrincipal), Chr(62), "&gt;")
      
      ECoccion = Replace(Trim(ECoccion), Chr(34), "&quot;")
      ECoccion = Replace(Trim(ECoccion), Chr(38), "&amp;")
      ECoccion = Replace(Trim(ECoccion), Chr(39), "&apos;")
      ECoccion = Replace(Trim(ECoccion), Chr(60), "&lt;")
      ECoccion = Replace(Trim(ECoccion), Chr(62), "&gt;")
      
      PSalsa = Replace(Trim(PSalsa), Chr(34), "&quot;")
      PSalsa = Replace(Trim(PSalsa), Chr(38), "&amp;")
      PSalsa = Replace(Trim(PSalsa), Chr(39), "&apos;")
      PSalsa = Replace(Trim(PSalsa), Chr(60), "&lt;")
      PSalsa = Replace(Trim(PSalsa), Chr(62), "&gt;")
      
      CodIntRecetaAMD = Replace(Trim(CodIntRecetaAMD), Chr(34), "&quot;")
      CodIntRecetaAMD = Replace(Trim(CodIntRecetaAMD), Chr(38), "&amp;")
      CodIntRecetaAMD = Replace(Trim(CodIntRecetaAMD), Chr(39), "&apos;")
      CodIntRecetaAMD = Replace(Trim(CodIntRecetaAMD), Chr(60), "&lt;")
      CodIntRecetaAMD = Replace(Trim(CodIntRecetaAMD), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Cd = " & Chr(34) & CodDiet & Chr(34)
      MyBuffer = MyBuffer & " Cp = " & Chr(34) & CodPla & Chr(34)
      MyBuffer = MyBuffer & " Nr = " & Chr(34) & nomrec & Chr(34)
      MyBuffer = MyBuffer & " Nfr = " & Chr(34) & NomFanRec & Chr(34)
      MyBuffer = MyBuffer & " Uni = " & Chr(34) & unidad & Chr(34)
      MyBuffer = MyBuffer & " Cost = " & Chr(34) & Costo & Chr(34)
      MyBuffer = MyBuffer & " Col = " & Chr(34) & Color & Chr(34)
      MyBuffer = MyBuffer & " MeC = " & Chr(34) & MetodoCoccion & Chr(34)
      MyBuffer = MyBuffer & " TIP = " & Chr(34) & TipoIngPrincipal & Chr(34)
      MyBuffer = MyBuffer & " ICG = " & Chr(34) & IngCruceGarnitura & Chr(34)
      MyBuffer = MyBuffer & " CaG = " & Chr(34) & CategorizacionCompleja & Chr(34)
      MyBuffer = MyBuffer & " Sel = " & Chr(34) & Sellos & Chr(34)
      MyBuffer = MyBuffer & " EfM = " & Chr(34) & EfectoMeteorizante & Chr(34)
      MyBuffer = MyBuffer & " THH = " & Chr(34) & TiempoHH & Chr(34)
      MyBuffer = MyBuffer & " TiC = " & Chr(34) & TiempoCoccion & Chr(34)
      MyBuffer = MyBuffer & " EtS = " & Chr(34) & EtiquetadoSello & Chr(34)
      MyBuffer = MyBuffer & " SiP = " & Chr(34) & SIPrincipal & Chr(34)
      MyBuffer = MyBuffer & " EC = " & Chr(34) & ECoccion & Chr(34)
      MyBuffer = MyBuffer & " PS = " & Chr(34) & PSalsa & Chr(34)
      MyBuffer = MyBuffer & " IRAMD = " & Chr(34) & CodIntRecetaAMD & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
    
    Case 2 'validar segunda hoja
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Receta" Or Trim(RsExcel.Fields(1).Name) <> "NOMBRE RECETA" Or _
         Trim(RsExcel.Fields(2).Name) <> "Numero_Línea" Or Trim(RsExcel.Fields(3).Name) <> "Cód#_Ingrediente" Or _
         Trim(RsExcel.Fields(4).Name) <> "Cantidad_Ingrediente" Or Trim(RsExcel.Fields(5).Name) <> "Costo_Ingrediente" Or _
         Trim(RsExcel.Fields(6).Name) <> "%_Aprovechamiento" Or Trim(RsExcel.Fields(7).Name) <> "%_Cocción" Or _
         Trim(RsExcel.Fields(8).Name) <> "%_Nutricional" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(2).Value) Or Trim(RsExcel.Fields(2).Value) = "" Or Not IsNumeric(RsExcel.Fields(2).Value) Then
   
         MsgBox "Valor n. línea esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(3).Value) Or Trim(RsExcel.Fields(3).Value) = "" Then
   
         MsgBox "Valor código ingrediente esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(4).Value) Or Trim(RsExcel.Fields(4).Value) = "" Or Not IsNumeric(RsExcel.Fields(4).Value) Then
   
         MsgBox "Valor cantidad ingrediente esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(6).Value) Or Trim(RsExcel.Fields(6).Value) = "" Or Not IsNumeric(RsExcel.Fields(6).Value) Then
   
         MsgBox "Valor % aprovechamiento esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(7).Value) Or Trim(RsExcel.Fields(7).Value) = "" Or Not IsNumeric(RsExcel.Fields(7).Value) Then
   
         MsgBox "Valor % cocción esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(8).Value) Or Trim(RsExcel.Fields(8).Value) = "" Or Not IsNumeric(RsExcel.Fields(8).Value) Then
   
         MsgBox "Valor % nutricional esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " DETALLE DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(0).Value
      NumLin = RsExcel.Fields(2).Value
      CodIng = RsExcel.Fields(3).Value
      cantidad = RsExcel.Fields(4).Value
      PorAprv = RsExcel.Fields(6).Value
      PorCoc = RsExcel.Fields(7).Value
      PorNut = RsExcel.Fields(8).Value
      
      NumLin = Replace(Trim(NumLin), Chr(34), "&quot;")
      NumLin = Replace(Trim(NumLin), Chr(38), "&amp;")
      NumLin = Replace(Trim(NumLin), Chr(39), "&apos;")
      NumLin = Replace(Trim(NumLin), Chr(60), "&lt;")
      NumLin = Replace(Trim(NumLin), Chr(62), "&gt;")
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodIng = Replace(Trim(CodIng), Chr(34), "&quot;")
      CodIng = Replace(Trim(CodIng), Chr(38), "&amp;")
      CodIng = Replace(Trim(CodIng), Chr(39), "&apos;")
      CodIng = Replace(Trim(CodIng), Chr(60), "&lt;")
      CodIng = Replace(Trim(CodIng), Chr(62), "&gt;")
      
      cantidad = Replace(Trim(cantidad), Chr(34), "&quot;")
      cantidad = Replace(Trim(cantidad), Chr(38), "&amp;")
      cantidad = Replace(Trim(cantidad), Chr(39), "&apos;")
      cantidad = Replace(Trim(cantidad), Chr(60), "&lt;")
      cantidad = Replace(Trim(cantidad), Chr(62), "&gt;")
      
      PorAprv = Replace(Trim(PorAprv), Chr(34), "&quot;")
      PorAprv = Replace(Trim(PorAprv), Chr(38), "&amp;")
      PorAprv = Replace(Trim(PorAprv), Chr(39), "&apos;")
      PorAprv = Replace(Trim(PorAprv), Chr(60), "&lt;")
      PorAprv = Replace(Trim(PorAprv), Chr(62), "&gt;")
      
      PorCoc = Replace(Trim(PorCoc), Chr(34), "&quot;")
      PorCoc = Replace(Trim(PorCoc), Chr(38), "&amp;")
      PorCoc = Replace(Trim(PorCoc), Chr(39), "&apos;")
      PorCoc = Replace(Trim(PorCoc), Chr(60), "&lt;")
      PorCoc = Replace(Trim(PorCoc), Chr(62), "&gt;")
      
      PorNut = Replace(Trim(PorNut), Chr(34), "&quot;")
      PorNut = Replace(Trim(PorNut), Chr(38), "&amp;")
      PorNut = Replace(Trim(PorNut), Chr(39), "&apos;")
      PorNut = Replace(Trim(PorNut), Chr(60), "&lt;")
      PorNut = Replace(Trim(PorNut), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Nl = " & Chr(34) & NumLin & Chr(34)
      MyBuffer = MyBuffer & " Ci = " & Chr(34) & CodIng & Chr(34)
      MyBuffer = MyBuffer & " Ca = " & Chr(34) & cantidad & Chr(34)
      MyBuffer = MyBuffer & " Pa = " & Chr(34) & PorAprv & Chr(34)
      MyBuffer = MyBuffer & " Pc = " & Chr(34) & PorCoc & Chr(34)
      MyBuffer = MyBuffer & " Pn = " & Chr(34) & PorNut & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
      
    Case 3  'validar tercera hoja oferta
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Oferta" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " Oferta Receta"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código oferta esta null o bien tiene datos  mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " Oferta Receta"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " Oferta Receta"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodOfe = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodOfe = Replace(Trim(CodOfe), Chr(34), "&quot;")
      CodOfe = Replace(Trim(CodOfe), Chr(38), "&amp;")
      CodOfe = Replace(Trim(CodOfe), Chr(39), "&apos;")
      CodOfe = Replace(Trim(CodOfe), Chr(60), "&lt;")
      CodOfe = Replace(Trim(CodOfe), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Co = " & Chr(34) & CodOfe & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
    
    Case 4 'validar cuarta hoja estacionalidad
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Estacionalidad" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ESTACIONALIDAD DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código estacionalidad esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ESTACIONALIDAD DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ESTACIONALIDAD DE RECETA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      codest = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      codest = Replace(Trim(codest), Chr(34), "&quot;")
      codest = Replace(Trim(codest), Chr(38), "&amp;")
      codest = Replace(Trim(codest), Chr(39), "&apos;")
      codest = Replace(Trim(codest), Chr(60), "&lt;")
      codest = Replace(Trim(codest), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Ce = " & Chr(34) & codest & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
   
    Case 5 'validar cuarta hoja tipo negocio
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Tipo Negocio" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " TIPO NEGOCIO"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código tipo negocio esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " TIPO NEGOCIO"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " TIPO NEGOCIO"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodTipoNegocio = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodTipoNegocio = Replace(Trim(CodTipoNegocio), Chr(34), "&quot;")
      CodTipoNegocio = Replace(Trim(CodTipoNegocio), Chr(38), "&amp;")
      CodTipoNegocio = Replace(Trim(CodTipoNegocio), Chr(39), "&apos;")
      CodTipoNegocio = Replace(Trim(CodTipoNegocio), Chr(60), "&lt;")
      CodTipoNegocio = Replace(Trim(CodTipoNegocio), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " TiN = " & Chr(34) & CodTipoNegocio & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
   
    Case 6 'validar cuarta hoja zona
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Zona" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ZONA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Then
   
         MsgBox "Valor código zona esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ZONA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ZONA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodZona = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodZona = Replace(Trim(CodZona), Chr(34), "&quot;")
      CodZona = Replace(Trim(CodZona), Chr(38), "&amp;")
      CodZona = Replace(Trim(CodZona), Chr(39), "&apos;")
      CodZona = Replace(Trim(CodZona), Chr(60), "&lt;")
      CodZona = Replace(Trim(CodZona), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Zon = " & Chr(34) & CodZona & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
   
    Case 7 'validar cuarta hoja intolerancia
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Intolerancia" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " INTOLERANCIA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código intolerancia esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " INTOLERANCIA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " INTOLERANCIA"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodIntolerancia = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodIntolerancia = Replace(Trim(CodIntolerancia), Chr(34), "&quot;")
      CodIntolerancia = Replace(Trim(CodIntolerancia), Chr(38), "&amp;")
      CodIntolerancia = Replace(Trim(CodIntolerancia), Chr(39), "&apos;")
      CodIntolerancia = Replace(Trim(CodIntolerancia), Chr(60), "&lt;")
      CodIntolerancia = Replace(Trim(CodIntolerancia), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Intol = " & Chr(34) & CodIntolerancia & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
    
    Case 8 'validar cuarta hoja alergeno
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Alergeno" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ALERGENO"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código alergeno esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ALERGENO"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ALERGENO"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodAlergeno = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodAlergeno = Replace(Trim(CodAlergeno), Chr(34), "&quot;")
      CodAlergeno = Replace(Trim(CodAlergeno), Chr(38), "&amp;")
      CodAlergeno = Replace(Trim(CodAlergeno), Chr(39), "&apos;")
      CodAlergeno = Replace(Trim(CodAlergeno), Chr(60), "&lt;")
      CodAlergeno = Replace(Trim(CodAlergeno), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Ale = " & Chr(34) & CodAlergeno & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
   
    Case 9 'validar cuarta hoja estilo alimentación
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Estilo Alimentación" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ESTILO ALIMENTACION"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código estilo alimentación esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ESTILO ALIMENTACION"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ESTILO ALIMENTACION"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodEstiloAlimentacion = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodEstiloAlimentacion = Replace(Trim(CodEstiloAlimentacion), Chr(34), "&quot;")
      CodEstiloAlimentacion = Replace(Trim(CodEstiloAlimentacion), Chr(38), "&amp;")
      CodEstiloAlimentacion = Replace(Trim(CodEstiloAlimentacion), Chr(39), "&apos;")
      CodEstiloAlimentacion = Replace(Trim(CodEstiloAlimentacion), Chr(60), "&lt;")
      CodEstiloAlimentacion = Replace(Trim(CodEstiloAlimentacion), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " EsA = " & Chr(34) & CodEstiloAlimentacion & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
    
    Case 10 'validar cuarta hoja parametro adicional 1
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Par 1" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " PARAMETRO ADICIONAL N°1"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código paraemtro adicional n°1 esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " EPARAMETRO ADICIONAL N°1"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " PARAMETRO ADICIONAL N°1"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodPar1 = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodPar1 = Replace(Trim(CodPar1), Chr(34), "&quot;")
      CodPar1 = Replace(Trim(CodPar1), Chr(38), "&amp;")
      CodPar1 = Replace(Trim(CodPar1), Chr(39), "&apos;")
      CodPar1 = Replace(Trim(CodPar1), Chr(60), "&lt;")
      CodPar1 = Replace(Trim(CodPar1), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Par1 = " & Chr(34) & CodPar1 & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
    
    Case 11 'validar cuarta hoja parametro adicional n°2
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód#_Par 2" Or Trim(RsExcel.Fields(1).Name) <> "Cód#_Receta" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " PARAMETRO ADICIONAL N°2"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor código paraemtro adicional n°2 esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " EPARAMETRO ADICIONAL N°2"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Or Not IsNumeric(RsExcel.Fields(1).Value) Then
   
         MsgBox "Valor código receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " PARAMETRO ADICIONAL N°2"
         ValidarPlantillaExcel = False
         Exit Do
      
      End If
   
      CodReceta = RsExcel.Fields(1).Value
      CodPar2 = RsExcel.Fields(0).Value
      
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
      
      CodPar2 = Replace(Trim(CodPar2), Chr(34), "&quot;")
      CodPar2 = Replace(Trim(CodPar2), Chr(38), "&amp;")
      CodPar2 = Replace(Trim(CodPar2), Chr(39), "&apos;")
      CodPar2 = Replace(Trim(CodPar2), Chr(60), "&lt;")
      CodPar2 = Replace(Trim(CodPar2), Chr(62), "&gt;")
      
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Par2 = " & Chr(34) & CodPar2 & Chr(34)
              
      MyBuffer = MyBuffer & "/>"
        
    Case 12 'Hoja 1 Metodo Preparacion
    
        CodReceta = RsExcel.Fields(0).Value
        MetPreparacion = IIf(IsNull(RsExcel.Fields(5).Value), "", RsExcel.Fields(5).Value)
        MetPreparacion = Replace(Trim(MetPreparacion), Chr(34), "&quot;")
        MetPreparacion = Replace(Trim(MetPreparacion), Chr(38), "&amp;")
        MetPreparacion = Replace(Trim(MetPreparacion), Chr(39), "&apos;")
        MetPreparacion = Replace(Trim(MetPreparacion), Chr(60), "&lt;")
        MetPreparacion = Replace(Trim(MetPreparacion), Chr(62), "&gt;")
    
        ConDelChef = IIf(IsNull(RsExcel.Fields(6).Value), "", RsExcel.Fields(6).Value)
        ConDelChef = Replace(Trim(ConDelChef), Chr(34), "&quot;")
        ConDelChef = Replace(Trim(ConDelChef), Chr(38), "&amp;")
        ConDelChef = Replace(Trim(ConDelChef), Chr(39), "&apos;")
        ConDelChef = Replace(Trim(ConDelChef), Chr(60), "&lt;")
        ConDelChef = Replace(Trim(ConDelChef), Chr(62), "&gt;")
        
        SugDelChef = IIf(IsNull(RsExcel.Fields(7).Value), "", RsExcel.Fields(7).Value)
        SugDelChef = Replace(Trim(SugDelChef), Chr(34), "&quot;")
        SugDelChef = Replace(Trim(SugDelChef), Chr(38), "&amp;")
        SugDelChef = Replace(Trim(SugDelChef), Chr(39), "&apos;")
        SugDelChef = Replace(Trim(SugDelChef), Chr(60), "&lt;")
        SugDelChef = Replace(Trim(SugDelChef), Chr(62), "&gt;")
    
        Set RS = vg_db.Execute("sgpadm_Upd_ExcelMetPrepa_ConDelChef_SugDelChef " & CodReceta & ", '" & MetPreparacion & "', '" & ConDelChef & "', '" & SugDelChef & "', '" & vg_NUsr & "', '" & cSpi & "'")
   
        If Not RS.EOF Then

            If RS(1) <> "" Then
   
               ValidarPlantillaExcel = False
      
               MsgBox RS(0) & " - " & RS(1), vbCritical, MsgTitulo
               
               RS.Close
               Set RS = Nothing
               
               Exit Function
   
            End If

        End If
        RS.Close
        Set RS = Nothing
   
   End Select
   
   DoEvents
           
   RsExcel.MoveNext
   
Loop
        
MyBuffer = MyBuffer & "</Rec>"

RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing

If ValidarPlantillaExcel Then

Select Case opSheet
   
    Case 1 'primera hoja
    
      If OpValIns Then
         
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

         Else
         
            Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarEncabezadoReceta_V05 '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
         End If
         
      ElseIf Not OpValIns Then
         
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelEncabezadoReceta_V05 '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If

    Case 2 'segunda hoja
    
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
            Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarDetalleReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
            
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelDetalleReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
      End If
      
    Case 3 'tercera hoja oferta
        
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

         Else
          
            Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarOfertaReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
        End If
        
      ElseIf Not OpValIns Then
      
          Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelOfertaReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 4 'cuarta hoja estacionalidad

      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarEstacionalidadReceta_V02 '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelEstacionalidadReceta_V02 '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 5 'quinta hoja tipo negocio

      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarTipoNegocioReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelTipoNegocioReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 6 'sexta hoja zona

      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarZonaReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelZonaReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 7 'septima hoja intolerancia
    
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarIntoleranciaReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelIntoleranciaReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 8 'octava hoja alergeno
    
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarAlergenoReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelAlergenoReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 9 'novena hoja estilo alimentación
    
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarEstiloAlimentacionReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelEstiloAlimentacionReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 10 'decima hoja parametro adicional n°1
    
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarParametroAdicional1Receta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelParametroAdicional1Receta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
    Case 11 'decima primera hoja parametro adicional n°2
    
      If OpValIns Then
      
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcel Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcel = False
            Exit Function

        Else
         
           Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarParametroAdicional2Receta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
         
        End If
      
      ElseIf Not OpValIns Then
      
         Set RS = vg_db.Execute("sgpadm_Ins_XmlExcelParametroAdicional2Receta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
      End If
      
End Select

If opSheet <> 12 Then
   If Not RS.EOF Then

      If RS(1) <> "" And Not OpValIns Then
   
         ValidarPlantillaExcel = False
      
         MsgBox RS(0) & " - " & RS(1), vbCritical, MsgTitulo
   
      ElseIf RS(0) = 0 And OpValIns Then
   
               
         ValidarPlantillaExcel = False
      
         MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo
            
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
         xlApp.Selection.CurrentRegion.Columns.AutoFit
         xlApp.Selection.CurrentRegion.Rows.AutoFit
    
         xlApp.Columns("A:B").Select
         xlApp.Selection.Delete Shift:=xlToLeft
  
         NomArchivoExcel = fg_ArchivoXls("ReporteError_InsercionRecetas")
                    
         xlWb.Close True, NomArchivoExcel

         Dim XL As New excel.Application 'Crea el objeto excel
         XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
         XL.Visible = True
         XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
         '-- Cerrar Excel
         xlApp.Quit
      
         '-------> Release Excel references
         Set xlWs = Nothing
         Set xlWb = Nothing
         Set xlApp = Nothing
      
      End If

   End If
   RS.Close
   Set RS = Nothing

End If

End If

Exit Function
Man_Error:

    Set RsExcel = Nothing
    Set cn = Nothing
    Set RS = Nothing
    fg_descarga
    ValidarPlantillaExcel = False
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Function

Public Function ValidarPlantillaExcelMetodo(NombreArchivo As String, NomSheet As String, opSheet As Integer, cSpi As Long, OpValIns As Boolean) As Boolean

On Error GoTo Man_Error

Dim i               As Long
Dim PathXls         As String
Dim File_Ext        As String
Dim NomHoja         As String
Dim dbexcel         As Database
Dim cn              As ADODB.Connection
Dim RS              As New ADODB.Recordset
Dim RsExcel         As ADODB.Recordset
Dim MyBuffer        As String
Dim MyBuffer_Orig   As String
Dim NomArchivoExcel As String

'Definición variables excel
Dim xlApp           As Object
Dim xlWb            As Object
Dim xlWs            As Object

Dim CodReceta       As Long
Dim nomrec          As String
Dim MetPreparacion  As String
Dim ConDelChef      As String
Dim SugDelChef      As String

Set RsExcel = New ADODB.Recordset
Set cn = New ADODB.Connection

ValidarPlantillaExcelMetodo = True
PathXls = Trim(NombreArchivo)
File_Ext = UCase(Right(CD.FileName, Len(CD.FileName) - (InStrRev(CD.FileName, "."))))

With cn
     
     Select Case File_Ext
        
        ' Excel 97/2003
        Case "XLS"
          
          .Provider = "Microsoft.Jet.OLEDB.4.0"
          .ConnectionString = "Data Source=" & PathXls & ";" & "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
          
        ' Excel 2010
        Case "XLSX"

          .Provider = "Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(PathXls) & ";"
          .ConnectionString = "Extended Properties=Excel 8.0;"
          .CursorLocation = 3
     
     End Select
     
     .Open

End With

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<Rec>"
Let MyBuffer_Orig = MyBuffer & "</Rec>"

RsExcel.CursorLocation = adUseClient
RsExcel.CursorType = adOpenKeyset
RsExcel.LockType = adLockBatchOptimistic

RsExcel.Open ("SELECT * FROM [" & NomSheet & "$]"), cn

If RsExcel.EOF Then Exit Function

RsExcel.MoveFirst

If RsExcel.Fields(0).Value = "*" Then
   
   ValidarPlantillaExcelMetodo = False
   MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
   
   Exit Function
   
End If

Do While RsExcel.EOF <> True
           
   If RsExcel.Fields(0).Value = "*" Or IsNull(RsExcel.Fields(0).Value) Then Exit Do
           
   Select Case opSheet
   
    Case 1 'validar primera hoja
   
      If Trim(RsExcel.Fields(0).Name) <> "Cód# Receta" Or Trim(RsExcel.Fields(2).Name) <> "Método Preparación" Or _
         Trim(RsExcel.Fields(3).Name) <> "Consejo del Chef" Or Trim(RsExcel.Fields(4).Name) <> "Sugerencia del chef" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcelMetodo = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor codigo receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcelMetodo = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(1).Value) Or Trim(RsExcel.Fields(1).Value) = "" Then
   
         MsgBox "Valor nombre receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcelMetodo = False
         Exit Do
      
      End If

      CodReceta = RsExcel.Fields(0).Value
      nomrec = RsExcel.Fields(1).Value
          
      CodReceta = Replace(Trim(CodReceta), Chr(34), "&quot;")
      CodReceta = Replace(Trim(CodReceta), Chr(38), "&amp;")
      CodReceta = Replace(Trim(CodReceta), Chr(39), "&apos;")
      CodReceta = Replace(Trim(CodReceta), Chr(60), "&lt;")
      CodReceta = Replace(Trim(CodReceta), Chr(62), "&gt;")
    
      nomrec = Replace(Trim(nomrec), Chr(34), "&quot;")
      nomrec = Replace(Trim(nomrec), Chr(38), "&amp;")
      nomrec = Replace(Trim(nomrec), Chr(39), "&apos;")
      nomrec = Replace(Trim(nomrec), Chr(60), "&lt;")
      nomrec = Replace(Trim(nomrec), Chr(62), "&gt;")
    
      MyBuffer = MyBuffer & " <RecD"
      MyBuffer = MyBuffer & " Rec = " & Chr(34) & CodReceta & Chr(34)
      MyBuffer = MyBuffer & " Nr = " & Chr(34) & nomrec & Chr(34)
      MyBuffer = MyBuffer & "/>"
    
    Case 12 'Hoja 1 Metodo Preparacion
    
      If Trim(RsExcel.Fields(0).Name) <> "Cód# Receta" Or Trim(RsExcel.Fields(2).Name) <> "Método Preparación" Or _
         Trim(RsExcel.Fields(3).Name) <> "Consejo del Chef" Or Trim(RsExcel.Fields(4).Name) <> "Sugerencia del chef" _
      Then
      
         MsgBox "Primeras columna no corresponde estandar. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcelMetodo = False
         Exit Do
      
      End If
      
      If IsNull(RsExcel.Fields(0).Value) Or Trim(RsExcel.Fields(0).Value) = "" Or Not IsNumeric(RsExcel.Fields(0).Value) Then
   
         MsgBox "Valor codigo receta esta null o bien tiene datos mal ingresado. Proceso cancelado ", vbCritical, MsgTitulo & " ENCABEZADO DE RECETA"
         ValidarPlantillaExcelMetodo = False
         Exit Do
      
      End If
        
      CodReceta = RsExcel.Fields(0).Value
      MetPreparacion = IIf(IsNull(RsExcel.Fields(2).Value), "", RsExcel.Fields(2).Value)
      MetPreparacion = Replace(Trim(MetPreparacion), Chr(34), "&quot;")
      MetPreparacion = Replace(Trim(MetPreparacion), Chr(38), "&amp;")
      MetPreparacion = Replace(Trim(MetPreparacion), Chr(39), "&apos;")
      MetPreparacion = Replace(Trim(MetPreparacion), Chr(60), "&lt;")
      MetPreparacion = Replace(Trim(MetPreparacion), Chr(62), "&gt;")
    
      ConDelChef = IIf(IsNull(RsExcel.Fields(3).Value), "", RsExcel.Fields(3).Value)
      ConDelChef = Replace(Trim(ConDelChef), Chr(34), "&quot;")
      ConDelChef = Replace(Trim(ConDelChef), Chr(38), "&amp;")
      ConDelChef = Replace(Trim(ConDelChef), Chr(39), "&apos;")
      ConDelChef = Replace(Trim(ConDelChef), Chr(60), "&lt;")
      ConDelChef = Replace(Trim(ConDelChef), Chr(62), "&gt;")
        
      SugDelChef = IIf(IsNull(RsExcel.Fields(4).Value), "", RsExcel.Fields(4).Value)
      SugDelChef = Replace(Trim(SugDelChef), Chr(34), "&quot;")
      SugDelChef = Replace(Trim(SugDelChef), Chr(38), "&amp;")
      SugDelChef = Replace(Trim(SugDelChef), Chr(39), "&apos;")
      SugDelChef = Replace(Trim(SugDelChef), Chr(60), "&lt;")
      SugDelChef = Replace(Trim(SugDelChef), Chr(62), "&gt;")
    
      Set RS = vg_db.Execute("sgpadm_Upd_ExcelMetPrepa_ConDelChef_SugDelChef_V01 " & CodReceta & ", '" & MetPreparacion & "', '" & ConDelChef & "', '" & SugDelChef & "', '" & vg_NUsr & "', '" & cSpi & "'")
   
      If Not RS.EOF Then

         If RS(1) <> "" Then
   
            ValidarPlantillaExcelMetodo = False
      
            MsgBox RS(0) & " - " & RS(1), vbCritical, MsgTitulo
               
            RS.Close
            Set RS = Nothing
               
            Exit Function
   
         End If

      End If
      RS.Close
      Set RS = Nothing
   
   End Select
   
   DoEvents
           
   RsExcel.MoveNext
   
Loop
        
MyBuffer = MyBuffer & "</Rec>"

RsExcel.Close
Set RsExcel = Nothing
    
cn.Close
Set cn = Nothing

If ValidarPlantillaExcelMetodo Then

Select Case opSheet
   
    Case 1 'primera hoja
    
      If OpValIns Then
         
         If Trim(MyBuffer) = Trim(MyBuffer_Orig) And ValidarPlantillaExcelMetodo Then

            MsgBox "Problema con la hoja " & NomSheet & ". Proceso cancelado ", vbCritical, MsgTitulo
            ValidarPlantillaExcelMetodo = False
            Exit Function

         Else
         
            Set RS = vg_db.Execute("sgpadm_Sel_XmlValidarEncReceta '" & MyBuffer & "', '', '" & vg_NUsr & "', '" & cSpi & "'")
      
         End If
         
      End If
      
End Select

If opSheet <> 12 Then

   If Not RS.EOF Then

      If RS(1) <> "" And Not OpValIns Then
   
         ValidarPlantillaExcelMetodo = False
      
         MsgBox RS(0) & " - " & RS(1), vbCritical, MsgTitulo
   
      ElseIf RS(0) = 0 And OpValIns Then
   
         ValidarPlantillaExcelMetodo = False
      
         MsgBox RS(1) & VgLinea & VgLinea & "            Adjunto archivo con error", vbCritical, MsgTitulo
            
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
         xlApp.Selection.CurrentRegion.Columns.AutoFit
         xlApp.Selection.CurrentRegion.Rows.AutoFit
    
         xlApp.Columns("A:B").Select
         xlApp.Selection.Delete Shift:=xlToLeft
  
         NomArchivoExcel = fg_ArchivoXls("ReporteError_RecetaMetodo")
                    
         xlWb.Close True, NomArchivoExcel

         Dim XL As New excel.Application 'Crea el objeto excel
         XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
         XL.Visible = True
         XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
         '-- Cerrar Excel
         xlApp.Quit
      
         '-------> Release Excel references
         Set xlWs = Nothing
         Set xlWb = Nothing
         Set xlApp = Nothing
      
      End If

   End If
   RS.Close
   Set RS = Nothing

End If

End If

Exit Function
Man_Error:

    Set RsExcel = Nothing
    Set cn = Nothing
    Set RS = Nothing
    fg_descarga
    ValidarPlantillaExcelMetodo = False
    MsgBox Err & ":  " & Err.Description, vbCritical, MsgTitulo

End Function

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim maxfila     As Long
Dim WsMaxFilas  As Long
Dim RS          As New ADODB.Recordset

With vaSpread1(1)
    
    Select Case Button.Index
    
    Case 1
        
        vg_nombre = ""
        vg_codigo = ""
        vg_left = fpayuda(2).Left + 550
        B_TabEst.LlenaDatos "b_ingrediente", "ing_", "Ingredientes", "AgregarIng"
        B_TabEst.Show 1
        If vg_codigo = "" Then Exit Sub
        
        '------- Revizar si existe ingredientes repetidos
        For i = 1 To .MaxRows
            
            .Row = i
            .Col = 1
            If Trim(.text) = vg_codigo Then
            
               MsgBox "Ingrediente Existe...", vbCritical, MsgTitulo
               Exit Sub
        
            End If
            
        Next i
        
        If vg_codigo <> "" And modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
        
           Gl_Ac_Botones Me, 3, 0, modo
           
           If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
           
              Hab_Des 0
              
           End If
           
        End If
        
        .Row = .ActiveRow
        .Col = 3
        
        If Val(.text) > 0 Then
           
           canpro1 = 0
           codpro1 = ""
           pctnut1 = 0
           
           .Col = 1
           codpro1 = .text
           
           .Col = 3
           canpro1 = .text
           
           .Col = 9 'actualizado col
           pctnut1 = .text
           
           .Col = 11 'actualizado col
           cospro = .text
           
           '------- Resta Aporte
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           Set RS = vg_db.Execute("sgpadm_Sel_CalculoAporteReceta '" & codpro1 & "', " & pctnut1 & ", " & canpro1 & ", " & Val(fpDouble1(0).Value) & "")
           If RS.EOF Then
              
              RS.Close
              Set RS = Nothing
              
           Else
              
              i = 1
              If RS!ing_indgrv = 1 Then
                 
                 Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
              
              End If
              
              Do While Not RS.EOF
                 
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then
                    
                    Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2))
                 
                 End If
                 
                 vaSpread2.Row = i
                 vaSpread2.Col = 3
                 vaSpread2.text = Format(CCur(vaSpread2.text - RS!canneta), fg_Pict(6, 2))
                 i = i + 1
                 RS.MoveNext
              
              Loop
              
              RS.Close
              Set RS = Nothing
           
           End If
           Label2(3).Caption = Format(CCur(Label2(3).Caption - cospro), fg_Pict(6, 2))
           CalTotalPavb
           
        End If
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Set RS = vg_db.Execute("sgpadm_Sel_DetalleIngredienteReceta_V02 '" & vg_codigo & "'")
        If RS.EOF Then
        
           RS.Close
           Set RS = Nothing
           .text = ""
           Exit Sub
           
        End If
        
        formatearcelda .Row, RS!ing_codigo, RS!ing_nombre & " - (" & IIf(RS!ing_indppr = "1", "Real)", "Propuesta)"), RS!unm_nomcor, 0, RS!ing_pctapr, RS!ing_pctcoc, RS!ing_pctnut, 0, 0, 0, RS!Huella_Carbono
        RS.Close
        Set RS = Nothing
        Me.Refresh
        .Row = .ActiveRow
        .SetFocus
        calnetoservido
        Toolbar1.Buttons(24).Enabled = True
    
    Case 2
        
        If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = False Or Toolbar1.Buttons(12).Visible = False Then Hab_Des 0
        For i = 1 To .MaxRows
            
            .Row = i
            .Col = 1
            If .text <> "" Then
            
               maxfila = .Row
                
            End If
            
        Next i
        .Row = .ActiveRow
        .Col = .ActiveCol
        If maxfila < .MaxRows Then
           
           maxfila = maxfila + 1
        
        Else
           
           Exit Sub
        
        End If
        '------- Insertar columna
        If .Row + 1 < 41 Then
           
           .MoveRange 1, (.ActiveRow), .maxcols, (.MaxRows - 1), 1, (.Row + 1)
           .ClearRange 1, .ActiveRow, .maxcols, .ActiveRow, False
        
        End If
    
    Case 3
        
        If MsgBox("Elimina Producto...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
        If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 3, 0, modo: If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then Hab_Des 0
        '------- Resta aporte
        .Row = .ActiveRow
        .Col = 1
        If .text = "" Then
        
           .DeleteRows .Row, 1
           .MaxRows = .MaxRows - 1
           .MaxRows = .MaxRows + 1
           WsMaxFilas = WsMaxFilas - 1
           .Row = .ActiveRow
           Exit Sub
           
        End If
        
        canpro1 = 0
        codpro1 = ""
        pctnut1 = 0
        .Col = 1
        codpro1 = .text
        
        .Col = 3
        canpro1 = .text
        
        .Col = 9 'actualizado col
        pctnut1 = .text
        
        If canpro1 > 0 Then
           
           If RS.State = 1 Then RS.Close
           RS.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient
'
           Set RS = vg_db.Execute("sgpadm_Sel_CalculoAporteReceta '" & codpro1 & "', " & pctnut1 & ", " & canpro1 & ", " & Val(fpDouble1(0).Value) & "")
           i = 1
           If Not RS.EOF Then
              
              If RS!ing_indgrv = 1 Then
              
                 Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
                 
              End If
              
              Do While Not RS.EOF
                 
                 If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then
                    
                    Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2)) 'RS!canneto
                    
                 End If
                 vaSpread2.Row = i
                 vaSpread2.Col = 3
                 vaSpread2.text = Format(CCur(vaSpread2.text - RS!canneta), fg_Pict(6, 2)) 'RS!canneto canneta
                 i = i + 1
                 RS.MoveNext
              
              Loop
              
           End If
           RS.Close
           Set RS = Nothing
           
           '------- Calcular aportes pavb
           CalTotalPavb
        
        End If
        
        .DeleteRows .Row, 1
        .MaxRows = .MaxRows - 1
        .MaxRows = .MaxRows + 1
        cosrec = 0
        
        For i = 1 To .MaxRows
            
            .Row = i
            .Col = 10
            
            If .text <> "" Then
            
               cosrec = CCur(cosrec + .text)
               
            End If
            
        Next i
        Label2(3).Caption = Format((cosrec), fg_Pict(6, 2))
        WsMaxFilas = WsMaxFilas - 1
        .Row = .ActiveRow
        calnetoservido
    
    Case 4
        
        .Row = .ActiveRow
        .Col = .ActiveCol
        If .Row > 1 Then
           
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
           
              Gl_Ac_Botones Me, 3, 0, modo
              
              If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
              
                 Hab_Des 0
                 
              End If
              
           End If
           '------- Copiar datos ultima fila
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow - 1), .maxcols, (.ActiveRow - 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .maxcols, (.ActiveRow - 1), False
           .MoveRange 1, (.ActiveRow), .maxcols, (.ActiveRow), 1, (.ActiveRow - 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .maxcols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .maxcols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow - 1
           .Col = 2
           .SetActiveCell .Col, .Row
        
        End If
    
    Case 5
        
        .Row = .ActiveRow
        .Col = .ActiveCol
        
        If .Row + 1 < 41 Then
           '------- Copiar datos ultima fila
           If modo = "M" And Toolbar1.Buttons(10).Visible = False And Toolbar1.Buttons(12).Visible = False Then
           
              Gl_Ac_Botones Me, 3, 0, modo
              
              If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
              
                 Hab_Des 0
                 
              End If
              
           End If
           .MaxRows = .MaxRows + 1
           .MoveRange 1, (.ActiveRow + 1), .maxcols, (.ActiveRow + 1), 1, .MaxRows
           '------- Copiar datos fila seleccionada
           .ClearRange 1, (.ActiveRow + 1), .maxcols, (.ActiveRow + 1), False
           .MoveRange 1, (.ActiveRow), .maxcols, (.ActiveRow), 1, (.ActiveRow + 1)
           '------- Devolver datos fila y restar ultima fila
           .ClearRange 1, .ActiveRow, .maxcols, .ActiveRow, False
           .MoveRange 1, .MaxRows, .maxcols, .MaxRows, 1, .ActiveRow
           .MaxRows = .MaxRows - 1
           .Row = .ActiveRow + 1
           .Col = 2
           .SetActiveCell .Col, .Row
           
        End If
        
    End Select
    
End With


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar3_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim izquierda As Integer
izquierda = 0

Select Case Button.Index

Case 1
    
    Toolbar3.Buttons(1).Value = 1
    Toolbar3.Buttons(1).Value = 0
    
    If RichTextBox1(Index).SelBold = True Then
       
       RichTextBox1(Index).SelBold = False
       Toolbar3.Buttons(1).Value = 1
       Toolbar3.Buttons(1).Value = 0
    
    ElseIf RichTextBox1(Index).SelBold = False Or IsNull(RichTextBox1(Index).SelBold) Then
       
       RichTextBox1(Index).SelBold = True
       Toolbar3.Buttons(1).Value = 0
       Toolbar3.Buttons(1).Value = 1
    
    End If

Case 2
    
    If RichTextBox1(Index).SelItalic = True Then
       
       RichTextBox1(Index).SelItalic = False
       Toolbar3.Buttons(2).Value = 1
       Toolbar3.Buttons(2).Value = 0
    
    ElseIf RichTextBox1(Index).SelItalic = False Or IsNull(RichTextBox1(Index).SelBold) Then
       
       RichTextBox1(Index).SelItalic = True
       Toolbar3.Buttons(2).Value = 0
       Toolbar3.Buttons(2).Value = 1
    
    End If

Case 3
    
    Toolbar3.Buttons(3).Value = 1
    If RichTextBox1(Index).SelUnderline = True Then
       
       RichTextBox1(Index).SelUnderline = False
       Toolbar3.Buttons(3).Value = 1
       Toolbar3.Buttons(3).Value = 0
    
    ElseIf RichTextBox1(Index).SelUnderline = False Or IsNull(RichTextBox1(Index).SelUnderline) Then
       
       RichTextBox1(Index).SelUnderline = True
       Toolbar3.Buttons(3).Value = 0
       Toolbar3.Buttons(3).Value = 1
    
    End If

Case 5
    
    Toolbar3.Buttons(6).Value = 1
    Toolbar3.Buttons(6).Value = 0
    Toolbar3.Buttons(5).Value = 1
    Toolbar3.Buttons(7).Value = 1
    Toolbar3.Buttons(7).Value = 0
    Toolbar3.Buttons(8).Value = 1
    Toolbar3.Buttons(8).Value = 0
    
    If izquierda = 0 Then
       
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(8).Value = 0
       Toolbar3.Buttons(8).Value = 1
       izquierda = 1
    
    Else
       
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(8).Value = 1
       Toolbar3.Buttons(8).Value = 0
       izquierda = 0
    
    End If
    RichTextBox1(Index).SelAlignment = 0

Case 6
    
    izquierda = 1
    Toolbar3.Buttons(5).Value = 1
    Toolbar3.Buttons(5).Value = 0
    Toolbar3.Buttons(6).Value = 1
    Toolbar3.Buttons(7).Value = 1
    Toolbar3.Buttons(7).Value = 0
    Toolbar3.Buttons(8).Value = 1
    Toolbar3.Buttons(8).Value = 0
    
    If RichTextBox1(Index).SelAlignment = 2 Then
       
       RichTextBox1(Index).SelAlignment = 0
       Toolbar3.Buttons(6).Value = 1
       Toolbar3.Buttons(6).Value = 0
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
    
    Else
       
       RichTextBox1(Index).SelAlignment = 2
       Toolbar3.Buttons(6).Value = 0
       Toolbar3.Buttons(6).Value = 1
    
    End If

Case 7
    
    izquierda = 1
    Toolbar3.Buttons(5).Value = 1
    Toolbar3.Buttons(5).Value = 0
    Toolbar3.Buttons(7).Value = 1
    Toolbar3.Buttons(6).Value = 1
    Toolbar3.Buttons(6).Value = 0
    Toolbar3.Buttons(8).Value = 1
    Toolbar3.Buttons(8).Value = 0
    
    If RichTextBox1(Index).SelAlignment = 1 Then
       
       RichTextBox1(Index).SelAlignment = 0
       Toolbar3.Buttons(7).Value = 1
       Toolbar3.Buttons(7).Value = 0
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
    
    Else
       
       RichTextBox1(Index).SelAlignment = 1
       Toolbar3.Buttons(7).Value = 0
       Toolbar3.Buttons(7).Value = 1
    
    End If

Case 8
    
    Toolbar3.Buttons(6).Value = 1
    Toolbar3.Buttons(6).Value = 0
    Toolbar3.Buttons(8).Value = 1
    Toolbar3.Buttons(7).Value = 1
    Toolbar3.Buttons(7).Value = 0
    Toolbar3.Buttons(5).Value = 1
    Toolbar3.Buttons(8).Value = 0
    
    If izquierda = 1 Then
       
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(8).Value = 0
       Toolbar3.Buttons(8).Value = 1
       izquierda = 0
    
    Else
       
       Toolbar3.Buttons(5).Value = 0
       Toolbar3.Buttons(5).Value = 1
       Toolbar3.Buttons(8).Value = 1
       Toolbar3.Buttons(8).Value = 0
       izquierda = 1
    
    End If
    RichTextBox1(Index).SelAlignment = 0

Case 10
    
    Toolbar3.Buttons(10).Value = 1
    Toolbar3.Buttons(10).Value = 0
    
    If RichTextBox1(Index).SelBullet = True Then
       
       RichTextBox1(Index).SelBullet = False
       Toolbar3.Buttons(10).Value = 1
       Toolbar3.Buttons(10).Value = 0
    
    ElseIf RichTextBox1(Index).SelBullet = False Then
       
       RichTextBox1(Index).SelBullet = True
       Toolbar3.Buttons(10).Value = 0
       Toolbar3.Buttons(10).Value = 1
    
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_Click(Index As Integer, ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If Row = 0 Then Exit Sub

Select Case Index

Case 0
    
    vaSpread1(0).Row = Row
    vaSpread1(0).Col = 1
    codigo = Val(vaSpread1(0).text)
    If vg_Indppr = 2 Then vaSpread1(0).Row = vaSpread1(0).ActiveRow: vaSpread1(0).Col = 7: Gl_Ac_BotonesRealPropuesta Me, 1, 1, modo, vg_Indppr, IIf(vaSpread1(0).text = "Real", "1", IIf(vaSpread1(0).text = "Propuesta", "2", "1")): MoverDatosPropuesta
    Toolbar1.Buttons(24).Enabled = False

Case 1
    
    vaSpread1(1).Row = Row
    vaSpread1(1).Col = 1
    
    If vaSpread1(1).text <> "" Then

          Toolbar1.Buttons(24).Enabled = True
    Else
       
       Toolbar1.Buttons(24).Enabled = False
    
    End If

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_EditMode(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
Dim OrgCompras As String

OrgCompras = fg_codigocbo(Combo3, 0, 4, "")

Select Case Index

Case 1
    
    If modo = "M" And ChangeMade = True Then
      
       If ((vg_Indppr = ComboValOrig) Or vg_Indppr = 3) And vg_PartePlani = False Then
        
          Gl_Ac_Botones Me, 3, 0, modo
          If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
          
              Hab_Des 0
              
          Else
          
             Exit Sub
          
          End If
          
          vg_modreceta = IIf(vg_Indppr <> "2", False, True)
          If vg_Indppr = 2 And vg_Indppr <> ComboValOrig Then
          
             vg_modreceta = True
             ConfiControlesReceta 1, False
             
          Else
             
             vg_modreceta = False
             ConfiControlesReceta 1, True
          
          End If
          
       Else
        
          Exit Sub
      
       End If
    
    End If
    
    Select Case Col
        
        Case 1
            
            vaSpread1(1).Row = Row
            vaSpread1(1).Col = 1
            
            If vaSpread1(1).text = "" Then Exit Sub
            If ChangeMade = False Then
            
               codpro2 = vaSpread1(1).text
               Exit Sub
               
            End If
            
            codpro1 = vaSpread1(1).text
            
            '-------> Revizar si existe ingredientes repetidos
            For i = 1 To vaSpread1(1).MaxRows
                
                vaSpread1(1).Row = i
                vaSpread1(1).Col = 1
                If Trim(vaSpread1(1).text) = codpro1 And Row <> i And Trim(vaSpread1(1).text) <> "" Then
                
                   MsgBox "Ingrediente Existe...", vbCritical, MsgTitulo
                   vaSpread1(1).Row = Row
                   vaSpread1(1).text = codpro2
                   Exit Sub
                   
                End If
                
            Next i
            vaSpread1(1).Row = Row
            vaSpread1(1).Col = 1
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_DetalleIngredienteRecetaRealProp_V02 '" & codpro1 & "', " & vg_Indppr & "")
            If RS.EOF Then
               
               RS.Close
               Set RS = Nothing
               codpro2 = ""
               vaSpread1(1).Col = 1
               vaSpread1(1).text = codpro2
               Exit Sub
            
            End If
            
            formatearcelda vaSpread1(1).Row, RS!ing_codigo, RS!ing_nombre & " - (" & IIf(RS!ing_indppr = 1, "Real)", "Propuesta)"), RS!unm_nomcor, 0, RS!ing_pctapr, RS!ing_pctcoc, RS!ing_pctnut, 0, 0, 0, RS!Huella_Carbono
            RS.Close: Set RS = Nothing
        
        Case 3
            
            vaSpread1(1).Row = Row
            vaSpread1(1).Col = Col
            
            If ChangeMade = False Then
               
               canpro2 = vaSpread1(1).text
               Exit Sub
               
            End If
            canpro1 = vaSpread1(1).text
            '-------> traer precio producto
            vaSpread1(1).Col = 1
            codpro1 = vaSpread1(1).text
            
            If vg_newestrec = False Then
                
                If VarSitioRemoto = False Then
                    
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
'                    Set RS = vg_db.Execute("sgpadm_Sel_TraerPrecioIngredienteReceta " & vg_codlpr & ", " & Val(Vg_FechaDesde) & ", '" & codpro1 & "'")
                    Set RS = vg_db.Execute("sgpadm_Sel_TraerPrecioIngredienteOrgCompras '" & codpro1 & "', '" & OrgCompras & "', " & Format(FpFecDesde, "yyyymmdd") & ", 1")
                
                Else
                    
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    Set RS = vg_db.Execute("SELECT  ing_precos  = SUM(b.red_canpro * i.cpi_precos)" & _
                        " FROM b_receta a with (nolock) " & _
                        "inner join b_recetadet b with (nolock) on a.rec_codigo= b.red_codigo " & _
                        "inner join b_ingrediente c with (nolock) on b.red_codpro  = c.ing_codigo " & _
                        "inner join cas_b_contlistpreing i with (nolock) on i.cpi_coding  = c.Ing_codigo " & _
                        " WHERE a.rec_indppr  = '1'" & _
                        " AND i.cpi_cecori  = '" & vg_codcasino & "'" & _
                        " AND i.cpi_cencos  = '" & vg_codcasino & "'" & _
                        " AND a.rec_codigo  = '" & codpro1 & "'" & _
                        " GROUP BY a.rec_codigo,  a.rec_nombre," & _
                        " a.rec_tippla , a.rec_fecvig, rec_indppr" & _
                        " ORDER BY rec_nombre")
                
                End If
            
            Else
                
                If VarSitioRemoto = False Then
                    
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    Set RS = vg_db.Execute("SELECT (b.dlp_precio/c.pro_facing) AS ing_precos " & _
                             "FROM b_productosing a with (nolock) " & _
                             "inner join b_productos c with (nolock) on a.pri_codpro = c.pro_codigo " & _
                             "inner join b_detlistaprecio b with (nolock) on  c.pro_codigo = b.dlp_codpro " & _
                             "WHERE b.dlp_codigo = " & vg_codlpr & " " & _
                             "AND   b.dlp_anomes = " & Val(Vg_FechaDesde) & " " & _
                             "AND   a.pri_coding = '" & codpro1 & "' AND a.pri_propre = 1")
                
                Else
                    
                    If RS.State = 1 Then RS.Close
                    RS.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    Set RS = vg_db.Execute("SELECT  ing_precos  = SUM(b.red_canpro * i.cpi_precos)" & _
                        " FROM b_receta a with (nolock) " & _
                        "inner join b_recetadet b with (nolock) on  a.rec_codigo= b.red_codigo " & _
                        "inner join b_ingrediente c with (nolock) on b.red_codpro  = c.ing_codigo " & _
                        "inner join cas_b_contlistpreing i with (nolock) on i.cpi_coding  = c.Ing_codigo " & _
                        " WHERE a.rec_indppr  = '1'" & _
                        " AND i.cpi_cecori  = '" & vg_codcasino & "'" & _
                        " AND i.cpi_cencos  = '" & vg_codcasino & "'" & _
                        " AND a.rec_codigo  = '" & codpro1 & "'" & _
                        " GROUP BY a.rec_codigo, a.rec_nombre," & _
                        " a.rec_tippla , a.rec_fecvig, rec_indppr" & _
                        " ORDER BY rec_nombre")
                
                End If
            
            End If
            
            If RS.EOF Then
               
               RS.Close
               Set RS = Nothing ': Exit Sub
            
            Else
                
                If VarSitioRemoto = False Then
                    
                    vaSpread1(1).Col = 11 'actualizado col
                    vaSpread1(1).text = Format(CCur(RS!ing_precos * canpro1), fg_Pict(6, 2))
                    
                    vaSpread1(1).Col = 14
                    vaSpread1(1).ColHidden = True
                    If vg_RecetaReal = 0 And VarSitioRemoto = False Then
   
                       vaSpread1(1).text = "Proveedor : " & RS!proveedor & Chr(13) & "Material SAP : " & RS!Material & Chr(13) & "Fecha Ini Convenio : " & RS!FIniCon & Chr(13) & "Fecha Fin Convenio : " & RS!FFinCon
   
                    End If

                
                Else
                    
                    Call vaSpread1(1).SetText(11, Row, Format(CCur(RS!ing_precos), fg_Pict(6, 2))) 'actualizado col
                
                End If
               
               RS.Close
               Set RS = Nothing
            
            End If
            
            If canpro2 <> canpro1 Then
               
               vaSpread1(1).Col = 1
               codpro1 = vaSpread1(1).text
               
               vaSpread1(1).Col = 9 'actualizado col
               pctnut1 = vaSpread1(1).text
               
               '-------> Calcular gramaje neto
               vaSpread1(1).Col = 10 'actualizado col
               vaSpread1(1).CellType = CellTypeStaticText
               vaSpread1(1).TypeHAlign = TypeHAlignRight
               vaSpread1(1).text = Format(CCur((pctnut1 / 100) * canpro1), fg_Pict(6, vg_RDCa))
               
               '-------> Calcular % limpieza & cocción
               vaSpread1(1).Col = 5
               pctapr1 = vaSpread1(1).text
               
               vaSpread1(1).Col = 7 'actualizado col
               pctcoc1 = vaSpread1(1).text
               
               vaSpread1(1).Col = 8 'actualizado col
               vaSpread1(1).CellType = CellTypeStaticText
               vaSpread1(1).TypeHAlign = TypeHAlignRight
               vaSpread1(1).text = Format(CCur(((pctapr1 / 100) * canpro1) * pctcoc1 / 100), fg_Pict(6, vg_RDCa))
               cosrec = 0
               
               For i = 1 To vaSpread1(1).MaxRows
                   
                   vaSpread1(1).Row = i
                   vaSpread1(1).Col = 11 'actualizado col
                   
                   If vaSpread1(1).text <> "" Then
                   
                      cosrec = CCur(cosrec + IIf(Val(vaSpread1(1).text) = 0, 0, vaSpread1(1).text))
                      
                   End If
               
               Next i
               Label2(3).Caption = Format(CCur(cosrec / Val(fpDouble1(0).Value)), fg_Pict(6, 2))
               
               '-------> Calcular total pavb
               CalTotalPavb
               calnetoservido
               
               '-------> Resta aporte
               If Val(fpDouble1(0).text) < 1 Then Exit Sub
               If canpro2 > 0 Then
               
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               Set RS = vg_db.Execute("sgpadm_Sel_CalculoAporteReceta '" & codpro1 & "'," & pctnut1 & "," & canpro2 & "," & Val(fpDouble1(0).Value) & "")
               If RS.EOF Then
                  
                  RS.Close
                  Set RS = Nothing
               
               Else
                  
                  i = 1
                  If RS!ing_indgrv = 1 Then
                     
                     Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
                
                  End If
                  
                  Do While Not RS.EOF
                     
                     If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then
                        
                        Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2)) 'RS!candiet
                        
                     End If
                     
                     vaSpread2.Row = i
                     vaSpread2.Col = 3
                     vaSpread2.text = Format(CCur(IIf(Val(vaSpread2.text) = 0, 0, vaSpread2.text) - RS!canneta), fg_Pict(6, 2)) 'RS!candiet
                     i = i + 1
                     RS.MoveNext
                     
                   Loop
                   RS.Close
                   Set RS = Nothing
                   
                End If
                
               End If
               '------- Sumar aporte
               If canpro1 < 0 Then Exit Sub
               
               If RS.State = 1 Then RS.Close
               RS.CursorLocation = adUseClient
               vg_db.CursorLocation = adUseClient
               
               Set RS = vg_db.Execute("sgpadm_Sel_CalculoAporteReceta '" & codpro1 & "'," & pctnut1 & "," & canpro1 & "," & Val(fpDouble1(0).Value) & "")
               If RS.EOF Then
               
                  RS.Close
                  Set RS = Nothing
               
               Else
               
                  i = 1
                  If RS!ing_indgrv = 1 Then
                     
                     Label2(4).Caption = Format(CCur(Label2(4).Caption + RS!cangrverneto), fg_Pict(6, 2))
                     
                  End If
                  
                  Do While Not RS.EOF
                     
                     If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then
                     
                        Label2(5).Caption = Format(CCur(Label2(5).Caption + RS!canneta), fg_Pict(6, 2)) 'RS!candiet
                        
                     End If
                     
                     vaSpread2.Row = i
                     vaSpread2.Col = 3
                     vaSpread2.text = Format(CCur(RS!canneta + IIf(Val(vaSpread2.text) = 0, 0, vaSpread2.text)), fg_Pict(6, 2)) 'RS!candiet
                     i = i + 1
                     RS.MoveNext
                  
                  Loop
                  RS.Close
                  Set RS = Nothing
                  
               End If
               
            End If
            
        Case 5, 7
            
            vaSpread1(1).Row = Row
            vaSpread1(1).Col = 1
            If vaSpread1(1).text = "" Then Exit Sub
            
            If ChangeMade = False Then
            
               vaSpread1(1).Col = 5
               pctapr2 = vaSpread1(1).text
               
               vaSpread1(1).Col = 7 'actualizado col
               pctcoc2 = vaSpread1(1).text
               Exit Sub
               
            End If
            vaSpread1(1).Col = 3
            canpro1 = vaSpread1(1).text
            
            vaSpread1(1).Col = 5
            pctapr1 = vaSpread1(1).text
            
            vaSpread1(1).Col = 7 'actualizado col
            pctcoc1 = vaSpread1(1).text
            
            If pctapr1 = 0 Then
            
               vaSpread1(1).Col = 5
               vaSpread1(1).text = pctapr2
               Exit Sub
               
            End If
            
            If pctcoc1 = 0 Then
            
               vaSpread1(1).Col = 7 'actualizado col
               vaSpread1(1).text = pctcoc2
               Exit Sub
               
            End If
            
            '------- Calcular % limpieza & cocción
            vaSpread1(1).Col = 8 'actualizado col
            vaSpread1(1).CellType = CellTypeStaticText
            vaSpread1(1).TypeHAlign = TypeHAlignRight
            vaSpread1(1).text = Format(CCur(((pctapr1 / 100) * canpro1) * (pctcoc1 / 100)), fg_Pict(6, vg_RDCa))
            
            '------- Calcular cantidad Neta % limpieza
            vaSpread1(1).Col = 6 'actualizado col
            vaSpread1(1).CellType = CellTypeStaticText
            vaSpread1(1).TypeHAlign = TypeHAlignRight
            vaSpread1(1).text = Format(CCur(((pctapr1 / 100) * canpro1)), fg_Pict(6, vg_RDCa))
            
            calnetoservido
            
        Case 9
            
            vaSpread1(1).Row = Row
            vaSpread1(1).Col = 1
            If vaSpread1(1).text = "" Then Exit Sub
            If ChangeMade = False Then
            
               vaSpread1(1).Col = 9 'actualizado col
               pctnut2 = vaSpread1(1).text
               Exit Sub
               
            End If
            
            vaSpread1(1).Col = 1
            codpro1 = vaSpread1(1).text
            
            vaSpread1(1).Col = 3
            canpro1 = vaSpread1(1).text
            
            vaSpread1(1).Col = 9 'actualizado col
            pctnut1 = vaSpread1(1).text
            
            If pctnut1 = 0 Then
            
               vaSpread1(1).Col = 9 'actualizado col
               vaSpread1(1).text = pctnut2
               Exit Sub
            
            End If
            
            '------- Resta aporte
            If canpro1 < 0 Then Exit Sub
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_CalculoAporteReceta '" & codpro1 & "'," & pctnut2 & "," & canpro1 & "," & Val(fpDouble1(0).Value) & "")
            If RS.EOF Then
               
               RS.Close
               Set RS = Nothing
               Exit Sub
               
            End If
            i = 1
            If RS!ing_indgrv = 1 Then
            
               Label2(4).Caption = Format(CCur(Label2(4).Caption - RS!cangrverneto), fg_Pict(6, 2))
               
            End If
            
            Do While Not RS.EOF
               
               vaSpread2.Row = i
               If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then
                  
                  Label2(5).Caption = Format(CCur(Label2(5).Caption - RS!canneta), fg_Pict(6, 2)) 'RS!candiet
               
               End If
               vaSpread2.Col = 3
               vaSpread2.text = Format(CCur(IIf(Val(vaSpread2.text) = 0, 0, vaSpread2.text) - RS!canneta), fg_Pict(6, 2)) 'RS!candiet
               i = i + 1
               RS.MoveNext
            
            Loop
            RS.Close
            Set RS = Nothing
            
            '------- Sumar aporte
            If canpro1 < 0 Then Exit Sub
            
            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Set RS = vg_db.Execute("sgpadm_Sel_CalculoAporteReceta '" & codpro1 & "'," & pctnut1 & "," & canpro1 & "," & Val(fpDouble1(0).Value) & "")
            If RS.EOF Then
            
               RS.Close
               Set RS = Nothing
               Exit Sub
            
            End If
            
            i = 1
            If RS!ing_indgrv = 1 Then
            
               Label2(4).Caption = Format(CCur(Label2(4).Caption + RS!cangrverneto), fg_Pict(6, 2))
            
            End If
            Do While Not RS.EOF
               
               vaSpread2.Row = i
               If RS!ing_indpav = 1 And RS!nut_codigo = 3 Then
               
                  Label2(5).Caption = Format(CCur(Label2(5).Caption + RS!canneta), fg_Pict(6, 2)) 'RS!candiet
                  
               End If
               vaSpread2.Col = 3
               vaSpread2.text = Format(CCur(RS!canneta + IIf(Val(vaSpread2.text) = 0, 0, vaSpread2.text)), fg_Pict(6, 2)) 'RS!candiet
               i = i + 1
               RS.MoveNext
            
            Loop
            RS.Close
            Set RS = Nothing
            
            '------- Calcular gramaje neto
            vaSpread1(1).Col = 10 'actualizado col
            vaSpread1(1).CellType = CellTypeStaticText
            vaSpread1(1).TypeHAlign = TypeHAlignRight
            vaSpread1(1).text = Format(CCur((pctnut1 / 100) * canpro1), fg_Pict(6, vg_RDCa))
            '------- Calcular total pavb
            CalTotalPavb
            calnetoservido
            
        Case 12
            
            If modo = "M" Then
                
                If ((vg_Indppr = ComboValOrig) Or vg_Indppr = 3) And vg_PartePlani = False Then
                    
                    Gl_Ac_Botones Me, 3, 0, modo
                    If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
                    
                       Hab_Des 0
                       
                    Else
                       Exit Sub
                       
                    End If
                    vg_modreceta = IIf(vg_Indppr <> "2", False, True)
                    If vg_Indppr = 2 And vg_Indppr <> ComboValOrig Then
                       
                       vg_modreceta = True
                       ConfiControlesReceta 1, False
                    Else
                       
                       vg_modreceta = False
                       ConfiControlesReceta 1, True
                       
                    End If
                
                Else
                    
                    Exit Sub
                
                End If
            
            End If
    
    End Select

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_LeaveCell(Index As Integer, ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)

On Error GoTo Man_Error

Select Case Index

    Case 0
        
        If vaSpread1(0).MaxRows < 1 Or NewRow = -1 Then Exit Sub
        vaSpread1(0).Row = NewRow
        vaSpread1(0).Col = 1
        codigo = Val(vaSpread1(0).text)
        modo = "M"
    
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub CargaMetodoReceta()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Frame4.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
Frame5(0).Enabled = Frame4.Enabled
itexto = 1
RichTextBox1(0).TextRTF = ""
metodoreceta = ""

If vg_newcodrec > 0 Then
   
   codigo = vg_newcodrec
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
            
   Set RS = vg_db.Execute("SELECT rec_nombre FROM b_receta with (nolock) WHERE rec_codigo = " & codigo & "")
   If Not RS.EOF Then
   
      Label3(2).Caption = "(" & codigo & ") " & Trim(RS!rec_nombre)
      
   End If
   RS.Close
   Set RS = Nothing

ElseIf vaSpread1(0).MaxRows > 0 Then
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
            
   Set RS = vg_db.Execute("SELECT rec_nombre FROM b_receta with (nolock) WHERE rec_codigo = " & codigo & "")
   If Not RS.EOF Then
      
      Label3(2).Caption = "(" & codigo & ") " & Trim(RS!rec_nombre)
   
   End If
   RS.Close
   Set RS = Nothing

End If

modo = "M"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT rec_metpre FROM b_receta with (nolock) WHERE rec_codigo = " & codigo & " AND rec_metpre Is Not Null")

If Not RS.EOF Then
   
   RichTextBox1(0).TextRTF = RS!rec_metpre
   metodoreceta = RichTextBox1(0).TextRTF 'fg_bcoenter(RichTextBox1.textRTF) 'LimpiaDato(ConSql!Rcpe_Mthd_Desc)

End If

RS.Close
Set RS = Nothing
itexto = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub CargaGrupoVulnerable()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

Frame4.Enabled = IIf(Mid(ValidarUsuario(Me), 1, 1) = "1" Or Mid(ValidarUsuario(Me), 2, 1) = "1", True, False)
'Frame5(1).Enabled = Frame4.Enabled
itexto = 1
RichTextBox1(1).TextRTF = ""
grupovulnerable = ""

If vg_newcodrec > 0 Then
   
   codigo = vg_newcodrec
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT rec_nombre FROM b_receta with (nolock) WHERE rec_codigo=" & codigo & "")
   If Not RS.EOF Then
   
      Label3(1).Caption = "(" & codigo & ") " & Trim(RS!rec_nombre)
      
   End If
   RS.Close
   Set RS = Nothing

Else
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT rec_nombre FROM b_receta with (nolock) WHERE rec_codigo=" & codigo & "")
   If Not RS.EOF Then
      
      Label3(1).Caption = "(" & codigo & ") " & Trim(RS!rec_nombre)
      
   End If
   RS.Close
   Set RS = Nothing

End If

modo = "M"
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT rec_gruvul FROM b_receta with (nolock) WHERE rec_codigo=" & codigo & " AND rec_gruvul Is Not Null") ', vg_db, adOpenStatic
If Not RS.EOF Then
   
   RichTextBox1(1).TextRTF = IIf(IsNull(RS!rec_gruvul), "", RS!rec_gruvul)
   grupovulnerable = RichTextBox1(1).TextRTF 'fg_bcoenter(RichTextBox1.textRTF) 'LimpiaDato(ConSql!Rcpe_Mthd_Desc)

End If
RS.Close
Set RS = Nothing
itexto = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub CargaHipersensabilidadAlimentaria()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

itexto = 1
RichTextBox1(2).TextRTF = "": HipersensabilidadAlimentaria = ""

If vg_newcodrec > 0 Then
   
   codigo = vg_newcodrec
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT rec_nombre with (nolock) FROM b_receta WHERE rec_codigo=" & codigo & "")
   If Not RS.EOF Then
   
      Label3(3).Caption = "(" & codigo & ") " & Trim(RS!rec_nombre)
    
   End If
   RS.Close
   Set RS = Nothing

Else
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("SELECT rec_nombre FROM b_receta with (nolock) WHERE rec_codigo=" & codigo & "")
   If Not RS.EOF Then
   
      Label3(3).Caption = "(" & codigo & ") " & Trim(RS!rec_nombre)
      
   End If
   RS.Close
   Set RS = Nothing

End If

modo = "M"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("SELECT rec_hipali FROM b_receta with (nolock) WHERE rec_codigo=" & codigo & " AND rec_hipali Is Not Null") ', vg_db, adOpenStatic

If Not RS.EOF Then
   
   RichTextBox1(2).TextRTF = IIf(IsNull(RS!rec_hipali), "", RS!rec_hipali)
   HipersensabilidadAlimentaria = RichTextBox1(2).TextRTF 'fg_bcoenter(RichTextBox1.textRTF) 'LimpiaDato(ConSql!Rcpe_Mthd_Desc)

End If
RS.Close
Set RS = Nothing
itexto = 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub calnetoservido()

On Error GoTo Man_Error

Dim totcservida        As Double
Dim totgneto           As Double
Dim totcbruta          As Double
Dim totgnetoapro       As Double
Dim TotalHuellaCarbono As Double
Dim CantBruta          As Double
Dim CantHuellaCarbono  As Double

With vaSpread1(1)

    If .MaxRows < 1 Then Exit Sub
    
    totcservida = 0
    totgneto = 0
    totgnetoapro = 0
    
    For i = 1 To .MaxRows
        
        .Row = i
        .Col = 3
        
        If Trim(.text) <> "" Then
        
           CantBruta = .text
           
           totcbruta = CCur(totcbruta + .text)
           
           .Col = 15
           If Trim(.text) <> "" Then
           
              CantHuellaCarbono = .text
              TotalHuellaCarbono = TotalHuellaCarbono + (CantBruta * CantHuellaCarbono)
           
           End If
           
           
        End If
        
        .Col = 6 'actualizado col
        If Trim(.text) <> "" Then
           
           totgnetoapro = CCur(totgnetoapro + .text)
        
        End If
        
        .Col = 8 'actualizado col
        If Trim(.text) <> "" Then
           
           totcservida = CCur(totcservida + .text)
        
        End If
        
        .Col = 10 'actualizado col
        If Trim(.text) <> "" Then
        
           totgneto = CCur(totgneto + .text)
           
        End If
    
    Next i
    
    fpDouble1(1).Value = totcservida
    fpDouble1(2).Value = totgneto
    fpDouble1(3).Value = totcbruta
    fpDouble1(6).Value = TotalHuellaCarbono
    
    If fpDouble1(4).text = "" Then
    
       fpDouble1(4).Value = totcservida
       
    End If

    fpDouble1(5).text = totgnetoapro
       
    
End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub CalTotalPavb()

On Error GoTo Man_Error

With vaSpread2

candiet = 0
For i = 1 To .MaxRows
    
    .Row = i
    .Col = 1
    
    If .text = "3" Then
       
       .Col = 3
       candiet = .text
       
       If candiet > 0 Then
          
          Label2(7).Caption = Format(CCur(((Label2(5).Caption / fpDouble1(0).Value) / candiet) * 100), fg_Pict(6, 2))
       
       Else
          
          Label2(7).Caption = Format(0, fg_Pict(6, 2))
       
       End If
       
       Exit For
    
    End If

Next i

End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub formatearcelda(Fila As Long, codpro As String, nompro As String, NomCor As String, canpro As Double, pctapr As Double, pctcoc As Double, pctnut As Double, canservida As Double, canneta As Double, cospro As Double, Optional DatoRrig As Boolean, Optional SumaTablaGramaje As String, Optional proveedor As String, Optional Material As String, Optional FIniConv As String, Optional FFinConv As String, Optional HuellaCarbono As Double)

On Error GoTo Man_Error

With vaSpread1(1)
    
    .Row = Fila
    .Col = 1
    .text = codpro
    If VarSitioRemoto = True Then
'        Let vg_modreceta = False
    End If
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 2
    .text = nompro
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 3
    .CellType = CellTypeCurrency
    .TypeCurrencyDecPlaces = vg_RDCa
    .TypeFloatMin = "-99999999"
    .TypeFloatMax = "99999999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(canpro, fg_Pict(6, vg_RDCa))
    .ForeColor = &HFF0000
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 4
    .TypeHAlign = TypeHAlignLeft
    .text = Trim(NomCor)
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 5
    .CellType = CellTypeCurrency
    .TypeCurrencyDecPlaces = vg_RDCa
    .TypeFloatMin = "0"
    .TypeFloatMax = "9999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(pctapr, fg_Pict(3, vg_RDCa))
    .ForeColor = &HFF0000
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 6
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    If canpro = 0 Then .text = "" Else .text = Format((canpro * pctapr) / 100, fg_Pict(6, vg_RDCa))
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 7
    .CellType = CellTypeCurrency
    .TypeCurrencyDecPlaces = vg_RDCa
    .TypeFloatMin = "0"
    .TypeFloatMax = "999999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(pctcoc, fg_Pict(6, vg_RDCa))
    .ForeColor = &HFF0000
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 8
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    If canservida = 0 Then .text = "" Else .text = Format(canservida, fg_Pict(6, vg_RDCa))
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 9
    .CellType = CellTypeNumber
    .TypeIntegerMin = 2
    .TypeIntegerMax = 100
    .TypeHAlign = 1
    .TypeSpin = False
    .TypeIntegerSpinInc = 1
    .TypeIntegerSpinWrap = False
    .TypeCurrencyShowSymbol = False
    .text = Format(pctnut, fg_Pict(3, vg_RDCa))
    .ForeColor = &HFF0000
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 10
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    If canneta = 0 Then .text = "" Else .text = Format(canneta, fg_Pict(6, vg_RDCa))
    .Lock = IIf(vg_modreceta, True, False)
    
    .Col = 11
    .CellType = CellTypeStaticText
    .TypeHAlign = TypeHAlignRight
    .text = Format(cospro, fg_Pict(7, vg_DPr))
    .Lock = IIf(vg_modreceta, True, False)
    
    If DatoRrig = True Then
       
       .Col = 12
       .text = codpro
    
    End If
    
    .Col = 13
    .ColHidden = True
    If vg_RecetaReal = 0 And VarSitioRemoto = False Then
       
       .CellType = CellTypeCheckBox
       .text = SumaTablaGramaje
       .TypeHAlign = TypeHAlignCenter
       .ColHidden = False
    
    End If
    .Lock = IIf(vg_modreceta, True, False)

    .Col = 15
    .CellType = CellTypeCurrency
    .TypeCurrencyDecPlaces = vg_RDCa
    .TypeFloatMin = "-99999999"
    .TypeFloatMax = "99999999"
    .TypeFloatMoney = False
    .TypeFloatSeparator = True
    .TypeHAlign = TypeHAlignRight
    .TypeFloatCurrencyChar = Asc("$")
    .TypeFloatDecimalChar = Asc(".")
    .TypeFloatSepChar = Asc(",")
    .TypeCurrencyShowSymbol = False
    .text = Format(HuellaCarbono, fg_Pict(6, vg_RDCa))
   
End With

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_TextTipFetch(Index As Integer, ByVal Col As Long, ByVal Row As Long, MultiLine As Integer, TipWidth As Long, TipText As String, ShowTip As Boolean)

On Error GoTo Man_Error

Select Case Index

    Case 0
    
        If vaSpread1(0).MaxRows < 1 Or Col = 1 Then Exit Sub
        ' Set tip to display and set tip's content
        vaSpread1(0).Row = Row
        TipWidth = 4000
        ShowTip = True
        MultiLine = 2
        Select Case Col
        
            Case 1
                
                vaSpread1(0).Col = Col
                TipText = "Código Recetas : " & vaSpread1(0).text
            
            Case 2
                
                vaSpread1(0).Col = Col
                TipText = "Descripción Recetas : " & Trim(vaSpread1(0).text)
            
            Case 3
                
                vaSpread1(0).Col = Col
                TipText = "Categoria Dietetica : " & Trim(vaSpread1(0).text)
            
            Case 4
                
                vaSpread1(0).Col = Col
                TipText = "Tipo Plato : " & Trim(vaSpread1(0).text)
            
            Case 5
                
                vaSpread1(0).Col = Col
                TipText = "Costo Receta : " & Trim(vaSpread1(0).text)
            
            Case 6
                
                vaSpread1(0).Col = Col
                TipText = "Metodo Preparación : " & Trim(vaSpread1(0).text)
            
            Case 7
                
                vaSpread1(0).Col = Col
                TipText = "Grupo Vulnerable : " & Trim(vaSpread1(0).text)
            
            Case 8
                
                vaSpread1(0).Col = Col
                TipText = "Tipo Receta : " & Trim(vaSpread1(0).text)
        
        End Select
    
    Case 1

        If vaSpread1(1).MaxRows < 1 Or Col > 3 Then Exit Sub
        ' Set tip to display and set tip's content
        vaSpread1(1).Row = Row
        TipWidth = 4000
        ShowTip = True
        MultiLine = 2
        Select Case Col
        
            Case 1, 2
                
                vaSpread1(1).Col = 14
                TipText = vaSpread1(1).text

        End Select
        
End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Zona_Click(Index As Integer)
     
On Error GoTo Man_Error

     Gl_Ac_Botones Me, 3, 0, modo
     
     If Toolbar1.Buttons(10).Visible = True Or Toolbar1.Buttons(12).Visible = True Then
        
        Hab_Des 0
        
     Else
     
        Exit Sub
   
     End If
     
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
