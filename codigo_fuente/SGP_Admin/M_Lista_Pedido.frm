VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form M_Lista_Pedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación Pedido"
   ClientHeight    =   10485
   ClientLeft      =   1200
   ClientTop       =   1470
   ClientWidth     =   18525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10485
   ScaleWidth      =   18525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   2730
      Top             =   10290
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9510
      Left            =   60
      TabIndex        =   0
      Top             =   720
      Width           =   18405
      _ExtentX        =   32464
      _ExtentY        =   16775
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Pedidos"
      TabPicture(0)   =   "M_Lista_Pedido.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "sombra(0)"
      Tab(0).Control(1)=   "sombra(1)"
      Tab(0).Control(2)=   "Label7(0)"
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(4)=   "Label1(0)"
      Tab(0).Control(5)=   "fpayuda(0)"
      Tab(0).Control(6)=   "Label2(0)"
      Tab(0).Control(7)=   "Image1(0)"
      Tab(0).Control(8)=   "Label2(1)"
      Tab(0).Control(9)=   "Label7(1)"
      Tab(0).Control(10)=   "sombra(3)"
      Tab(0).Control(11)=   "lbl_proceso"
      Tab(0).Control(12)=   "ProgressBar1"
      Tab(0).Control(13)=   "fpText(1)"
      Tab(0).Control(14)=   "fpText(0)"
      Tab(0).Control(15)=   "fpDateTime1(1)"
      Tab(0).Control(16)=   "fpDateTime1(0)"
      Tab(0).Control(17)=   "vaSpread1"
      Tab(0).Control(18)=   "Combo1(0)"
      Tab(0).Control(19)=   "Combo1(1)"
      Tab(0).Control(20)=   "Frame1"
      Tab(0).Control(21)=   "Frame3"
      Tab(0).ControlCount=   22
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Lista_Pedido.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label5(0)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Lbl_provedor"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Lbl_familuia"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblestado"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fpDateTime1(3)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "fpDateTime1(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "vaSpread2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Frame2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "fpText2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "fpText3"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "fpProveedor"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "fpFamilia"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Frame4"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).ControlCount=   16
      Begin VB.Frame Frame4 
         Caption         =   "Nota"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   11025
         TabIndex        =   55
         Top             =   8610
         Width           =   7050
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Excep. F. Compras"
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
            Left            =   3840
            TabIndex        =   58
            Top             =   420
            Width           =   1515
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   2
            Left            =   3480
            Top             =   450
            Width           =   300
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   2025
            Top             =   450
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "S/Ruta"
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
            Left            =   2385
            TabIndex        =   57
            Top             =   420
            Width           =   600
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0080FF80&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   210
            Top             =   450
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "S/Convenio"
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
            Left            =   570
            TabIndex        =   56
            Top             =   420
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Fecha Limite Confirmacion"
         Height          =   2295
         Left            =   -69960
         TabIndex        =   47
         Top             =   2520
         Width           =   7575
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   315
            Left            =   2400
            TabIndex        =   54
            Top             =   480
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy HH:mm "
            Format          =   62914563
            UpDown          =   -1  'True
            CurrentDate     =   41698
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Actualizar Fecha"
            Height          =   375
            Left            =   1800
            TabIndex        =   53
            Top             =   1680
            Width           =   1095
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   3360
            TabIndex        =   52
            Top             =   1680
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Todos Los Pedido con fecha :"
            Height          =   255
            Left            =   2400
            TabIndex        =   50
            Top             =   1080
            Width           =   2535
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pedido Selecionado"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   1080
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label lbl_fecha 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   4680
            TabIndex        =   51
            Top             =   1080
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha Limite Confirmación"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Index           =   2
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   1935
         End
      End
      Begin EditLib.fpText fpFamilia 
         Height          =   315
         Left            =   6945
         TabIndex        =   40
         Top             =   8700
         Width           =   3045
         _Version        =   196608
         _ExtentX        =   5371
         _ExtentY        =   556
         Enabled         =   0   'False
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpProveedor 
         Height          =   315
         Left            =   1725
         TabIndex        =   38
         Top             =   8745
         Width           =   4170
         _Version        =   196608
         _ExtentX        =   7355
         _ExtentY        =   556
         Enabled         =   0   'False
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText3 
         Height          =   315
         Left            =   6750
         TabIndex        =   30
         Top             =   645
         Width           =   3240
         _Version        =   196608
         _ExtentX        =   5715
         _ExtentY        =   556
         Enabled         =   0   'False
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText2 
         Height          =   315
         Left            =   2535
         TabIndex        =   28
         Top             =   660
         Width           =   1560
         _Version        =   196608
         _ExtentX        =   2752
         _ExtentY        =   556
         Enabled         =   0   'False
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
         ThreeDInsideStyle=   1
         ThreeDInsideHighlightColor=   -2147483633
         ThreeDInsideShadowColor=   -2147483642
         ThreeDInsideWidth=   1
         ThreeDOutsideStyle=   1
         ThreeDOutsideHighlightColor=   -2147483628
         ThreeDOutsideShadowColor=   -2147483632
         ThreeDOutsideWidth=   1
         ThreeDFrameWidth=   0
         BorderStyle     =   0
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
         ThreeDTextHighlightColor=   -2147483633
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
         InvalidColor    =   -2147483637
         InvalidOption   =   0
         MarginLeft      =   3
         MarginTop       =   3
         MarginRight     =   3
         MarginBottom    =   3
         NullColor       =   -2147483637
         OnFocusAlignH   =   0
         OnFocusAlignV   =   0
         OnFocusNoSelect =   0   'False
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
         ThreeDFrameColor=   -2147483633
         Appearance      =   2
         BorderDropShadow=   0
         BorderDropShadowColor=   -2147483632
         BorderDropShadowWidth=   3
         ButtonColor     =   -2147483633
         AutoMenu        =   0   'False
         ButtonAlign     =   0
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Frame Frame2 
         Caption         =   "Materiales Asociados al Ingredientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   4965
         Left            =   960
         TabIndex        =   23
         Top             =   3240
         Visible         =   0   'False
         Width           =   16335
         Begin EditLib.fpText fp_descripcion 
            Height          =   315
            Left            =   2835
            TabIndex        =   45
            Top             =   510
            Width           =   5265
            _Version        =   196608
            _ExtentX        =   9287
            _ExtentY        =   556
            Enabled         =   0   'False
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
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
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
            ThreeDTextHighlightColor=   -2147483633
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
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText fp_codigo 
            Height          =   315
            Left            =   1650
            TabIndex        =   44
            Top             =   495
            Width           =   1125
            _Version        =   196608
            _ExtentX        =   1984
            _ExtentY        =   556
            Enabled         =   0   'False
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
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
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
            ThreeDTextHighlightColor=   -2147483633
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
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin EditLib.fpText Fp_proveedor 
            Height          =   315
            Left            =   9255
            TabIndex        =   42
            Top             =   540
            Width           =   5610
            _Version        =   196608
            _ExtentX        =   9895
            _ExtentY        =   556
            Enabled         =   0   'False
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
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
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
            ThreeDTextHighlightColor=   -2147483633
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
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Cancelar"
            Height          =   405
            Left            =   13605
            TabIndex        =   25
            Top             =   4275
            Width           =   1335
         End
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   2940
            Left            =   180
            TabIndex        =   24
            Top             =   1080
            Width           =   15990
            _Version        =   393216
            _ExtentX        =   28205
            _ExtentY        =   5186
            _StockProps     =   64
            AutoClipboard   =   0   'False
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
            MaxCols         =   10
            SpreadDesigner  =   "M_Lista_Pedido.frx":0038
         End
         Begin VB.Label Label8 
            Caption         =   "Material Sap"
            Height          =   345
            Left            =   435
            TabIndex        =   43
            Top             =   570
            Width           =   1050
         End
         Begin VB.Label Label6 
            Caption         =   "Proveedor"
            Height          =   285
            Left            =   8310
            TabIndex        =   41
            Top             =   570
            Width           =   930
         End
         Begin VB.Label Label3 
            Caption         =   "Al hacer Doble Click Selecionara el Producto Actualizado en la Grilla"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   285
            Left            =   2160
            TabIndex        =   26
            Top             =   4410
            Width           =   7665
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Errores"
         Height          =   5250
         Left            =   -71850
         TabIndex        =   20
         Top             =   2895
         Width           =   9495
         Begin VB.CommandButton Command2 
            Caption         =   "Aceptar"
            Height          =   405
            Left            =   7455
            TabIndex        =   22
            Top             =   4575
            Width           =   1335
         End
         Begin EditLib.fpText fpText1 
            Height          =   3720
            Left            =   435
            TabIndex        =   21
            Top             =   630
            Width           =   8430
            _Version        =   196608
            _ExtentX        =   14870
            _ExtentY        =   6562
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
            ThreeDInsideStyle=   1
            ThreeDInsideHighlightColor=   -2147483633
            ThreeDInsideShadowColor=   -2147483642
            ThreeDInsideWidth=   1
            ThreeDOutsideStyle=   1
            ThreeDOutsideHighlightColor=   -2147483628
            ThreeDOutsideShadowColor=   -2147483632
            ThreeDOutsideWidth=   1
            ThreeDFrameWidth=   0
            BorderStyle     =   0
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
            ThreeDTextHighlightColor=   -2147483633
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
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "fpText1"
            CharValidationText=   ""
            MaxLength       =   255
            MultiLine       =   -1  'True
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
            ThreeDOnFocusInvert=   0   'False
            ThreeDFrameColor=   -2147483633
            Appearance      =   2
            BorderDropShadow=   0
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            ButtonColor     =   -2147483633
            AutoMenu        =   0   'False
            ButtonAlign     =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   -67980
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   690
         Width           =   2085
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   -73425
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   645
         Width           =   2085
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6750
         Left            =   -75000
         TabIndex        =   1
         Top             =   2040
         Width           =   18285
         _Version        =   393216
         _ExtentX        =   32253
         _ExtentY        =   11906
         _StockProps     =   64
         AutoClipboard   =   0   'False
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
         MaxCols         =   16
         ScrollBars      =   2
         SpreadDesigner  =   "M_Lista_Pedido.frx":1B7B
      End
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   6660
         Left            =   120
         TabIndex        =   2
         Top             =   1785
         Width           =   18255
         _Version        =   393216
         _ExtentX        =   32200
         _ExtentY        =   11748
         _StockProps     =   64
         AutoClipboard   =   0   'False
         ButtonDrawMode  =   1
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   25
         MaxRows         =   1
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_Lista_Pedido.frx":390B
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   -73425
         TabIndex        =   7
         Top             =   1155
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   3
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   1
         Left            =   -70830
         TabIndex        =   8
         Top             =   1155
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
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
         ButtonStyle     =   3
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   -67965
         TabIndex        =   11
         Top             =   1200
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   -67950
         TabIndex        =   15
         Top             =   1605
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         AutoCase        =   1
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   2
         Left            =   2475
         TabIndex        =   31
         Top             =   1230
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   0   'False
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
         ButtonStyle     =   3
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   3
         Left            =   6795
         TabIndex        =   32
         Top             =   1230
         Width           =   1305
         _Version        =   196608
         _ExtentX        =   2302
         _ExtentY        =   556
         Enabled         =   0   'False
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
         ButtonStyle     =   3
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
         Text            =   "13/07/2004"
         DateCalcMethod  =   3
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   210
         Left            =   -74820
         TabIndex        =   35
         Top             =   8790
         Width           =   16620
         _ExtentX        =   29316
         _ExtentY        =   370
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblestado 
         Height          =   495
         Left            =   10800
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label Lbl_familuia 
         Caption         =   "Familia"
         Height          =   225
         Left            =   6165
         TabIndex        =   39
         Top             =   8745
         Width           =   480
      End
      Begin VB.Label Lbl_provedor 
         Caption         =   "Proveedor"
         Height          =   240
         Left            =   330
         TabIndex        =   37
         Top             =   8745
         Width           =   1155
      End
      Begin VB.Label lbl_proceso 
         Height          =   210
         Left            =   -67980
         TabIndex        =   36
         Top             =   9060
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emisión Desde"
         Height          =   480
         Index           =   3
         Left            =   615
         TabIndex        =   34
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emisión Hasta"
         Height          =   480
         Index           =   2
         Left            =   4800
         TabIndex        =   33
         Top             =   1260
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo de Pedido"
         Height          =   360
         Index           =   0
         Left            =   4770
         TabIndex        =   29
         Top             =   780
         Width           =   1470
      End
      Begin VB.Label Label4 
         Caption         =   "Numero Pedido"
         Height          =   315
         Left            =   660
         TabIndex        =   27
         Top             =   765
         Width           =   1515
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   3
         Left            =   -67935
         TabIndex        =   19
         Top             =   765
         Width           =   2070
      End
      Begin VB.Label Label7 
         Caption         =   "Estado Pedido"
         Height          =   255
         Index           =   1
         Left            =   -69225
         TabIndex        =   18
         Top             =   765
         Width           =   1350
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Org.Compra"
         Height          =   195
         Index           =   1
         Left            =   -69195
         TabIndex        =   16
         Top             =   1650
         Width           =   840
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   -66645
         Picture         =   "M_Lista_Pedido.frx":452D
         Top             =   1080
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
         Height          =   195
         Index           =   0
         Left            =   -69210
         TabIndex        =   13
         Top             =   1245
         Width           =   1140
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -66225
         TabIndex        =   12
         Top             =   1155
         Width           =   5655
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emisión Hasta"
         Height          =   480
         Index           =   0
         Left            =   -71955
         TabIndex        =   10
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Emisión Desde"
         Height          =   480
         Index           =   1
         Left            =   -74655
         TabIndex        =   9
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo Pedido"
         Height          =   255
         Index           =   0
         Left            =   -74685
         TabIndex        =   6
         Top             =   720
         Width           =   1350
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   -73380
         TabIndex        =   5
         Top             =   720
         Width           =   2070
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   -66195
         TabIndex        =   14
         Top             =   1185
         Width           =   5655
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   18525
      _ExtentX        =   32676
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_Lista_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Sql                  As String
Dim BtnX                 As Variant
Dim RS                   As New ADODB.Recordset
Dim Ingrediente_ACAMBIAR As Integer
Dim cantidad_despacho    As Double
Dim cantidad_despacho_anterior As Double
Dim perfil_redondeo      As Double
Dim rowanterior          As Integer
Dim tipopedido           As String
Dim REGISTROS_ELIMINADOS As String
Dim PermisoPAP           As Boolean
Dim PermisoCD            As Boolean
Dim Permiso              As Boolean
Dim tippedido            As Integer
Dim nestado              As Integer
Dim pedido               As Integer
Dim fdesde               As String
Dim fhasta               As String
Dim fechaanterior        As String
Dim fechapedido          As String
Dim fecha_limiteanterior As String

Private Sub Command1_Click()
 
 Frame3.Visible = False

End Sub

Private Sub Command2_Click()

On Error GoTo Man_Error
    
    Frame1.Visible = False
    Toolbar1.Buttons(1).Visible = True
    Toolbar1.Buttons(2).Visible = False
    Toolbar1.Buttons(7).Enabled = True
    Toolbar1.Buttons(4).Enabled = True
  '  Toolbar1.Buttons(8).Enabled = False
    Call limpia_grilla

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Command3_Click()

Frame2.Visible = False

End Sub

Private Sub Command4_Click()

On Error GoTo Man_Error

Dim fechalimite As String

Dim fechadesde As String
Dim fechahasta As String

Dim tipo As Integer
tipo = 0

fechalimite = Format(DTPicker1, "YYYYMMDD HH:MM:SS")

fechaanterior = Format(lbl_fecha.Caption, "YYYYMMDD HH:MM:SS")
fechapedido = Format(fechaanterior, "YYYYMMDD HH:MM:SS")


fechadesde = Format(fdesde, "YYYYMMDD HH:MM:SS")
fechahasta = Format(fhasta, "YYYYMMDD HH:MM:SS")

If Format(fechalimite, ("YYYYMMDD")) >= Format(fechapedido, "YYYYMMDD") Then
  If Format(fechalimite, ("YYYYMMDD")) < Format(fechadesde, "YYYYMMDD") Then
     
 Else
    MsgBox "La fecha de confirmación no puede ser mayor o igual a la fecha desde...", vbExclamation
    Exit Sub
End If
Else
    MsgBox "La fecha de confirmación no puede ser menor a la fecha del Pedido...", vbExclamation
    Exit Sub
End If

If Option2.Value = True Then tipo = 1
 
      Sql = " sgpadm_iu_ActualizacionFechaLimite "
      Sql = Sql & pedido & ","
      Sql = Sql & " '" & fecha_limiteanterior & "',"
        Sql = Sql & " '" & fechalimite & "',"
      Sql = Sql & " '" & fechadesde & "',"
      Sql = Sql & " '" & fechahasta & "',"
      Sql = Sql & tipo
      Set RS = vg_db.Execute(Sql)
      MsgBox "Se Actualizo la Fecha Correctamente", vbExclamation
      Frame3.Visible = False
      Call busca_encabezado

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Load()
   
   On Error GoTo Man_Error
  
  SSTab1.Tab = 0
  ProgressBar1.Visible = False
  Dim X As Variant
  Frame1.Visible = False
  fg_centra Me
  Toolbar1.ImageList = Partida.IL1
  
  Set BtnX = Toolbar1.Buttons.Add(, "A_Incluir  ", , tbrDefault, "A_Incluir  "): BtnX.Visible = True: BtnX.ToolTipText = "Nuevo Pedido"
  Set BtnX = Toolbar1.Buttons.Add(, "I_Incluir  ", , tbrDefault, "I_Incluir  "): BtnX.Visible = False: BtnX.ToolTipText = ""
  Set BtnX = Toolbar1.Buttons.Add(, "A_CopiarD", , tbrDefault, "A_CopiarD"): BtnX.Visible = True: BtnX.ToolTipText = "Enviando Minuta a Sitio" '"Enviar a PEL"
  Set BtnX = Toolbar1.Buttons.Add(, "A_Filtro", , tbrDefault, "A_Filtro"): BtnX.Visible = True: BtnX.ToolTipText = "Filtrar"
  Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = "Confirmar "
  Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = False: BtnX.ToolTipText = ""
  Set BtnX = Toolbar1.Buttons.Add(, "A_Buscar", , tbrDefault, "A_Buscar"): BtnX.Visible = False: BtnX.ToolTipText = "Detalle Pedido"
  Set BtnX = Toolbar1.Buttons.Add(, "A_VerErrores", , tbrDefault, "A_VerErrores"): BtnX.Visible = False: BtnX.ToolTipText = "Errores PEL"
  Set BtnX = Toolbar1.Buttons.Add(, "A_Reporcesar", , tbrDefault, "A_Reporcesar"): BtnX.Visible = True: BtnX.ToolTipText = "Reprocesar"
  Set BtnX = Toolbar1.Buttons.Add(, "A_Deshacer", , tbrDefault, "A_Deshacer"): BtnX.Visible = True: BtnX.ToolTipText = "Desmarcar"
  Set BtnX = Toolbar1.Buttons.Add(, "excel", , tbrDefault, "excel"): BtnX.Visible = True: BtnX.ToolTipText = "Exportar a Excel "
  Set BtnX = Toolbar1.Buttons.Add(, "A_Agregar", , tbrDefault, "A_Agregar"): BtnX.Visible = True: BtnX.ToolTipText = "Agregar Ingrediente"
  Set BtnX = Toolbar1.Buttons.Add(, "A_Calendario", , tbrDefault, "A_Calendario"): BtnX.Visible = False: BtnX.ToolTipText = "Calendario"
  Set BtnX = Toolbar1.Buttons.Add(, "A_VerConvenio", , tbrDefault, "A_VerConvenio"): BtnX.Visible = True: BtnX.ToolTipText = "Buscar Convenios"
  Set BtnX = Toolbar1.Buttons.Add(, "A_Alzas", , tbrDefault, "A_Alzas"): BtnX.Visible = True: BtnX.ToolTipText = "Saldo Mayor Formato de Compra"
  Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
  Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  
  Sql = " sgpadm_Sel_leemaximofecha "
  Set RS = vg_db.Execute(Sql)
  If Not RS.EOF Then
    
    fpDateTime1(0).text = Format(RS(0), "dd/mm/yyyy")
    fpDateTime1(1).text = Format(RS(1), "dd/mm/yyyy")
  
  End If
  
  '-------> Llenar combo Tipo Pedido
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  
  PermisoPAP = False
  PermisoCD = False
  Sql = ""
  Sql = " sgpadm_Sel_TipoPedido "
  Set RS = vg_db.Execute(Sql)
  Combo1(0).Clear
  Combo1(0).AddItem "Todos" & Space(150) & "(0)"
  Do While Not RS.EOF
     
     '-------> Permiso pedido PAP
     Permiso = True
     
     If (RS(0) = 1 Or RS(0) = 3) Then
        
        Permiso = False
        Me.HelpContextID = IIf(RS(0) = 1, 1191000, 1192000)
        
        If Mid(ValidaPerfil(Me), 1, 1) = "1" Then
           
           Permiso = True
           
           If Me.HelpContextID = 1191000 Then
              
              PermisoPAP = True
           
           End If
           
           If Me.HelpContextID = 1192000 Then
              
              PermisoCD = True
           
           End If
        
        End If
     
     End If
     
     If Permiso Then
        
        Combo1(0).AddItem Trim(RS(1)) & Space(150) & "(" & Trim(RS(0)) & ")"
     
     End If
     
     RS.MoveNext
  
  Loop
  RS.Close: Set RS = Nothing
  Combo1(0).ListIndex = 0
  
  '-------> Llenar combo Tipo Pedido
  
  Combo1(1).Clear
  Combo1(1).AddItem "Todos" & Space(150) & "(0"
  If RS.State = 1 Then RS.Close
  RS.CursorLocation = adUseClient
  vg_db.CursorLocation = adUseClient
  
  Sql = ""
  Sql = " sgpadm_Sel_EstadoPedidos "
  Set RS = vg_db.Execute(Sql)
  Do While Not RS.EOF
     
     Combo1(1).AddItem Trim(RS(1)) & Space(150) & "(" & Trim(RS(0)) & ")"
     RS.MoveNext
  
  Loop
  Combo1(1).ListIndex = 0
  Combo1(1).ListIndex = 0
  
 ' Toolbar1.Buttons(3).Enabled = False
  Toolbar1.Buttons(7).Enabled = False
 ' Toolbar1.Buttons(8).Enabled = False
  Toolbar1.Buttons(10).Enabled = False
  Toolbar1.Buttons(11).Enabled = True '20140513False
  Toolbar1.Buttons(12).Enabled = False

  ' Control displays text tips aligned to pointer with focus
  vaSpread2.TextTip = 2
  vaSpread2.TextTipDelay = 250
  X = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
  
  Frame3.Visible = False
  
  Me.HelpContextID = vg_OpcM
  Call busca_encabezado

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)
  
End Sub

Private Sub busca_encabezado()

On Error GoTo Man_Error
 
  If fpText(0) <> "" Then
      
      If Not IsNumeric(fpText(0)) Then
         
         MsgBox "El Centro de Costo debe ser Numerico", vbCritical, MsgTitulo
         Exit Sub
      
      End If
  
  End If
  
  If fpDateTime1(1).DateValue < fpDateTime1(0).DateValue Then
     
     MsgBox "La fecha de emision hasta no puede se menor que la desde", vbCritical, MsgTitulo
     Exit Sub
  
  End If
   
  'registrar Log sistema
  Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Filtrar"), Me.HelpContextID, "", "", "")
  
  Call lee_encabezado
  
Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub lee_encabezado()

Dim tipopedido As Integer
Dim estado     As Double
Dim fechai     As String
Dim fechaF     As String
Dim estadoviz  As Boolean

 On Error GoTo Man_Error
 
    tipopedido = IIf(Combo1(0).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 0, 1, "")))
    estado = IIf(Combo1(1).ListIndex = -1, "Null", Val(fg_codigocbo(Combo1, 1, 1, "")))
    estadoviz = True
    

    Toolbar1.Buttons(7).Enabled = False
    
    fechai = Format(fpDateTime1(0), "YYYYMMDD") + " " + "00:00"
    fechaF = Format(fpDateTime1(1), "YYYYMMDD") + " " + "00:00"
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = " sgpadm_sel_seleccionaPedidos_V01 "
    Sql = Sql & tipopedido & ","
    Sql = Sql & estado & ","
    Sql = Sql & " '" & fpText(1) & "',"
    Sql = Sql & " '" & fpText(0) & "',"
    Sql = Sql & " '" & fechai & "',"
    Sql = Sql & " '" & fechaF & "'"
    
    Set RS = vg_db.Execute(Sql)
    
   '-------> Inicio LLenar grilla
   
    vaSpread1.MaxRows = 0
 
    Do While Not RS.EOF
        estadoviz = True
        '------> validar si tipo pedido PAP y CD este Activado
        '------> validar permiso PAP
        If RS(9) = "Pedido PAP" And Not PermisoPAP Then
           
           estadoviz = False
        
        End If
        
        '------> validar permiso CD
        If RS(9) = "Pedido CD" And Not PermisoCD Then
           
           estadoviz = False
        
        End If
        
        If estadoviz Then
           
           vaSpread1.MaxRows = vaSpread1.MaxRows + 1
           vaSpread1.Row = vaSpread1.MaxRows
        
           vaSpread1.Col = 2 ' IdPedido
           vaSpread1.text = Val(RS(0))
           vaSpread1.TypeHAlign = TypeHAlignCenter
        
           vaSpread1.Col = 3 ' Celo
           vaSpread1.text = RS(1)
           vaSpread1.TypeHAlign = TypeHAlignCenter
        
           vaSpread1.Col = 4 ' Cencos
           vaSpread1.text = IIf(IsNull(RS(2)), " ", RS(2))
           vaSpread1.TypeHAlign = TypeHAlignCenter
        
           vaSpread1.Col = 5 ' Nombre Cencos
           vaSpread1.text = IIf(IsNull(RS(3)), " ", RS(3))
       
           vaSpread1.Col = 6 ' Observacoion
           vaSpread1.text = IIf(IsNull(RS(4)), " ", RS(3))
        
           vaSpread1.Col = 7 ' Fecha desde
           vaSpread1.text = Format(RS(5), "DD/MM/YYYY")
           vaSpread1.TypeHAlign = TypeHAlignCenter
        
           vaSpread1.Col = 8 ' Fecha Hasta
           vaSpread1.text = Format(RS(6), "DD/MM/YYYY")
           vaSpread1.TypeHAlign = TypeHAlignCenter
        
           vaSpread1.Col = 9 ' Estado
           vaSpread1.text = RS(8)
        
           vaSpread1.Col = 10 ' Tipo Pedido
           vaSpread1.text = RS(11)
        
           vaSpread1.Col = 11 ' Id Pedido Pel
           vaSpread1.text = IIf(IsNull(RS(7)), " ", RS(7))
       
           vaSpread1.Col = 12 ' Estado del Pedido
           vaSpread1.text = IIf(IsNull(RS(9)), " ", RS(9))
        
           vaSpread1.Col = 13 ' Tipo del Pedido
           vaSpread1.text = IIf(IsNull(RS(10)), " ", RS(10))
       
           vaSpread1.Col = 14 ' Fecha Linite
           vaSpread1.text = IIf(IsNull(RS(13)), " ", Format(RS(13), "DD/MM/YYYY hh:mm"))
              
           vaSpread1.Col = 15 ' Fecha Pedido
           vaSpread1.text = IIf(IsNull(RS(14)), " ", Format(RS(14), "DD/MM/YYYY hh:mm"))
              
           vaSpread1.Col = 16 ' Fecha Pedido
           vaSpread1.text = Format(IIf(IsNull(RS(15)), " ", RS(15)), fg_Pict(6, 2))
              
        End If
        
        RS.MoveNext
    Loop
        vaSpread1.SetActiveCell 1, rowanterior

Exit Sub
Man_Error:
  fg_descarga
  MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
  ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Form_Unload(Cancel As Integer)

'registrar Log sistema salir
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Me.HelpContextID, "", "", "")

End Sub

Private Sub fpText_Change(Index As Integer)
 
 On Error GoTo Man_Error
 
    Dim RS1 As New ADODB.Recordset
    Dim Sql As String
    If fpText(0).text = "" Then fpayuda(0).Caption = "": Exit Sub
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = Trim(LimpiaDato(fpText(0).text))
    Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 29, '" & Sql & "', ''")
    If RS1.EOF Then RS1.Close: Set RS1 = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS1!Cli_nombre)
    RS1.Close
    Set RS1 = Nothing
    
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Image1_Click(Index As Integer)
 
 On Error GoTo Man_Error
    
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Casino", "Clientesimap"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText(0).text = vg_codigo: fpayuda(0).Caption = vg_nombre
    If Me.Visible Then fpDateTime1(0).SetFocus

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Reprocesar()
 
 Dim RS            As New ADODB.Recordset
 Dim seleccion     As Integer
 Dim pedido        As Long
 Dim i             As Long
 Dim Prod_CD       As Integer
 Dim Prod_PAP      As Integer
 Dim xmlfamilia    As String
 Dim centrocosto   As String
 
 Dim FechaInicial  As String
 Dim FechaFinal    As String
 Dim tipopedido    As Integer
 Dim familia       As String
 Dim OrgCompra     As String
 Dim cecos         As String
 Dim detallePedido As String
 Dim estado        As Integer
 
 Dim Conta         As Integer
 
Dim EstadoGenerarArrastreSaldo As String

 On Error GoTo Man_Error
 
 Conta = 0
 For i = 1 To vaSpread1.MaxRows
          
      vaSpread1.Row = i
      
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
      If seleccion = 1 Then
        Conta = Conta + 1
      End If
      
 Next i
    
 If Conta = 0 Then
     
    MsgBox "Debe haber un pedido seleccionado por lo menos para Reprocesar a PEL ", 16
    Call busca_encabezado
    fg_descarga
    Exit Sub
    
 End If
  
 If tipopedido = 2 Then
    
    If MsgBox("Esta Seguro Reprocesar el Pedido...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
       
       Exit Sub
    
    End If
 
 Else
    
    If MsgBox("Al Reprocesar el Pedidos (CD - PAP) y de Existir Pedidos Posteriores se Reprocesara" & VgLinea & VgLinea & "                                        Esta Seguro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then
       
       Exit Sub
    
    End If
 
 End If
 
    DoEvents
    Screen.MousePointer = 11
    DoEvents
   
    'registrar Log sistema
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Procesar"), Me.HelpContextID, "", "", "")
   
   Conta = 0
   For i = 1 To M_Lista_Pedido.vaSpread1.MaxRows
        
        M_Lista_Pedido.vaSpread1.Row = i
        M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
        seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        If seleccion = 1 Then
           
           Conta = Conta + 1
        
        End If
         
    Next i
    ProgressBar1.Scrolling = ccScrollingSmooth
    ProgressBar1.Max = Conta
    
    ProgressBar1.Visible = True
    ProgressBar1.Value = 0
    
   For i = 1 To M_Lista_Pedido.vaSpread1.MaxRows
        
'        DoEvents
        
        M_Lista_Pedido.vaSpread1.Row = i
        M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
        seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        If seleccion = 1 Then
          
           M_Lista_Pedido.vaSpread1.Col = 2 'Pedido
           pedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
             
           M_Lista_Pedido.vaSpread1.Col = 13 'Cecos
           tippedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
          
           M_Lista_Pedido.vaSpread1.Col = 4 'Cecos
           cecos = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
                   
          
           M_Lista_Pedido.vaSpread1.Col = 10 'Tipo de Pedido
           detallePedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
          
           M_Lista_Pedido.vaSpread1.Col = 12 'Estado de pedido
           estado = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
   
          
           If estado = 2 Or estado = 6 Or estado = 5 Then
              
              MsgBox "El Pedido Numero : " + CStr(pedido) + ", no puede ser Reprocesado, Verifique su Estado ", 16
              Call busca_encabezado
              fg_descarga
              Exit Sub
           
           End If
          
'          sql = " sgpadm_Sel_mayorpedidoparaReprocesar " & "'" & cecos & "'," & tippedido
'          Set RS = vg_db.Execute(sql)
'
'          If Not RS.EOF Then
'             If pedido <> RS(0) Then
'              Screen.MousePointer = 1
'              DoEvents
'              MsgBox " Solo se puede Reprocesar el Ultimo Pedido realizado para este Ceco " + CStr(cecos), vbExclamation
'              Exit Sub
'             End If
'          End If
          
         
          ' Recupera los Valores del Pedido para volver a procesar
          
           Sql = " sgpadm_Sel_BuscaParametrosPedidos " & pedido
           Set RS = vg_db.Execute(Sql)
          
           Do While Not RS.EOF
              
              If RS("TipoPar") = "@cecosXmlData" Then centrocosto = RS("Valor")
              If RS("TipoPar") = "@Ceco" Then centrocosto = RS("Valor")
              If RS("TipoPar") = "@FechaInicial" Then FechaInicial = RS("Valor")
              If RS("TipoPar") = "@FechaFinal" Then FechaFinal = RS("Valor")
              If RS("TipoPar") = "@TipoPedido" Then tipopedido = RS("Valor")
              If RS("TipoPar") = "@FamiliasXmlData" Then xmlfamilia = RS("Valor")
              If RS("TipoPar") = "@Incluir_Prod_CD" Then Prod_CD = RS("Valor")
              If RS("TipoPar") = "@Incluir_Prod_PAP" Then Prod_PAP = RS("Valor")
              If RS("TipoPar") = "@IDPedido" Then pedido = RS("Valor")
              If RS("TipoPar") = "@OrgCompra" Then OrgCompra = RS("Valor")
              EstadoGenerarArrastreSaldo = RS("EstadoGenerarArrastreSaldo")
            
              RS.MoveNext
          
           Loop
           RS.Close: Set RS = Nothing
           
           If tipopedido = 2 Then
              
              Sql = " sgpadm_iu_Elimina_Pedidodetalle "
              Sql = Sql & pedido
              Set RS = vg_db.Execute(Sql)
              Set RS = Nothing
             
              Dim s As String
              Dim Arreglo() As String
              Dim r As Integer
              
              s = centrocosto
              Arreglo = Split(s, ",")
          
              For r = 1 To UBound(Arreglo)
                  
                  centrocosto = Arreglo(r)
'                  EstadoGenerarArrastreSaldo = 1
                  Sql = " sgpadm_Sel_GeneracionPedido_FDespacho_V09 "
                  Sql = Sql & " '" & centrocosto & "',"
                  Sql = Sql & " '" & FechaInicial & "',"
                  Sql = Sql & " '" & FechaFinal & "',"
                  Sql = Sql & tipopedido & ","
                  Sql = Sql & " '" & xmlfamilia & "',"
                  Sql = Sql & Prod_CD & ","
                  Sql = Sql & Prod_PAP & ","
                  Sql = Sql & pedido & ","
                  Sql = Sql & " '" & OrgCompra & "',"
                  Sql = Sql & " '" & "" & "',"
                  Sql = Sql & " '" & EstadoGenerarArrastreSaldo & "'"
                  Set RS = vg_db.Execute(Sql)
                  RS.Close: Set RS = Nothing
              
              Next r
 
           Else
             
              Sql = " sgpadm_Sel_GeneracionPedido_FDespacho_V09 "
              Sql = Sql & " '" & centrocosto & "',"
              Sql = Sql & " '" & FechaInicial & "',"
              Sql = Sql & " '" & FechaFinal & "',"
              Sql = Sql & tipopedido & ","
              Sql = Sql & " '" & xmlfamilia & "',"
              Sql = Sql & Prod_CD & ","
              Sql = Sql & Prod_PAP & ","
              Sql = Sql & pedido & ","
              Sql = Sql & " '" & OrgCompra & "',"
              Sql = Sql & " '" & "" & "',"
              Sql = Sql & " '" & 1 & "'"
              Set RS = vg_db.Execute(Sql)
              RS.Close: Set RS = Nothing
              
              '--> Reprocesar Pedidos posteriores
              Sql = ""
              Sql = " sgpadm_Pro_ReprocesarPedidosPosteriores "
              Sql = Sql & " '" & centrocosto & "',"
              Sql = Sql & pedido & " "
              Set RS = vg_db.Execute(Sql)

           End If
             
           If tipopedido = 2 Then
               
              Sql = " sgpadm_iu_VadidaEstadoPedido "
              Sql = Sql & pedido
              Set RS = vg_db.Execute(Sql)
              Set RS = Nothing
            End If
            
            ProgressBar1.Value = ProgressBar1.Value + 1
            lbl_proceso.Caption = CLng((ProgressBar1.Value * 100) / ProgressBar1.Max) & " %"
            DoEvents
             
        End If
    Next i
    
    Screen.MousePointer = 0
    DoEvents
    ProgressBar1.Visible = False
    lbl_proceso.Caption = ""
    MsgBox "Termino el Reproceso Correctamente", vbExclamation
    
    Call limpia_grilla
    Call busca_encabezado
      
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Envio_A_Pel()

 Dim seleccion As Integer
 Dim pedido As Long
 Dim i As Integer
 Dim estado As Integer
 Dim Conta As Integer
 
On Error GoTo Man_Error
  
  Conta = 0
  For i = 1 To vaSpread1.MaxRows
          
      vaSpread1.Row = i
      
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
      If seleccion = 1 Then
        
        Conta = Conta + 1
      
      End If
      
    Next i
    
    If Conta = 0 Then
        
        MsgBox "Debe haber un pedido seleccionado por lo menos para Enviando Minuta a Sitio ", 16
        Call busca_encabezado
        fg_descarga
        Exit Sub
    
    End If
  
 If MsgBox("Este Pedido sera Enviando Minuta a Sitio esta Seguro ...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
  
'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Envio_Minuta_Pedido"), Me.HelpContextID, "", "", "")
  
  For i = 1 To M_Lista_Pedido.vaSpread1.MaxRows
        
        M_Lista_Pedido.vaSpread1.Row = i
        M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
        seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        If seleccion = 1 Then
            M_Lista_Pedido.vaSpread1.Col = 2 'Pedido
            pedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
             
           M_Lista_Pedido.vaSpread1.Col = 12 'Estado de pedido
           estado = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
           
'Inicio 20150219 comenta
'          If estado <> 1 Then '5
'             MsgBox "El Pedido Numero : " + CStr(pedido) + ", no puede Enviando Minuta a Sitio, Verifique su Estado ", 16
'             Call busca_encabezado
'             fg_descarga
'             Exit Sub
'           End If
             
             
'            Set RS = vg_db.Execute("sgpadm_sel_seleccionaDetallePedidoActivo " & pedido)
'            If Not RS.EOF Then
'               MsgBox "El Pedido Numero : " + CStr(pedido) + ", no puede ser Enviando Minuta a Sitio, debido a que tiene items sin rutas o convenios ", 16
'               Exit Sub
'               RS.Close: Set RS = Nothing
'
'            End If
'Fin 20150219 comenta

            Sql = ""
            Sql = " sgpadm_iu_actualizaestado "
            Sql = Sql & pedido & "," & 2
            Set RS = vg_db.Execute(Sql)
        End If
    Next i

    Call busca_encabezado

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Revisar_el_Mensaje()
  
  Dim seleccion As Integer
  Dim pedido As Integer
  Dim Conta As Integer

On Error GoTo Man_Error
  
   
  'Toolbar1.Buttons(1).Visible = False
  'Toolbar1.Buttons(2).Visible = True
  'Toolbar1.Buttons(4).Enabled = False
  Toolbar1.Buttons(10).Enabled = False
  Toolbar1.Buttons(11).Enabled = True 'False
  Toolbar1.Buttons(12).Enabled = False
  
 Dim i As Integer
  
  Conta = 0
  For i = 1 To vaSpread1.MaxRows
          
      vaSpread1.Row = i
      
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
      If seleccion = 1 Then
        Conta = Conta + 1
      End If
      
    Next i
    
    If Conta = 0 Then
        
        MsgBox "Debe haber un pedido seleccionado por lo menos para Ver los Mensaje ", 16
        Call busca_encabezado
        fg_descarga
        Exit Sub
    
    End If
    
 For i = 1 To M_Lista_Pedido.vaSpread1.MaxRows
        
        M_Lista_Pedido.vaSpread1.Row = i
        M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
        seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        If seleccion = 1 Then
            
            M_Lista_Pedido.vaSpread1.Col = 2 'Pedido
            pedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
                 
            '------------Detalle del Pedido
                         
            Sql = " sgpadm_Sel_Error_Pedido "
            Sql = Sql & pedido
            Set RS = vg_db.Execute(Sql)
            Frame1.Visible = True
            fpText1 = IIf(IsNull(RS(0)), "", Trim(RS(0)))
                        
            RS.Close: Set RS = Nothing
        
        End If

Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Detalle_del_Pedido()

On Error GoTo Man_Error

  Dim RS              As New ADODB.Recordset
  Dim seleccion       As Integer
  Dim pedido          As Long
  Dim TipodePedido    As String
  Dim GlosaDetPedido  As String
  Dim cecos           As String
  Dim estado          As String
  
  Dim NomArchivoExcel As String
  Dim Extension       As String
  
  Frame2.Visible = False
  
  Toolbar1.Buttons(1).Visible = False
  Toolbar1.Buttons(2).Visible = True
  Toolbar1.Buttons(4).Enabled = False
  Toolbar1.Buttons(7).Enabled = False
  Toolbar1.Buttons(10).Enabled = False
  Toolbar1.Buttons(11).Enabled = True
  
  SSTab1.Tab = 1
  
 ' Rescata el Pedido del cual se va a ver el Detalle
 
 Dim i As Integer
 
For i = 1 To M_Lista_Pedido.vaSpread1.MaxRows

    M_Lista_Pedido.vaSpread1.Row = i
    M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
    If seleccion = 1 Then
        
        M_Lista_Pedido.vaSpread1.Col = 2 'Pedido
        pedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        M_Lista_Pedido.vaSpread1.Col = 9 'Estado del Pedio
        estado = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        M_Lista_Pedido.vaSpread1.Col = 10 'Estado del Pedio
        TipodePedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        M_Lista_Pedido.vaSpread1.Col = 12 'Estado del Pedio
        nestado = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        M_Lista_Pedido.vaSpread1.Col = 13 'Tipo de Pedido
        tippedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        M_Lista_Pedido.vaSpread1.Col = 4 'Cencos
        vg_cencos = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
        
        fpText2 = pedido
        fpText3 = TipodePedido
        lblestado.Caption = estado
        
        '------------Detalle del Pedido
        
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        Sql = " sgpadm_sel_seleccionaDetallePedido_NRegistro "
        Sql = Sql & pedido
        Set RS = vg_db.Execute(Sql)
        
        If RS!nReg > 45000 Then
            
           RS.Close
           Set RS = Nothing
          
           If MsgBox("Excede numero registro para visualizar sistema. " & VgLinea & VgLinea & "          Desea Exportar Excel ?? ", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then

              '-------> Guardar nombre archivo excel
              NomArchivoExcel = ""
              CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
              CD.Filter = "Todos los archivos *.xls,*.xlsx"
              CD.ShowSave
              If CD.FileName = "" Then
                
                 MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
                 Exit Sub
             
              Else
                
                 Extension = ""
                 Extension = Right(CD.FileName, Len(CD.FileName) - (InStr(CD.FileName, ".")))
                
                 If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
                   
                    MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
                    Exit Sub
                
                 End If
                 NomArchivoExcel = CD.FileName  'Dir(CD.Filename)
             
              End If
             
              Dim xlApp As Object
              Dim xlWb  As Object
              Dim xlWs  As Object
    
              If Dir(NomArchivoExcel) <> "" Then Kill NomArchivoExcel
    
              '-------> Create an instance of Excel and add a workbook
              Set xlApp = CreateObject("Excel.Application")
              Set xlWb = xlApp.Workbooks.Add
              Set xlWs = xlWb.Worksheets("Hoja1")
  
              '-------> Display Excel and give user control of Excel's lifetime
              xlApp.UserControl = True
    
              If RS.State = 1 Then RS.Close
              RS.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              Sql = " sgpadm_sel_seleccionaDetallePedido_FormatoExcel "
              Sql = Sql & pedido
              Set RS = vg_db.Execute(Sql)
              
              If RS.EOF Then
              
                 RS.Close
                 Set RS = Nothing
              
                 MsgBox "No exite información. proceso cancelado", vbInformation
              
                 Exit Sub
                 
              End If
              
              '-------> Check version of Excel
              Call encabezado(RS, xlWs)
          
              xlWs.Cells(2, 1).CopyFromRecordset RS
              '-------> Auto-fit the column widths and row heights
              xlApp.Selection.CurrentRegion.Columns.AutoFit
              xlApp.Selection.CurrentRegion.Rows.AutoFit
    
              xlWb.Close True, NomArchivoExcel

              Dim XL As New excel.Application 'Crea el objeto excel
              XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
              XL.Visible = True
              XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
              '-------> Close ADO objects
              RS.Close
              Set RS = Nothing
    
              '-- Cerrar Excel
              xlApp.Quit
              '-------> Release Excel references
              Set xlWs = Nothing
              Set xlWb = Nothing
              Set xlApp = Nothing
            
              MsgBox "Exportación realizada con exito", vbInformation
            
              Exit Sub
            
           End If
            
        
        End If
        '-------> Inicio LLenar grilla
        vg_Archxls = fg_ArchivoTxt
        Open vg_Archxls For Output As #1
        vaSpread2.MaxRows = 0
        GlosaDetPedido = ""
        
        RS.Close
        Set RS = Nothing
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Sql = " sgpadm_sel_seleccionaDetallePedido "
        Sql = Sql & pedido
        Set RS = vg_db.Execute(Sql)
        
        Do While Not RS.EOF
        
            vaSpread2.MaxRows = vaSpread2.MaxRows + 1
            vaSpread2.Row = vaSpread2.MaxRows
            
            '------> generar encabezado text
            If vaSpread2.MaxRows = 1 Then
               
               Print #1, "Ingrediente" & "|" & "Descripcion" & "|" & "Proveedor" & "|" & "Familia" & "|" & "C.Costos" & "|" & "Producto" & "|" & "Descripcion" & "|" & "Unidad" & "|" & "Fecha Despacho" & "|" & "Cant.Despacho" & "|" & "Saldo Consumido" & "|" & "Saldo Ingrediente" & "|" & "Cant.Planif." & "|" & "Perfil de Redobdeo" & "|" & "Und. Ingrediente" & "|" & "Factor Conversion" & "|" & "Cantidad Original" & "|" & "Precio Convenio" & "|" & "Estado Pedido"
            
            End If
            '-------> Print detalle pedido
            Print #1, Val(RS(1)) & "|" & RS(2) & "|" & IIf(RS(3) <> "" And RS(3) <> "0", RS(3) + " - " + IIf(IsNull(RS(14)), " ", RS(14)), "0") & "|" & RS(4) & "|" & RS(5) & "|" & RS(6) & "|" & RS(7) & "|" & IIf(IsNull(RS(8)), " ", RS(8)) & "|" & " " & Format(RS(10), "DD/MM/YYYY") & "|" & RS(9) & "|" & IIf(IsNull(RS(15)), "0", RS(15)) & "|" & IIf(IsNull(RS(16)), "0", RS(16)) & "|" & IIf(IsNull(RS(17)), "0", RS(17)) & "|" & IIf(IsNull(RS(18)), "0", RS(18)) & "|" & IIf(IsNull(RS(19)), "0", RS(19)) & "|" & IIf(IsNull(RS(20)), "0", RS(20)) & "|" & IIf(IsNull(RS(22)), " ", RS(22)) & "|" & IIf(IsNull(RS(23)), " ", RS(23)) & "|" & IIf(RS(24) = "2", "Sin Convenios", IIf(RS(24) = "3", "Sin Ruta o Convenios", IIf(RS(24) = "4", "Excepción Formato Compras", "Correcto")))
            
            vaSpread2.Col = 2 ' Codigo Ingrediente
            vaSpread2.text = Val(RS(1))
            vaSpread2.TypeHAlign = TypeHAlignLeft
            
            vaSpread2.Col = 3 ' Nombre Ingrediente
            vaSpread2.TypeHAlign = TypeHAlignLeft
            vaSpread2.text = RS(2)
            
            vaSpread2.Col = 4 ' Proveedor
            
            If RS(3) <> "" And RS(3) <> "0" Then
              
              vaSpread2.text = RS(3) + " - " + IIf(IsNull(RS(14)), " ", RS(14))
            
            Else
              
              vaSpread2.text = "0"
            
            End If
            
            vaSpread2.Col = 5 ' Familia Producto
            vaSpread2.text = RS(4)
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 6 ' Cecos
            vaSpread2.text = RS(5)
            cecos = RS(5)
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 7 ' Producto Sap
            vaSpread2.text = RS(6)
            
            vaSpread2.Col = 8 ' Des. Producto Sap
            vaSpread2.text = RS(7)
            
            vaSpread2.Col = 9 ' Unidad
            vaSpread2.text = IIf(IsNull(RS(8)), " ", RS(8))
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 10 ' Fecha de Despacho
            vaSpread2.text = Format(RS(10), "DD/MM/YYYY")
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 11 ' Cantidad
            vaSpread2.text = Format(IIf(IsNull(RS(9)), 0, RS(9)), fg_Pict(6, 2))
            vaSpread2.TypeHAlign = TypeHAlignRight
           
            vaSpread2.Col = 12 ' Cantidad Consumida
            vaSpread2.text = IIf(IsNull(RS(15)), "0", RS(15))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 13 ' Cantidad Ingrediente
            vaSpread2.text = IIf(IsNull(RS(16)), "0", RS(16))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 14 ' Cantidad Ingresada
            vaSpread2.text = IIf(IsNull(RS(17)), "0", RS(17))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 15 ' Formato Convenio
            vaSpread2.text = IIf(IsNull(RS(18)), "0", RS(18))
            vaSpread2.TypeHAlign = TypeHAlignRight
        
        
            vaSpread2.Col = 16 ' Formato Convenio
            vaSpread2.text = IIf(IsNull(RS(19)), "0", RS(19))
           
            vaSpread2.Col = 17 ' Formato Convenio
            vaSpread2.text = IIf(IsNull(RS(20)), "0", RS(20))
            vaSpread2.TypeHAlign = TypeHAlignRight
        
            vaSpread2.Col = 18 ' Numero de Linea
            vaSpread2.text = Val(RS(12))
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 19 ' Tipo Pedido
            vaSpread2.text = Val(RS(13))
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 20 ' Marca
            vaSpread2.text = 1
            vaSpread2.TypeHAlign = TypeHAlignCenter
        
            vaSpread2.Col = 21 ' IdPedido
            vaSpread2.text = IIf(IsNull(RS(21)), "0", RS(21))
            vaSpread2.TypeHAlign = TypeHAlignCenter
             
            vaSpread2.Col = 22 ' Cantidad Original
            vaSpread2.text = IIf(IsNull(RS(22)), " ", RS(22))
            vaSpread2.TypeHAlign = TypeHAlignRight
         
        
            vaSpread2.Col = 23 ' Precio Convenios
            vaSpread2.text = IIf(IsNull(RS(23)), " ", RS(23))
            vaSpread2.TypeHAlign = TypeHAlignRight
                    
            If Trim(RS(24)) <> "" Then
               vaSpread2.Col = 2 ' Observación
               
               If Trim(RS(24)) = "2" Then
                  
                  vaSpread2.BackColor = Shape1(1).FillColor
               
               ElseIf Trim(RS(24)) = "3" Then
                  
                  vaSpread2.BackColor = Shape1(0).FillColor
               
               ElseIf Trim(RS(24)) = "4" Then
                  
                  vaSpread2.BackColor = Shape1(2).FillColor
               
               End If
            
            End If
            
            vaSpread2.Col = 25 ' Saldo carro anteior
            vaSpread2.text = Format(IIf(IsNull(RS(25)), " ", RS(25)), fg_Pict(6, 2))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
        RS.MoveNext
        Loop
       RS.Close: Set RS = Nothing
       Close #1
    
    End If

Next i
    
   'valida si se pueden hacer cambio al detalle del pedido
   
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Sql = " sgpadm_Sel_mayorpedidoparaReprocesar " & "'" & cecos & "'," & tippedido
   Set RS = vg_db.Execute(Sql)
  
   If Not RS.EOF Then

'       If (pedido <> RS(0)) Then
'           vaSpread2.Row = -1
'           vaSpread2.Col = -1
'           vaSpread2.Lock = True
'           Toolbar1.Buttons(12).Enabled = False
'        Else
         
         If (nestado <> 1) And (nestado <> 3) And (nestado <> 4) Then
           
           vaSpread2.Row = -1
           vaSpread2.Col = -1
           vaSpread2.Lock = True
           Toolbar1.Buttons(12).Enabled = False
         
         End If

'       End If
   
   End If
   
  If tippedido <> 2 Then
    
    vaSpread2.SetActiveCell 1, rowanterior
  
  If rowanterior <> 0 Then
    
    If REGISTROS_ELIMINADOS = "" Then
      
      vaSpread2.Row = rowanterior: vaSpread2.Col = -1: vaSpread2.BackColor = &HE0FEFE
    
    End If
     
     rowanterior = 0
  End If

End If

Exit Sub
Man_Error:
    fg_descarga
    
    If Err = 438 Or Err = 70 Or Err = 1004 Then
       
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    
    End If
    
    Close #1
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Detalle_del_Pedido1()

  Dim seleccion As Integer
  Dim pedido As Integer
  Dim TipodePedido As String
  Dim GlosaDetPedido As String
  Dim cecos As String
  
On Error GoTo Man_Error

 ' Rescata el Pedido del cual se va a ver el Detalle
   
        '------------Detalle del Pedido
        
        If RS.State = 1 Then RS.Close
        RS.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        Sql = " sgpadm_sel_seleccionaDetallePedido "
        Sql = Sql & fpText2
        Set RS = vg_db.Execute(Sql)
        
        '-------> Inicio LLenar grilla
        vg_Archxls = fg_ArchivoTxt
        Open vg_Archxls For Output As #1
        vaSpread2.MaxRows = 0
        GlosaDetPedido = ""
        
        Do While Not RS.EOF
        
            vaSpread2.MaxRows = vaSpread2.MaxRows + 1
            vaSpread2.Row = vaSpread2.MaxRows
            
            '------> generar encabezado text
            If vaSpread2.MaxRows = 1 Then
               Print #1, "Ingrediente" & "|" & "Descripcion" & "|" & "Proveedor" & "|" & "Familia" & "|" & "C.Costos" & "|" & "Producto" & "|" & "Descripcion" & "|" & "Unidad" & "|" & "Fecha Despacho" & "|" & "Cant.Despacho" & "|" & "Saldo Consumido" & "|" & "Saldo Ingrediente" & "|" & "Cant.Planif." & "|" & "Fmto. Convenio" & "|" & "Und. Ingrediente" & "|" & "Factor Conversion" & "|" & "Cantidad Original" & "|" & "Precio Convenio"
            End If
            '-------> Print detalle pedido
            Print #1, Val(RS(1)) & "|" & RS(2) & "|" & IIf(RS(3) <> "" And RS(3) <> "0", RS(3) + " - " + IIf(IsNull(RS(14)), " ", RS(14)), "0") & "|" & RS(4) & "|" & RS(5) & "|" & RS(6) & "|" & RS(7) & "|" & IIf(IsNull(RS(8)), " ", RS(8)) & "|" & " " & Format(RS(10), "DD/MM/YYYY") & "|" & Format(RS(9), "###,00") & "|" & IIf(IsNull(RS(15)), "0", RS(15)) & "|" & IIf(IsNull(RS(16)), "0", RS(16)) & "|" & IIf(IsNull(RS(17)), "0", RS(17)) & "|" & IIf(IsNull(RS(18)), "0", RS(18)) & "|" & IIf(IsNull(RS(19)), "0", RS(19)) & "|" & IIf(IsNull(RS(20)), "0", RS(20)) & "|" & IIf(IsNull(RS(22)), " ", RS(22)) & "|" & IIf(IsNull(RS(23)), " ", RS(23))
            
            vaSpread2.Col = 2 ' Codigo Ingrediente
            vaSpread2.text = Val(RS(1))
            vaSpread2.TypeHAlign = TypeHAlignLeft
            
            vaSpread2.Col = 3 ' Nombre Ingrediente
            vaSpread2.TypeHAlign = TypeHAlignLeft
            vaSpread2.text = RS(2)
            
            vaSpread2.Col = 4 ' Proveedor
            If RS(3) <> "" And RS(3) <> "0" Then
              
              vaSpread2.text = RS(3) + " - " + IIf(IsNull(RS(14)), " ", RS(14))
            
            Else
              
              vaSpread2.text = "0"
            
            End If
            
            vaSpread2.Col = 5 ' Familia Producto
            vaSpread2.text = RS(4)
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 6 ' Cecos
            vaSpread2.text = RS(5)
            cecos = RS(5)
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 7 ' Producto Sap
            vaSpread2.text = RS(6)
            
            vaSpread2.Col = 8 ' Des. Producto Sap
            vaSpread2.text = RS(7)
            
            vaSpread2.Col = 9 ' Unidad
            vaSpread2.text = IIf(IsNull(RS(8)), " ", RS(8))
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 10 ' Fecha de Despacho
            vaSpread2.text = Format(RS(10), "DD/MM/YYYY")
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 11 ' Cantidad
            vaSpread2.text = Format(RS(9), "###,00")
            vaSpread2.TypeHAlign = TypeHAlignRight
           
            vaSpread2.Col = 12 ' Cantidad Consumida
            vaSpread2.text = IIf(IsNull(RS(15)), "0", RS(15))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 13 ' Cantidad Ingrediente
            vaSpread2.text = IIf(IsNull(RS(16)), "0", RS(16))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 14 ' Cantidad Ingresada
            vaSpread2.text = IIf(IsNull(RS(17)), "0", RS(17))
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 15 ' Formato Convenio
            vaSpread2.text = IIf(IsNull(RS(18)), "0", RS(18))
            vaSpread2.TypeHAlign = TypeHAlignRight
        
            vaSpread2.Col = 16 ' Formato Convenio
            vaSpread2.text = IIf(IsNull(RS(19)), "0", RS(19))
           
        
            vaSpread2.Col = 17 ' Formato Convenio
            vaSpread2.text = IIf(IsNull(RS(20)), "0", RS(20))
            vaSpread2.TypeHAlign = TypeHAlignRight
        
            vaSpread2.Col = 18 ' Numero de Linea
            vaSpread2.text = Val(RS(12))
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 19 ' Tipo Pedido
            vaSpread2.text = Val(RS(13))
            vaSpread2.TypeHAlign = TypeHAlignCenter
            
            vaSpread2.Col = 20 ' Marca
            vaSpread2.text = 1
            vaSpread2.TypeHAlign = TypeHAlignCenter
        
            vaSpread2.Col = 21 ' IdPedido
            vaSpread2.text = IIf(IsNull(RS(21)), "0", RS(21))
            vaSpread2.TypeHAlign = TypeHAlignCenter
             
            vaSpread2.Col = 22 ' Cantidad Original
            vaSpread2.text = IIf(IsNull(RS(22)), " ", RS(22))
            vaSpread2.TypeHAlign = TypeHAlignRight
         
           vaSpread2.Col = 23 ' Cantidad Original
            vaSpread2.text = IIf(IsNull(RS(23)), " ", RS(23))
            vaSpread2.TypeHAlign = TypeHAlignRight
         
        
        RS.MoveNext
        
        Loop
       RS.Close: Set RS = Nothing
       Close #1
    
  
   'valida si se pueden hacer cambio al detalle del pedido
   
   
   Dim tipo As Integer
   
   If tippedido <> 2 Then
    
    vaSpread2.SetActiveCell 1, rowanterior
  
  If rowanterior <> 0 Then
    
    If REGISTROS_ELIMINADOS = "" Then
      
      vaSpread2.Row = rowanterior: vaSpread2.Col = -1: vaSpread2.BackColor = &HE0FEFE
    
    End If
     
    rowanterior = 0
  
  End If

End If

Exit Sub
Man_Error:
    fg_descarga
    Close #1
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Genera_pedido()
' Permite Generar Pedido Nuevo

M_Generacion_Pedido.Show 1, Partida

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

If tippedido = 2 Then
   
   vaSpread2.Row = -1
   vaSpread2.Col = -1
   vaSpread2.Lock = True

Else
   
   vaSpread2.Row = -1
   vaSpread2.Col = -1
   vaSpread2.Lock = False

End If
            
fpProveedor = ""
fpFamilia = ""

On Error GoTo Man_Error

rowanterior = 0
Dim Conta As Integer
Dim i As Integer
Dim seleccion As Integer
Conta = 0
vaSpread2.MaxRows = 0

    If SSTab1.Tab = 0 Then
        
        Toolbar1.Buttons(1).Visible = True
        Toolbar1.Buttons(3).Enabled = True
        Toolbar1.Buttons(8).Enabled = True
        Toolbar1.Buttons(9).Enabled = True
        
        Toolbar1.Buttons(2).Visible = False
        Toolbar1.Buttons(7).Enabled = False
        Toolbar1.Buttons(4).Enabled = True
        Toolbar1.Buttons(10).Enabled = False
        Toolbar1.Buttons(11).Enabled = True 'False
        Toolbar1.Buttons(12).Enabled = False
        Toolbar1.Buttons(13).Enabled = True
        Call limpia_grilla
    
    Else
      
      Toolbar1.Buttons(13).Enabled = False
      
      If tippedido = 2 Then
        
        Toolbar1.Buttons(12).Enabled = False
      
      Else
        
        Toolbar1.Buttons(12).Enabled = True
        Toolbar1.Buttons(3).Enabled = False
        Toolbar1.Buttons(8).Enabled = False
        Toolbar1.Buttons(9).Enabled = False
       
      End If
    
    For i = 1 To vaSpread1.MaxRows
          
      vaSpread1.Row = i
      
      vaSpread1.Col = 1 'Seleccion
      seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
      
      If seleccion = 1 Then
        
        Conta = Conta + 1
      
      End If
      
    Next i
    
    If Conta = 1 Then
        
        Toolbar1.Buttons(1).Visible = False
        Toolbar1.Buttons(2).Visible = True
        
        Call Detalle_del_Pedido
    
    Else
       
       MsgBox "Debe seleccionar un Pedido"
       SSTab1.Tab = 0
       Toolbar1.Buttons(1).Visible = True
       Toolbar1.Buttons(2).Visible = False
       Toolbar1.Buttons(7).Enabled = False
       Toolbar1.Buttons(4).Enabled = True
      ' Toolbar1.Buttons(8).Enabled = False
       Toolbar1.Buttons(10).Enabled = False
       Toolbar1.Buttons(11).Enabled = True 'False
       Toolbar1.Buttons(12).Enabled = False
       Call limpia_grilla
    
    End If
    
    End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub limpia_grilla()
  
On Error GoTo Man_Error
   
   Dim i As Integer

    For i = 1 To M_Lista_Pedido.vaSpread1.MaxRows
        
        M_Lista_Pedido.vaSpread1.Row = i
        vaSpread1.Col = 1 ' Seleccion
        vaSpread1.text = 0
        vaSpread1.TypeHAlign = TypeHAlignCenter
    
    Next i

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error
    
Dim i As Long
Dim seleccion As String
Dim tipopedido As String
Dim EstadoPedido As String
Dim EstSel As Boolean
Dim NomArchivoExcel As String
Dim estados As String
Dim Extension As String
Dim TipoPedidoCDPAP As String
Dim TipoPedidoProy  As String

    Select Case Button.Index
    
    Case 1 ' Genera Pedcido Nuevo por Pedido o Proyectado
        
        'registrar Log sistema
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Ingreso_Pedido"), Me.HelpContextID, "", "", "")
        
        Call Genera_pedido
        
    Case 4 ' Filtra los Pedidos
        
        Call busca_encabezado
       
    Case 3  'Enviando Minuta a Sitio se cambia concepto Enviado a Pel
        
        Call Envio_A_Pel
       
    Case 7  'Detalle del Pedido
        
        Call Detalle_del_Pedido
       
    Case 8 'Revisar el Mensaje
        
        Call Revisar_el_Mensaje
       
    Case 9 'Reprocesar
        
        Call Reprocesar
    
    Case 10 'Deshacer
        
        Call Deshacer
       
    Case 11 'Lleva a Excel
        
        If SSTab1.Tab = 0 Then
           
           '-------> validar si existe seleccionado grilla
           EstSel = False
           For i = 1 To vaSpread1.MaxRows
                
                vaSpread1.Row = i
                vaSpread1.Col = 1
                seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                vaSpread1.Col = 12
                EstadoPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                vaSpread1.Col = 13
                tipopedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)

                If seleccion = 1 Then

                   EstSel = True
                   Exit For

                End If
          
          Next i

          If Not EstSel Then

             MsgBox "No existe registros seleccionado: " & vbCrLf, vbExclamation
             Exit Sub

          End If
          
          '-------> Validar que no existan tipo pedido CD o PAP seleccionado Proyectado
          TipoPedidoCDPAP = ""
          TipoPedidoProy = ""
       
           For i = 1 To vaSpread1.MaxRows
                
                vaSpread1.Row = i
                vaSpread1.Col = 1
                seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                vaSpread1.Col = 13
                tipopedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
                
                If seleccion = "1" Then
                   If tipopedido = "1" Or tipopedido = "3" Then
                   
                      TipoPedidoCDPAP = "1"
                
                   ElseIf tipopedido = "2" Then
                
                      TipoPedidoProy = "1"
                   
                   End If
                   
               End If
               
          Next i

          If TipoPedidoCDPAP = "1" And TipoPedidoProy = "1" Then

             MsgBox "Existen pedidos (CD o PAP) mezclado con Proyectado. solo debe selecionar por separado " & vbCrLf, vbExclamation
             Exit Sub
             
          End If
          
          '-------> Guardar nombre archivo excel
          NomArchivoExcel = ""
          CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
'          CD.Filter = "Todos los archivos (*.xls)|*.xls,*.xlsx"
          CD.Filter = "Todos los archivos *.xls,*.xlsx"
          CD.ShowSave
          If CD.FileName = "" Then
             
             MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
             Exit Sub
          
          Else
             
             Extension = ""
             Extension = Right(CD.FileName, Len(CD.FileName) - (InStr(CD.FileName, ".")))
             If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
                
                MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
                Exit Sub
             
             End If
             NomArchivoExcel = CD.FileName  'Dir(CD.Filename)
          
          End If
          
          'registrar Log sistema
          Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_Excel"), CStr(Me.HelpContextID), "", "", "")
          
          If Not Exportar_PedidoExcelMasivo(NomArchivoExcel) Then
             
             MsgBox "Ocurrio un error al exportar", vbCritical
          
          Else
             
             MsgBox "Exportación realizada con exito", vbInformation
          
          End If
    
          If MsgBox("Desea Generar Pedido Ordenado Familia Producto ??", vbQuestion + vbYesNo, MsgTitulo) = vbYes Then

             '-------> Guardar nombre archivo excel
             NomArchivoExcel = ""
             CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
             CD.Filter = "Todos los archivos *.xls,*.xlsx"
             CD.ShowSave
             If CD.FileName = "" Then
                
                MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
                Exit Sub
             
             Else
                
                Extension = ""
                Extension = Right(CD.FileName, Len(CD.FileName) - (InStr(CD.FileName, ".")))
                
                If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
                   
                   MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
                   Exit Sub
                
                End If
                NomArchivoExcel = CD.FileName  'Dir(CD.Filename)
             
             End If
             
             If Not Exportar_PedidoExcelMasivoxFamilia(NomArchivoExcel) Then
                
                MsgBox "Ocurrio un error al exportar", vbCritical
             
             Else
                
                MsgBox "Exportación realizada con exito", vbInformation
             
             End If

          End If
          
          
        Else
           
           'registrar Log sistema
           Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Exporta_Excel"), CStr(Me.HelpContextID), "", "", "")
           
           Call lleva_excel
        
        End If
    
    Case 12 'Agergar Ingredientes
        
        'registrar Log sistema
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Agregar_Ingrediente_Pedido"), CStr(Me.HelpContextID), "", "", "")
       
        Call Agregar_Ingrediente
        Call Detalle_del_Pedido1
       
    
    Case 13 ' Cambio de Fecha
        
        Call Cambio_fecha
        
    Case 14 'Convenios

        'registrar Log sistema
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Listar_Convenios"), CStr(Me.HelpContextID), "", "", "")
 
        C_Convenios.Show 0, Me
        
    Case 15 'Saldo Mayor Unidad de Formato de Compra
        
        'registrar Log sistema
        Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Listar_Saldo_Mayor"), CStr(Me.HelpContextID), "", "", "")
        
        Exp_Excel_Saldo_Mayor_FormatoCompras
    
    Case 17 ' Salir del Programa
        
        Me.Hide
        Unload Me
    
    End Select
        
        'Call Deshacer

Exit Sub
Man_Error:
    fg_descarga
    
    If Err = 438 Or Err = 70 Then
       
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    
    End If
   
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Agregar_Ingrediente()
  
  M_Agregar_Ingrediente_al_Pedido.Show 1, Partida

End Sub

Private Sub Cambio_fecha()
   
   Call valida_selecccion

End Sub

Sub Exportar_PedidoTextoMasino(NomArchivoTexto As String)

On Error GoTo Man_Error

    'Exportar_Excel = False
    Dim rst As New ADODB.Recordset
    
    Dim seleccion As Long
    Dim NPedido As Long
    Dim MyBuffer    As String
    Dim str_Dato As String
    Dim i As Long
    Dim Path_Csv As String
    Dim NomArcTexto As String
    Dim tipopedido  As String
    Dim EstadoPedido  As String
    
    '-------> Armar xml
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<NewDataSet>"
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        
        '-------> estado pedido
        vaSpread1.Col = 12
        EstadoPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> tipo pedido
        vaSpread1.Col = 13
        tipopedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)

        If seleccion = 1 And (tipopedido = "1" Or tipopedido = "3") And (EstadoPedido = "2" Or EstadoPedido = "5" Or EstadoPedido = "97") Then
           
           vaSpread1.Col = 2 'Nº Pedido
           NPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           MyBuffer = MyBuffer & " <Np"
           MyBuffer = MyBuffer & " Np = " & Chr(34) & NPedido & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next i
    MyBuffer = MyBuffer & "</NewDataSet>"

    '-------> Lectura
    If rst.State = 1 Then rst.Close
    rst.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = ""
    Sql = " sgpadm_Sel_XmlExportarPedidoTexto "
    Sql = Sql & " '" & MyBuffer & "'"

    Set rst = vg_db.Execute(Sql)
    
    If Not rst.EOF Then

'        str_ = ""
'        For i = 0 To RS.Fields.count
'            str_ = str_ + RS.Fields.item(i).Name
'        Next i
'        str_ = str_ + vbCrLf
         str_Dato = rst.GetString(adClipString, -1, ";", vbCrLf, "")
         '-------> Abre y Crea un archivo de texto para escribir los datos
         'NomArcTexto = "Pedido" & "-" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".txt"
         Path_Csv = NomArchivoTexto 'dir_trabajo & NomArcTexto
         If Dir(Path_Csv) <> "" Then Kill Path_Csv   ' borrar base datos si existe

         Open Path_Csv For Output As #1
         '-------> escribe los datos
         Print #1, str_Dato
         '-------> cierra
         Close #1
         '-------> Ok

    End If
    
    rst.Close
    Set rst = Nothing
    
Exit Sub
Man_Error:
    fg_descarga
    
    If Err = 438 Or Err = 70 Or Err = 1004 Then
       
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    
    End If
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Function Exportar_PedidoExcelMasivo(NomArchivoExcel As String) As Boolean

On Error GoTo Man_Error

    Exportar_PedidoExcelMasivo = False
    
    Dim rst As New ADODB.Recordset
    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    
    Dim recArray As Variant
    
    Dim strDB As String
    Dim fldCount As Integer
    Dim recCount As Long
    Dim icol As Integer
    Dim IRow As Integer

    Dim seleccion As Long
    Dim NPedido As Long
    Dim MyBuffer    As String
    Dim i As Long
    Dim tipopedido  As String
    Dim EstadoPedido  As String
    
    If Dir(NomArchivoExcel) <> "" Then Kill NomArchivoExcel
    
    '-------> Armar xml
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<NewDataSet>"
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> estado pedido
        vaSpread1.Col = 12
        EstadoPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> tipo pedido
        vaSpread1.Col = 13
        tipopedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)

        If seleccion = 1 Then

'        And (tipopedido = "1" Or tipopedido = "3") And (EstadoPedido = "2" Or EstadoPedido = "5" Or EstadoPedido = "97") Then
           vaSpread1.Col = 2 'Nº Pedido
           NPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           MyBuffer = MyBuffer & " <Np"
           MyBuffer = MyBuffer & " Np = " & Chr(34) & NPedido & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next i
    MyBuffer = MyBuffer & "</NewDataSet>"

    '-------> Lectura
    If rst.State = 1 Then rst.Close
    rst.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = ""
    Sql = " sgpadm_Sel_XmlExportarPedidoExcel_V03 "
    Sql = Sql & " '" & MyBuffer & "'"

    Set rst = vg_db.Execute(Sql)

    '-------> Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Hoja1")
  
    '-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
    xlApp.UserControl = True
    
    '-------> Check version of Excel
    Call encabezado(rst, xlWs)
          
    xlWs.Cells(2, 1).CopyFromRecordset rst
    '-------> Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    
    xlApp.Columns("F:F").Select
    xlApp.Selection.NumberFormat = "0,00"
    xlApp.Columns("F:F").Select
    With xlApp.Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    'Poner color amarillo la columna D - H
    xlApp.Range("J1").Select
    With xlApp.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    xlApp.Range("K1").Select
    With xlApp.Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    xlWb.Close True, NomArchivoExcel

    Dim XL As New excel.Application 'Crea el objeto excel
    XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    XL.Visible = True
    XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
    '-------> Close ADO objects
    rst.Close
    Set rst = Nothing
    
    ' -- Cerrar Excel
    xlApp.Quit
    '-------> Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
'    EstadoPedido = Len(Right(NomArchivoExcel, Len(NomArchivoExcel) - (InStr(NomArchivoExcel, "."))))
'    Call Exportar_PedidoTextoMasino(Mid(NomArchivoExcel, 1, Len(NomArchivoExcel) - 3) & "txt")
    Call Exportar_PedidoTextoMasino(Mid(NomArchivoExcel, 1, Len(NomArchivoExcel) - Len(Right(NomArchivoExcel, Len(NomArchivoExcel) - (InStr(NomArchivoExcel, "."))))) & "txt")
    Exportar_PedidoExcelMasivo = True
    
Exit Function
Man_Error:
    fg_descarga
    Exportar_PedidoExcelMasivo = False
    If Err = 438 Or Err = 70 Or Err = 1004 Then
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Function
    End If
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
    
End Function

Sub encabezado(ByRef rst As ADODB.Recordset, ByRef xlWs As Object)

On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = rst.Fields.count
For icol = 1 To fldCount
    
    xlWs.Cells(1, icol).Value = rst.Fields(icol - 1).Name

Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, MsgTitulo

End Sub

Function Exportar_PedidoExcelMasivoxFamilia(NomArchivoExcel As String) As Boolean

On Error GoTo Man_Error

    Exportar_PedidoExcelMasivoxFamilia = False
    
    Dim RS    As New ADODB.Recordset
    Dim xlApp As Object
    Dim xlWb  As Object
    Dim xlWs  As Object
    
    Dim seleccion    As Long
    Dim NPedido      As Long
    Dim MyBuffer     As String
    Dim i            As Long
    Dim tipopedido   As String
    Dim EstadoPedido As String
    
    If Dir(NomArchivoExcel) <> "" Then Kill NomArchivoExcel
    
    '-------> Armar xml
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<NewDataSet>"
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = 1
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> estado pedido
        vaSpread1.Col = 12
        EstadoPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> tipo pedido
        vaSpread1.Col = 13
        tipopedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)

        If seleccion = 1 Then
'        And (tipopedido = "1" Or tipopedido = "3") And (EstadoPedido = "2" Or EstadoPedido = "5" Or EstadoPedido = "97") Then
           vaSpread1.Col = 2 'Nº Pedido
           NPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           MyBuffer = MyBuffer & " <Np"
           MyBuffer = MyBuffer & " Np = " & Chr(34) & NPedido & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
        End If
    
    Next i
    MyBuffer = MyBuffer & "</NewDataSet>"

    '-------> Lectura
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = ""
    Sql = " sgpadm_Sel_XmlExportarFamiliaPedidoExcel_V01 "
    Sql = Sql & " '" & MyBuffer & "'"

    Set RS = vg_db.Execute(Sql)

    '-------> Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Hoja1")
  
    '-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
    xlApp.UserControl = True
    
    '-------> Check version of Excel
    Call encabezado(RS, xlWs)
          
    xlWs.Cells(2, 1).CopyFromRecordset RS
    '-------> Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    
    xlWb.Close True, NomArchivoExcel

    Dim XL As New excel.Application 'Crea el objeto excel
    XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    XL.Visible = True
    XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
    '-------> Close ADO objects
    RS.Close
    Set RS = Nothing
    
    ' -- Cerrar Excel
    xlApp.Quit
    '-------> Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    
    Exportar_PedidoExcelMasivoxFamilia = True

Exit Function
Man_Error:
    
    fg_descarga
    Exportar_PedidoExcelMasivoxFamilia = False
    
    If Err = 438 Or Err = 70 Or Err = 1004 Then
       
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Function
    
    End If
    
    MsgBox Err.Description, vbCritical, MsgTitulo

End Function

Private Sub lleva_excel()

On Error GoTo Man_Error

If vaSpread2.MaxRows < 1 Then Exit Sub
  
    Screen.MousePointer = 11
    DoEvents
  
    Dim X As Boolean
    vaSpread2.Row = -1
    vaSpread2.Col = -1
    vaSpread2.RowHidden = False
    
    
    ' Export Excel file and set result to x
    Dim XL As excel.Application
    Set XL = CreateObject("Excel.application")
    XL.Visible = True
    XL.Workbooks.OpenText vg_Archxls, , 1, 1, , , , , , , True, "|"

    Screen.MousePointer = 1
    DoEvents

Exit Sub
Man_Error:
'    XL.Close
    fg_descarga
    
    If Err = 438 Or Err = 70 Or Err = 1004 Then
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    End If
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub Deshacer()

On Error GoTo Man_Error
    
Toolbar1.Buttons(10).Enabled = False
    
'registrar Log sistema
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Desmarcar"), CStr(Me.HelpContextID), "", "", "")
    
Call limpia_grilla

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
        vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
    Next

End Select

End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

    ' Validad si hay mas de Un Click para la fecha
 
Dim estado As Integer
Dim seleccion As Integer


On Error GoTo Man_Error

If vaSpread1.MaxRows = 0 Then Exit Sub


    vaSpread1.Row = vaSpread1.ActiveRow
    rowanterior = vaSpread1.ActiveRow
   
    M_Lista_Pedido.vaSpread1.Col = 12 'Estado de Pedido
    estado = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
    M_Lista_Pedido.vaSpread1.Col = 13 'Tipo de Pedido
    tippedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
    M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
    seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)

    M_Lista_Pedido.vaSpread1.Col = 7 'Seleccion
    fpDateTime1(2).text = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
    M_Lista_Pedido.vaSpread1.Col = 8 'Seleccion
    fpDateTime1(3).text = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
    M_Lista_Pedido.vaSpread1.Col = 14 'Fecha Confirmacion
    fechapedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    If fechapedido <> " " Then
        DTPicker1 = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    End If
    fecha_limiteanterior = Format(DTPicker1, "YYYYMMDD HH:MM:SS")
    M_Lista_Pedido.vaSpread1.Col = 15 'Fecha Pedido
    lbl_fecha = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
   
    Toolbar1.Buttons(7).Enabled = True
     
    If tippedido = 2 Then
         Toolbar1.Buttons(13).Enabled = False
    Else
         Toolbar1.Buttons(13).Enabled = True
    End If
      
    
 Toolbar1.Buttons(10).Enabled = True
      
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub valida_selecccion()

Dim contador As Integer
Dim seleccion As Integer
Dim i As Integer
Dim estado As Integer
Dim cecos As String
Dim pedido1 As Integer
Dim pedido2 As Integer
Dim Conta As Integer
Conta = 0
pedido1 = 0
pedido1 = 0
contador = 0

On Error GoTo Man_Error

 vaSpread1.Row = vaSpread1.ActiveRow
 For i = 1 To vaSpread1.MaxRows
     
     vaSpread1.Row = i
     M_Lista_Pedido.vaSpread1.Col = 1 'Seleccion
     seleccion = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
  If seleccion = 1 Then
     
     contador = contador + 1
     M_Lista_Pedido.vaSpread1.Col = 2 'Pedido
     pedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
     
     M_Lista_Pedido.vaSpread1.Col = 12 'Seleccion
     estado = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
     
 
     M_Lista_Pedido.vaSpread1.Col = 7 'fecha desdes
     fdesde = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
    
     M_Lista_Pedido.vaSpread1.Col = 8 'fecha hasta
     fhasta = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
     
     M_Lista_Pedido.vaSpread1.Col = 13 'Cecos
     tippedido = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
          
     M_Lista_Pedido.vaSpread1.Col = 4 'Cecos
     cecos = IIf(M_Lista_Pedido.vaSpread1.text = "", 0, M_Lista_Pedido.vaSpread1.text)
     
     'Recorre para buscar los 2 Ultimos pedido para poder cambiar la fecha
     
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     
     Sql = " sgpadm_Sel_penultimopedidoparaActualizarfecha " & "'" & cecos & "'," & tippedido
     Set RS = vg_db.Execute(Sql)
     
     If Not RS.EOF Then
         
         Conta = Conta + 1
         Do While Not RS.EOF
            
            If Conta = 1 Then
               
               pedido1 = RS(0)
            
            End If
            
            If Conta = 2 Then
               
               pedido2 = RS(0)
            
            End If
            
            RS.MoveNext
            Conta = Conta + 1
      
      Loop
     
     End If
     RS.Close
        
       
       If contador > 1 Then
            
            MsgBox " Debe seleccionarse un Pedido para Modificar Fechas ", vbExclamation
            Exit Sub
       
       End If
      
       
       If pedido2 = 0 Then
          
          If pedido <> pedido1 Then
             
             Screen.MousePointer = 1
             DoEvents
             MsgBox " Solo se puede cambiar fecha a los 2 Ultimo Pedido realizado para este Ceco " + CStr(cecos), vbExclamation
            Exit Sub
          
          End If
       
       Else
          
          If pedido <> pedido1 And pedido <> pedido2 Then
             
             Screen.MousePointer = 1
             DoEvents
             MsgBox " Solo se puede cambiar fecha a los 2 Ultimo Pedido realizado para este Ceco " + CStr(cecos), vbExclamation
            Exit Sub
          
          End If
       
       End If
     
     'fin de la rutina para validar
 
 End If
 
 Next i
 
 If contador = 1 Then
   
   If estado = 1 Or estado = 3 Then
      
      Frame3.Visible = True
   
   Else
      
      MsgBox " El Estado del Pedido No Permite Modificar Fecha", vbExclamation
      Exit Sub
   
   End If
 
 Else
   
   MsgBox " Debe seleccionarse un Pedido para Modificar Fechas ", vbExclamation
   Exit Sub
 
 End If

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)
  
On Error GoTo Man_Error

Dim proveedor As String
Dim familia As String

vaSpread2.Row = vaSpread2.ActiveRow
vaSpread2.Col = vaSpread2.ActiveCol

 With Me.vaSpread2

            .Row = Row
            .Col = 4
            proveedor = IIf(Trim(.text) = "", -1, .text)

            .Row = Row
            .Col = 5
            familia = IIf(Trim(.text) = "", -1, .text)

            Lbl_provedor.Visible = True
            fpProveedor.Visible = True
            Lbl_familuia.Visible = True
            fpFamilia.Visible = True

            fpProveedor = UCase(proveedor)
            fpFamilia = familia

  End With


Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)


End Sub

Private Sub vaSpread2_DblClick(ByVal Col As Long, ByVal Row As Long)
If vaSpread2.MaxRows < 1 Then Exit Sub

   On Error GoTo Man_Error
    
    Dim Costo As String
    Dim Ingrediente As Integer
    Dim proveedor As String
    Dim familia As String
    Dim codmaterial As String
    Dim desmaterial As String
     
    cantidad_despacho_anterior = 0
    cantidad_despacho = 0
    
    vaSpread2.Row = vaSpread2.ActiveRow
    vaSpread2.Col = vaSpread2.ActiveCol
    rowanterior = vaSpread2.ActiveRow
    
    
    'valida si se pueden hacer cambio al detalle del pedido

    vaSpread2.Row = Row
    vaSpread2.Col = 6
    Costo = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)
            
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Sql = " sgpadm_Sel_mayorpedidoparaReprocesar " & "'" & Costo & "'," & tippedido
   Set RS = vg_db.Execute(Sql)
  
   If Not RS.EOF Then
'       If (fpText2 = RS(0)) Then
         If (nestado = 1) Or (nestado = 3) Or (nestado = 4) Then
       
    If (vaSpread2.Col = 7 Or vaSpread2.Col = 8) Or vaSpread2.Row > 0 Then
     
        With Me.vaSpread2
        
            .Row = Row
            .Col = 2
            Ingrediente = IIf(Trim(.text) = "", -1, .text)
            Ingrediente_ACAMBIAR = IIf(Trim(.text) = "", -1, .text)
            .Row = Row
            .Col = 4
            proveedor = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 5
            familia = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 7
            codmaterial = IIf(Trim(.text) = "", -1, .text)
            .Row = Row
            .Col = 8
            desmaterial = IIf(Trim(.text) = "", -1, .text)
            
            Fp_proveedor = UCase(proveedor)
            fp_codigo = codmaterial
            fp_descripcion = desmaterial
                  
            Lbl_provedor.Visible = True
            fpProveedor.Visible = True
            Lbl_familuia.Visible = True
            fpFamilia.Visible = True
            fpProveedor = UCase(proveedor)
            fpFamilia = familia
            
            .Row = Row
            .Col = 6
            Costo = IIf(Trim(.text) = "", -1, .text)
                 
        ' Rescata los Codigo Sap Asociados al Ingrediente

            If RS.State = 1 Then RS.Close
            RS.CursorLocation = adUseClient
            vg_db.CursorLocation = adUseClient
            
            Sql = "sgpadm_Sel_materialsap_sinPref_Nuevo "
            Sql = Sql & "'" & Costo & "',"
            Sql = Sql & "'" & Ingrediente & "',"
            Sql = Sql & "'" & Format(Me.fpDateTime1(0).text, "yyyymmdd") & "',"
            Sql = Sql & "'" & Format(Me.fpDateTime1(1).text, "yyyymmdd") & "'"
            Set RS = vg_db.Execute(Sql)
                
       '-------> Inicio LLenar grilla
            Frame2.Visible = False
            vaSpread3.MaxRows = 0
            If Not RS.EOF Then
            Frame2.Visible = True
            
            Do While Not RS.EOF
                
                vaSpread3.MaxRows = vaSpread3.MaxRows + 1
                vaSpread3.Row = vaSpread3.MaxRows
                
                vaSpread3.Col = 1 ' Codigo Ingrediente
                vaSpread3.text = RS(0) + " - " + RS(4)
                
                vaSpread3.Col = 2 ' Nombre Ingrediente
                vaSpread3.text = RS(1)
                
                vaSpread3.Col = 3 ' Proveedor
                vaSpread3.text = RS(2)
                 
                vaSpread3.Col = 4 ' Descripcio Material
                vaSpread3.text = RS(3)
                
                vaSpread3.Col = 5 ' Descripcion Medida
                vaSpread3.text = RS(5)
                
                vaSpread3.Col = 6 ' Formato Convenio
                vaSpread3.text = RS(6)
                
                vaSpread3.Col = 7 ' Profuctos GSP
                vaSpread3.text = RS(7)
                
                vaSpread3.Col = 8 ' Precio Convenio
                vaSpread3.text = Format(RS(8), fg_Pict(6, 2))
                
                vaSpread3.Col = 9 ' Convenio
                vaSpread3.text = Format(RS(9), fg_Pict(6, 2))
                
                vaSpread3.Col = 10 ' Minima Pedido
                vaSpread3.text = Format(RS(12), fg_Pict(6, 2))
                
                
                RS.MoveNext
           Loop
        Else
            
            MsgBox "No se encuentra producto alternativo, para las fecha de despacho entre : " & Me.fpDateTime1(0).text & " - " & Me.fpDateTime1(1).text
        
        End If
        RS.Close: Set RS = Nothing
                
    End With
    

'End If

       End If
       End If
   End If
  
Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Private Sub vaSpread2_EditMode(ByVal Col As Long, ByVal Row As Long, ByVal Mode As Integer, ByVal ChangeMade As Boolean)

On Error GoTo Man_Error

If vaSpread2.MaxRows < 1 Then Exit Sub

Dim codest As Long
Dim i As Long
Dim icol As Long
Dim numlinea As Integer
Dim pedido As Integer
Dim actualizacion As String
    
    vaSpread2.SetFocus
    vaSpread2.Row = vaSpread2.ActiveRow
    rowanterior = vaSpread2.ActiveRow
    vaSpread2.Col = 22
    cantidad_despacho_anterior = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)
   
    vaSpread2.Col = 15
    perfil_redondeo = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)

If Mode = 1 Then

   If cantidad_despacho_anterior = 0 Then
        
        MsgBox "Cantidad Calculada de Despacho es 0 no se puede Cambiar", vbExclamation
        vaSpread2.Col = 11 ' Cantidad Despacho
        vaSpread2.text = 0
        vaSpread2.SetFocus
        vaSpread2.Row = rowanterior
       Exit Sub
   
   End If


   If perfil_redondeo = 0 Then
       
       MsgBox "El Perfil de Redondeo debe ser mayor a 0", vbExclamation
       Exit Sub
   
   End If
 
 End If



If ChangeMade = True Then
    
    vaSpread2.Col = 11
    cantidad_despacho = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)
 
 
 If cantidad_despacho = 0 Then

          vaSpread2.Col = 18 'Numero de Linea
          numlinea = IIf(vaSpread2.text = "", 0, vaSpread2.text)

          vaSpread2.Col = 19 'Pedido
          pedido = IIf(vaSpread2.text = "", 0, vaSpread2.text)

          Sql = " sgpadm_iu_ActualizaDetalleCantidad "
          Sql = Sql & numlinea & ","
          Sql = Sql & pedido & ","
          Sql = Sql & cantidad_despacho
          Set RS = vg_db.Execute(Sql)


          vaSpread2.Col = 11 ' Cantidad de Despacho
          vaSpread2.text = cantidad_despacho

          MsgBox "Se Actualizo la Cantidad del Detalle Correctamente", vbExclamation
          cantidad_despacho = 0
          vaSpread2.SetFocus
          vaSpread2.SetActiveCell 11, rowanterior
          Exit Sub
 End If
 
 If cantidad_despacho_anterior <> 0 And cantidad_despacho <> 0 Then

    If cantidad_despacho - (perfil_redondeo * Fix(cantidad_despacho / perfil_redondeo)) = 0 Then
        
        If cantidad_despacho <= cantidad_despacho_anterior Then
        
            vaSpread2.Col = 18 'Numero de Linea
            numlinea = IIf(vaSpread2.text = "", 0, vaSpread2.text)
        
            vaSpread2.Col = 19 'Pedido
            pedido = IIf(vaSpread2.text = "", 0, vaSpread2.text)
        
            Sql = " sgpadm_iu_ActualizaDetalleCantidad "
            Sql = Sql & numlinea & ","
            Sql = Sql & pedido & ","
            Sql = Sql & cantidad_despacho
            Set RS = vg_db.Execute(Sql)
            
            vaSpread2.Col = 11 ' Cantidad de Despacho
            vaSpread2.text = cantidad_despacho
        
            MsgBox "Se Actualizo la Cantidad del Detalle  Correctamente", vbExclamation
           
            cantidad_despacho_anterior = 0
            cantidad_despacho = 0
            vaSpread2.SetFocus
            vaSpread2.SetActiveCell 11, rowanterior
        
        Else
            vaSpread2.Col = 11 ' Cantidad de Despacho
            vaSpread2.text = cantidad_despacho_anterior
            
            MsgBox "La Cantidad Despachada debe ser menor o igual al Despacho Original " & CStr(cantidad_despacho_anterior), vbExclamation
            cantidad_despacho_anterior = 0
            cantidad_despacho = 0
            vaSpread2.SetFocus
            vaSpread2.SetActiveCell 11, rowanterior
        End If
       Else
             vaSpread2.Col = 11 ' Cantidad de Despacho
             vaSpread2.text = cantidad_despacho_anterior
             MsgBox "La Cantidad Despachada debe ser Multiplo del Perfil de Redondeo", vbExclamation
             cantidad_despacho_anterior = 0
             cantidad_despacho = 0
             vaSpread2.SetFocus
             vaSpread2.SetActiveCell 11, rowanterior
             Exit Sub
      
      
       End If
 End If

End If

Man_Error:
    fg_descarga
    If Err.Number = 11 Then
      MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
     ins_log_error Date & Time & Err & ":  " & Error$(Err)
    End If

 End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
   
   On Error GoTo Man_Error

   Dim proveedor As String
   Dim familia As String
   
   vaSpread2.Row = vaSpread2.ActiveRow

 With Me.vaSpread2


            .Row = Row
            .Col = 4
            proveedor = IIf(Trim(.text) = "", -1, .text)

            .Row = Row
            .Col = 5
            familia = IIf(Trim(.text) = "", -1, .text)

            Lbl_provedor.Visible = True
            fpProveedor.Visible = True
            Lbl_familuia.Visible = True
            fpFamilia.Visible = True

            fpProveedor = UCase(proveedor)
            fpFamilia = familia

  End With

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

Private Sub vaSpread2_ScriptTextTipFetch(ByVal Col As Long, ByVal Row As Long, MultiLine As Variant, TipWidth As Variant, TipText As Variant, ShowTip As Variant)

'If vaSpread2.MaxRows < 1 Or Col = 1 Then Exit Sub
If vaSpread2.MaxRows < 1 Or Col = 0 Then Exit Sub
' Set tip to display and set tip's content
vaSpread2.Row = Row
TipWidth = 1000
ShowTip = True
MultiLine = 2

Select Case Col

    Case 2
        
        vaSpread2.Col = Col
        TipText = "Nombre Ingrediente : " & vaSpread2.text
    
    Case 3
        
        vaSpread2.Col = Col
        TipText = "Proveedor : " & vaSpread2.text
    
    Case 7
        
        vaSpread2.Col = Col
        TipText = "Nombre Producto : " & vaSpread2.text

End Select

End Sub

Private Sub vaSpread3_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

    Dim proveedor As String
    Dim familia As String
    Dim codmaterial As String
    Dim desmaterial As String
    Dim desunidad As String
    Dim perfilredondeo As Double
    Dim productosgp As String
    Dim precio As Double
    Dim formatoconvenio  As Double
    Dim MinimaPedido As Double
    
' Actualiza la Grilla de Detalle de Producto con el Producto Nuevo
    
    With Me.vaSpread3
            
            .Row = .ActiveRow
            .Col = 1
            proveedor = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 2
            familia = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 3
            codmaterial = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 4
            desmaterial = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 5
            desunidad = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 6
            perfilredondeo = IIf(Trim(.text) = "", -1, .text)
    
            .Row = Row
            .Col = 7
             productosgp = IIf(Trim(.text) = "", -1, .text)
    
    
            .Row = Row
            .Col = 8
            precio = IIf(Trim(.text) = "", -1, .text)
    
            .Row = Row
            .Col = 9
            formatoconvenio = IIf(Trim(.text) = "", -1, .text)
            
            .Row = Row
            .Col = 10
            MinimaPedido = IIf(Trim(.text) = "", -1, .text)
     
    End With

Dim i As Integer
Dim Ingrediente_actual As String
Dim CodGrupoEst As Long
Dim RowGrupo As Long
Dim colgrupo  As Long
Dim Proveedoractual As String
Dim pedido As Long
Dim numlinea As Integer
Dim ccosto As String
Dim Ingrediente As String

Dim swpregunta As String
Dim CONTADORpregunta As Integer
CONTADORpregunta = 0
swpregunta = ""
REGISTROS_ELIMINADOS = ""
CodGrupoEst = Ingrediente_ACAMBIAR
'-------> columna de grupo de estructura y Encabezado
'colgrupo = vaSpread2.GetColFromID("Ingrediente") + 1
'-------> fila a buscar
RowGrupo = vaSpread2.SearchCol(2, 0, -1, CodGrupoEst, SearchFlagsValue)

For i = RowGrupo To vaSpread2.MaxRows
  
  CONTADORpregunta = CONTADORpregunta + 1
  vaSpread2.Row = i
  vaSpread2.Col = 2
  Ingrediente_actual = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)

If Ingrediente_ACAMBIAR <> Ingrediente_actual Then Exit For
    
    vaSpread2.Col = 6
    ccosto = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)
    
    vaSpread2.Col = 2
    Ingrediente = IIf(Trim(vaSpread2.text) = "", -1, vaSpread2.text)
    
    vaSpread2.Col = 4 ' Proveedor
    vaSpread2.text = proveedor
    Proveedoractual = Mid(proveedor, 1, 9)
   
    vaSpread2.Col = 5 ' Familia Producto
    vaSpread2.text = familia
    vaSpread2.TypeHAlign = TypeHAlignCenter
    
    
    vaSpread2.Col = 6 ' C.Costo
    vaSpread2.text = ccosto
    vaSpread2.TypeHAlign = TypeHAlignCenter
    
    vaSpread2.Col = 2 ' Ingrediente
    vaSpread2.text = Ingrediente
       
    vaSpread2.Col = 7 ' Producto
    vaSpread2.text = codmaterial
    vaSpread2.TypeHAlign = TypeHAlignCenter
    
    vaSpread2.Col = 8 ' Des. Material
    vaSpread2.text = desmaterial
    
    vaSpread2.Col = 9 ' Unidad
    vaSpread2.text = desunidad
    
    vaSpread2.Col = 15 ' perfil de Redondeo
    vaSpread2.text = perfilredondeo
    
    vaSpread2.Col = 18 'Numero de Linea
    numlinea = IIf(vaSpread2.text = "", 0, vaSpread2.text)
    
    vaSpread2.Col = 19 'Pedido
    pedido = IIf(vaSpread2.text = "", 0, vaSpread2.text)
   
   'VAlidacion del RuT Proveedor es CD
    
    Dim Rut_ES_Cd As Integer
   
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
   
    Sql = " sgpadm_sel_ExisteRutCD "
    Sql = Sql & " '" & Proveedoractual & "'"
    Set RS = vg_db.Execute(Sql)
    
    If Not RS.EOF Then
    
      Rut_ES_Cd = RS(0)
    
    
    End If
    
    RS.Close
   
   
   If tippedido = 1 Then
    If Rut_ES_Cd = 1 Then ' El Rut de Proveedor Actual es CD
    
     If CONTADORpregunta = 1 Then
            If MsgBox(" ¿ El Proveedor no coincide con Tipo de Pedido ?, si Presiona SI " + Chr(13) + " - Elimina el Ingrediente del Pedido " + Chr(13) + " - Se devuelve el Saldo ", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
                swpregunta = "SI"
                REGISTROS_ELIMINADOS = "SI"

                Sql = " sgpadm_del_DetallePedidoNuevo "
                Sql = Sql & "'" & pedido & "',"
                Sql = Sql & "'" & numlinea & "'"
                
                Set RS = vg_db.Execute(Sql)
            
       End If
     Else
       Sql = ""
       Sql = " sgpadm_iu_ActualizaDetallePedidoNuevo "
       Sql = Sql & " '" & Proveedoractual & "',"
       Sql = Sql & " '" & familia & "',"
       Sql = Sql & " '" & codmaterial & "',"
       Sql = Sql & " '" & desmaterial & "',"
       Sql = Sql & " '" & numlinea & "',"
       Sql = Sql & " '" & pedido & "',"
       Sql = Sql & " '" & desunidad & "',"
       Sql = Sql & " '" & perfilredondeo & "',"
       Sql = Sql & " '" & productosgp & "',"
       Sql = Sql & " '" & precio & "',"
       Sql = Sql & " '" & formatoconvenio & "', "
       Sql = Sql & " '" & MinimaPedido & "'"
       
       Set RS = vg_db.Execute(Sql)
       
       '--> Ejecutar Detalle pedidos posteriores
       Sql = ""
       Sql = " sgpadm_Pro_ReprocesarDetPedidosPosteriores "
       Sql = Sql & " '" & ccosto & "',"
       Sql = Sql & " " & pedido & ","
       Sql = Sql & " '" & Ingrediente & "'"
       Set RS = vg_db.Execute(Sql)
              
     End If
       
       Sql = " sgpadm_INS_excepcionformatoPedido "
       Sql = Sql & " '" & ccosto & "',"
       Sql = Sql & " '" & Ingrediente & "',"
       Sql = Sql & " '" & codmaterial & "',"
       Sql = Sql & " '" & Proveedoractual & "'"
       Set RS = vg_db.Execute(Sql)

   End If
   
   If tippedido = 3 Then
     If Rut_ES_Cd = 0 Then ' El Rut de Proveedor Actual no es CD
       If CONTADORpregunta = 1 Then
           If MsgBox(" ¿ El Proveedor no coincide con Tipo de Pedido ?, si Presiona SI " + Chr(13) + " - Elimina el Ingrediente del Pedido " + Chr(13) + " - Se devuelve el Saldo ", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
             swpregunta = "SI"
             REGISTROS_ELIMINADOS = "SI"
             
             Sql = " sgpadm_del_DetallePedidoNuevo "
             Sql = Sql & "'" & pedido & "',"
             Sql = Sql & "'" & numlinea & "'"
            
             Set RS = vg_db.Execute(Sql)
        
     End If
   Else
       
       Sql = ""
       Sql = " sgpadm_iu_ActualizaDetallePedidoNuevo "
       Sql = Sql & " '" & Proveedoractual & "',"
       Sql = Sql & " '" & familia & "',"
       Sql = Sql & " '" & codmaterial & "',"
       Sql = Sql & " '" & desmaterial & "',"
       Sql = Sql & " '" & numlinea & "',"
       Sql = Sql & " '" & pedido & "',"
       Sql = Sql & " '" & desunidad & "',"
       Sql = Sql & " '" & perfilredondeo & "',"
       Sql = Sql & " '" & productosgp & "',"
       Sql = Sql & " '" & precio & "',"
       Sql = Sql & " '" & formatoconvenio & "', "
       Sql = Sql & " '" & MinimaPedido & "'"
    
       Set RS = vg_db.Execute(Sql)
       
       '--> Ejecutar Detalle pedidos posteriores
       Sql = ""
       Sql = " sgpadm_Pro_ReprocesarDetPedidosPosteriores "
       Sql = Sql & " '" & ccosto & "',"
       Sql = Sql & " " & pedido & ","
       Sql = Sql & " '" & Ingrediente & "'"
       Set RS = vg_db.Execute(Sql)
       
   End If
       
       Sql = " sgpadm_INS_excepcionformatoPedido "
       Sql = Sql & " '" & ccosto & "',"
       Sql = Sql & " '" & Ingrediente & "',"
       Sql = Sql & " '" & codmaterial & "',"
       Sql = Sql & " '" & Proveedoractual & "'"
       Set RS = vg_db.Execute(Sql)
  
  End If
 
Next i
     Frame2.Visible = False
     MsgBox "Se Actualizo el Detalle del Pedido OK"
     Call Detalle_del_Pedido1

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub Exp_Excel_Saldo_Mayor_FormatoCompras()

On Error GoTo Man_Error

    Dim RS       As New ADODB.Recordset
    Dim Sql      As String
    Dim xlApp    As Object
    Dim xlWb     As Object
    Dim xlWs     As Object
    
    Dim recArray As Variant
    
    Dim strDB           As String
    Dim NomArchivoExcel As String
    Dim Extension       As String
    Dim fldCount        As Integer
    Dim recCount        As Long


    Dim seleccion As Long
    Dim NPedido   As Long
    Dim MyBuffer  As String
    Dim i         As Long
    Dim tipopedido As String
    Dim EstadoPedido As String
    
    
    '-------> Guardar nombre archivo excel
    NomArchivoExcel = ""
    CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    CD.Filter = "Todos los archivos *.xls,*.xlsx"
    CD.ShowSave
    If CD.FileName = "" Then
             MsgBox "Debe seleccionar la ruta y nombre de archivo", vbExclamation
             Exit Sub
    Else
       Extension = ""
       Extension = Right(CD.FileName, Len(CD.FileName) - (InStr(CD.FileName, ".")))
       If UCase(Extension) <> "XLS" And UCase(Extension) <> "XLSX" Then
          MsgBox "La extensión del archivo debe ser (*.xls,*.xlsx)", vbCritical
          Exit Sub
       End If
       NomArchivoExcel = CD.FileName  'Dir(CD.Filename)
    End If
          
    If Dir(NomArchivoExcel) <> "" Then Kill NomArchivoExcel
    
    '-------> Armar xml
    Let MyBuffer = ""
    Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
    Let MyBuffer = MyBuffer & "<NewDataSet>"
    
    For i = 1 To vaSpread1.MaxRows
        vaSpread1.Row = i
        vaSpread1.Col = 1
        seleccion = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> estado pedido
        vaSpread1.Col = 12
        EstadoPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
        '-------> tipo pedido
        vaSpread1.Col = 13
        tipopedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)

        If seleccion = 1 Then
           vaSpread1.Col = 2 'Nº Pedido
           NPedido = IIf(vaSpread1.text = "", 0, vaSpread1.text)
           
           MyBuffer = MyBuffer & " <Np"
           MyBuffer = MyBuffer & " Np = " & Chr(34) & NPedido & Chr(34)
           MyBuffer = MyBuffer & "/>"
        
        End If
    Next i
    MyBuffer = MyBuffer & "</NewDataSet>"

    '-------> Lectura
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    Sql = ""
    Sql = " sgpadm_Sel_XmlExpExcelSaldoMayorFormatoCompra "
    Sql = Sql & " '" & MyBuffer & "'"

    Set RS = vg_db.Execute(Sql)

    '-------> Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Hoja1")
  
    '-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
    xlApp.UserControl = True
    
    '-------> Check version of Excel
    Call encabezado(RS, xlWs)
          
    xlWs.Cells(2, 1).CopyFromRecordset RS
    '-------> Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    
    xlWb.Close True, NomArchivoExcel

    Dim XL As New excel.Application 'Crea el objeto excel
    XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    XL.Visible = True
    XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
    '-------> Close ADO objects
    RS.Close
    Set RS = Nothing
    
    ' -- Cerrar Excel
    xlApp.Quit
    '-------> Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing

    MsgBox "Exportación realizada con exito", vbInformation
    
Exit Sub
Man_Error:
    fg_descarga
    If Err = 438 Or Err = 70 Or Err = 1004 Then
       MsgBox "Hay un Excel Abierto debe cerrarlo para poder bajar el nuevo", vbExclamation
       Exit Sub
    End If
    
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)

End Sub

