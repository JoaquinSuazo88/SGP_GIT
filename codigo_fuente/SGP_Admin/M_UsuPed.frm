VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_UsuPed 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario Pedido"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9630
   ScaleWidth      =   12870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12615
      Begin TabDlg.SSTab SSTab1 
         Height          =   8655
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   12255
         _ExtentX        =   21616
         _ExtentY        =   15266
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Listar Usuarios"
         TabPicture(0)   =   "M_UsuPed.frx":0000
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "vaSpread1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame2"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "Usuarios"
         TabPicture(1)   =   "M_UsuPed.frx":001C
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).ControlCount=   1
         Begin VB.Frame Frame3 
            Height          =   8055
            Left            =   -74760
            TabIndex        =   10
            Top             =   480
            Width           =   11775
            Begin VB.Frame Frame6 
               Caption         =   "Tipo de Usuarios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4335
               Left            =   240
               TabIndex        =   36
               Top             =   3600
               Width           =   5535
               Begin VB.Frame Frame14 
                  Height          =   435
                  Left            =   600
                  TabIndex        =   37
                  Top             =   3840
                  Width           =   4485
                  Begin VB.TextBox TextCan1 
                     Appearance      =   0  'Flat
                     BorderStyle     =   0  'None
                     Height          =   240
                     Index           =   3
                     Left            =   45
                     TabIndex        =   38
                     Top             =   135
                     Width           =   4380
                  End
               End
               Begin FPSpread.vaSpread vaSpread2 
                  Height          =   3600
                  Left            =   240
                  TabIndex        =   21
                  Top             =   240
                  Width           =   5055
                  _Version        =   393216
                  _ExtentX        =   8916
                  _ExtentY        =   6350
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
                  MaxCols         =   3
                  SpreadDesigner  =   "M_UsuPed.frx":0038
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Asignar Contratos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   4335
               Left            =   5880
               TabIndex        =   31
               Top             =   3600
               Width           =   5655
               Begin VB.Frame Frame13 
                  Height          =   435
                  Left            =   240
                  TabIndex        =   34
                  Top             =   3840
                  Width           =   915
                  Begin VB.TextBox TextCai1 
                     Appearance      =   0  'Flat
                     BorderStyle     =   0  'None
                     Height          =   240
                     Index           =   2
                     Left            =   45
                     TabIndex        =   35
                     Top             =   135
                     Width           =   810
                  End
               End
               Begin VB.Frame Frame16 
                  Height          =   435
                  Left            =   1170
                  TabIndex        =   32
                  Top             =   3840
                  Width           =   4005
                  Begin VB.TextBox TextCai1 
                     Appearance      =   0  'Flat
                     BorderStyle     =   0  'None
                     Height          =   240
                     Index           =   3
                     Left            =   45
                     TabIndex        =   33
                     Top             =   135
                     Width           =   3900
                  End
               End
               Begin FPSpread.vaSpread vaSpread3 
                  Height          =   3600
                  Left            =   240
                  TabIndex        =   22
                  Top             =   240
                  Width           =   5175
                  _Version        =   393216
                  _ExtentX        =   9128
                  _ExtentY        =   6350
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
                  MaxCols         =   4
                  OperationMode   =   4
                  SelectBlockOptions=   0
                  SpreadDesigner  =   "M_UsuPed.frx":1907
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Datos Usuarios"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3135
               Left            =   1320
               TabIndex        =   11
               Top             =   240
               Width           =   8775
               Begin VB.CheckBox Check1 
                  Caption         =   "Pedidos Express"
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
                  Left            =   6720
                  TabIndex        =   20
                  Top             =   600
                  Width           =   1815
               End
               Begin EditLib.fpText fpText1 
                  Height          =   315
                  Index           =   0
                  Left            =   240
                  TabIndex        =   12
                  Top             =   600
                  Width           =   1605
                  _Version        =   196608
                  _ExtentX        =   2831
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
                  MaxLength       =   12
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
                  Index           =   2
                  Left            =   240
                  TabIndex        =   13
                  Top             =   1320
                  Width           =   4005
                  _Version        =   196608
                  _ExtentX        =   7064
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
                  MaxLength       =   30
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
                  Index           =   3
                  Left            =   4440
                  TabIndex        =   14
                  Top             =   1320
                  Width           =   4005
                  _Version        =   196608
                  _ExtentX        =   7064
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
                  MaxLength       =   30
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
                  Index           =   4
                  Left            =   240
                  TabIndex        =   15
                  Top             =   2040
                  Width           =   1605
                  _Version        =   196608
                  _ExtentX        =   2831
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
                  MaxLength       =   12
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
                  Index           =   5
                  Left            =   2400
                  TabIndex        =   16
                  Top             =   2040
                  Width           =   1605
                  _Version        =   196608
                  _ExtentX        =   2831
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
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin EditLib.fpText fpText1 
                  Height          =   315
                  Index           =   6
                  Left            =   4680
                  TabIndex        =   17
                  Top             =   2040
                  Width           =   1605
                  _Version        =   196608
                  _ExtentX        =   2831
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
                  MaxLength       =   12
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
                  Index           =   7
                  Left            =   240
                  TabIndex        =   19
                  Top             =   2760
                  Width           =   8205
                  _Version        =   196608
                  _ExtentX        =   14473
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
                  MaxLength       =   100
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
               Begin EditLib.fpLongInteger fpLongInteger1 
                  Height          =   315
                  Index           =   0
                  Left            =   7080
                  TabIndex        =   18
                  Top             =   2040
                  Width           =   645
                  _Version        =   196608
                  _ExtentX        =   1138
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
                  BackColor       =   16777215
                  ForeColor       =   -2147483640
                  ThreeDInsideStyle=   0
                  ThreeDInsideHighlightColor=   -2147483633
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
                  ThreeDTextHighlightColor=   -2147483633
                  ThreeDTextShadowColor=   -2147483632
                  ThreeDTextOffset=   1
                  AlignTextH      =   2
                  AlignTextV      =   0
                  AllowNull       =   -1  'True
                  NoSpecialKeys   =   3
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
                  MaxValue        =   "2147483647"
                  MinValue        =   "-2147483647"
                  NegFormat       =   1
                  NegToggle       =   0   'False
                  Separator       =   ""
                  UseSeparator    =   0   'False
                  IncInt          =   1
                  BorderGrayAreaColor=   -2147483637
                  ThreeDOnFocusInvert=   0   'False
                  ThreeDFrameColor=   -2147483633
                  Appearance      =   1
                  BorderDropShadow=   1
                  BorderDropShadowColor=   -2147483632
                  BorderDropShadowWidth=   3
                  ButtonColor     =   -2147483633
                  AutoMenu        =   0   'False
                  ButtonAlign     =   0
                  OLEDropMode     =   0
                  OLEDragMode     =   0
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Días Tope S/R"
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
                  Left            =   7080
                  TabIndex        =   30
                  Top             =   1800
                  Width           =   1320
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "E-Mail"
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
                  Index           =   8
                  Left            =   240
                  TabIndex        =   29
                  Top             =   2520
                  Width           =   540
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Creado Por"
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
                  Left            =   4680
                  TabIndex        =   28
                  Top             =   1800
                  Width           =   960
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Clave"
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
                  Left            =   2400
                  TabIndex        =   27
                  Top             =   1800
                  Width           =   495
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "RUT"
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
                  Left            =   240
                  TabIndex        =   26
                  Top             =   1800
                  Width           =   405
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Apellidos"
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
                  Left            =   4440
                  TabIndex        =   25
                  Top             =   1065
                  Width           =   780
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Usuario"
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
                  Left            =   240
                  TabIndex        =   24
                  Top             =   315
                  Width           =   660
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Nombre"
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
                  Left            =   240
                  TabIndex        =   23
                  Top             =   1035
                  Width           =   660
               End
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            ForeColor       =   &H80000008&
            Height          =   971
            Left            =   2640
            TabIndex        =   3
            Top             =   480
            Width           =   7335
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "M_UsuPed.frx":3244
               Left            =   2010
               List            =   "M_UsuPed.frx":3257
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   240
               Width           =   2500
            End
            Begin EditLib.fpText fpText1 
               Height          =   315
               Index           =   1
               Left            =   2010
               TabIndex        =   5
               Top             =   555
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
               AllowNull       =   0   'False
               NoSpecialKeys   =   0
               AutoAdvance     =   -1  'True
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Buscar Columna"
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
               Left            =   525
               TabIndex        =   8
               Top             =   345
               Width           =   1380
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
               Left            =   525
               TabIndex        =   7
               Top             =   645
               Width           =   1140
            End
            Begin VB.Label Label2 
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
               Left            =   4590
               TabIndex        =   6
               Top             =   645
               Width           =   585
            End
         End
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   6285
            Left            =   960
            TabIndex        =   9
            Top             =   1560
            Width           =   10245
            _Version        =   393216
            _ExtentX        =   18071
            _ExtentY        =   11086
            _StockProps     =   64
            AllowCellOverflow=   -1  'True
            AutoCalc        =   0   'False
            DisplayRowHeaders=   0   'False
            EditEnterAction =   5
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
            FormulaSync     =   0   'False
            MaxCols         =   5
            ProcessTab      =   -1  'True
            SpreadDesigner  =   "M_UsuPed.frx":3283
            ScrollBarTrack  =   3
            ClipboardOptions=   0
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12870
      _ExtentX        =   22701
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_UsuPed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, MsgTitulo As String
Dim Est As Boolean

Private Sub Check1_Click()
If Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 10125
Me.Width = 13035
MsgTitulo = "Usuario Pedidos"
fg_centra Me
SSTab1.Tab = 0
modo = ""
Est = True
Gl_Mo_Botones Me, 14
Toolbar1.Buttons.item(15).ButtonMenus(1).Visible = False
Toolbar1.Buttons.item(15).ButtonMenus(2).Visible = False
Toolbar1.Buttons.item(15).Visible = False
Gl_Ac_Botones Me, 14, 1, modo
Combo1.ListIndex = 1
MoverDatosGrilla
MoverDatosUsuarios
'MoverDatosListadePreciosCasinoAsignados
Gl_Ac_Botones Me, 14, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverDatosGrilla()
fg_carga ""
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = -1
vaSpread1.Lock = True
Set RS = vg_dbpedweb.Execute("pedweb_s_usuarios 1, '', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = Trim(RS!usuario)
   vaSpread1.Col = 2
   vaSpread1.text = Trim(RS!Nombres)
   vaSpread1.Col = 3
   vaSpread1.text = Trim(RS!apellidos)
   vaSpread1.Col = 4
   vaSpread1.text = fg_PintaRut(RS!rut)
   vaSpread1.Col = 5
   vaSpread1.text = Trim(RS!email)
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
If vaSpread1.MaxRows > 0 Then
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = ""
   codigo = Val(vaSpread1.text)
   vaSpread1.SetActiveCell 1, 1 ': vaSpread1.SetFocus
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
fg_descarga
End Sub

Sub MoverDatosUsuarios()
fg_carga ""
Est = True
Limpia 1
Set RS = vg_dbpedweb.Execute("pedweb_s_usuarios 2, '" & codigo & "', ''")
If Not RS.EOF Then
   fpText1(0).text = Trim(RS!usuario)
   fpText1(2).text = Trim(IIf(IsNull(RS!Nombres), "", RS!Nombres))
   fpText1(3).text = Trim(IIf(IsNull(RS!apellidos), "", RS!apellidos))
   fpText1(4).text = Trim(IIf(IsNull(RS!rut), "", fg_PintaRut(RS!rut)))
   fpText1(5).text = Trim(IIf(IsNull(RS!clave), "", RS!clave))
   fpText1(6).text = Trim(IIf(IsNull(RS!creadopor), "", RS!creadopor))
   fpText1(7).text = Trim(IIf(IsNull(RS!email), "", RS!email))
   fpLongInteger1(0).Value = IIf(IsNull(RS!dtope), 0, RS!dtope)
   Check1.Value = IIf(IsNull(RS!pedexpress) Or RS!pedexpress = 0, 0, 1)
End If
RS.Close: Set RS = Nothing

Limpia 2
If modo <> "A" Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text
End If
vaSpread2.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_tipodeusuarios 1, '" & codigo & "', ''")
Do While Not RS.EOF
   vaSpread2.MaxRows = vaSpread2.MaxRows + 1
   vaSpread2.Row = vaSpread2.MaxRows
   vaSpread2.Col = 1
   vaSpread2.text = IIf(IsNull(RS!tipusu), "0", "1")
   vaSpread2.Col = 2
   vaSpread2.text = IIf(IsNull(RS!tipo_usuario), "", RS!tipo_usuario)
   vaSpread2.Col = 3
   vaSpread2.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread2.Visible = True

vaSpread3.Visible = False
Set RS = vg_dbpedweb.Execute("pedweb_s_clientes 1, '" & codigo & "', ''")
Do While Not RS.EOF
   vaSpread3.MaxRows = vaSpread3.MaxRows + 1
   vaSpread3.Row = vaSpread3.MaxRows
   If RS!cco = 1 Then vaSpread3.SelModeSelected = True
   vaSpread3.Col = 2
   vaSpread3.text = IIf(IsNull(RS!centrocosto), "", RS!centrocosto)
   vaSpread3.Col = 3
   vaSpread3.text = Trim(IIf(IsNull(RS!Nombre), "", RS!Nombre))
   vaSpread3.Col = 4
   vaSpread3.text = Trim(IIf(IsNull(RS!codigo), "", RS!codigo))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread3.Visible = True
Est = False
fg_descarga
End Sub

Sub Limpia(Op As Integer)
Select Case Op
Case 1
   fpText1(0).text = ""
   fpText1(0).Enabled = IIf(modo = "A", True, False)
   fpText1(2).text = ""
   fpText1(3).text = ""
   fpText1(4).text = ""
   fpText1(5).text = ""
   fpText1(6).text = Trim(LimpiaDato(vg_NUsr))
   fpText1(6).Enabled = False
   fpText1(7).text = ""
   fpLongInteger1(0).Value = 0
   Check1.Value = 0
Case 2
    vaSpread2.MaxRows = 0
    vaSpread3.MaxRows = 0
End Select
End Sub

Private Sub fpLongInteger1_Change(Index As Integer)
If Est Then Exit Sub
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 14, 0, modo
End Sub

Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
'If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_Change(Index As Integer)
Select Case Index
Case 0, 2, 3, 4, 5, 6, 7
    If Est Then Exit Sub
    If modo = "" Then modo = "M"
    Gl_Ac_Botones Me, 14, 0, modo
Case 1
    If LimpiaDato(Trim(fpText1(1).text)) & Chr(KeyAscii) = "" Then Exit Sub
    vaSpread1.Visible = False
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_usuarios 3, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_usuarios 4, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 2 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_usuarios 5, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 3 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_usuarios 6, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 4 Then
       Set RS2 = vg_dbpedweb.Execute("pedweb_s_usuarios 7, '', '%" & UCase(LimpiaDato(fpText1(1).text)) & "%'")
    End If
    If RS2.EOF Then vaSpread1.MaxRows = 0 Else vaSpread1.MaxRows = RS2!nReg
    i = 1
    If Not RS2.EOF Then
       Do While Not RS2.EOF
          vaSpread1.Row = i: i = i + 1
          vaSpread1.Col = 1
          vaSpread1.TypeHAlign = 1
          vaSpread1.text = Trim(RS2!usuario)
          vaSpread1.Col = 2
          vaSpread1.text = Trim(RS2!Nombres)
          vaSpread1.Col = 3
          vaSpread1.text = Trim(RS2!apellidos)
          vaSpread1.Col = 4
          vaSpread1.text = fg_PintaRut(RS2!rut)
          vaSpread1.Col = 5
          vaSpread1.text = Trim(RS2!email)
          RS2.MoveNext
        Loop
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        Gl_Ac_Botones Me, 14, 1, modo
    Else
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = False
    End If
    RS2.Close: Set RS2 = Nothing
    vaSpread1.Col = 1: vaSpread1.col2 = vaSpread1.maxcols: vaSpread1.Row = 1: vaSpread1.row2 = vaSpread1.MaxRows
    vaSpread1.SetActiveCell 1, 1
    vaSpread1.Visible = True
    If fpText1(1).text = "" Then Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro" Else Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Reg. Enc."
End Select

End Sub

Private Sub fpText1_GotFocus(Index As Integer)
Select Case Index
Case 4
    If Trim(fpText1(4).text) = "" Or vg_Dig = "N" Then Exit Sub
    Est = True
    fpText1(4).text = fg_DespintaRut(fpText1(4).text)
    fpText1(4).text = Mid(fpText1(4).text, 1, Len(Trim(fpText1(4).text)) - 1)
    Est = False
End Select
End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)
If Index = 0 Then KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
End Sub

Private Sub fpText1_LostFocus(Index As Integer)
Select Case Index
Case 4
    Est = True
    fpText1(Index).text = UCase(fpText1(Index).text)
    If Trim(fpText1(4).text) = "" Or vg_Dig = "N" Then Exit Sub
    fpText1(4).text = fg_RutDig(Trim(fpText1(4).text))
    fpText1(4).text = fg_PintaRut(fpText1(4).text)
    Est = False
End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
If modo <> "A" Then
   vaSpread1.Row = vaSpread1.ActiveRow
   vaSpread1.Col = 1
   codigo = vaSpread1.text
End If
Select Case SSTab1.Tab
Case 1
    MoverDatosUsuarios
End Select
End Sub

Private Sub TextCai1_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3, 4
    vaSpread3.Visible = False
    If Trim(TextCai1(Index).text) <> "" Then
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           vaSpread3.Col = Index: nom = UCase(Trim(vaSpread3.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCai1(Index).text) & "*"
           vaSpread3.Col = 2
           If indactivo = -1 And Trim(vaSpread3.text) <> "" Then
              If vaSpread3.RowHidden = True Then vaSpread3.RowHidden = False
           Else
              If vaSpread3.RowHidden = False Then vaSpread3.RowHidden = True
           End If
        Next i
        vaSpread3.SetActiveCell Index, 1
    End If
    vaSpread3.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread3.ColUserSortIndicator(IIf(Trim(TextCai1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread3.SortKey(1) = IIf(Trim(TextCai1(Index).text) = "", 0, 0): vaSpread3.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread3.Sort -1, -1, vaSpread3.maxcols, vaSpread3.MaxRows, SortByRow
    If Trim(TextCai1(Index).text) = "" Then
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           If vaSpread3.RowHidden = True Then vaSpread3.RowHidden = False
       Next
       vaSpread3.SetActiveCell Index, vaSpread3.SearchCol(Index, 0, vaSpread3.MaxRows, Trim(TextCai1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread3.SetActiveCell Index, 1
    End If
    vaSpread3.Visible = True
End Select
End Sub

Private Sub TextCan1_Change(Index As Integer)
Dim i As Long, nom As String
Select Case Index
Case 2, 3, 4
    vaSpread2.Visible = False
    If Trim(TextCan1(Index).text) <> "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = Index: nom = UCase(Trim(vaSpread2.text))
           indactivo = UCase(Trim(nom)) Like "*" & UCase(TextCan1(Index).text) & "*"
           vaSpread2.Col = 2
           If indactivo = -1 And Trim(vaSpread2.text) <> "" Then
              If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
           Else
              If vaSpread2.RowHidden = False Then vaSpread2.RowHidden = True
           End If
        Next i
        vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
    vaSpread2.ColUserSortIndicator(IIf(Trim(TextCan1(Index).text) = "", 0, 0)) = ColUserSortIndicatorAscending
    vaSpread2.SortKey(1) = IIf(Trim(TextCan1(Index).text) = "", 0, 0): vaSpread2.SortKeyOrder(1) = SortKeyOrderAscending
    vaSpread2.Sort -1, -1, vaSpread2.maxcols, vaSpread2.MaxRows, SortByRow
    If Trim(TextCan1(Index).text) = "" Then
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           If vaSpread2.RowHidden = True Then vaSpread2.RowHidden = False
       Next
       vaSpread2.SetActiveCell Index, vaSpread2.SearchCol(Index, 0, vaSpread2.MaxRows, Trim(TextCan1(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread2.SetActiveCell Index, 1
    End If
    vaSpread2.Visible = True
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1 '-------> Incluir nuevos registros
    modo = "A"
    codigo = ""
    MoverDatosUsuarios
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    '-------> Traer ultimo registro
    Limpia 1
    fpText1(0).SetFocus
    vg_codigo = "x"
    If vg_codigo <> "" Then Gl_Ac_Botones Me, 14, 0, modo
Case 3 '-------> Alterar registro
    modo = "M"
    Gl_Ac_Botones Me, 14, 0, modo
    SSTab1.Tab = 1
    SSTab1.TabEnabled(0) = False
    SSTab1.TabEnabled(1) = True
    fpText1(2).SetFocus
Case 5 '-------> Eliminar Registro
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
    If MsgBox("Elimina registro y todas sus relaciones...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    '-------> borrar ruta
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    vg_dbpedweb.Execute ("pedweb_d_usuarios " & codigo & "")
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    codigo = 0
    If vaSpread1.MaxRows > 0 Then
       vaSpread1.SetActiveCell 1, vaSpread1.MaxRows
       vaSpread1.Row = vaSpread1.ActiveRow
       vaSpread1.Col = 1
       codigo = vaSpread1.text
    End If
    MoverDatosUsuarios
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 7 '-------> Actualizar lista
    Select Case SSTab1.Tab
    Case 0
        fpText1(1).text = ""
        MoverDatosGrilla
    Case 1
        MoverDatosUsuarios
    End Select
Case 10 '-------> Cancelar Información
    If MsgBox("Cancela...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
    SSTab1.Tab = 1
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1
    codigo = vaSpread1.text
    MoverDatosUsuarios
    '-------> Desbloquear hojas
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 12 '-------> grabaRegistro
    Dim estmar As Boolean, i As Long
    v_rut = fg_DespintaRut(fpText1(4).text)
    If LimpiaDato(Trim(fpText1(0).text)) = "" Or LimpiaDato(Trim(fpText1(2).text)) = "" Or LimpiaDato(Trim(fpText1(3).text)) = "" _
       Or LimpiaDato(Trim(fpText1(5).text)) = "" Or LimpiaDato(Trim(fpText1(6).text)) = "" Or LimpiaDato(Trim(fpText1(7).text)) = "" _
       Or fpLongInteger1(0).Value = 0 Then MsgBox "Debe ingresar información...", vbCritical, MsgTitulo: Exit Sub
    If Not fg_Check_Rut(v_rut) Then MsgBox "El rut no es valido...", vbExclamation + vbOKOnly, "Valida rut": Exit Sub
    estmar = False
    For i = 1 To vaSpread2.MaxRows
        vaSpread2.Row = i
        vaSpread2.Col = 1
        If vaSpread2.text = "1" Then estmar = True
    Next i
    If Not estmar Then MsgBox "Debe seleccinar a lo menos un tipo usuarios...", vbCritical, MsgTitulo: Exit Sub
    estmar = False
    For i = 1 To vaSpread3.MaxRows
        vaSpread3.Row = i
        If vaSpread3.SelModeSelected = True Then estmar = True
    Next i
    If Not estmar Then MsgBox "Debe seleccinar a lo menos un contrato...", vbCritical, MsgTitulo: Exit Sub
    If modo = "A" Then
       codigo = 0
       Set RS = vg_dbpedweb.Execute("pedweb_iu_usuarios 'A', '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & LimpiaDato(Trim(fpText1(2).text)) & "', '" & LimpiaDato(Trim(fpText1(3).text)) & "', '" & v_rut & "', '" & LimpiaDato(Trim(fpText1(5).text)) & "', '" & LimpiaDato(Trim(fpText1(6).text)) & "', '" & LimpiaDato(Trim(fpText1(7).text)) & "', " & fpLongInteger1(0).Value & ", " & IIf(Check1.Value = 1, 1, 0) & "")
       If Not RS.EOF Then
          codigo = RS!indice
          For i = 1 To vaSpread2.MaxRows
              vaSpread2.Row = i
              vaSpread2.Col = 1
              If vaSpread2.text = "1" Then
                 vaSpread2.Col = 2
                 vg_dbpedweb.Execute ("pedweb_iu_tipousuarios '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & Trim(LimpiaDato(vaSpread2.text)) & "'")
              End If
          Next i
          For i = 1 To vaSpread3.MaxRows
             vaSpread3.Row = i
             If vaSpread3.SelModeSelected = True Then
                vaSpread3.Col = 4
                vg_dbpedweb.Execute ("pedweb_iu_clienteusuarios '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & Trim(LimpiaDato(vaSpread3.text)) & "'")
             End If
         Next i
       End If
       RS.Close: Set RS = Nothing
       vaSpread1.MaxRows = vaSpread1.MaxRows + 1: vaSpread1.Row = vaSpread1.MaxRows
       vaSpread1.SetActiveCell 1, vaSpread1.Row
    Else
       vg_dbpedweb.Execute ("pedweb_iu_usuarios 'M', '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & LimpiaDato(Trim(fpText1(2).text)) & "', '" & LimpiaDato(Trim(fpText1(3).text)) & "', '" & v_rut & "', '" & LimpiaDato(Trim(fpText1(5).text)) & "', '" & LimpiaDato(Trim(fpText1(6).text)) & "', '" & LimpiaDato(Trim(fpText1(7).text)) & "', " & fpLongInteger1(0).Value & ", " & IIf(Check1.Value = 1, 1, 0) & "")
       
       vg_dbpedweb.Execute "pedweb_d_tipousuarios 1, '" & LimpiaDato(Trim(fpText1(0).text)) & "'"
       For i = 1 To vaSpread2.MaxRows
           vaSpread2.Row = i
           vaSpread2.Col = 1
           If vaSpread2.text = "1" Then
              vaSpread2.Col = 2
              vg_dbpedweb.Execute ("pedweb_iu_tipousuarios '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & Trim(LimpiaDato(vaSpread2.text)) & "'")
           End If
       Next i
       
       vg_dbpedweb.Execute "pedweb_d_clienteusuarios '" & LimpiaDato(Trim(fpText1(0).text)) & "'"
       For i = 1 To vaSpread3.MaxRows
           vaSpread3.Row = i
           If vaSpread3.SelModeSelected = True Then
              vaSpread3.Col = 4
              vg_dbpedweb.Execute ("pedweb_iu_clienteusuarios '" & LimpiaDato(Trim(fpText1(0).text)) & "', '" & Trim(LimpiaDato(vaSpread3.text)) & "'")
           End If
       Next i
    End If
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(0).text))
    vaSpread1.Col = 2: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(2).text))
    vaSpread1.Col = 3: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(3).text))
    vaSpread1.Col = 4: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = fg_PintaRut(v_rut)
    vaSpread1.Col = 5: vaSpread1.TypeHAlign = TypeHAlignLeft: vaSpread1.text = LimpiaDato(Trim(fpText1(7).text))
    SSTab1.TabEnabled(0) = True
    SSTab1.TabEnabled(1) = True
    modo = "": Gl_Ac_Botones Me, 14, 1, modo
Case 19 '------> impresion
    I_UsuariosWeb
Case 22
    Me.Hide
    Unload Me
End Select
End Sub
