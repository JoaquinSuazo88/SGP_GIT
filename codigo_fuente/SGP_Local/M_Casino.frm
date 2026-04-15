VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_Casino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Contrato"
   ClientHeight    =   8265
   ClientLeft      =   2655
   ClientTop       =   2505
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7785
      Left            =   0
      TabIndex        =   4
      Top             =   375
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   13732
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabMaxWidth     =   4
      OLEDropMode     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Casinos"
      TabPicture(0)   =   "M_Casino.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame2"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "M_Casino.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame3 
         Height          =   7095
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   7575
         Begin VB.OptionButton Option1 
            Caption         =   "Traspaso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   6120
            TabIndex        =   35
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Operación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   4080
            TabIndex        =   34
            Top             =   600
            Width           =   1215
         End
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   1200
            Width           =   7125
            _Version        =   196608
            _ExtentX        =   12568
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   50
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   12
            Top             =   1800
            Width           =   7125
            _Version        =   196608
            _ExtentX        =   12568
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   50
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   7
            Left            =   5085
            TabIndex        =   13
            Top             =   3000
            Width           =   2265
            _Version        =   196608
            _ExtentX        =   3995
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   15
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   4
            Left            =   4155
            TabIndex        =   14
            Top             =   2400
            Width           =   3180
            _Version        =   196608
            _ExtentX        =   5609
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   15
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   2400
            Width           =   3180
            _Version        =   196608
            _ExtentX        =   5609
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   15
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   16
            Top             =   600
            Width           =   1740
            _Version        =   196608
            _ExtentX        =   3069
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   10
            Left            =   240
            TabIndex        =   17
            Top             =   4800
            Width           =   7125
            _Version        =   196608
            _ExtentX        =   12568
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   50
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   5
            Left            =   240
            TabIndex        =   18
            Top             =   3000
            Width           =   2265
            _Version        =   196608
            _ExtentX        =   3995
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   15
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   8
            Left            =   240
            TabIndex        =   19
            Top             =   3600
            Width           =   7125
            _Version        =   196608
            _ExtentX        =   12568
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   50
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   9
            Left            =   240
            TabIndex        =   20
            Top             =   4200
            Width           =   7125
            _Version        =   196608
            _ExtentX        =   12568
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   50
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   6
            Left            =   2655
            TabIndex        =   21
            Top             =   3000
            Width           =   2265
            _Version        =   196608
            _ExtentX        =   3995
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   15
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   0
            Left            =   240
            TabIndex        =   36
            Top             =   5415
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   1
            Left            =   3960
            TabIndex        =   39
            Top             =   5415
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
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
         Begin EditLib.fpLongInteger fpLongInteger1 
            Height          =   315
            Index           =   2
            Left            =   240
            TabIndex        =   43
            Top             =   6015
            Width           =   555
            _Version        =   196608
            _ExtentX        =   979
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   11
            Left            =   3960
            TabIndex        =   47
            Top             =   6000
            Width           =   1260
            _Version        =   196608
            _ExtentX        =   2222
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            MaxLength       =   4
            MultiLine       =   0   'False
            PasswordChar    =   ""
            IncHoriz        =   0.25
            BorderGrayAreaColor=   -2147483637
            NoPrefix        =   0   'False
            ScrollV         =   0   'False
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
         Begin EditLib.fpText fpText 
            Height          =   315
            Index           =   12
            Left            =   240
            TabIndex        =   50
            Top             =   6600
            Width           =   3180
            _Version        =   196608
            _ExtentX        =   5609
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
            AlignTextH      =   0
            AlignTextV      =   2
            AllowNull       =   0   'False
            NoSpecialKeys   =   3
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
            Caption         =   "Resolución"
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
            Left            =   240
            TabIndex        =   49
            Top             =   6360
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sociedad SAP"
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
            Left            =   3960
            TabIndex        =   48
            Top             =   5760
            Width           =   1230
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Segmento"
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
            Left            =   240
            TabIndex        =   45
            Top             =   5760
            Width           =   855
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   2
            Left            =   720
            Picture         =   "M_Casino.frx":0038
            Top             =   5910
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   3
            Left            =   1170
            TabIndex        =   44
            Top             =   6015
            Width           =   2445
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Servicio"
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
            Index           =   14
            Left            =   3960
            TabIndex        =   41
            Top             =   5160
            Width           =   1140
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   1
            Left            =   4440
            Picture         =   "M_Casino.frx":0342
            Top             =   5310
            Width           =   480
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   1
            Left            =   4890
            TabIndex        =   40
            Top             =   5415
            Width           =   2445
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   1170
            TabIndex        =   37
            Top             =   5415
            Width           =   2445
         End
         Begin VB.Image Image1 
            Height          =   480
            Index           =   0
            Left            =   720
            Picture         =   "M_Casino.frx":064C
            Top             =   5310
            Width           =   480
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bodega"
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
            Index           =   13
            Left            =   240
            TabIndex        =   33
            Top             =   5160
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            Left            =   240
            TabIndex        =   32
            Top             =   360
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fono Nº 1"
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
            Left            =   240
            TabIndex        =   31
            Top             =   2760
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ciudad"
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
            Left            =   4155
            TabIndex        =   30
            Top             =   2160
            Width           =   600
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Comuna"
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
            Left            =   240
            TabIndex        =   29
            Top             =   2160
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Dirección"
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
            TabIndex        =   28
            Top             =   1560
            Width           =   825
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
            Index           =   2
            Left            =   240
            TabIndex        =   27
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fono Nº 2"
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
            Left            =   2655
            TabIndex        =   26
            Top             =   2760
            Width           =   870
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
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
            Left            =   5085
            TabIndex        =   25
            Top             =   2760
            Width           =   315
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Persona de Contactos"
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
            Index           =   9
            Left            =   240
            TabIndex        =   24
            Top             =   3360
            Width           =   1890
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Giro"
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
            Left            =   240
            TabIndex        =   23
            Top             =   3960
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Email"
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
            Left            =   240
            TabIndex        =   22
            Top             =   4560
            Width           =   465
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   6
            Left            =   1185
            TabIndex        =   38
            Top             =   5460
            Width           =   2475
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   2
            Left            =   4905
            TabIndex        =   42
            Top             =   5460
            Width           =   2475
         End
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   4
            Left            =   1185
            TabIndex        =   46
            Top             =   6060
            Width           =   2475
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   -74640
         TabIndex        =   6
         Top             =   600
         Width           =   6615
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "M_Casino.frx":0956
            Left            =   1680
            List            =   "M_Casino.frx":0960
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   240
            Width           =   2865
         End
         Begin EditLib.fpText fptnombre 
            Height          =   315
            Left            =   1680
            TabIndex        =   1
            Top             =   600
            Width           =   2895
            _Version        =   196608
            _ExtentX        =   5106
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
            NoSpecialKeys   =   3
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
            Height          =   315
            Index           =   0
            Left            =   150
            TabIndex        =   9
            Top             =   675
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "B"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   4680
            TabIndex        =   8
            Top             =   675
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Label Label1 
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
            Height          =   315
            Index           =   11
            Left            =   150
            TabIndex        =   7
            Top             =   345
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   5055
         Left            =   -74880
         TabIndex        =   5
         Top             =   1800
         Width           =   7185
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   4610
            Left            =   240
            TabIndex        =   2
            Top             =   240
            Width           =   6765
            _Version        =   393216
            _ExtentX        =   11933
            _ExtentY        =   8132
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
            MaxRows         =   20
            OperationMode   =   3
            ScrollBars      =   2
            SelectBlockOptions=   0
            SpreadDesigner  =   "M_Casino.frx":0974
            VisibleCols     =   2
            VisibleRows     =   15
            ScrollBarTrack  =   1
         End
      End
   End
End
Attribute VB_Name = "M_Casino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ibusca As Long, i As Long
Dim itab As Integer, est As Boolean
Dim modo As String, codigo As String

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Me.HelpContextID = vg_OpcM
Me.Height = 8700
Me.Width = 8025
est = True
MsgTitulo = "Contrato"
fg_centra Me
SSTab1.Tab = 0
modo = ""
Combo1.ListIndex = 1
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
MoverDatosGrilla
est = False

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Form_Resize()

On Error GoTo Man_Error

If Me.WindowState <> 1 Then SSTab1.Move 0, Toolbar1.Height, ScaleWidth, ScaleHeight - Toolbar1.Height

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpLongInteger1_Change(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

If est Then Exit Sub
itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Select Case Index

Case 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open "SELECT a.* FROM a_bodega a, b_clientes b WHERE (a.bod_codigo NOT IN (SELECT DISTINCT cli_codbod FROM b_clientes WHERE (cli_codbod) Is Not Null  AND cli_codbod>0)  OR a.bod_codigo=b.cli_codbod) AND a.bod_codigo=" & Val(fpLongInteger1(0).Value) & " AND b.cli_codigo='" & Trim(fpText(0).text) & "'", vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(0).Caption = "": Exit Sub
    fpayuda(0).Caption = Trim(RS!bod_nombre)
    RS.Close: Set RS = Nothing

Case 1
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.TipoServicio(2, Val(fpLongInteger1(1).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(1).Caption = "": Exit Sub
    fpayuda(1).Caption = Trim(RS!tis_nombre)
    RS.Close: Set RS = Nothing

Case 2
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Segmento(2, Val(fpLongInteger1(2).Value), ""), vg_db, adOpenStatic
    If RS.EOF Then RS.Close: Set RS = Nothing: fpayuda(2).Caption = "": Exit Sub
    fpayuda(2).Caption = Trim(RS!seg_nombre)
    RS.Close: Set RS = Nothing

End Select

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpText_Change(Index As Integer)

On Error GoTo Man_Error

If est Then Exit Sub
itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpText_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Or est Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub fpTnombre_Change()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

If LimpiaDato(Trim(fptnombre.text)) & Chr(KeyAscii) = "" Then Exit Sub

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

If Combo1.ItemData(Combo1.ListIndex) = 0 Then
   
   RS.Open RutinaLectura.Cliente(6, "", UCase(LimpiaDato(fptnombre.text))), vg_db, adOpenStatic

ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
   
   RS.Open RutinaLectura.Cliente(7, "", UCase(LimpiaDato(fptnombre.text))), vg_db, adOpenStatic

End If

With vaSpread1
    
    ibusca = RS.RecordCount: .MaxRows = RS.RecordCount: i = 1
    
    If Not RS.EOF Then
       Do While Not RS.EOF
          
          .Row = i
          i = i + 1
          .Col = 1: .text = RS!cli_codigo
          .Col = 2: .TypeHAlign = TypeHAlignLeft: .text = Trim(RS!cli_nombre)
          .Col = 3: .text = IIf(RS!cli_tipo = 0, "Operación", "Traspaso")
          
          RS.MoveNext
       
       Loop
       SSTab1.TabEnabled(1) = True
       modo = ""
       Gl_Ac_Botones Me, 1, 1, modo
    
    Else
       
       SSTab1.TabEnabled(1) = False
    
    End If
    RS.Close: Set RS = Nothing
    Label1(1).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"

End With

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 0
    
    If Option1(1).Value = True Then Exit Sub
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_bodega", Trim(fpText(0).text), "Bodega", "Bod"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(0).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre

Case 1
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_tiposervicio", "tis_", "Tipo Servicio", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(1).Caption = vg_nombre

Case 2
    
    vg_left = fpayuda(1).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "a_segmento", "seg_", "Segmento", "Gen"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(2).Value = Val(vg_codigo)
    fpayuda(2).Caption = vg_nombre

End Select

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

If est Then Exit Sub
itab = 1
SSTab1.Tab = 1
SSTab1.TabEnabled(1) = True
Gl_Ac_Botones Me, 1, 0, modo
itab = 0
fpLongInteger1(0).Value = ""
fpayuda(0).Caption = ""
Select Case Index

Case 0
    
    fpLongInteger1(0).Enabled = True
    Image1(0).Enabled = True

Case 1
    
    fpLongInteger1(0).Enabled = False
    Image1(0).Enabled = False

End Select

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

On Error GoTo Man_Error

With SSTab1
    
    Select Case .Tab
    
    Case 0
        
        Frame3.Enabled = False
        Combo1.Enabled = True: fptnombre.Enabled = True
    
    Case 1
        
        If vaSpread1.MaxRows > 0 And itab = 0 Then
           
           modo = "M"
           .TabEnabled(0) = True
           .Tab = 1
           .TabEnabled(1) = True
           est = True
           MoverDatos
           est = False
        
        End If
    
    End Select

End With

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index

Case 1
    
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False: itab = 1
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    est = True
    Frame3.Enabled = True
    fpLongInteger1(0).text = ""
    fpLongInteger1(1).text = ""
    fpLongInteger1(2).text = ""
    fpayuda(0).Caption = ""
    fpayuda(1).Caption = ""
    fpayuda(3).Caption = ""
    For i = 0 To 11
        If i < 11 Then fpText(i).text = ""
    Next i
    fpText(0).Enabled = True
    Option1(0).Enabled = True: Option1(1).Enabled = True
    Option1(0).Value = False: Option1(1).Value = True: est = False: itab = 0

Case 3
    
    If vaSpread1.MaxRows < 1 Then Exit Sub
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(0) = False: itab = 1
    SSTab1.Tab = 1: SSTab1.TabEnabled(1) = True
    est = True
    MoverDatos
    est = False: itab = 0

Case 5
    
    Borra_Datos

Case 7
    
    modo = ""
    SSTab1.Tab = 0
    MoverDatosGrilla

Case 10
    
    Cancela_Datos

Case 12
    
    Actualiza_Datos

Case 15
    
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, , MsgTitulo: Exit Sub
    I_Contratos

Case 18
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatosGrilla()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

fg_carga ""
With vaSpread1
    
    .MaxRows = 0
    itab = 0
    
    If RS.State = 1 Then RS.Close
    RS.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    RS.Open RutinaLectura.Cliente(3, "", ""), vg_db, adOpenStatic
    
    If Not RS.EOF Then
       
       Do While Not RS.EOF
          
          .MaxRows = .MaxRows + 1
          .Row = .MaxRows
                  
          .Col = 1
          .text = RS!cli_codigo
    
          .Col = 2
          .TypeHAlign = TypeHAlignLeft
          .text = Trim(RS!cli_nombre)
          
          .Col = 3
          .text = IIf(RS!cli_tipo = 0, "Operación", "Traspaso")
                 
          RS.MoveNext
       Loop
       SSTab1.TabEnabled(1) = True
    
    Else
       
       SSTab1.Tab = 0
       SSTab1.TabEnabled(1) = False
       modo = "NE"
       Gl_Ac_Botones Me, 1, 2, modo
    
    End If
    RS.Close: Set RS = Nothing
    Label1(1).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
    fptnombre.text = ""

End With

fg_descarga
Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub MoverDatos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

fg_carga ""
Combo1.Enabled = False: fptnombre.Enabled = False
Frame3.Enabled = True
fpLongInteger1(0).text = ""
fpayuda(0).Caption = ""

For i = 0 To 12
    
    If i < 12 Then fpText(i).text = ""

Next i

est = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
fpText(0).Enabled = False

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open RutinaLectura.Cliente(5, codigo, ""), vg_db, adOpenStatic
If Not RS.EOF Then
   fpText(0).text = RS!cli_codigo
   fpText(1).text = IIf(IsNull(RS!cli_nombre), "", Trim(RS!cli_nombre))
   fpText(2).text = IIf(IsNull(RS!cli_direccion), "", Trim(RS!cli_direccion))
   fpText(3).text = IIf(IsNull(RS!cli_comuna), "", Trim(RS!cli_comuna))
   fpText(4).text = IIf(IsNull(RS!cli_ciudad), "", Trim(RS!cli_ciudad))
   fpText(5).text = IIf(IsNull(RS!cli_fono1), "", Trim(RS!cli_fono1))
   fpText(6).text = IIf(IsNull(RS!cli_fono2), "", Trim(RS!cli_fono2))
   fpText(7).text = IIf(IsNull(RS!cli_fax), "", Trim(RS!cli_fax))
   fpText(8).text = IIf(IsNull(RS!cli_percon), "", Trim(RS!cli_percon))
   fpText(9).text = IIf(IsNull(RS!cli_giro), "", Trim(RS!cli_giro))
   fpText(10).text = IIf(IsNull(RS!cli_email), "", Trim(RS!cli_email))
   fpText(11).text = IIf(IsNull(RS!cli_socsap), "", Trim(RS!cli_socsap))
   fpText(12).text = IIf(IsNull(RS!cli_resolucion), "", Trim(RS!cli_resolucion))
   Image1(0).Enabled = IIf(RS!cli_tipo = 2, False, True)
   fpLongInteger1(0).Enabled = IIf(RS!cli_tipo = 2, False, True)
   If RS!cli_tipo = 2 Then
      
      Option1(0).Value = False
      Option1(1).Value = True
   
   ElseIf RS!cli_tipo = 0 Then
      
      Option1(0).Value = True
      Option1(1).Value = False
   
   End If
   fpLongInteger1(0).Value = ""
   fpayuda(0).Caption = ""
   '------- traer bodega asignada
   If IsNull(RS!cli_codbod) Or RS!cli_codbod = 0 Or RS!cli_tipo = 2 Then
      
      fpLongInteger1(0).Enabled = IIf(RS!cli_tipo = 2, False, True)
      Image1(0).Enabled = IIf(RS!cli_tipo = 2, False, True)
   
   Else
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient
      
      RS1.Open RutinaLectura.Bodega(3, RS!cli_codbod, ""), vg_db, adOpenStatic
      If Not RS1.EOF Then
         
         fpLongInteger1(0).Enabled = IIf(ValidarBodega(RS!cli_codbod, 0), False, True)
         fpLongInteger1(0).Value = RS1!bod_codigo
         fpayuda(0).Caption = Trim(RS1!bod_nombre)
         Image1(0).Enabled = IIf(ValidarBodega(RS!cli_codbod, 0), False, True)
         Option1(0).Enabled = IIf(ValidarBodega(RS!cli_codbod, 0), False, True)
         Option1(1).Enabled = IIf(ValidarBodega(RS!cli_codbod, 0), False, True)
      
      End If
      RS1.Close: Set RS1 = Nothing
   
   End If
   '------- traer tipo servicio
   If IsNull(RS!cli_codtis) Or RS!cli_codtis = 0 Then
      
      fpLongInteger1(1).Value = "": fpayuda(1).Caption = ""
   
   Else
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      RS1.Open RutinaLectura.TipoServicio(2, RS!cli_codtis, ""), vg_db, adOpenStatic
      If Not RS1.EOF Then
         
         fpLongInteger1(1).Value = RS1!tis_codigo
         fpayuda(1).Caption = Trim(RS1!tis_nombre)
      
      End If
      RS1.Close: Set RS1 = Nothing
   
   End If
   '------- traer segmento
   If IsNull(RS!cli_codseg) Or RS!cli_codseg = 0 Then
      fpLongInteger1(2).Value = "": fpayuda(2).Caption = ""
   
   Else
      
      If RS1.State = 1 Then RS1.Close
      RS1.CursorLocation = adUseClient
      vg_db.CursorLocation = adUseClient

      RS1.Open RutinaLectura.Segmento(2, RS!cli_codseg, ""), vg_db, adOpenStatic
      If Not RS1.EOF Then
         
         fpLongInteger1(2).Value = RS1!seg_codigo
         fpayuda(2).Caption = Trim(RS1!seg_nombre)
      
      End If
      RS1.Close: Set RS1 = Nothing
   
   End If
   fpText(11).text = IIf(IsNull(RS!cli_socsap), "", Trim(RS!cli_socsap))

End If
RS.Close: Set RS = Nothing
fg_descarga

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Borra_Datos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim codigo As String

With vaSpread1
    
    If .MaxRows < 1 Then Exit Sub
    .Row = .ActiveRow
    .Col = 1: codigo = .text

End With

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open RutinaLectura.Cliente(5, codigo, ""), vg_db, adOpenStatic
If Not RS.EOF Then
   
   If RS!cli_codbod > 0 Or RS!cli_tipo = 0 Then
      
      If ValidarBodega(RS!cli_codbod, 0) Then RS.Close: Set RS = Nothing: MsgBox "Existen documento asociado, para este contrato...", vbCritical, "Error": Exit Sub
   
   ElseIf RS!cli_tipo = 2 Then
      
      RS1.Open "SELECT DISTINCT * FROM b_totventas WHERE tov_codcas='" & codigo & "'", vg_db, adOpenStatic
      
      If Not RS1.EOF Then RS.Close: Set RS = Nothing: RS1.Close: Set RS1 = Nothing: MsgBox "Existen documento asociado, para este contrato...", vbCritical, "Error": Exit Sub
      RS1.Close: Set RS1 = Nothing
   
   End If

End If
RS.Close: Set RS = Nothing

If MsgBox("Elimina registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

vg_db.Execute "DELETE b_costopatron FROM b_costopatron WHERE cpa_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_paramdesp FROM b_paramdesp WHERE pad_cencos = '" & codigo & "'"
vg_db.Execute "DELETE a_serviciorac FROM a_serviciorac WHERE sra_cencos = '" & codigo & "'"
vg_db.Execute "DELETE a_estservicio FROM a_estservicio WHERE ess_cencos = '" & codigo & "'"
vg_db.Execute "DELETE a_param FROM a_param WHERE par_cencos = '" & codigo & "'"
vg_db.Execute "DELETE a_infcfcfofi FROM a_infcfcfofi WHERE inf_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_cierreperiodo FROM b_cierreperiodo WHERE cie_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_contlistpreing FROM b_contlistpreing WHERE cpi_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_casinotipoactividades FROM b_casinotipoactividades WHERE cta_cencos = '" & codigo & "'"
vg_db.Execute "DELETE b_clientes FROM b_clientes WHERE cli_codigo = '" & codigo & "' AND (cli_tipo = 0 OR cli_tipo = 2)"
vg_db.Execute "DELETE b_recetadet FROM b_recetadet WHERE red_cencos = '" & codigo & "'"

With vaSpread1
    
    .Row = .ActiveRow
    .DeleteRows .Row, 1
    .MaxRows = .MaxRows - 1
    .Row = .MaxRows
    Label1(1).Caption = Format(.MaxRows, fg_Pict(7, 0)) & " Registros"
    
    If .MaxRows < 1 Then
       
       SSTab1.TabEnabled(1) = False
       SSTab1.Tab = 0
       modo = "NE"
    
    Else
       
       modo = ""
       SSTab1.TabEnabled(1) = True
       SSTab1.Tab = 0
    
    End If

End With

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
If Err = 3034 Then Exit Sub

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub Cancela_Datos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

If MsgBox("Cancela registro...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub
SSTab1.TabEnabled(0) = True

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS.Open RutinaLectura.Cliente(8, "", UCase((""))), vg_db, adOpenStatic
If RS.EOF Or RS!nreg = 0 Then RS.Close: Set RS = Nothing: SSTab1.TabEnabled(1) = False: modo = "NE": SSTab1.Tab = 0: Gl_Ac_Botones Me, 1, 2, modo: Exit Sub
RS.Close: Set RS = Nothing

If vaSpread1.MaxRows > 0 Then
   
   SSTab1.TabEnabled(1) = True

Else
   
   SSTab1.TabEnabled(1) = False

End If

SSTab1.Tab = 0
modo = ""
Gl_Ac_Botones Me, 1, 1, modo

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub

Private Sub Actualiza_Datos()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim Tipo As Integer
Dim codbod As Long

Tipo = IIf(Option1(0).Value = False, 2, 0)
codbod = 0
codbod = IIf(Val(fpLongInteger1(0).Value) > 0 And Trim(fpayuda(0).Caption) <> "", fpLongInteger1(0).Value, 0)
  
  If modo = "A" Then
     
     If Trim(fpText(0).text) = "" Or Trim(fpText(1).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, "Maestro de Clientes": Exit Sub
     
     If Option1(0).Value = True Then
     
        If Trim(fpText(2).text) = "" Then
        
           MsgBox "Faltan ingresar dirección del cliente...", vbExclamation + vbOKOnly, "Maestro de Clientes"
           Exit Sub
     
        End If
        
        If Trim(fpText(12).text) = "" Then
        
           MsgBox "Faltan ingresar resolución del cliente...", vbExclamation + vbOKOnly, "Maestro de Clientes"
           Exit Sub
     
        End If
        
     End If
     
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     
     RS.Open RutinaLectura.Cliente(4, LimpiaDato(Trim(fpText(0).text)), ""), vg_db, adOpenStatic
     If Not RS.EOF Then RS.Close: Set RS = Nothing: MsgBox "Cliente existe", vbExclamation + vbOKOnly, "Maestro de Clientes": Exit Sub
     RS.Close: Set RS = Nothing
     vg_db.Execute "INSERT INTO b_clientes (cli_codigo, cli_nombre, cli_direccion, " & _
                   "cli_comuna, cli_ciudad, cli_fono1, cli_fono2, cli_fax, cli_percon, " & _
                   "cli_giro, cli_email, cli_tipo, cli_socsap, cli_activo, cli_resolucion) VALUES ('" & LimpiaDato(Trim(fpText(0).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(1).text)) & "', '" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(3).text)) & "', '" & LimpiaDato(Trim(fpText(4).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(5).text)) & "', '" & LimpiaDato(Trim(fpText(6).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(7).text)) & "', '" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                   "'" & LimpiaDato(Trim(fpText(9).text)) & "', '" & LimpiaDato(Trim(fpText(10).text)) & "', " & _
                   "" & Tipo & ", '" & LimpiaDato(Trim(fpText(11).text)) & "', '1', '" & LimpiaDato(Trim(fpText(12).text)) & "')"
     '------- Actualizar codigo bodega
     vg_db.Execute "UPDATE b_clientes SET cli_codbod=" & codbod & " WHERE cli_codigo='" & LimpiaDato(Trim(fpText(0).text)) & "'"
     With vaSpread1
        
        .MaxRows = .MaxRows + 1: .Row = .MaxRows
        .Col = 1: .TypeHAlign = 0: .Value = LimpiaDato(Trim(fpText(0).text))
        .Col = 2: .TypeHAlign = 0: .Value = LimpiaDato(Trim(fpText(1).text))
        .Col = 3: .Value = IIf(Tipo = 0, "Operación", "Traspaso")
    
    End With
  
  Else
     
     If Trim(fpText(1).text) = "" Then MsgBox "Faltan datos importantes para identificar el cliente...", vbExclamation + vbOKOnly, "Maestro de Clientes": Exit Sub
     
     If Option1(0).Value = True Then
     
        If Trim(fpText(2).text) = "" Then
        
           MsgBox "Faltan ingresar dirección del cliente...", vbExclamation + vbOKOnly, "Maestro de Clientes"
           Exit Sub
     
        End If
        
        If Trim(fpText(12).text) = "" Then
        
           MsgBox "Faltan ingresar resolución del cliente...", vbExclamation + vbOKOnly, "Maestro de Clientes"
           Exit Sub
     
        End If
        
     End If
     
     
     '------- Validar si el contrato esta asociado usuario centro de costo
     If RS.State = 1 Then RS.Close
     RS.CursorLocation = adUseClient
     vg_db.CursorLocation = adUseClient
     
     RS.Open "SELECT DISTINCT a.* FROM b_clientes a, b_usuariocontratos b WHERE a.cli_codigo=b.uco_codcon AND a.cli_codigo='" & LimpiaDato(Trim(fpText(0).text)) & "' AND (a.cli_tipo=0 OR a.cli_tipo=2) AND a.cli_codbod>0", vg_db, adOpenStatic
     If Not RS.EOF Then
        If (Not IsNull(RS!cli_codbod) Or RS!cli_codbod <> 0 Or RS!cli_tipo = 2) And codbod = 0 Then
           RS.Close: Set RS = Nothing: MsgBox "Centro de costo esta asociado, en mantenedor usuario contrato...", vbCritical, "Error": Exit Sub
        End If
     End If
     RS.Close: Set RS = Nothing
     '------- Actualizar clientes

     vg_db.Execute "UPDATE b_clientes SET cli_nombre='" & LimpiaDato(Trim(fpText(1).text)) & "', cli_direccion='" & LimpiaDato(Trim(fpText(2).text)) & "', " & _
                   "cli_comuna='" & LimpiaDato(Trim(fpText(3).text)) & "', cli_ciudad='" & LimpiaDato(Trim(fpText(4).text)) & "', cli_fono1='" + LimpiaDato(Trim(fpText(5).text)) & "', " & _
                   "cli_fono2='" & LimpiaDato(Trim(fpText(6).text)) & "', cli_fax='" & LimpiaDato(Trim(fpText(7).text)) & "', cli_percon='" & LimpiaDato(Trim(fpText(8).text)) & "', " & _
                   "cli_giro='" & LimpiaDato(Trim(fpText(9).text)) & "', cli_email='" & LimpiaDato(Trim(fpText(10).text)) & "', cli_tipo=" & Tipo & ", cli_socsap='" & LimpiaDato(Trim(fpText(11).text)) & "', cli_resolucion='" & LimpiaDato(Trim(fpText(12).text)) & "' " & _
                   "WHERE cli_codigo='" & LimpiaDato(Trim(fpText(0).text)) & "'"
     '------- Actualizar codigo bodega
     vg_db.Execute "UPDATE b_clientes SET cli_codbod=" & codbod & " WHERE cli_codigo='" & LimpiaDato(Trim(fpText(0).text)) & "'"
     With vaSpread1
         .Col = 2
         .Value = LimpiaDato(Trim(fpText(1).text))
         .Col = 3
         .Value = IIf(Tipo = 0, "Operaración", "Traspaso")
     End With

  End If
  
  '------- si el contrato tiene asignado codigo bodega mover datos al contrato
  If codbod > 0 And Option1(0).Value = True Then
     
     MoverDatoNuevoContrato False, LimpiaDato(Trim(fpText(0).text))
  
     If modo = "A" Then
     
        vg_db.Execute ("sgp_Ins_TipoActividad 6")
        
     End If
     
  End If
  With vaSpread1
    .SortKey(1) = 2
    .SortKeyOrder(1) = 1
    .Sort 1, 1, .MaxCols, .MaxRows, SortByRow
  End With
  
  SSTab1.TabEnabled(0) = True
  If vaSpread1.MaxRows < 1 Then
     SSTab1.TabEnabled(1) = False
  Else
     SSTab1.TabEnabled(1) = True
     SSTab1.Tab = 0
  End If
  est = True
  modo = ""
  Gl_Ac_Botones Me, 1, 1, modo
  Label1(1).Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registros"

Exit Sub
Man_Error:
If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error":    Exit Sub
If Err = 3034 Then Exit Sub

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
ins_log_error Date & Time & Err & ":  " & error$(Err)
End Sub

Private Sub vaSpread1_Click(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

With vaSpread1
    
    If .MaxRows < 1 Then Exit Sub
    .Row = Row: .Col = 1: codigo = .text

End With

Exit Sub
Man_Error:

fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Cliente"

End Sub
