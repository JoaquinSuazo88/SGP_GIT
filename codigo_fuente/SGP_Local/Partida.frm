VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm Partida 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Gestión Contrato"
   ClientHeight    =   5070
   ClientLeft      =   2325
   ClientTop       =   3735
   ClientWidth     =   9510
   Icon            =   "Partida.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Fondo 
      Align           =   3  'Align Left
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   0
      ScaleHeight     =   4695
      ScaleWidth      =   20295
      TabIndex        =   1
      Top             =   0
      Width           =   20295
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Height          =   2535
         Left            =   6120
         TabIndex        =   18
         Top             =   1440
         Width           =   5895
         Begin VB.Label Label4 
            Height          =   1935
            Left            =   240
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   5415
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Height          =   1335
         Left            =   6120
         TabIndex        =   16
         Top             =   0
         Visible         =   0   'False
         Width           =   5295
         Begin VB.Label Label3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   735
            Left            =   360
            TabIndex        =   17
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Height          =   1065
         Left            =   14490
         TabIndex        =   13
         Top             =   105
         Width           =   2955
         Begin VB.CommandButton Command1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Actualizar Ahora"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   630
            TabIndex        =   14
            Top             =   630
            Width           =   1695
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Existe una nueva versión SGP"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   330
            Left            =   210
            TabIndex        =   15
            Top             =   210
            Width           =   2640
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Log Minuta Bloque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   0
         TabIndex        =   11
         Top             =   7080
         Width           =   9975
         Begin FPSpread.vaSpread vaSpread1 
            Height          =   2415
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Width           =   9735
            _Version        =   393216
            _ExtentX        =   17171
            _ExtentY        =   4260
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
            MaxCols         =   3
            SpreadDesigner  =   "Partida.frx":1CCA
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Log Factura PEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   10200
         TabIndex        =   9
         Top             =   7080
         Width           =   10335
         Begin FPSpread.vaSpread vaSpread3 
            Height          =   2415
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   10095
            _Version        =   393216
            _ExtentX        =   17806
            _ExtentY        =   4260
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
            MaxCols         =   6
            MaxRows         =   0
            SpreadDesigner  =   "Partida.frx":355B
            TextTip         =   3
         End
      End
      Begin VB.Frame Frame2 
         Height          =   3705
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   6015
         Begin FPSpread.vaSpread vaSpread2 
            Height          =   2325
            Left            =   210
            TabIndex        =   3
            Top             =   1140
            Width           =   5595
            _Version        =   393216
            _ExtentX        =   9869
            _ExtentY        =   4101
            _StockProps     =   64
            AllowCellOverflow=   -1  'True
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
            MaxCols         =   14
            MaxRows         =   6
            ScrollBars      =   0
            SpreadDesigner  =   "Partida.frx":38E3
            UserResize      =   0
         End
         Begin EditLib.fpDateTime fpDateTime1 
            Height          =   315
            Left            =   1635
            TabIndex        =   4
            Top             =   420
            Width           =   945
            _Version        =   196608
            _ExtentX        =   1658
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
            ButtonStyle     =   1
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
            Text            =   "09/2025"
            DateCalcMethod  =   4
            DateTimeFormat  =   5
            UserDefinedFormat=   "mm/yyyy"
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
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   2
            Left            =   3255
            Top             =   585
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Día Cerrado y enviado"
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
            Left            =   3600
            TabIndex        =   8
            Top             =   570
            Width           =   1935
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Día Cerrado y no enviado"
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
            Left            =   3600
            TabIndex        =   7
            Top             =   840
            Width           =   2205
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H008080FF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   1
            Left            =   3255
            Top             =   870
            Width           =   300
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Día Habilitado"
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
            Left            =   3615
            TabIndex        =   6
            Top             =   240
            Width           =   1245
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00FFFFFF&
            FillColor       =   &H00C0FFFF&
            FillStyle       =   0  'Solid
            Height          =   135
            Index           =   0
            Left            =   3255
            Top             =   270
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha(mm/aa)"
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
            Left            =   210
            TabIndex        =   5
            Top             =   495
            Width           =   1230
         End
      End
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   120
         Top             =   4560
      End
      Begin VB.Image Tapiz 
         Height          =   7710
         Left            =   3000
         Picture         =   "Partida.frx":3FCF
         Stretch         =   -1  'True
         Top             =   1680
         Width           =   15300
      End
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   240
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4695
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2011
            MinWidth        =   882
            TextSave        =   "08/09/2025"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   882
            TextSave        =   "16:02"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3069
            MinWidth        =   3069
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu Main 
      Caption         =   "&Minutas"
      Index           =   0
      Begin VB.Menu Minutas 
         Caption         =   "Productos"
         HelpContextID   =   1010000
         Index           =   0
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu Minutas 
         Caption         =   "Recetas"
         HelpContextID   =   1020000
         Index           =   10
      End
      Begin VB.Menu Minutas 
         Caption         =   "Bloque Minuta"
         HelpContextID   =   1022000
         Index           =   12
      End
      Begin VB.Menu Minutas 
         Caption         =   "Planificación Teorica"
         HelpContextID   =   1030000
         Index           =   20
      End
      Begin VB.Menu Minutas 
         Caption         =   "Planificación Real"
         HelpContextID   =   1040000
         Index           =   30
      End
      Begin VB.Menu Minutas 
         Caption         =   "Estructura Fija Servicio"
         HelpContextID   =   1050000
         Index           =   40
      End
      Begin VB.Menu Minutas 
         Caption         =   "Solicitud de Pedido Mensual"
         HelpContextID   =   1060000
         Index           =   50
      End
      Begin VB.Menu Minutas 
         Caption         =   "Solicitud de Adicionales y Anulaciones"
         HelpContextID   =   1070000
         Index           =   60
      End
      Begin VB.Menu Minutas 
         Caption         =   "Envio Bloque Minuta"
         HelpContextID   =   1080000
         Index           =   70
      End
      Begin VB.Menu Minutas 
         Caption         =   "Pedidos"
         Index           =   80
         Begin VB.Menu Pedidos 
            Caption         =   "Crear Pedidos"
            HelpContextID   =   1080000
            Index           =   10
         End
         Begin VB.Menu Pedidos 
            Caption         =   "Anular Pedidos"
            HelpContextID   =   1090000
            Index           =   20
         End
         Begin VB.Menu Pedidos 
            Caption         =   "Pedido Adicional"
            HelpContextID   =   1100000
            Index           =   30
         End
         Begin VB.Menu Pedidos 
            Caption         =   "-"
            Index           =   40
         End
      End
      Begin VB.Menu Minutas 
         Caption         =   "Pedido Mensual Ruta"
         HelpContextID   =   1100000
         Index           =   90
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Pacientes"
      Index           =   5
      Begin VB.Menu Paciente 
         Caption         =   "Horario Servicio"
         HelpContextID   =   5010000
         Index           =   0
      End
      Begin VB.Menu Paciente 
         Caption         =   "Grupo Paciente"
         HelpContextID   =   5020000
         Index           =   10
      End
      Begin VB.Menu Paciente 
         Caption         =   "Usuario Grupo Paciente"
         HelpContextID   =   5030000
         Index           =   20
      End
      Begin VB.Menu Paciente 
         Caption         =   "Paciente"
         HelpContextID   =   5040000
         Index           =   30
      End
      Begin VB.Menu Paciente 
         Caption         =   "-"
         Index           =   35
      End
      Begin VB.Menu Paciente 
         Caption         =   "Toma Pedido"
         HelpContextID   =   5050000
         Index           =   40
      End
      Begin VB.Menu Paciente 
         Caption         =   "Control de Ingesta"
         HelpContextID   =   5060000
         Index           =   50
      End
      Begin VB.Menu Paciente 
         Caption         =   "-"
         Index           =   55
      End
      Begin VB.Menu Paciente 
         Caption         =   "Informe Aporte Nutricional"
         HelpContextID   =   5070000
         Index           =   60
      End
      Begin VB.Menu Paciente 
         Caption         =   "Informe de Producción"
         HelpContextID   =   5080000
         Index           =   70
      End
      Begin VB.Menu Paciente 
         Caption         =   "Informe Detalle de Consumo"
         HelpContextID   =   5090000
         Index           =   80
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Inventario"
      Index           =   10
      Begin VB.Menu Inventario 
         Caption         =   "Proveedores"
         HelpContextID   =   2000000
         Index           =   0
      End
      Begin VB.Menu Inventario 
         Caption         =   "Ingreso Documento Proveedor"
         HelpContextID   =   2010000
         Index           =   10
      End
      Begin VB.Menu Inventario 
         Caption         =   "Salida a Producción"
         HelpContextID   =   2020000
         Index           =   20
      End
      Begin VB.Menu Inventario 
         Caption         =   "Devolución de Producción"
         HelpContextID   =   2030000
         Index           =   30
      End
      Begin VB.Menu Inventario 
         Caption         =   "Traspasos"
         HelpContextID   =   2040000
         Index           =   40
      End
      Begin VB.Menu Inventario 
         Caption         =   "Mermas"
         HelpContextID   =   2050000
         Index           =   50
      End
      Begin VB.Menu Inventario 
         Caption         =   "Raciones no Vendidas"
         HelpContextID   =   2060000
         Index           =   60
      End
      Begin VB.Menu Inventario 
         Caption         =   "Venta Directa"
         HelpContextID   =   2070000
         Index           =   70
      End
      Begin VB.Menu Inventario 
         Caption         =   "Control de Raciones "
         HelpContextID   =   2080000
         Index           =   80
      End
      Begin VB.Menu Inventario 
         Caption         =   "Precio Venta Cliente"
         HelpContextID   =   2090000
         Index           =   90
      End
      Begin VB.Menu Inventario 
         Caption         =   "Otros Costos A13"
         HelpContextID   =   2100000
         Index           =   100
      End
      Begin VB.Menu Inventario 
         Caption         =   "Venta Servicio Contado"
         HelpContextID   =   2110000
         Index           =   110
      End
      Begin VB.Menu Inventario 
         Caption         =   "-"
         Index           =   115
         Visible         =   0   'False
      End
      Begin VB.Menu Inventario 
         Caption         =   "Lista de Precio Cafetería"
         HelpContextID   =   2120000
         Index           =   120
         Visible         =   0   'False
      End
      Begin VB.Menu Inventario 
         Caption         =   "Registro de Venta Cafetería"
         HelpContextID   =   2130000
         Index           =   130
         Visible         =   0   'False
      End
      Begin VB.Menu Inventario 
         Caption         =   "Generar Guía Venta SAP"
         HelpContextID   =   2140000
         Index           =   140
      End
      Begin VB.Menu Inventario 
         Caption         =   "Cierre Diario"
         HelpContextID   =   2150000
         Index           =   150
      End
      Begin VB.Menu Inventario 
         Caption         =   "-"
         Index           =   155
      End
      Begin VB.Menu Inventario 
         Caption         =   "Presupuesto y Proyección"
         HelpContextID   =   2160000
         Index           =   160
      End
      Begin VB.Menu Inventario 
         Caption         =   "-"
         Index           =   165
      End
      Begin VB.Menu Inventario 
         Caption         =   "Recalculo Precio Prom. Ponderado"
         HelpContextID   =   2170000
         Index           =   170
      End
      Begin VB.Menu Inventario 
         Caption         =   "Toma de Inventario"
         HelpContextID   =   2180000
         Index           =   180
      End
      Begin VB.Menu Inventario 
         Caption         =   "-"
         Index           =   185
      End
      Begin VB.Menu Inventario 
         Caption         =   "Control Facturas Compras"
         HelpContextID   =   2190000
         Index           =   190
      End
      Begin VB.Menu Inventario 
         Caption         =   "Control Traspasos Entre Casino"
         HelpContextID   =   2200000
         Index           =   200
      End
      Begin VB.Menu Inventario 
         Caption         =   "Control Fondo Fijo (Fofi)"
         HelpContextID   =   2210000
         Index           =   210
      End
      Begin VB.Menu Inventario 
         Caption         =   "Resultado Operacionales Mensual o A13"
         HelpContextID   =   2220000
         Index           =   220
      End
      Begin VB.Menu Inventario 
         Caption         =   "Cartola Inventario"
         HelpContextID   =   2230000
         Index           =   230
      End
      Begin VB.Menu Inventario 
         Caption         =   "Control Facturas Compras(Cierre de Mes)"
         HelpContextID   =   2240000
         Index           =   240
      End
      Begin VB.Menu Inventario 
         Caption         =   "-"
         Index           =   250
      End
      Begin VB.Menu Inventario 
         Caption         =   "Salida Venta Servicios Especiales"
         HelpContextID   =   2260000
         Index           =   260
      End
      Begin VB.Menu Inventario 
         Caption         =   "Devolución Venta Servicios Especiales"
         HelpContextID   =   2270000
         Index           =   270
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Lectura Vales"
      Index           =   15
      Begin VB.Menu LecturaVales 
         Caption         =   "Punto Atención"
         HelpContextID   =   6010000
         Index           =   0
      End
      Begin VB.Menu LecturaVales 
         Caption         =   "Personal"
         HelpContextID   =   6020000
         Index           =   20
      End
      Begin VB.Menu LecturaVales 
         Caption         =   "Punto Lectura de Vales"
         HelpContextID   =   6030000
         Index           =   30
      End
      Begin VB.Menu LecturaVales 
         Caption         =   "-"
         Index           =   40
      End
      Begin VB.Menu LecturaVales 
         Caption         =   "Lectura Vales"
         HelpContextID   =   6080000
         Index           =   80
      End
      Begin VB.Menu LecturaVales 
         Caption         =   "-"
         Index           =   100
      End
      Begin VB.Menu LecturaVales 
         Caption         =   "Reporte Generico"
         HelpContextID   =   6150000
         Index           =   150
      End
   End
   Begin VB.Menu Main 
      Caption         =   "I&nforme"
      Index           =   20
      Begin VB.Menu Informe 
         Caption         =   "Planificación Teórica"
         HelpContextID   =   3010000
         Index           =   0
      End
      Begin VB.Menu Informe 
         Caption         =   "Planificación Real"
         HelpContextID   =   3020000
         Index           =   10
      End
      Begin VB.Menu Informe 
         Caption         =   "Etiquetado de Receta"
         HelpContextID   =   3021000
         Index           =   11
      End
      Begin VB.Menu Informe 
         Caption         =   "Costos Totales del Período"
         HelpContextID   =   3022000
         Index           =   12
      End
      Begin VB.Menu Informe 
         Caption         =   "Costo Detalle Periodo Realizado"
         HelpContextID   =   3024000
         Index           =   14
      End
      Begin VB.Menu Informe 
         Caption         =   "Curva ABC"
         HelpContextID   =   3026000
         Index           =   16
      End
      Begin VB.Menu Informe 
         Caption         =   "Comparativo Curva ABC"
         HelpContextID   =   3027000
         Index           =   17
      End
      Begin VB.Menu Informe 
         Caption         =   "Inflación Interna"
         HelpContextID   =   3028000
         Index           =   18
      End
      Begin VB.Menu Informe 
         Caption         =   "Analisis de Consumo Precio Fijo"
         HelpContextID   =   3029000
         Index           =   19
      End
      Begin VB.Menu Informe 
         Caption         =   "Costo Plan. Teórico -  Plan. Real - Realizado"
         HelpContextID   =   3030000
         Index           =   20
      End
      Begin VB.Menu Informe 
         Caption         =   "Comparativo Costo Teórico vs Negociado"
         HelpContextID   =   3031000
         Index           =   21
      End
      Begin VB.Menu Informe 
         Caption         =   "Food Cost"
         HelpContextID   =   3032000
         Index           =   22
      End
      Begin VB.Menu Informe 
         Caption         =   "Costo x Sector"
         HelpContextID   =   3034000
         Index           =   24
      End
      Begin VB.Menu Informe 
         Caption         =   "Comparativo de Raciones"
         HelpContextID   =   3036000
         Index           =   26
      End
      Begin VB.Menu Informe 
         Caption         =   "Facturación Clientes"
         HelpContextID   =   3040000
         Index           =   30
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   33
      End
      Begin VB.Menu Informe 
         Caption         =   "Stock"
         HelpContextID   =   3045000
         Index           =   35
      End
      Begin VB.Menu Informe 
         Caption         =   "Ficha Stock"
         HelpContextID   =   3047000
         Index           =   37
      End
      Begin VB.Menu Informe 
         Caption         =   "Producto Sin Movimiento"
         HelpContextID   =   3048000
         Index           =   38
      End
      Begin VB.Menu Informe 
         Caption         =   "Detalle Cartola de Inventario"
         HelpContextID   =   3049000
         Index           =   39
      End
      Begin VB.Menu Informe 
         Caption         =   "Compras por Período"
         HelpContextID   =   3050000
         Index           =   40
      End
      Begin VB.Menu Informe 
         Caption         =   "Detalle de Compras por Período"
         HelpContextID   =   3060000
         Index           =   50
      End
      Begin VB.Menu Informe 
         Caption         =   "Documentos Pendientes Proveedor"
         HelpContextID   =   3065000
         Index           =   55
      End
      Begin VB.Menu Informe 
         Caption         =   "Traspasos"
         HelpContextID   =   3067000
         Index           =   57
      End
      Begin VB.Menu Informe 
         Caption         =   "Salida y Devolución de Producción"
         HelpContextID   =   3070000
         Index           =   60
      End
      Begin VB.Menu Informe 
         Caption         =   "Mermas por Período"
         HelpContextID   =   3080000
         Index           =   70
      End
      Begin VB.Menu Informe 
         Caption         =   "Mermas por Preparación"
         HelpContextID   =   3082000
         Index           =   72
      End
      Begin VB.Menu Informe 
         Caption         =   "Merma Desconche"
         HelpContextID   =   3083000
         Index           =   73
      End
      Begin VB.Menu Informe 
         Caption         =   "Venta Directa"
         HelpContextID   =   3090000
         Index           =   80
      End
      Begin VB.Menu Informe 
         Caption         =   "Informe Consulta Salida o Devolución a Bodega"
         HelpContextID   =   3100000
         Index           =   90
      End
      Begin VB.Menu Informe 
         Caption         =   "Insumos no Planificados en Salida Bodega"
         HelpContextID   =   3112000
         Index           =   92
      End
      Begin VB.Menu Informe 
         Caption         =   "Ajuste Inventario"
         HelpContextID   =   3114000
         Index           =   94
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   99
      End
      Begin VB.Menu Informe 
         Caption         =   "Ventas por artículos de cafetería"
         HelpContextID   =   3110000
         Index           =   100
         Visible         =   0   'False
      End
      Begin VB.Menu Informe 
         Caption         =   "Ventas de cafetería por cliente y centro de costo"
         HelpContextID   =   3120000
         Index           =   110
         Visible         =   0   'False
      End
      Begin VB.Menu Informe 
         Caption         =   "Ventas de cafetería por cliente y centro de costo detallado"
         HelpContextID   =   3130000
         Index           =   120
         Visible         =   0   'False
      End
      Begin VB.Menu Informe 
         Caption         =   "Salida de bodega por ventas de cafetería"
         HelpContextID   =   3140000
         Index           =   130
         Visible         =   0   'False
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   140
      End
      Begin VB.Menu Informe 
         Caption         =   "Planificación Minutas Sansis"
         HelpContextID   =   3150000
         Index           =   150
      End
      Begin VB.Menu Informe 
         Caption         =   "Frecuencia de recetas o Gramos Producto Mensual Sansis"
         HelpContextID   =   3160000
         Index           =   160
      End
      Begin VB.Menu Informe 
         Caption         =   "Aporte Nutricionales Sansis"
         HelpContextID   =   3170000
         Index           =   170
      End
      Begin VB.Menu Informe 
         Caption         =   "Composición Minutas Sansis"
         HelpContextID   =   3180000
         Index           =   180
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   190
      End
      Begin VB.Menu Informe 
         Caption         =   "Importación Guías de CD "
         HelpContextID   =   3200000
         Index           =   200
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   210
      End
      Begin VB.Menu Informe 
         Caption         =   "Datos No Integrados (FLMStoSGP)"
         HelpContextID   =   3400000
         Index           =   220
      End
      Begin VB.Menu Informe 
         Caption         =   "Datos Pendientes de Integración (FLMStoSGP)"
         HelpContextID   =   3500000
         Index           =   230
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&General"
      Index           =   30
      Begin VB.Menu General 
         Caption         =   "Cambio de Contrato"
         HelpContextID   =   4011000
         Index           =   0
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu General 
         Caption         =   "Familia de Producto"
         HelpContextID   =   4010000
         Index           =   2
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Despacho"
         HelpContextID   =   4012000
         Index           =   3
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Medida"
         HelpContextID   =   4015000
         Index           =   5
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Stock"
         HelpContextID   =   4020000
         Index           =   10
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Embalaje"
         HelpContextID   =   4030000
         Index           =   20
      End
      Begin VB.Menu General 
         Caption         =   "Nutriente"
         HelpContextID   =   4040000
         Index           =   30
      End
      Begin VB.Menu General 
         Caption         =   "Impuestos"
         HelpContextID   =   4050000
         Index           =   40
      End
      Begin VB.Menu General 
         Caption         =   "Cuenta Contable"
         HelpContextID   =   4060000
         Index           =   42
      End
      Begin VB.Menu General 
         Caption         =   "Tipo Documento"
         HelpContextID   =   4061000
         Index           =   44
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   45
      End
      Begin VB.Menu General 
         Caption         =   "Categoria Dietética"
         HelpContextID   =   4070000
         Index           =   50
      End
      Begin VB.Menu General 
         Caption         =   "Tipo de Plato"
         HelpContextID   =   4080000
         Index           =   60
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   62
      End
      Begin VB.Menu General 
         Caption         =   "Tipo de Servicio"
         HelpContextID   =   4084000
         Index           =   64
      End
      Begin VB.Menu General 
         Caption         =   "Segmento"
         HelpContextID   =   4086000
         Index           =   66
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   68
      End
      Begin VB.Menu General 
         Caption         =   "Regimen"
         HelpContextID   =   4090000
         Index           =   70
      End
      Begin VB.Menu General 
         Caption         =   "Servicio"
         HelpContextID   =   4100000
         Index           =   80
      End
      Begin VB.Menu General 
         Caption         =   "Contratos"
         HelpContextID   =   4110000
         Index           =   90
      End
      Begin VB.Menu General 
         Caption         =   "Clientes"
         HelpContextID   =   4115000
         Index           =   95
      End
      Begin VB.Menu General 
         Caption         =   "Sector"
         HelpContextID   =   4117000
         Index           =   97
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   99
      End
      Begin VB.Menu General 
         Caption         =   "Bodegas"
         HelpContextID   =   4120000
         Index           =   100
      End
      Begin VB.Menu General 
         Caption         =   "Tipos de Merma"
         HelpContextID   =   4130000
         Index           =   110
      End
      Begin VB.Menu General 
         Caption         =   "Tipos de Ajuste"
         HelpContextID   =   4140000
         Index           =   120
      End
      Begin VB.Menu General 
         Caption         =   "Curva ABC"
         HelpContextID   =   4142000
         Index           =   122
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   128
      End
      Begin VB.Menu General 
         Caption         =   "Actualizar Base de Datos"
         HelpContextID   =   4149000
         Index           =   129
      End
      Begin VB.Menu General 
         Caption         =   "Parámetros Generales"
         HelpContextID   =   4150000
         Index           =   130
      End
      Begin VB.Menu General 
         Caption         =   "Limpiar Base de Dato"
         HelpContextID   =   4152000
         Index           =   132
      End
      Begin VB.Menu General 
         Caption         =   "Calendario de Cierres de Mes"
         HelpContextID   =   4160000
         Index           =   135
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   790
      End
      Begin VB.Menu General 
         Caption         =   "Casino MVI"
         HelpContextID   =   4170000
         Index           =   795
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   799
      End
      Begin VB.Menu General 
         Caption         =   "Perfiles de Acceso"
         HelpContextID   =   4800000
         Index           =   800
      End
      Begin VB.Menu General 
         Caption         =   "Usuarios"
         HelpContextID   =   4810000
         Index           =   810
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Salir"
      Index           =   40
   End
End
Attribute VB_Name = "Partida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" ( _
'ByVal hWnd As Long, _
'ByVal lpOperation As String, _
'ByVal lpFile As String, _
'ByVal lpParameters As String, _
'ByVal lpDirectory As String, _
'ByVal nShowCmd As Long) As Long
'
'Private Const BCM_SETSHIELD As Long = &H160C&
'
'Private Declare Function SendMessage Lib "user32" _
'    Alias "SendMessageA" ( _
'    ByVal hWnd As Long, _
'    ByVal wMsg As Long, _
'    ByVal wParam As Long, _
'    lParam As Any) As Long

Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As Long
    lpFile As Long
    lpParameters As Long
    lpDirectory As Long
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As Long
    hkeyClass As Long
    dwHotKey As Long
    hIcon_Or_hMonitor As Long
    hProcess As Long
End Type

Private Declare Function ShellExecuteEx Lib "shell32.dll" Alias "ShellExecuteExW" (pExecInfo As SHELLEXECUTEINFO) As Long
Private Declare Function WaitForSingleObject Lib "kernel32.dll" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const SEE_MASK_NOCLOSEPROCESS   As Long = &H40&
Private Const SW_SHOWNORMAL             As Long = 1&
Private Const INFINITE                  As Long = -1
Private Const STILL_ACTIVE              As Long = &H103
Private Const WAIT_FAILED               As Long = -1
Private Const ERROR_SUCCESS             As Long = 0&


Dim SwSalir As Integer
Dim est     As Boolean
Dim TemSeg  As Long
Dim IntMin  As Long

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim Glosa As String
Dim Ruta As String
Dim vRet  As Long

Dim sei     As SHELLEXECUTEINFO
Dim Pos     As Long
Dim sDir    As String
Dim Verb    As String
Dim sPath   As String

Dim fso               As Object
Dim Ruta_WsSgp        As String

'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
Set fso = CreateObject("Scripting.FileSystemObject")

Ruta_WsSgp = Environ("PROGRAMFILES") & "\wssgp\"

If Not isNetwork(NETWORK_ALIVE_LAN) Then

   MsgBox "No hay conexión a internet, intentelo mas tarde. Proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
   
   '-------> Ocultar Actualizador
   Frame4.Visible = False
   Set fso = Nothing
   Exit Sub
   
End If

'Validar que exista archivo
If Not fso.FileExists(Trim(dir_trabajo) & "Push.exe") Then
   
   MsgBox "No existe archivo de actualización (PUSH). " & VgLinea & _
          "            Proceso cancelado...", vbCritical + vbOKOnly, MsgTitulo
   
   Set fso = Nothing
   Exit Sub
   
End If
'Set fso = Nothing

If Version > CLng(App.Major & App.Minor & App.Revision) Then
          
   Glosa = "Para continuar con la actualización, si tiene mas de un PC con SGP recomienda aplicar la actualización" & vbCrLf & vbCrLf & Space(40) & "¿Desea continuar?"
          
ElseIf VersionSGPSDX > VersionSGPSDXPar And fso.FileExists(Ruta_WsSgp & "SgpSDX.exe") Then
       
   Glosa = "Esta actualización se aplicara solo en el PC Madre" & vbCrLf & vbCrLf & Space(40) & "¿Desea continuar?"
       
       
End If

'Set fso = Nothing

If MsgBox(Glosa, vbYesNo, "Actualizando") = vbYes Then

     '-------> Llamar programa que realiza actualización de sistema
'     Shell (Trim(dir_trabajo) & "Push.exe " & vg_RutaActualizacion)
     Dim RutaEjecutable As String
     Dim VersionSGPLocal As Long
     Dim VarPaso As String
     
     VersionSGPLocal = CLng(App.Major & App.Minor & App.Revision)
     RutaEjecutable = Trim(dir_trabajo) & "Push.exe"
     VarPaso = vg_RutaActualizacion & ";" & CStr(VersionSGPLocal) & ";" & CStr(VersionSGPSDXPar)
     'NombreEjecutable = ";" & NombreEjecutable & ";"
     'VarPaso = "-" & vg_RutaActualizacion & ";" & CStr(VersionSGPLocal) & ";" & CStr(VersionSGPSDXPar) & "-"
     'Shell "C:\windows\syswow64\calc.exe", 1
     'Shell Replace(NombreEjecutable, ";", """") & Replace(VarPaso, "-", """"), 1
     'Call Shell(NombreEjecutable & VarPaso, 0)  ' & VarPaso) '& vg_RutaActualizacion & " " & CStr(VersionSGPLocal) & " " & CStr(VersionSGPSDXPar))
     '-------> Borrar concepto descarga
'     vg_db.Execute ("DELETE  a_param WHERE par_codigo = 'Descarga'")
'    Call ShellExecute(0&, vbNullString, "Push.exe", VarPaso, RutaEjecutable, 1&)
'    Call ShellExecute(hwnd, "Open", "Push.exe", VarPaso, RutaEjecutable, vbNormalFocus)
    
    Verb = ""
    sPath = RutaEjecutable
    
    With sei
        .cbSize = LenB(sei)
        .fMask = SEE_MASK_NOCLOSEPROCESS
        .hWnd = Me.hWnd
        .lpVerb = StrPtr(Verb) 'runas
        .lpFile = StrPtr(sPath)
        .lpParameters = StrPtr(VarPaso)
        Pos = InStrRev(sPath, "\")
        If Pos <> 0 Then sDir = Left$(sPath, Pos - 1)
        .lpDirectory = StrPtr(sDir)
        .nShow = SW_SHOWNORMAL
    End With

    ShellExecuteEx sei
    
    End

End If

Exit Sub
Man_Error:
MsgBox (Err.Description)

End Sub

Private Sub fpDateTime1_Change()

On Error GoTo Man_Error

If Trim(fpDateTime1.text) = "" Or Not IsDate(fpDateTime1.text) Or est Then Exit Sub
ArmarCalendario

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub General_Click(Index As Integer)

On Error GoTo Man_Error

vg_OpcM = General.Item(Index).HelpContextID

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

Case 0 '-------> Cambio contrato
    
    M_CamCon.Show 0, Partida

Case 2 '-------> Familia producto
    
    T_FamPro.Show 0, Partida

Case 3 '-------> Parametro despacho
    
    M_PamDes.Show 0, Partida

Case 5 '-------> Unidad medida
    
    T_Unimed.Show 0, Partida

Case 10
    
    T_Unienv.Show 0, Partida

Case 20 '-------> unidad embalaje
    
    T_Uniemb.Show 0, Partida

Case 30 '-------> Nutrientes
    
    T_Nutrie.Show 0, Partida

Case 40 '-------> Impuesto
    
    T_Impues.Show 0, Partida

Case 42 '-------> Cuenta contable
    
    T_CtaCon.Show 0, Partida

Case 44 '-------> Tipo Documento
    
    T_TipDoc.Show 0, Partida

Case 50 '-------> Categoria dietetica
    
    T_CatDie.Show 0, Partida

Case 60 '-------> Tipo de plato
    
    T_TipPla.Show 0, Partida

Case 64 '-------> Tipo servicio
    
    T_TipSer.Show 0, Partida

Case 66 '-------> Segmento
    
    T_Segmen.Show 0, Partida

Case 70 '-------> Regimen
    
    T_Regime.Show 0, Partida

Case 80 '-------> Servicio
    
    vg_modpac = False
    T_Servic.Show 0, Partida

Case 90 '--------> Contrato
    
    M_Casino.Show 0, Partida

Case 95 '-------> Cliente
    
    M_Client.Show 0, Partida

Case 97 '-------> Sector
    
    T_Sector.Show 0, Partida

Case 100 '-------> Bodega
    
    T_Bodega.Show 0, Partida

Case 110 '-------> Tipo merma
    
    T_TipMer.Show 0, Partida

Case 120 '-------> Tipo Ajsute
    
    T_TipAju.Show 0, Partida

Case 122 '-------> Curva abc
    
    T_CurAbc.Show 0, Partida

Case 129 '-------> Actualizar base de dato
    
    M_ActuBD.Show 0, Partida

Case 130 '-------> Parametros generales
    
    M_Parame.Show 0, Partida

Case 132 '-------> Limpiar base de dato
    
    P_LimDat.Show 0, Partida

Case 133 '-------> Proceso Recalculo Día
    
    P_RCaDia.Show 0, Partida

Case 135 '-------> Calendario cierre de mes
    
    M_CiePer.Show 0, Partida

Case 800 '-------> Perfil acceso
    
    M_Perfil.Show 0, Partida

Case 795 '--------> Contrato MVI
    
    'M_Casino_MVI.Show 0, Partida

Case 810
    
    vg_modpac = False
    If Not formAbierto("Usuari") Then
       
       Dim Usuari As New M_Usuari
       Usuari.lc_Aux = "Usuari"
       Usuari.Tag = "Usuari"
       Usuari.Show 0, Partida
       Set Usuari = Nothing
    
    End If

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Informe_Click(Index As Integer)

On Error GoTo Man_Error

vg_OpcM = Informe.Item(Index).HelpContextID

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

Case 0
    
    If Not formAbierto("PlaTei") Then
       
       Dim PlaTei As New I_MenTeo
       PlaTei.lc_Aux = "PlaTei"
       PlaTei.Tag = "PlaTei"
       PlaTei.Show 0, Partida
       Set PlaTei = Nothing
    
    End If

Case 10
    
    If Not formAbierto("PlaRei") Then
       
       Dim PlaRei As New I_MenTeo
       PlaRei.lc_Aux = "PlaRei"
       PlaRei.Tag = "PlaRei"
       PlaRei.Show 0, Partida
       Set PlaRei = Nothing
    
    End If

Case 11 'etiquetado receta
    
    If Not formAbierto("EtiquetaReceta") Then
       
       Dim EtiquetaReceta As New I_EtiquetaReceta
       EtiquetaReceta.lc_Aux = "EtiquetaReceta"
       EtiquetaReceta.Tag = "EtiquetaReceta"
       EtiquetaReceta.Show 0, Partida
       Set EtiquetaReceta = Nothing
    
    End If

Case 12
    
    If Not formAbierto("CosTot") Then
       
       Dim CosTot As New I_FCost
       CosTot.lc_Aux = "CosTot"
       CosTot.Tag = "CosTot"
       CosTot.Show 0, Partida
       Set CosTot = Nothing
    
    End If

Case 14 '-------> Costo del periodo realizado
    
    If Not formAbierto("CosPer") Then
       
       Dim CosPer As New I_FCost
       CosPer.lc_Aux = "CosPer"
       CosPer.Tag = "CosPer"
       CosPer.Show 0, Partida
       Set CosPer = Nothing
    
    End If

Case 16 '------> Curva ABC
    
    If Not formAbierto("CurABC") Then
       
       Dim CurABC As New I_FCost
       CurABC.lc_Aux = "CurABC"
       CurABC.Tag = "CurABC"
       CurABC.Show 0, Partida
       Set CurABC = Nothing
    
    End If

Case 17 '------> Comparativo curva ABC
    
    If Not formAbierto("CocABC") Then
       
       Dim CocABC As New I_FCost
       CocABC.lc_Aux = "CocABC"
       CocABC.Tag = "CocABC"
       CocABC.Show 0, Partida
       Set CocABC = Nothing
    
    End If

Case 18 '------> Inflación interna
    
    If Not formAbierto("InfInt") Then
       
       Dim InfInt As New I_FicSto
       InfInt.lc_Aux = "InfInt"
       InfInt.Tag = "InfInt"
       InfInt.Show 0, Partida
       Set InfInt = Nothing
    
    End If

Case 19 '------> Analisis de Consumos Precio Fijo
    
    If Not formAbierto("AnaCpf") Then
       
       Dim AnaCpf As New I_FicSto
       AnaCpf.lc_Aux = "AnaCpf"
       AnaCpf.Tag = "AnaCpf"
       AnaCpf.Show 0, Partida
       Set AnaCpf = Nothing
    
    End If

Case 20 '-------> Costo planificación teórica - real - realizado
'    I_CoteRe.Inicio "Costo Plan. Teórico - Plan. Real - Realizado", "I"
'    I_CoteRe.Show 0, Partida
    
    If Not formAbierto("CoTeRe") Then
       
       Dim CoTeRe As New I_CoteRe
       CoTeRe.lc_Aux = "CoTeRe"
       CoTeRe.Tag = "CoTeRe"
       CoTeRe.Show 0, Partida
       Set CoTeRe = Nothing
    
    End If

Case 21 '-------> Comparativo Planificación Teorica Vs Negociado
    
    If Not formAbierto("CPlanTeoNeg") Then
       
       Dim CPlanTeoNeg As New I_CoteRe
       CPlanTeoNeg.lc_Aux = "CPlanTeoNeg"
       CPlanTeoNeg.Tag = "CPlanTeoNeg"
       CPlanTeoNeg.Show 0, Partida
       Set CPlanTeoNeg = Nothing
    
    End If

Case 22 '------> Food cost
    
    If Not formAbierto("FooCos") Then
       
       Dim FooCos As New I_FCost
       FooCos.lc_Aux = "FooCos"
       FooCos.Tag = "FooCos"
       FooCos.Show 0, Partida
       Set FooCos = Nothing
    
    End If

Case 24 '-------> Costo sector
    
    If Not formAbierto("CosSec") Then
       
       Dim CosSec As New I_FCost
       CosSec.lc_Aux = "CosSec"
       CosSec.Tag = "CosSec"
       CosSec.Show 0, Partida
       Set CosSec = Nothing
    
    End If

Case 26 '-------> Comparativo de raciones
    
    If Not formAbierto("ConRac") Then
       
       Dim ConRac As New I_FCost
       ConRac.lc_Aux = "ConRac"
       ConRac.Tag = "ConRac"
       ConRac.Show 0, Partida
       Set ConRac = Nothing
    
    End If

Case 30 '-------> Facturación Clientes
    
    I_FacCli.Show 0, Partida

Case 35 'Informe de Stock
    
    I_Stock.Show 0, Partida

Case 37 '-------> Movimiento Stock

'    I_FicSto.Show 0, Partida
    If Not formAbierto("FicSto") Then
       
       Dim FicSto As New I_FicSto
       FicSto.lc_Aux = "FicSto"
       FicSto.Tag = "FicSto"
       FicSto.Show 0, Partida
       Set FicSto = Nothing
    
    End If

Case 38 '-------> Producto Sin Movimiento
    
    If Not formAbierto("ProMov") Then
       
       Dim ProMov As New I_FicSto
       ProMov.lc_Aux = "ProMov"
       ProMov.Tag = "ProMov"
       ProMov.Show 0, Partida
       Set ProMov = Nothing
    
    End If

Case 39 '-------> Detalle Cartola Inventario
    
    If Not formAbierto("DetCarInv") Then
       
       Dim DetCarInv As New I_FicSto
       DetCarInv.lc_Aux = "DetCarInv"
       DetCarInv.Tag = "DetCarInv"
       DetCarInv.Show 0, Partida
       Set DetCarInv = Nothing
    
    End If

Case 40 '-------> Control Facturas Compras MSP
    
    I_ComPer.Show 0, Partida

Case 50 '------> Control Facturas Compras MSP
    
    I_DetCom.Show 0, Partida

Case 55 'Control Facturas Compras MSP
    
    I_DocPen.Show 0, Partida

Case 57 '-------> Traspasos
    
    I_Traspa.Show 0, Partida

Case 60 'Salidas de Bodega MSP
    
    I_SalBod.Show 0, Partida

Case 70 '-------> Mermas por Período MSP
'    I_MerPed.Show 0, Partida
    
    If Not formAbierto("MerPed") Then
       
       Dim MerPed As New I_MerPed
       MerPed.lc_Aux = "MerPed"
       MerPed.Tag = "MerPed"
       MerPed.Show 0, Partida
       Set MerPed = Nothing
    
    End If

Case 72 '-------> Mermas por Preparación
    
    If Not formAbierto("MerPre") Then
       
       Dim MerPre As New I_FCost
       MerPre.lc_Aux = "MerPre"
       MerPre.Tag = "MerPre"
       MerPre.Show 0, Partida
       Set MerPre = Nothing
    
    End If

Case 73 '-------> Mermas Desconche
    
    If Not formAbierto("MerDes") Then
       
       Dim MerDes As New E_MermaDesconcheProduccion
       MerDes.lc_Aux = "MerDes"
       MerDes.Tag = "MerDes"
       MerDes.Show 0, Partida
       Set MerDes = Nothing
    
    End If

Case 80 '-------> Venta Directa MSP
    
    I_VenDir.Show 0, Partida

Case 90
    
    C_SaDebo.Show 0, Partida

Case 92
    
    If Not formAbierto("InNPla") Then
       
       Dim InNPla As New I_FCost
       InNPla.lc_Aux = "InNPla"
       InNPla.Tag = "InNPla"
       InNPla.Show 0, Partida
       Set InNPla = Nothing
    
    End If

Case 94 '-------> Detalle ajuste inventario
    
    If Not formAbierto("AjuInv") Then
       
       Dim AjuInv As New I_MerPed
       AjuInv.lc_Aux = "AjuInv"
       AjuInv.Tag = "AjuInv"
       AjuInv.Show 0, Partida
       Set AjuInv = Nothing
    
    End If

Case 100
    
    If Not formAbierto("VenCaf1") Then
        
        Dim VenCaf1 As New I_VenCaf
        VenCaf1.lc_Aux = "VenCaf1"
        VenCaf1.Tag = "VenCaf1"
        VenCaf1.Show 0, Partida
        Set VenCaf1 = Nothing
    
    End If

Case 110
    
    If Not formAbierto("VenCaf2") Then
        
        Dim VenCaf2 As New I_VenCaf
        VenCaf2.lc_Aux = "VenCaf2"
        VenCaf2.Tag = "VenCaf2"
        VenCaf2.Show 0, Partida
        Set VenCaf2 = Nothing
    
    End If

Case 120
    
    If Not formAbierto("VenCaf3") Then
        
        Dim VenCaf3 As New I_VenCaf
        VenCaf3.lc_Aux = "VenCaf3"
        VenCaf3.Tag = "VenCaf3"
        VenCaf3.Show 0, Partida
        Set VenCaf3 = Nothing
    
    End If

Case 130
    
    If Not formAbierto("VenCaf4") Then
        
        Dim VenCaf4 As New I_VenCaf
        VenCaf4.lc_Aux = "VenCaf4"
        VenCaf4.Tag = "VenCaf4"
        VenCaf4.Show 0, Partida
        Set VenCaf4 = Nothing
    
    End If

Case 150 'Planificción Minutas Sansis

    If Not formAbierto("SetPlaSansis") Then
        
        Dim SetPlaSansis As New I_SetPlaSansis
        SetPlaSansis.lc_Aux = "SetPlaSansis"
        SetPlaSansis.Tag = "SetPlaSansis"
        SetPlaSansis.Show 0, Partida
        Set SetPlaSansis = Nothing
    
    End If

Case 160 ' Frecuencia de recetas o Gramos producto mensual Sansis

    If Not formAbierto("FreGrP") Then
        
        Dim FreGrP As New I_FreGrP
        FreGrP.lc_Aux = "FreGrP"
        FreGrP.Tag = "FreGrP"
        FreGrP.Show 0, Partida
        Set FreGrP = Nothing
    
    End If

Case 170 'Aporte Nutricionales Sansis

    If Not formAbierto("ApoNutSansis") Then
        
        Dim ApoNutSansis As New I_ApoNutSansis
        ApoNutSansis.lc_Aux = "ApoNutSansis"
        ApoNutSansis.Tag = "ApoNutSansis"
        ApoNutSansis.Show 0, Partida
        Set ApoNutSansis = Nothing
    
    End If

Case 180 'Composición Minuta Sansis

    If Not formAbierto("E_ComposicionMinutasSansis") Then
        
        Dim ComposicionMinutasSansis As New E_ComposicionMinutasSansis
        ComposicionMinutasSansis.lc_Aux = "ComposicionMinutasSansis"
        ComposicionMinutasSansis.Tag = "ComposicionMinutasSansis"
        ComposicionMinutasSansis.Show 0, Partida
        Set ComposicionMinutasSansis = Nothing
    
    End If

Case 200 'Importación Guia Cd

    If Not formAbierto("I_ImportacionGuiaCd") Then
        
        Dim I_ImportacionGuiaCd As New I_ImportacionGuiaCd
        I_ImportacionGuiaCd.lc_Aux = "I_ImportacionGuiaCd"
        I_ImportacionGuiaCd.Tag = "I_ImportacionGuiaCd"
        I_ImportacionGuiaCd.Show 0, Partida
        Set I_ImportacionGuiaCd = Nothing
    
    End If
    
Case 220 'reporte de datos no integrados

    If Not formAbierto("I_FLMS_NoIntegrado") Then
        
        Dim I_FLMS_NoIntegrado As New I_FLMS_NoIntegrado
        I_FLMS_NoIntegrado.lc_Aux = "I_FLMS_NoIntegrado"
        I_FLMS_NoIntegrado.Tag = "I_FLMS_NoIntegrado"
        I_FLMS_NoIntegrado.Show 0, Partida
        Set I_FLMS_NoIntegrado = Nothing
    
    End If
    
Case 230 'reporte de datos pendientes de integración

    If Not formAbierto("I_FLMS_PendienteIntegrar") Then
        
        Dim I_FLMS_PendienteIntegrar As New I_FLMS_PendienteIntegrar
        I_FLMS_PendienteIntegrar.lc_Aux = "I_FLMS_PendienteIntegrar"
        I_FLMS_PendienteIntegrar.Tag = "I_FLMS_PendienteIntegrar"
        I_FLMS_PendienteIntegrar.Show 0, Partida
        Set I_FLMS_PendienteIntegrar = Nothing
    
    End If

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Inventario_Click(Index As Integer)

On Error GoTo Man_Error

vg_OpcM = Inventario.Item(Index).HelpContextID

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

Case 0 '-------> Proveedor
    
    M_Provee.Show 0, Partida

Case 10 '------> Documento proveedor
    
    M_DocPro.Show 0, Partida

Case 20 '-------> Salida a bodega
    
    M_SalBod.Show 0, Partida

Case 30 '------> Devolución a bodega
    
    M_DevBod.Show 0, Partida

Case 40 '------> Traspaso
    
    M_Traspa.Show 0, Partida

Case 50 '------> Mermas
    
    M_Mermas.Show 0, Partida

Case 60 '------> Raciones no Vendidas
    
    M_MerPre.Show 0, Partida

Case 70 '------> Venta directa
    
    M_VenDir.Show 0, Partida

Case 80 '-------> Control de Raciones Administrador
    
    M_ConRac.Show 0, Partida

Case 90 '-------> Precio Venta Cliente
    
    M_PVtaCl.Show 0, Partida

Case 100 '-------> Gastos A13
    
    M_GasA13.Show 0, Partida

Case 110 '-------> Venta Servicio Contado
    
    M_VtaCon.Show 0, Partida

Case 120 '-------> Lista de Precio Cafetería
    
    T_LiPrCa.Show 0, Partida

Case 130 '-------> Registro de Venta Cafetería
    
    M_VenCaf.Show 0, Partida

Case 140 '-------> Generar Guías Ventas (SAP)
    
    M_GuiVta.Show 0, Partida

Case 150 '-------> Cierre Diario
    
    M_RCDiar.Show 1, Partida

Case 160 '-------> Presupuesto y Proyección
    
    M_PrePro.Show 0, Partida

Case 170 '-------> Recalculo Precio Prom. Ponderado
    
    P_RecPPP.Show 0, Partida

Case 180 '-------> Toma de Inventario
    
    vg_invrot = "0"
   
    M_TomInv.Show 1, Partida

Case 190 '-------> Control Facturas Compras
    
    I_CtrFCo.Inicio "Control Facturas Compras", "C"
    I_CtrFCo.Show 0, Partida

Case 200 '-------> Control traspasos entre Contratos
    
    I_CtrFCo.Inicio "Control Traspasos Entre Contratos", "T"
    I_CtrFCo.Show 0, Partida

Case 210 '-------> Control Fondo Fijo (Fofi)
    
    I_CtrFCo.Inicio "Control Fondo Fijo (Fofi)", "F"
    I_CtrFCo.Show 0, Partida

Case 220 '-------> Resultado Operacionales Mensual o A13
    
    I_A13.Show 0, Partida

Case 230 '-------> Cartola Inventario
    
    I_CarInv.Show 0, Partida

Case 240 '-------> Control Facturas Compras (Cierre de Mes)
    
    I_CfcCie.Show 0, Partida

Case 260 '-------> Salida Venta Servicios Especiales
    
    M_SalidaServicioEspeciales.Show 0, Partida

Case 270 '-------> Devolución Venta Servicios Especiales
    
    M_DevolucionServicioEspeciales.Show 0, Partida

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub LecturaVales_Click(Index As Integer)

On Error GoTo Man_Error

vg_OpcM = LecturaVales.Item(Index).HelpContextID

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

Case 0 '-------> Mantenedor Punto Atención
    
    T_PunVen.Show 0, Partida

Case 20 '-------> Mantenedor Personal
    
    M_Personas.Show 0, Partida

Case 30 '-------> Punto Lectura de Vales
    
    M_PtoLecVal.Show 0, Partida

Case 80 '-------> Lectura Vales
    
    M_LecVal.Show 0, Partida

Case 150 '------> Informe Lectura Vales
    
    If Not formAbierto("LecVal") Then
       
       Dim lecval As New I_LecVal
       lecval.lc_Aux = "LecVal"
       lecval.Tag = "LecVal"
       lecval.Show 0, Partida
       Set lecval = Nothing
    
    End If

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Main_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index

Case 40
    
    Dim frm As Form
    
    If (Forms.count > 1) Then
       
       For Each frm In Forms
           
           If frm.Name = "M_ActuBD" Then
              
              MsgBox "Esta activo el ítem actualizar base de datos, espera hasta que termine el proceso y luego puede salir...", vbCritical + vbOKOnly, "Menú Principal": Exit Sub
           
           End If
        
        Next
    
    End If
    
    MDIForm_Unload (0)

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MainPa_Click(Index As Integer)

On Error GoTo Man_Error

vg_OpcM = Inventario.Item(Index).HelpContextID
Select Case Index

Case 10
    
    M_TomPed.Show 0, Partida

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MDIForm_Activate()

On Error GoTo Man_Error

TraerFechaCierre
VerConfReg
vg_CSep = ","
vg_CDec = "."
'ArmarCalendario
est = False

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MDIForm_Deactivate()

On Error GoTo Man_Error

TraerFechaCierre
'ArmarCalendario
est = False

sendmessage Command1.hWnd, BCM_SETSHIELD, 0&, 1&

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Public Sub MDIForm_Load()

Dim i    As Long, cCas As String, nCas As String, cBody As String, sql1 As String
Dim RS  As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset

On Error GoTo ManError

TemSeg = 0
IntMin = 10

'-------> Ocultar Actualizador
Frame4.Visible = False
'-------> Ocultar aviso raciones SPRS digitado un día que el sitio no trabaja
Frame6.Visible = False

'-------> Validar si existe parametro envio datos
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

vg_bloenv = 0
Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'parbloenv'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'parbloenv', 'Parametro bloque envio datos', 'N', '100', '" & vg_contra & "'")
   vg_bloenv = 100

Else
   
   vg_bloenv = RS!par_valor

End If
RS.Close: Set RS = Nothing

'-------> Inicio insert ftp y correo si no existe
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'ftpdir'")

If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'ftpdir', 'Ftp Directorio', 'C', 'y®­À·½¿Äµ¸', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'ftppas'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'ftppas', 'Ftp Password', 'C', 't¾°Å½†…ƒŠ}', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'ftppue'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'ftppue', 'ftp Puerto',  'C', '||', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'ftpser'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'ftpser', 'Ftp Servidor', 'C', '½²¼{Á¾´¶Ê»Ã¸¾ÀÄ¾ˆ¾È', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'ftpusu'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'ftpusu', 'Ftp Usuario', 'C', '¿¾±¿´ÃÀ', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'corcum'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'corcum','Correo Cuenta Mail','C', '«¯¹¶¼±´´³Æ½ÃÅ—ËÈ¾ÀÔÅÍÃÍ', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'corpas'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'corpas','Correo Password', 'C',    '«¯¹¶¼²±Ä»ÁÃÈ', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'corser'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'corser','Correo Cuenta Mail', 'C', '{{z}|€€€……', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & vg_contra & "', 'corusu'")
If RS.EOF Then
   
   vg_db.Execute ("sgp_Ins_Param 'corusu','Correo Usuario', 'C', '«¯¹¶¼±´´³Æ½ÃÅÊ', '" & vg_contra & "'")

End If
RS.Close: Set RS = Nothing
'-------> Fin insert ftp y correo si no existe

Me.Caption = "SGP v " & Trim(Str(App.Major)) & "." & Trim(Str(App.Minor)) & "." & Trim(Str(App.Revision))
StatusBar1.Panels(3).text = "Contrato : " & vg_contra 'Trim(GetParametro("casino")) & " "
vg_tipser = False

'-------> Traer datos del contrato
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Clientes 1, '" & vg_contra & "'")
If Not RS.EOF Then
   
   cCas = Trim(RS!cli_codigo)
   nCas = Trim(RS!cli_nombre)
   vg_tipser = IIf(IsNull(RS!cli_codtis) Or RS!cli_codtis = 2, True, False)
   StatusBar1.Panels(4).text = Trim(RS!cli_nombre) & " "

End If
RS.Close: Set RS = Nothing

StatusBar1.Panels(5).text = "Usuario : " & Trim(vg_NUsr) & " "
StatusBar1.Panels(6).text = vg_codbod & " " & vg_nombod

If vg_tipbase = "1" Then
   
   StatusBar1.Panels(9).text = "Base Access"

Else
   
   StatusBar1.Panels(9).text = "Sql Server"

End If

'-------> Traer Pais
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Pais 1, '" & vg_pais & "'")
If Not RS.EOF Then
   
   StatusBar1.Panels(10).text = "Pais : " & Trim(RS!pai_nombre) & " "

End If
RS.Close: Set RS = Nothing

'-------> Traer periodo
StatusBar1.Panels(7).text = "Periodo : "
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
If Not RS.EOF Then
   
   StatusBar1.Panels(7).text = "Periodo : " & Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)

End If
RS.Close: Set RS = Nothing
cBody = cCas & Chr(9) & nCas & Chr(9) & Format(Date, "dd/MM/yyyy") & Chr(9) & Format(Time, "HH:mm") & Chr(9) & Trim(vg_NUsr)

'-------> Mover datos un nuevo contrato creado
MoverDatoNuevoContrato True, vg_contra

'-------> Mover Cantidad decimales
vg_DCa = IIf(IsNull(GetParametro("parcandec")), 3, GetParametro("parcandec"))
If vg_DCa = 0 Then vg_DCa = 3

'-------> Mover Precios decimales
vg_DPr = 0
'amorgado vg_DPr = IIf(IsNull(GetParametro("parpredec")), 2, GetParametro("parpredec"))
'If vg_DPr = 0 Then vg_DPr = 2

'-------> Consultar si el contrato tiene permiso de bloqueo en productos y planificación
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Let vg_modprod = False: vg_modrec = False: vg_5etapas = False: vg_modprove = False: Vg_MinSre = True
Set RS = vg_db.Execute("sgp_Sel_Param 2, '" & MuestraCasino(1) & "', ''")
If Not RS.EOF Then
   
   Do While Not RS.EOF
      
      If RS!par_codigo = "modpro" And RS!par_valor = 1 Then
         
         vg_modprod = True
      
      ElseIf RS!par_codigo = "modrec" And RS!par_valor = 1 Then
         
         vg_modrec = True
      
      ElseIf RS!par_codigo = "modprove" And RS!par_valor = 1 Then
         
         vg_modprove = True
      
      ElseIf RS!par_codigo = "minsre" And RS!par_valor = 1 Then
         
         Let Vg_MinSre = False
         vaSpread1.Visible = True
         Frame3.Visible = True
      
      End If
      
      RS.MoveNext
   
   Loop

End If
Vg_MinSre = True

If ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 2) Then
   
   Vg_MinSre = False
   vaSpread1.Visible = True
   Frame3.Visible = True

End If

If Vg_MinSre = False Then
   
   Let vg_5etapas = False

Else
   
   Let vg_5etapas = True

End If
RS.Close: Set RS = Nothing

'-------> Traer encabezado clave contable
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'claencsap'")
If Not RS.EOF Then vg_claencsap = RS!par_valor
RS.Close: Set RS = Nothing

'-------> Traer detalle clave contable
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'cladetsap'")
If Not RS.EOF Then vg_cladetsap = RS!par_valor
RS.Close: Set RS = Nothing

'-------> Traer documento exento impuesto
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'docexento'")
If Not RS.EOF Then vg_docexento = Trim(RS!par_valor)
RS.Close: Set RS = Nothing

'-------> Traer documento exento impuesto
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'docafecto'")
If Not RS.EOF Then vg_docafecto = Trim(RS!par_valor)
RS.Close: Set RS = Nothing

'-------> Traer activar opcion paciente
vg_modpac = False
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 3, '" & MuestraCasino(1) & "', 'modpac'")
If Not RS.EOF Then vg_modpac = True
RS.Close: Set RS = Nothing

'-------> Traer cantidad decimales recetas
vg_RDCa = 2
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'parrcandec'")
If Not RS.EOF Then vg_RDCa = IIf(IsNull(RS!par_valor) Or Trim(RS!par_valor) = "", 2, RS!par_valor)
RS.Close: Set RS = Nothing

'-------> Traer calcular ditigo verificador
vg_Dig = "S"
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Param 1, '" & MuestraCasino(1) & "', 'parcaldig'")
If Not RS.EOF Then vg_Dig = IIf(IsNull(RS!par_valor) Or Trim(RS!par_valor) = "S", 2, RS!par_valor)
RS.Close: Set RS = Nothing

If "S" = fg_CambiaChar(GetParametro("5etapas"), ";", "','") Then
   vg_db.Execute ("sgp_Upd_Param 1, '" & MuestraCasino(1) & "', 'ingpedmen', '', '', '1'")
End If

'SendMail oMail, "SGP : (" & cCas & ") " & nCas, cBody, "", "Alexis Morgado", "amorgado@sodexho.cl"

vg_db.Execute "UPDATE b_bodegas set bod_canmer = 0 WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 0"

Dim imgX As ListImage
IL1.ListImages.Clear
IL1.ImageHeight = 16
IL1.ImageWidth = 16
Set imgX = IL1.ListImages.Add(, "A_Incluir  ", LoadResPicture(101, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Incluir  ", LoadResPicture(102, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Alterar", LoadResPicture(103, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Alterar", LoadResPicture(104, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Borrar ", LoadResPicture(105, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Borrar ", LoadResPicture(106, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Actualizar", LoadResPicture(107, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Actualizar", LoadResPicture(108, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Cancelar ", LoadResPicture(109, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Cancelar ", LoadResPicture(110, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Confirmar ", LoadResPicture(111, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Confirmar ", LoadResPicture(112, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Imprimir ", LoadResPicture(113, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Imprimir ", LoadResPicture(114, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Salir    ", LoadResPicture(115, vbResIcon))
Set imgX = IL1.ListImages.Add(, "excel", LoadResPicture(116, vbResIcon))
Set imgX = IL1.ListImages.Add(, "word", LoadResPicture(117, vbResIcon))
Set imgX = IL1.ListImages.Add(, "acrobat", LoadResPicture(118, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_VerReceta", LoadResPicture(119, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Grabar   ", LoadResPicture(120, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Grabar   ", LoadResPicture(121, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Enviar", LoadResPicture(122, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Previa   ", LoadResPicture(123, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Historico", LoadResPicture(124, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Cortar", LoadResPicture(125, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Copiar", LoadResPicture(126, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Pegar", LoadResPicture(127, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Pegar", LoadResPicture(128, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_InsertarF", LoadResPicture(129, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_EliminarF", LoadResPicture(130, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_SubirF", LoadResPicture(131, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_BajarF", LoadResPicture(132, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_CopiarD", LoadResPicture(133, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Aportes", LoadResPicture(134, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_BuscarPro", LoadResPicture(135, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Frecuencia", LoadResPicture(136, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Filtro", LoadResPicture(137, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Grafico", LoadResPicture(138, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Costo", LoadResPicture(139, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Ajuste", LoadResPicture(140, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_AnulaAjuste", LoadResPicture(141, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Deshacer", LoadResPicture(142, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ActCostoReceta", LoadResPicture(143, vbResIcon))
Set imgX = IL1.ListImages.Add(, "ActuBD", LoadResPicture(144, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_PegadoEspecial", LoadResPicture(145, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_PegadoEspecial", LoadResPicture(146, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Candado", LoadResPicture(147, vbResIcon))
Set imgX = IL1.ListImages.Add(, "I_Llave", LoadResPicture(148, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ExporReceta", LoadResPicture(149, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Grafico1", LoadResPicture(150, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Planificacion", LoadResPicture(151, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_EstFijaDía", LoadResPicture(152, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Buscar", LoadResPicture(153, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ExportaPlanif", LoadResPicture(154, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ImportaPlanif", LoadResPicture(155, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_NoEnviar", LoadResPicture(156, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ExportarInventario", LoadResPicture(158, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ImportarInventario", LoadResPicture(159, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ExploradorWindows", LoadResPicture(160, vbResIcon))

Set imgX = IL1.ListImages.Add(, "A_GenerarArchivo", LoadResPicture(102, vbResBitmap))

For Each Item In Minutas
    
    If Item.Caption <> "-" Then Item.Visible = False

Next

For Each Item In Paciente
    
    If Item.Caption <> "-" Then Item.Visible = False

Next

For Each Item In Inventario
    
    If Item.Caption <> "-" Then Item.Visible = False

Next

For Each Item In LecturaVales
    
    If Item.Caption <> "-" Then Item.Visible = False

Next

For Each Item In Informe
    
    If Item.Caption <> "-" Then Item.Visible = False

Next

For Each Item In General
    
    If Item.Caption <> "-" Then Item.Visible = False

Next
Main.Item(0).Visible = False
Main.Item(5).Visible = False
Main.Item(10).Visible = False
Main.Item(15).Visible = False
Main.Item(20).Visible = False
Main.Item(30).Visible = False

Dim nVer1 As Long
Dim aVer1 As Long

nVer1 = CLng(App.Major & App.Minor & App.Revision)
aVer1 = TipoDato(GetParametro("version"), 0)

If nVer1 = aVer1 Then

   If RS1.State = 1 Then RS1.Close
   RS1.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   If Not vg_tipser Then
   
'      RS1.Open "SELECT * FROM a_derechosperfil WHERE dpe_codper = " & vg_CPer & " AND dpe_codopc <> 4140000", vg_db, adOpenStatic
      Set RS1 = vg_db.Execute("sgp_Sel_DerechosPerfiles " & vg_CPer & "")
   
   Else
   
'      RS1.Open "SELECT * FROM a_derechosperfil WHERE dpe_codper = " & vg_CPer & " AND dpe_codopc <> 4140000 and dpe_codopc in (1010000,2000000,2010000, 2020000,2030000,2040000, " & _
'               "2050000,2100000,2150000,2160000,2180000,2190000,2200000,2210000,2220000,2230000,2240000,3045000,3047000,3050000,3060000,3065000,3067000,3090000,3100000,3114000," & _
'               "4011000,4010000,4015000,4020000,4030000,4050000,4060000,4084000,4086000,4110000,4120000,4130000,4149000,4150000,4160000,4800000,4810000)", vg_db, adOpenStatic
     Set RS1 = vg_db.Execute("sgp_Sel_DerechosPerfilesII " & vg_CPer & "")
     
   End If

    Do While Not RS1.EOF
       
       Select Case Mid(Trim(Str(RS1!dpe_codopc)), 1, 1)
       
       Case 1
            
            For Each Item In Minutas
                
                If Not ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 2) Then
                  
                  Let Partida.Minutas.Item(12).Visible = False
                  Let Partida.Minutas.Item(70).Visible = False
               End If
               
               If Not ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
                  
                  Let Partida.Minutas.Item(50).Visible = False
               
               End If
           
               If RS1!dpe_codopc = 1070000 And "S" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")) Then Exit For
               If RS1!dpe_codopc = 1070000 Then Exit For
               If Item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
                  
                  Item.Visible = True: Main.Item(0).Visible = True
    
                  If Not ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 2) Then
                     
                     Let Partida.Minutas.Item(12).Visible = False
                     Let Partida.Minutas.Item(70).Visible = False
                  
                  End If
               
               If Not ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
                  
                  Let Partida.Minutas.Item(50).Visible = False
               
               End If
                  
                  Exit For
               
               End If
    
               If Not ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 2) Then
                  
                  Let Partida.Minutas.Item(12).Visible = False
                  Let Partida.Minutas.Item(70).Visible = False
               
               End If
               
               If Not ValidarAccesoMinutaBloqueyBloqueo(MuestraCasino(1), 1) Then
                  
                  Let Partida.Minutas.Item(50).Visible = False
               
               End If
               
           Next
       
       Case 2
           
           For Each Item In Inventario
               
               If Item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then Item.Visible = True: Main.Item(10).Visible = True: Exit For
           
           Next
       
       Case 3
           
           For Each Item In Informe
               
               If Item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then Item.Visible = True: Main.Item(20).Visible = True: Exit For
           
           Next
       
       Case 4
           
           For Each Item In General
               
               If Item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then Item.Visible = True: Main.Item(30).Visible = True: Exit For
           
           Next
       
       Case 5 And vg_modpac
           
           For Each Item In Paciente
               
               If Item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then Item.Visible = True: Main.Item(5).Visible = True: Exit For
           
           Next
       
       Case 6
           
           For Each Item In LecturaVales
               
               If Item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then Item.Visible = True: Main.Item(15).Visible = True: Exit For
           
           Next
       
       End Select
       
       RS1.MoveNext
    
    Loop
    RS1.Close
    Set RS1 = Nothing

End If

Main.Item(40).Visible = True

'-------> Armar Calendario
est = True
fpDateTime1.DateTimeFormat = UserDefined
fpDateTime1.UserDefinedFormat = "mm/yyyy"

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_CierrePeriodo 1, '" & MuestraCasino(1) & "'")
If Not RS.EOF Then
   
   fpDateTime1.text = Mid(RS!cie_periodo, 5, 2) & "/" & Mid(RS!cie_periodo, 1, 4)

Else
   
   fpDateTime1.text = Format(Date, "mm/yyyy")

End If
RS.Close
Set RS = Nothing

est = False
TraerFechaCierre

'-------> Armar calendario
ArmarCalendario

'------> Carga evento sitio remoto
CargaEventoSitioRemoto

'------> Carga Factura PEL
CargaEventoLogFacturaPEL

est = False

vg_Block_Botton_Actua_Receta_MVI = False 'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18
vg_Clave_MVI = "" 'MVA - MVI - BLOQUEO BOTON TOOLBAR ACTUALIZAR RECETA - 2013-01-18

Exit Sub
ManError:
If Err.Number = 340 Then Resume Next

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

On Error GoTo Man_Error

Dim msg, Response   '-------> Declara variables.
If SwSalir = 1 Then Exit Sub
msg = "¿Esta Seguro Salir?"
Response = MsgBox(msg, 4 + 32, "Sistema Gestión")
Select Case Response

Case 2  '-------> No permite cerrar.
    
    Cancel = -1
    msg = "El comando ha sido cancelado."

Case 6
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir_Sistema"), "SGP", "", "", "")
    
    SwSalir = 1
    'vg_db.Close
    Me.Hide
    Unload Me
    End

Case 7
    
    Cancel = -1

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub MDIForm_Resize()

On Error GoTo Man_Error

'If Me.WindowState = 0 Then
'   Fondo.Height = Me.Height - 1200
'Else
'   Fondo.Height = Me.Height
'End If

'Fondo.Height = Me.Height
Fondo.Width = Me.Width - 120
Tapiz.Top = Fondo.Top
Tapiz.Left = Fondo.Left
Tapiz.Height = Fondo.Height
Tapiz.Width = Fondo.Width

Frame3.Top = ScaleHeight - 2800 'IIf(Me.WindowState = 2, 11300, ScaleHeight - 2000)
Frame1.Top = ScaleHeight - 2800
Frame3.Left = Fondo.Left
'Frame4.Left = ScaleHeight + 6200 '- 2800 'ScaleLeft + 16000
''vaSpread1.Height = Fondo.Height
'vaSpread1.Width = Fondo.Width

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Minutas_Click(Index As Integer)

On Error GoTo Man_Error

    vg_OpcM = Minutas.Item(Index).HelpContextID

    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
    
    Select Case Index
        
        Case 0
            
            M_Produc.Show 0, Partida
        
        Case 10
            
            vg_newcodrec = 0: vg_newnomrec = "": vg_tiprec = -2: vg_newestrec = False
            vg_5etapas = IIf("N" = (fg_CambiaChar(GetParametro("5etapas"), ";", "','")), False, True)
            M_Receta.Show 0, Partida
        
        Case 12 ' bloque minuta
            
            vg_tipmin = True
            Call M_MinSR1.Show(0, Partida)
        
        Case 20
            
            vg_tipmin = False
            Dim PlaTeo As New M_Plami1
            If Not formAbierto("PlaTeo") Then
               
               PlaTeo.lc_Aux = "PlaTeo"
               PlaTeo.Tag = "PlaTeo"
               PlaTeo.Show 0, Partida
               Set PlaTeo = Nothing
            End If

        Case 30
            
            vg_tipmin = False
            If Not formAbierto("PlaRea") Then
                Dim PlaRea As New M_Plami1
                PlaRea.lc_Aux = "PlaRea"
                PlaRea.Tag = "PlaRea"
                PlaRea.Show 0, Partida
                Set PlaRea = Nothing
            End If
        
        Case 40
            
            M_EstFij.Show 0, Partida
        
        Case 50
                
            M_Pedido.Show 0, Partida
        
        Case 60
            
            M_GeAdAn.Show 0, Partida
        
        Case 70
            
            Call M_EnMinRem.Show(0, Partida)

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Paciente_Click(Index As Integer)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

'-------> Traer activar opcion paciente
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

vg_modpac = False
Set RS = vg_db.Execute("sgp_Sel_Param 3, '" & MuestraCasino(1) & "', 'modpac'")
If Not RS.EOF Then

   vg_modpac = True
   
End If
RS.Close: Set RS = Nothing
vg_OpcM = Paciente.Item(Index).HelpContextID

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

Case 0

    If Not formAbierto("SerPac") Then
       
       Dim SerPac As New T_Servic
       SerPac.lc_Aux = "SerPac"
       SerPac.Tag = "SerPac"
       SerPac.Show 0, Partida
       Set SerPac = Nothing
    
    End If

Case 10
    
    T_GruPac.Show 0, Partida

Case 20
    
    If Not formAbierto("UsuGPa") Then
       
       Dim UsuGPa As New M_Usuari
       UsuGPa.lc_Aux = "UsuGPa"
       UsuGPa.Tag = "UsuGPa"
       UsuGPa.Show 0, Partida
       Set UsuGPa = Nothing
    
    End If

Case 30 '-------> Paciente
    
    M_Pacien.Show 0, Partida

Case 40 'Toma pedido paciente
    
    M_TomPed.Show 0, Partida

Case 50 '-------> Control de Ingesta
    
    M_CtrIng.Show 0, Partida

Case 60 'Informe aporte nutriconal
    
    If Not formAbierto("ANutPa") Then
       
       Dim ANutPa As New I_PedPac
       ANutPa.lc_Aux = "ANutPa"
       ANutPa.Tag = "ANutPa"
       ANutPa.Show 0, Partida
       Set ANutPa = Nothing
    
    End If

Case 70 '-------> Informe producción
    
    If Not formAbierto("Produc") Then
       
       Dim Produc As New I_PedPac
       Produc.lc_Aux = "Produc"
       Produc.Tag = "Produc"
       Produc.Show 0, Partida
       Set Produc = Nothing
    
    End If

Case 80
    
    If Not formAbierto("DetCon") Then
       
       Dim DetCon As New I_PedPac
       DetCon.lc_Aux = "DetCon"
       DetCon.Tag = "DetCon"
       DetCon.Show 0, Partida
       Set DetCon = Nothing
    
    End If

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Pedidos_Click(Index As Integer)

On Error GoTo Man_Error

vg_OpcM = Minutas.Item(Index).HelpContextID
Select Case Index

Case 10 '-------> Crear Pedidos
    
    M_CrePed.Show 0, Partida

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Sub ArmarCalendario()

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim sql1 As String
'-------> Armar calendario
Dim i As Long, j As Long, nrosem As Long, diafin As Long, indexit As Boolean
vaSpread2.TextTip = 2 'SS_TEXTTIP_FLOATINGFOCUSONLY
' Control displays text tips after 250 milliseconds
vaSpread2.TextTipDelay = 0
vaSpread2.Row = -1: vaSpread2.Col = -1:
vaSpread2.BackColor = &H8000000F
diafin = fg_mes(Format(fpDateTime1.text, "mm") & Format(fpDateTime1.text, "yyyy"))
nrosem = 1
vaSpread2.Visible = False

For i = 1 To 6
    
    For j = 1 To 14
        
        vaSpread2.Row = i
        vaSpread2.Col = j
        vaSpread2.text = ""
    
    Next j

Next i

For i = 1 To diafin
    
    Select Case fg_Dia(Format(fpDateTime1.text, "yyyymm") & fg_pone_cero(i, 2))
    
    Case 1
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 7 'domingo
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
        nrosem = nrosem + 1
    
    Case 2
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 1 'lunes
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
    
    Case 3
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 2 'martes
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
    
    Case 4
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 3 'miercoles
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
    
    Case 5
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 4 'jueves
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
    
    Case 6
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 5 'viernes
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
    
    Case 7
        
        vaSpread2.Row = nrosem
        vaSpread2.Col = 6 'sabado
        vaSpread2.BackColor = IIf(Format(fg_pone_cero(i, 2) & "/" & (Format(fpDateTime1.text, "mm/yyyy")), "yyyymmdd") < Format(vg_ciedia, "yyyymmdd"), Shape1(2).FillColor, Shape1(0).FillColor)
        vaSpread2.CellType = CellTypeStaticText
        vaSpread2.TypeHAlign = TypeHAlignCenter
        vaSpread2.text = CStr(i)
    
    End Select
    
    If i = 1 Then

'       Toolbar1.Buttons(1).Visible = IIf(vaSpread1.BackColor = Shape1(1).FillColor, False, True)
'       Toolbar1.Buttons(2).Visible = IIf(vaSpread1.BackColor = Shape1(1).FillColor, True, False)
    
    End If

Next i


If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

vaSpread2.RetainSelBlock = False
Set RS1 = vg_db.Execute("sgp_Sel_Log_EnvioCierreDiario 1, '" & MuestraCasino(1) & "', " & Format(fpDateTime1.text, "yyyymm") & "")
indexit = False

If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      indexit = False
      
      For j = 1 To 6
          
          For i = 1 To 7
              
              vaSpread2.Row = j
              vaSpread2.Col = i
              
              If Val(vaSpread2.text) = Val(Mid(RS1!Fecha, 1, 2)) And RS1!Fecha <= (CDate(vg_ciedia) - 1) Then
                 
                 vaSpread2.BackColor = IIf(RS1!estenv = "0", Shape1(1).FillColor, Shape1(2).FillColor)
                 vaSpread2.Col = i + 7
                 vaSpread2.text = IIf(IsNull(RS1!fecsub) Or Trim(RS1!fecsub) = "", "", Trim(RS1!fecsub))
                 indexit = True: Exit For
              
              End If
          
          Next i
          
          If indexit Then Exit For
      
      Next j
      
      RS1.MoveNext
   
   Loop

End If
RS1.Close: Set RS1 = Nothing
vaSpread2.Visible = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub CargaEventoSitioRemoto()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Log_EnvioMinutaSitioRemoto 1, '" & MuestraCasino(1) & "'")
vaSpread1.MaxRows = 0
If RS.EOF Then
   
   Frame3.Visible = False: vaSpread1.Visible = False

Else
   
   Frame3.Visible = True: vaSpread1.Visible = True

End If
Do While Not RS.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   Let vaSpread1.text = IIf(IsNull(RS!fecpro), "", RS!fecpro)
   vaSpread1.Col = 2
   Let vaSpread1.text = IIf(IsNull(RS!FecRec), "", RS!FecRec)
   vaSpread1.Col = 3
   Let vaSpread1.text = IIf(IsNull(RS!cencos), "", Trim(RS!cencos)) & " - " & IIf(IsNull(RS!cli_nombre), "", RS!cli_nombre) & " - " & IIf(IsNull(RS!estado), "", Trim(RS!estado)) & " - " & IIf(IsNull(RS!mensaje), "", Trim(RS!mensaje)) '& VgLinea
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub CargaEventoLogFacturaPEL()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_Log_FacturaSap 1, '" & MuestraCasino(1) & "'")
vaSpread3.MaxRows = 0
Frame1.Left = 10000
If RS.EOF Then
   
   Frame1.Visible = False: vaSpread3.Visible = False

Else
   
   Frame1.Visible = True: vaSpread3.Visible = True

End If

Do While Not RS.EOF
   
   vaSpread3.MaxRows = vaSpread3.MaxRows + 1
   vaSpread3.Row = vaSpread3.MaxRows
   vaSpread3.Col = 1
   Let vaSpread3.text = RS!prv_codigo
   vaSpread3.Col = 2
   Let vaSpread3.text = RS!prv_nombre
   vaSpread3.Col = 3
   Let vaSpread3.text = RS!NumeroFactura
   vaSpread3.Col = 4
   Let vaSpread3.text = RS!TipoDocumento
   vaSpread3.Col = 5
   Let vaSpread3.text = RS!Fecha
   vaSpread3.Col = 6
   Let vaSpread3.text = RS!estado & " - " & Trim(RS!Observacion)
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub CargarMensajeInvCalendarizado()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

'-------> Ocultar Actualizador
Frame5.Visible = False

'Traer parametro versión SGPSDX
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_MesajeSemanalInvCalendarizado '" & MuestraCasino(1) & "'")

If Not RS.EOF Then
   
   If Trim(RS!Glosa) <> "" Then
   
      Frame5.Visible = True
      Label3.Caption = Trim(RS!Glosa)
      
   End If

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub ActualizarVersion()

On Error GoTo Man_Error

Dim RS                As New ADODB.Recordset
Dim Ruta              As String
Dim vRet              As Long
Dim Descarga()        As String
'Dim Version           As Long
'Dim VersionSGPSDX     As Long
'Dim VersionSGPSDXPar  As Long

Dim fso               As Object
Dim Ruta_WsSgp        As String

'Instanciar el objeto FSO para poder usar las funciones FileExists y FolderExists
Set fso = CreateObject("Scripting.FileSystemObject")

Ruta_WsSgp = Environ("PROGRAMFILES") & "\wssgp\"

Version = 0
VersionSGPSDX = 0
VersionSGPSDXPar = 0

If Not isNetwork(NETWORK_ALIVE_LAN) Then

   Set fso = Nothing
   Exit Sub
   
End If
'-------> Ocultar Actualizador
Frame4.Visible = False

'Traer parametro versión SGPSDX
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("select isnull(par_valor,'') as par_valor from a_param where par_codigo = 'VersionSDX'")
If Not RS.EOF Then
   
   If Trim(RS!par_valor) <> "" Then
   
      VersionSGPSDXPar = Trim(RS!par_valor)
    
   Else
   
      VersionSGPSDXPar = 0
      
   End If

End If
RS.Close
Set RS = Nothing

' Traer parametro actualización
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("select isnull(par_valor,'') as par_valor from a_param where par_codigo = 'DescargaR'")
If Not RS.EOF Then
   
'nVer < aVer
   If Trim(RS!par_valor) <> "" Then
   
       Descarga = Split(Trim(RS!par_valor), ";")
       Version = Descarga(1)
       VersionSGPSDX = Descarga(2)
       
       If Version > CLng(App.Major & App.Minor & App.Revision) Then
          
          Label2.Caption = "Existe una nueva versión SGP"
          Frame4.Visible = True
      
          vg_RutaActualizacion = ""
          vg_RutaActualizacion = Descarga(0)
          
          'ValidaPCServidor = False
       ElseIf VersionSGPSDX > VersionSGPSDXPar And fso.FileExists(Ruta_WsSgp & "SgpSDX.exe") Then
       
          Label2.Caption = "Existe una nueva versión SGPSDX"
          Frame4.Visible = True
      
          vg_RutaActualizacion = ""
          vg_RutaActualizacion = Descarga(0)
       
       End If
       
   End If

End If
RS.Close
Set RS = Nothing

Set fso = Nothing

Exit Sub
Man_Error:
    
    fg_descarga
    
    If Err = 13 Then

       MsgBox Err & ":  " & " Existe problema en la ruta descarga, informe al area IT...", vbCritical + vbOKOnly, MsgTitulo
    
    Else
    
       MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
       ins_log_error Date & Time & Err & ":  " & error$(Err)
  
    End If
    
End Sub

Private Sub VerificarTomaInventario()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

' Traer parametro actualización
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_VerificarTomaInventario '" & MuestraCasino(1) & "'")
If Not RS.EOF Then

   P_CierreDiarioInventario.GeneraMDBInven RS!Fecha
   P_CierreDiarioInventario.Show 1, Partida

End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Timer1_Timer()

On Error GoTo Man_Error

Dim RS         As New ADODB.Recordset
'Dim nVer       As Long
'Dim aVer       As Long
Dim NomMaquina As String

Timer1.Enabled = False

   '-------> restablecer proceso del demonio
   If Not ConsultaProcess("sgpsdx.exe") And Not IsFormLoaded(M_RCDiar) Then
      
      On Error Resume Next
      vRet = Shell(Environ("PROGRAMFILES") & "\wssgp\" & "sgpsdx.exe")

   End If
   
   '-------> Insertar y actualizar descarga
   If RS.State = 1 Then RS.Close
   RS.CursorLocation = adUseClient
   vg_db.CursorLocation = adUseClient
   
   Set RS = vg_db.Execute("sgp_InsUpd_VersionDescarga")
   If Not RS.EOF Then
   
      If RS(0) > 0 Then
                  
         MsgBox RS(0) & " " & RS(1), vbCritical + vbOKOnly, MsgTitulo
               
      End If
      
   End If
   RS.Close
   Set RS = Nothing
 
   '-------> reestablece
   ActualizarVersion
   ArmarCalendario
   CargaEventoSitioRemoto
   CargaEventoLogFacturaPEL
   CargarMensajeInvCalendarizado
   
   '--> Proceso bloqueo & eliminacion de cuenta usuario
   ProcesoBloqueoEliminacionCuentaUsuario

   '--> Proceso cambio contraseña desde ADMSGP
   ProcesoCambioContrasenaADMSGP

   
   '--> Proceso de integración entre FLMS y SGP Casino
   Integracion_SGPtoFLMS


   If ValidaPCServidor Then
   
      VerificarTomaInventario
      
   End If


   '--> Validar raciones SPRS ingresada un dia que no tiene minuta
   ValidarRaciones_SPRS
   
'   nVer = CLng(App.Major & App.Minor & App.Revision)
'   aVer = TipoDato(GetParametro("version"), 0)
'
'    '-------> Validar PC Servidor
'    If ValidaPCServidor = False Then
'
'        NomMaquina = "Hijo"
'
'    Else
'
'        NomMaquina = "Madre"
'
'    End If
'
'    If nVer < aVer Then
'
'      MsgBox "Debe realizar la actualización de Versión " & aVer & " en SGP " & NomMaquina & VgLinea & VgLinea & "        Se Cerrara sistema ...", vbCritical + vbOKOnly, "SGP"
'      End
'
'    End If

Timer1.Enabled = True

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub


Private Sub Integracion_SGPtoFLMS()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'--> Ejecuta proceso de integración entre FLMS y SGP Casino
vg_db.Execute ("EXEC ProcesoDeIntegracion_FLMStoSGP")


Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub


Private Sub ProcesoBloqueoEliminacionCuentaUsuario()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'Ejecutar Proceso Bloqueo Cuenta Usuario 45 dias
vg_db.Execute ("exec sgp_Upd_ProcesoBloqueoCuentaUsuario")

'Ejecutar Proceso eliminación Cuenta Usuario 90 dias
vg_db.Execute ("exec sgp_DelIns_ProcesoEliminacionCuentaUsuario")

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub


Private Sub ProcesoCambioContrasenaADMSGP()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

'--> Cambio de contraseña desde ADMSGP a SGP local del ADMSGPLOCAL
vg_db.Execute ("exec sgp_Upd_ContrasenaSGPADM_Usuario_ADMSGPLOCAL")

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

Private Sub ValidarRaciones_SPRS()

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset
Dim Glosa As String

Frame6.Visible = False
      
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgp_Sel_ValidarRacionesSPRS '" & MuestraCasino(1) & "'")

If Not RS.EOF Then
     
      Frame6.Visible = True
      Label4.Visible = True
      
      Glosa = ""
      Glosa = "Existen raciones vendida SPRS, en días que no hay Planificación : "
        
      Do While Not RS.EOF
           
         Glosa = Glosa & VgLinea & RS(2) & " - Reg. " & RS(0) & "- Ser. " & RS(1)
           
         RS.MoveNext
        
      Loop
      
      Label4.Caption = Glosa
      
End If
RS.Close
Set RS = Nothing

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo
    ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub

   

