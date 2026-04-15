VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.MDIForm Partida 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema de Gestión Contrato"
   ClientHeight    =   7605
   ClientLeft      =   2295
   ClientTop       =   3225
   ClientWidth     =   9510
   Icon            =   "Partida.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Fondo 
      Align           =   3  'Align Left
      Height          =   7230
      Left            =   0
      ScaleHeight     =   7170
      ScaleWidth      =   19020
      TabIndex        =   1
      Top             =   0
      Width           =   19080
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar Lista Minuta Bloque"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   3
         Top             =   5640
         Width           =   1815
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   1935
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   16095
         _Version        =   393216
         _ExtentX        =   28390
         _ExtentY        =   3413
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
         SpreadDesigner  =   "Partida.frx":030A
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
      Top             =   7230
      Width           =   9510
      _ExtentX        =   16775
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2011
            MinWidth        =   882
            TextSave        =   "13/10/2025"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            TextSave        =   "13:59"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4313
            MinWidth        =   4304
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2725
            MinWidth        =   2716
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
         HelpContextID   =   1000000
         Index           =   0
      End
      Begin VB.Menu Minutas 
         Caption         =   "Generar Ing. - Nut. - Rec. - Minuta OPTIMUM"
         HelpContextID   =   1013000
         Index           =   2
      End
      Begin VB.Menu Minutas 
         Caption         =   "Asociar Productos SAC vs SGP"
         HelpContextID   =   1010000
         Index           =   10
      End
      Begin VB.Menu Minutas 
         Caption         =   "Asociar Productos SAP Vs SGP"
         HelpContextID   =   1012000
         Index           =   12
      End
      Begin VB.Menu Minutas 
         Caption         =   "Proveedores"
         HelpContextID   =   1020000
         Index           =   20
      End
      Begin VB.Menu Minutas 
         Caption         =   "Generación Archivos Planos Productos"
         HelpContextID   =   1030000
         Index           =   30
      End
      Begin VB.Menu Minutas 
         Caption         =   "Generación Archivos Planos Lista Precios Sac"
         HelpContextID   =   1032000
         Index           =   32
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   35
      End
      Begin VB.Menu Minutas 
         Caption         =   "Lista de Precio"
         HelpContextID   =   1040000
         Index           =   40
      End
      Begin VB.Menu Minutas 
         Caption         =   "Costo Precio Ingrediente Comercial"
         HelpContextID   =   1045000
         Index           =   45
      End
      Begin VB.Menu Minutas 
         Caption         =   "Importar Lista de Precio Desde SAC"
         HelpContextID   =   1050000
         Index           =   50
      End
      Begin VB.Menu Minutas 
         Caption         =   "Importar Lista de Precio Desde Excel"
         HelpContextID   =   1060000
         Index           =   60
      End
      Begin VB.Menu Minutas 
         Caption         =   "Asociar Lista de Precio"
         HelpContextID   =   1070000
         Index           =   70
      End
      Begin VB.Menu Minutas 
         Caption         =   "Actualizar Lista de Precio Planificación"
         HelpContextID   =   1080000
         Index           =   80
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   85
      End
      Begin VB.Menu Minutas 
         Caption         =   "Recetas"
         HelpContextID   =   1090000
         Index           =   90
      End
      Begin VB.Menu Minutas 
         Caption         =   "Generación Archivos Planos Recetas"
         HelpContextID   =   1100000
         Index           =   100
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   105
      End
      Begin VB.Menu Minutas 
         Caption         =   "Planificación Minutas Segmento"
         HelpContextID   =   1110000
         Index           =   110
      End
      Begin VB.Menu Minutas 
         Caption         =   "Cambio Planificación Minuta De Propuesta a Real"
         HelpContextID   =   1111000
         Index           =   111
      End
      Begin VB.Menu Minutas 
         Caption         =   "Minuta Real Casino"
         HelpContextID   =   1112000
         Index           =   112
      End
      Begin VB.Menu Minutas 
         Caption         =   "Tabla Gramaje"
         HelpContextID   =   1120000
         Index           =   120
      End
      Begin VB.Menu Minutas 
         Caption         =   "Actualiza - Copia Minuta Lideres a Seguidores"
         HelpContextID   =   1122000
         Index           =   122
      End
      Begin VB.Menu Minutas 
         Caption         =   "Generación Archivos Planos Planificación"
         HelpContextID   =   1130000
         Index           =   130
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   131
      End
      Begin VB.Menu Minutas 
         Caption         =   "Borrar Sitios Masivos"
         HelpContextID   =   1132000
         Index           =   132
      End
      Begin VB.Menu Minutas 
         Caption         =   "Copia Minuta Bloque Estandar x Ceco"
         HelpContextID   =   1133000
         Index           =   133
      End
      Begin VB.Menu Minutas 
         Caption         =   "Minuta Bloque Costo Bandeja x Servicios"
         HelpContextID   =   1134000
         Index           =   134
      End
      Begin VB.Menu Minutas 
         Caption         =   "Copia Minuta Bloque x Ceco"
         HelpContextID   =   1135000
         Index           =   135
      End
      Begin VB.Menu Minutas 
         Caption         =   "Minuta Bloque"
         HelpContextID   =   1136000
         Index           =   136
      End
      Begin VB.Menu Minutas 
         Caption         =   "Envio Minuta Bloque"
         HelpContextID   =   1137000
         Index           =   137
      End
      Begin VB.Menu Minutas 
         Caption         =   "Liberación Minuta Bloque"
         HelpContextID   =   1138000
         Index           =   138
      End
      Begin VB.Menu Minutas 
         Caption         =   "Actualizaciones Varias"
         Index           =   139
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Raciones & Ponderaciones Excel"
            HelpContextID   =   1139000
            Index           =   0
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Estado de Minuta"
            HelpContextID   =   1139100
            Index           =   10
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Porcentaje Ponderación"
            HelpContextID   =   1139200
            Index           =   20
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "-"
            Index           =   29
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Cambiar Pedido Proyecto a CD o PAP"
            HelpContextID   =   1139300
            Index           =   30
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Cambiar Estado Pedido"
            HelpContextID   =   1139320
            Index           =   32
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Asignar Lista Convenios Ceco Propuesta"
            HelpContextID   =   1139400
            Index           =   40
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Eliminar Carros de Compras"
            HelpContextID   =   1139500
            Index           =   50
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "-"
            Index           =   51
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Cambio de Recetas Minuta Bloque"
            HelpContextID   =   1139600
            Index           =   60
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "-"
            Index           =   61
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Exportar Tabla Gramaje & Bach - Input"
            HelpContextID   =   1139700
            Index           =   70
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "-"
            Index           =   71
         End
         Begin VB.Menu ActualizarMinutas 
            Caption         =   "Ajuste Estacionales Recetas"
            HelpContextID   =   1139800
            Index           =   80
         End
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   140
      End
      Begin VB.Menu Minutas 
         Caption         =   "Parametrizar Nº Recetas 5 Etapas"
         HelpContextID   =   1140000
         Index           =   141
      End
      Begin VB.Menu Minutas 
         Caption         =   "Parametrizar Costo Patrón Piso 5 Etapas"
         HelpContextID   =   1150000
         Index           =   150
      End
      Begin VB.Menu Minutas 
         Caption         =   "Parametrizar Costo Patrón Techo 5 Etapas"
         HelpContextID   =   1160000
         Index           =   160
      End
      Begin VB.Menu Minutas 
         Caption         =   "Gramaje Familia Producto 5 Etapas"
         HelpContextID   =   1170000
         Index           =   170
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   180
      End
      Begin VB.Menu Minutas 
         Caption         =   "Pedido"
         Index           =   190
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Generación Pedido"
            HelpContextID   =   1190000
            Index           =   0
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "-"
            Index           =   10
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Excepción Formato de Compra Centro Costo"
            HelpContextID   =   1182000
            Index           =   20
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Producto que no Arrastran Saldo"
            HelpContextID   =   1193000
            Index           =   30
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Excluir Ingrediente en Pedido"
            HelpContextID   =   1194000
            Index           =   40
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Ingredientes Saldo Abastecimiento Exceso"
            HelpContextID   =   1194200
            Index           =   42
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "-"
            Index           =   50
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Parametro Fecha Despacho x Ceco"
            HelpContextID   =   1196000
            Index           =   60
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Parametro Fecha Despacho x Proveedor"
            HelpContextID   =   1197000
            Index           =   70
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Asociar Familia SGP & Grupo Despacho"
            HelpContextID   =   1198000
            Index           =   80
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "-"
            Index           =   82
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Consultar convenios Pel"
            HelpContextID   =   1199000
            Index           =   90
         End
         Begin VB.Menu GeneracionPedido 
            Caption         =   "Mover Cero Arrastre de Saldo"
            HelpContextID   =   1199400
            Index           =   94
         End
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   199
      End
      Begin VB.Menu Minutas 
         Caption         =   "Diwo"
         Index           =   200
         Begin VB.Menu Diwo 
            Caption         =   "Parametrizar Ceco Estructura Servicio"
            HelpContextID   =   1210000
            Index           =   10
         End
         Begin VB.Menu Diwo 
            Caption         =   "Formato Salida "
            HelpContextID   =   1220000
            Index           =   20
         End
         Begin VB.Menu Diwo 
            Caption         =   "-"
            Index           =   30
         End
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   201
      End
      Begin VB.Menu Minutas 
         Caption         =   "Pantalla Led"
         Index           =   220
         Begin VB.Menu PanLed 
            Caption         =   "Parametrizar Ceco Estructura Servicio"
            HelpContextID   =   1221000
            Index           =   10
         End
         Begin VB.Menu PanLed 
            Caption         =   "-"
            Index           =   20
         End
      End
      Begin VB.Menu Minutas 
         Caption         =   "-"
         Index           =   230
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Informe"
      Index           =   20
      Begin VB.Menu Informe 
         Caption         =   "Planificación Minutas Segmento"
         HelpContextID   =   2010000
         Index           =   0
      End
      Begin VB.Menu Informe 
         Caption         =   "Costo Minutas"
         HelpContextID   =   2012000
         Index           =   2
      End
      Begin VB.Menu Informe 
         Caption         =   "Planificación Minuta Bloque"
         HelpContextID   =   2013000
         Index           =   3
      End
      Begin VB.Menu Informe 
         Caption         =   "Consumo Ingredientes"
         HelpContextID   =   2020000
         Index           =   10
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu Informe 
         Caption         =   "Comsumo Ingrediente Minuta Bloque"
         HelpContextID   =   2030000
         Index           =   20
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel Minuta Bloque"
         HelpContextID   =   2040000
         Index           =   30
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel Detalle Minuta Bloque"
         HelpContextID   =   2050000
         Index           =   40
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel Detalle Minuta II"
         HelpContextID   =   2060000
         Index           =   50
      End
      Begin VB.Menu Informe 
         Caption         =   "Template Minuta Bloque"
         HelpContextID   =   2062000
         Index           =   52
      End
      Begin VB.Menu Informe 
         Caption         =   "Historial Trabajos por Lotes"
         HelpContextID   =   2065000
         Index           =   60
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   65
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel Q Sitios"
         HelpContextID   =   2070000
         Index           =   70
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   80
      End
      Begin VB.Menu Informe 
         Caption         =   "Aportes Nutricional Sansis"
         HelpContextID   =   2080000
         Index           =   90
      End
      Begin VB.Menu Informe 
         Caption         =   "Planificación Minutas Sansis"
         HelpContextID   =   2090000
         Index           =   100
      End
      Begin VB.Menu Informe 
         Caption         =   "Frecuencia De Recetas o Gramos Producto Mensual"
         HelpContextID   =   2092000
         Index           =   110
      End
      Begin VB.Menu Informe 
         Caption         =   "Composición Minutas Sansis"
         HelpContextID   =   2093000
         Index           =   120
      End
      Begin VB.Menu Informe 
         Caption         =   "Minuta Costo Sansis"
         HelpContextID   =   2094000
         Index           =   122
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   130
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel Varios"
         HelpContextID   =   2095000
         Index           =   140
      End
      Begin VB.Menu Informe 
         Caption         =   "-"
         Index           =   150
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Plaf. Teorica - Real - Realizada"
         HelpContextID   =   2160000
         Index           =   160
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Costo Merma Desconche"
         HelpContextID   =   2170000
         Index           =   170
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel Ingrediente No Poseen Precio Vigente"
         HelpContextID   =   2180000
         Index           =   180
      End
      Begin VB.Menu Informe 
         Caption         =   "Exportar Excel So Health"
         HelpContextID   =   2190000
         Index           =   190
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&General"
      Index           =   30
      Begin VB.Menu General 
         Caption         =   "Familia de Producto"
         HelpContextID   =   4000000
         Index           =   0
      End
      Begin VB.Menu General 
         Caption         =   "Carga de Archivos"
         HelpContextID   =   4000700
         Index           =   5
      End
      Begin VB.Menu General 
         Caption         =   "Grupo de Cambio Ingrediente en Receta SGP"
         HelpContextID   =   4010000
         Index           =   10
      End
      Begin VB.Menu General 
         Caption         =   "Habilitar Cambio Ingrediente en Receta SGP"
         HelpContextID   =   4020000
         Index           =   20
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Medida"
         HelpContextID   =   4030000
         Index           =   30
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Stock"
         HelpContextID   =   4040000
         Index           =   40
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Embalaje"
         HelpContextID   =   4050000
         Index           =   50
      End
      Begin VB.Menu General 
         Caption         =   "Nutriente"
         HelpContextID   =   4060000
         Index           =   60
      End
      Begin VB.Menu General 
         Caption         =   "Impuestos"
         HelpContextID   =   4070000
         Index           =   70
      End
      Begin VB.Menu General 
         Caption         =   "Cuenta Contable"
         HelpContextID   =   4080000
         Index           =   80
      End
      Begin VB.Menu General 
         Caption         =   "Tipo Documento"
         HelpContextID   =   4090000
         Index           =   90
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   95
      End
      Begin VB.Menu General 
         Caption         =   "Categoria Dietética"
         HelpContextID   =   4100000
         Index           =   100
      End
      Begin VB.Menu General 
         Caption         =   "Tipo de Plato"
         HelpContextID   =   4110000
         Index           =   110
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Recetas"
         Index           =   112
         Begin VB.Menu ParametroReceta 
            Caption         =   "-"
            Index           =   1
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Color"
            HelpContextID   =   4111000
            Index           =   10
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Costo"
            HelpContextID   =   4112000
            Index           =   20
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Tipo Ingrediente Principal"
            HelpContextID   =   4113000
            Index           =   30
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Metodo Cocción"
            HelpContextID   =   4114000
            Index           =   40
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Categorización Complejidad"
            HelpContextID   =   4115000
            Index           =   50
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Ingrediente de Cruce Garnitura"
            HelpContextID   =   4116000
            Index           =   60
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Estacionalidad"
            HelpContextID   =   4117000
            Index           =   70
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Efecto Meteorizante"
            HelpContextID   =   4118000
            Index           =   80
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Sellos"
            HelpContextID   =   4119000
            Index           =   90
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Tipo Negocio"
            HelpContextID   =   4111100
            Index           =   100
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Intolerancia"
            HelpContextID   =   4111110
            Index           =   110
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Tiempo HH"
            HelpContextID   =   4111120
            Index           =   120
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Tiempo Cocción"
            HelpContextID   =   4111130
            Index           =   130
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Alergeno"
            HelpContextID   =   4111140
            Index           =   140
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Estilo Alimentación"
            HelpContextID   =   4111150
            Index           =   150
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Parámetro Adicional N°1"
            HelpContextID   =   4111160
            Index           =   160
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Parámetro Adicional N°2"
            HelpContextID   =   4111170
            Index           =   170
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Etiquetado Sello"
            HelpContextID   =   4111180
            Index           =   180
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Grupo Ingrediente Principal"
            HelpContextID   =   4111190
            Index           =   190
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Parametro Salsa"
            HelpContextID   =   4111200
            Index           =   200
         End
         Begin VB.Menu ParametroReceta 
            Caption         =   "Equipamiento Cocción"
            HelpContextID   =   4111210
            Index           =   210
         End
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   115
      End
      Begin VB.Menu General 
         Caption         =   "Casino"
         HelpContextID   =   4120000
         Index           =   120
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Despachos"
         HelpContextID   =   4122000
         Index           =   122
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Servicio Principal"
         HelpContextID   =   4123000
         Index           =   123
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Inventario Calendarizado"
         HelpContextID   =   4124000
         Index           =   124
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Pvta. Cliente Calendarizado"
         HelpContextID   =   4125000
         Index           =   125
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Categoría Dietética"
         HelpContextID   =   4126000
         Index           =   126
      End
      Begin VB.Menu General 
         Caption         =   "Parametro Costo Mermas"
         HelpContextID   =   4127000
         Index           =   127
      End
      Begin VB.Menu General 
         Caption         =   "Subsegmento"
         HelpContextID   =   4130000
         Index           =   130
      End
      Begin VB.Menu General 
         Caption         =   "Servicio"
         HelpContextID   =   4140000
         Index           =   140
      End
      Begin VB.Menu General 
         Caption         =   "Grupo Estructura"
         HelpContextID   =   4141000
         Index           =   141
      End
      Begin VB.Menu General 
         Caption         =   "Regimen"
         HelpContextID   =   4150000
         Index           =   150
      End
      Begin VB.Menu General 
         Caption         =   "Zona"
         HelpContextID   =   4160000
         Index           =   160
      End
      Begin VB.Menu General 
         Caption         =   "Tipo de Servicio"
         HelpContextID   =   4170000
         Index           =   170
      End
      Begin VB.Menu General 
         Caption         =   "Segmento"
         HelpContextID   =   4180000
         Index           =   180
      End
      Begin VB.Menu General 
         Caption         =   "Tipo Interfaz"
         HelpContextID   =   4190000
         Index           =   190
      End
      Begin VB.Menu General 
         Caption         =   "Tipo Actividad"
         HelpContextID   =   4192000
         Index           =   192
      End
      Begin VB.Menu General 
         Caption         =   "Municipio"
         HelpContextID   =   4194000
         Index           =   194
      End
      Begin VB.Menu General 
         Caption         =   "Región"
         HelpContextID   =   4195000
         Index           =   195
      End
      Begin VB.Menu General 
         Caption         =   "Retención en la Fuente"
         HelpContextID   =   4196000
         Index           =   196
      End
      Begin VB.Menu General 
         Caption         =   "Retención ICA"
         HelpContextID   =   4198000
         Index           =   198
      End
      Begin VB.Menu General 
         Caption         =   "Ofertas"
         HelpContextID   =   4199000
         Index           =   199
      End
      Begin VB.Menu General 
         Caption         =   "Unidad Receta"
         HelpContextID   =   4199300
         Index           =   200
      End
      Begin VB.Menu General 
         Caption         =   "Días Feriados"
         HelpContextID   =   4199400
         Index           =   201
      End
      Begin VB.Menu General 
         Caption         =   "Estacionalidad"
         HelpContextID   =   4199500
         Index           =   202
      End
      Begin VB.Menu General 
         Caption         =   "Ajuste Estacional Receta"
         HelpContextID   =   4199600
         Index           =   203
      End
      Begin VB.Menu General 
         Caption         =   "Homologación FoodUp"
         HelpContextID   =   4199700
         Index           =   204
      End
      Begin VB.Menu General 
         Caption         =   "Tipo de Mermas"
         HelpContextID   =   4199800
         Index           =   205
      End
      Begin VB.Menu General 
         Caption         =   "-"
         Index           =   206
      End
      Begin VB.Menu General 
         Caption         =   "Parámetros Sistemas"
         HelpContextID   =   4200000
         Index           =   207
      End
      Begin VB.Menu General 
         Caption         =   "Parámetros Recetas"
         HelpContextID   =   4210000
         Index           =   210
      End
      Begin VB.Menu General 
         Caption         =   "Parámetros Código Barra"
         HelpContextID   =   4215000
         Index           =   215
      End
      Begin VB.Menu General 
         Caption         =   "Parárametro SGP LOCAL"
         HelpContextID   =   4217000
         Index           =   217
      End
      Begin VB.Menu General 
         Caption         =   "Perfiles de Acceso"
         HelpContextID   =   4220000
         Index           =   220
      End
      Begin VB.Menu General 
         Caption         =   "Usuarios"
         HelpContextID   =   4230000
         Index           =   230
      End
   End
   Begin VB.Menu Main 
      Caption         =   "&Pedido Web"
      Index           =   35
      Begin VB.Menu PedidoWeb 
         Caption         =   "Mantenedor Ruta"
         HelpContextID   =   5000000
         Index           =   0
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "Calendarios Días Feriados"
         HelpContextID   =   5010000
         Index           =   10
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "Agregar Cantidad y Lista Cantidad Productos"
         HelpContextID   =   5020000
         Index           =   20
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "Reglas de Negocios"
         HelpContextID   =   5030000
         Index           =   30
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "Lista de Precios"
         HelpContextID   =   5040000
         Index           =   40
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "Definir Pedidos"
         HelpContextID   =   5050000
         Index           =   50
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "-"
         Index           =   215
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "WebToMDB"
         HelpContextID   =   5220000
         Index           =   220
      End
      Begin VB.Menu PedidoWeb 
         Caption         =   "SACToWeb"
         HelpContextID   =   5230000
         Index           =   230
      End
   End
   Begin VB.Menu Main 
      Caption         =   "Servicio Logistico"
      Index           =   37
      Begin VB.Menu ServicioLogistico 
         Caption         =   "Descuentos por Volumen"
         HelpContextID   =   6010000
         Index           =   10
      End
      Begin VB.Menu ServicioLogistico 
         Caption         =   "Indice Precio de Alimento por Periodo"
         HelpContextID   =   6020000
         Index           =   20
      End
      Begin VB.Menu ServicioLogistico 
         Caption         =   "Porcentaje Costo por Servicios"
         HelpContextID   =   6030000
         Index           =   30
      End
      Begin VB.Menu ServicioLogistico 
         Caption         =   "Precio Referencia"
         HelpContextID   =   6040000
         Index           =   40
      End
      Begin VB.Menu ServicioLogistico 
         Caption         =   "Nota Venta"
         HelpContextID   =   6050000
         Index           =   50
      End
      Begin VB.Menu ServicioLogistico 
         Caption         =   "-"
         Index           =   60
      End
      Begin VB.Menu ServicioLogistico 
         Caption         =   "Informe Servicio Logistico"
         HelpContextID   =   6070000
         Index           =   70
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
Option Explicit
Option Compare Text

Private SwSalir     As Integer
Private Inventario  As Variant
Private item        As Variant
Dim TemSeg          As Long
Dim IntMin          As Long

Private Sub ActualizarMinutas_Click(Index As Integer)

vg_OpcM = ActualizarMinutas.item(Index).HelpContextID

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

Case 0 'Actualizar Raciones & Ponderación Excel
    
    Call P_ActComExcel.Show(0, Partida)

Case 10 ' Actualizar estado minuta
    
    If Not formAbierto("P_CamEstMin") Then
       
       Dim P_CamEstMin As New P_CamEstMin
       P_CamEstMin.lc_Aux = "CamEstMin"
       P_CamEstMin.Tag = "CamEstMin"
       P_CamEstMin.Show 0, Partida
       Set P_CamEstMin = Nothing
    
    End If

Case 20 ' Actualizar % ponderación
    
    If Not formAbierto("PoncentajePonderacion") Then
       
       Dim PoncentajePonderacion As New P_CamEstMin
       PoncentajePonderacion.lc_Aux = "PoncentajePonderacion"
       PoncentajePonderacion.Tag = "PoncentajePonderacion"
       PoncentajePonderacion.Show 0, Partida
       Set PoncentajePonderacion = Nothing
    
    End If
    
Case 30 ' Cambiar Pedido Proyecta a CD o PAP
    
    If Not formAbierto("P_CambioCarro") Then
       
       Dim P_CambioCarro As New P_CambioCarro
       P_CambioCarro.lc_Aux = "P_CambioCarro"
       P_CambioCarro.Tag = "P_CambioCarro"
       P_CambioCarro.Show 0, Partida
       Set P_CambioCarro = Nothing
    
    End If
    
Case 32 ' Cambiar Estado Pedido
    
    If Not formAbierto("P_CambioCarro") Then
       
       Dim P_CambioPedido As New P_CambioCarro
       P_CambioPedido.lc_Aux = "P_CambioPedido"
       P_CambioPedido.Tag = "P_CambioPedido"
       P_CambioPedido.Show 0, Partida
       Set P_CambioPedido = Nothing
    
    End If

Case 40 ' Asignar Lista Convenios Ceco Propuesta
    
    If Not formAbierto("P_AsigListaPrecioProp") Then
       
       Dim P_AsigListaPrecioProp As New P_AsigListaPrecioProp
       P_AsigListaPrecioProp.lc_Aux = "P_AsigListaPrecioProp"
       P_AsigListaPrecioProp.Tag = "P_AsigListaPrecioProp"
       P_AsigListaPrecioProp.Show 0, Partida
       Set P_AsigListaPrecioProp = Nothing
    
    End If
    
Case 50 ' Eliminar Carros de Compras
    
    If Not formAbierto("P_EliminarCarroCompras") Then
       
       Dim P_EliminarCarroCompras As New P_EliminarCarroCompras
       P_EliminarCarroCompras.lc_Aux = "P_EliminarCarroCompras"
       P_EliminarCarroCompras.Tag = "P_EliminarCarroCompras"
       P_EliminarCarroCompras.Show 0, Partida
       Set P_EliminarCarroCompras = Nothing
    
    End If
    
Case 60 ' Cambio de recetas minuta bloque
    
    If Not formAbierto("P_CambioRecetaMinBloque") Then
       
       Dim P_CambioRecetaMinBloque As New P_CambioRecetaMinBloque
       P_CambioRecetaMinBloque.lc_Aux = "P_CambioRecetaMinBloque"
       P_CambioRecetaMinBloque.Tag = "P_CambioRecetaMinBloque"
       P_CambioRecetaMinBloque.Show 0, Partida
       Set P_CambioRecetaMinBloque = Nothing
    
    End If
    
Case 70 ' Exportar Tabla de Gramaje & Bach - Input
    
    If Not formAbierto("P_ExpTGranejeBachInput") Then
       
       Dim P_ExpTGranejeBachInput As New P_ExpTGranejeBachInput
       P_ExpTGranejeBachInput.lc_Aux = "P_ExpTGranejeBachInput"
       P_ExpTGranejeBachInput.Tag = "P_ExpTGranejeBachInput"
       P_ExpTGranejeBachInput.Show 0, Partida
       Set P_ExpTGranejeBachInput = Nothing
    
    End If
    
Case 80 ' Ajuste Estacionales Recetas
    
    If Not formAbierto("P_ActualizarAjusteEstacionales") Then
       
       Dim P_ActualizarAjusteEstacionales As New P_ActualizarAjusteEstacionales
       P_ActualizarAjusteEstacionales.lc_Aux = "P_ActualizarAjusteEstacionales"
       P_ActualizarAjusteEstacionales.Tag = "P_ActualizarAjusteEstacionales"
       P_ActualizarAjusteEstacionales.Show 0, Partida
       Set P_ActualizarAjusteEstacionales = Nothing
    
    End If
    
End Select

End Sub

Private Sub Command1_Click()

MonitoriarSitioRemoto

End Sub

Private Sub Diwo_Click(Index As Integer)

vg_OpcM = Diwo.item(Index).HelpContextID
Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

Select Case Index

    Case 10 'Parametrización ceco estructura servicio

        If Not formAbierto("M_estructuraserviciodiwo") Then
           
           Dim M_EstructuraServicioDiwo As New M_EstructuraServicioDiwo
           M_EstructuraServicioDiwo.lc_Aux = "M_estructuraserviciodiwo"
           M_EstructuraServicioDiwo.Tag = "M_estructuraserviciodiwo"
           M_EstructuraServicioDiwo.Show 0, Partida
           Set M_EstructuraServicioDiwo = Nothing
        
        End If

    Case 20 'Formato salida
    
        If Not formAbierto("I_FormatoSalidaDiwo") Then
           
           Dim I_FormatoSalidaDiwo As New I_FormatoSalidaDiwo
           I_FormatoSalidaDiwo.lc_Aux = "I_FormatoSalidaDiwo"
           I_FormatoSalidaDiwo.Tag = "I_FormatoSalidaDiwo"
           I_FormatoSalidaDiwo.Show 0, Partida
           Set I_FormatoSalidaDiwo = Nothing
        
        End If
        
End Select

End Sub

Private Sub GeneracionPedido_Click(Index As Integer)
    
    vg_OpcM = GeneracionPedido.item(Index).HelpContextID
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
    
    Select Case Index
        
        Case 0 'Generación Pedidos
            
            M_Lista_Pedido.Show 0, Partida
        
        Case 20 '------->Excepciones Formato Compra
            
            M_ForComPrexCeCo.Show 0, Partida
        
        Case 30 '-------> Producto que no Arrastran Saldo
            
            M_ProNOArrrastreSaldo.Show 0, Partida
        
        Case 40 '-------> Excluir Ingrediente en Pedido
            
            M_IngExe.Show 0, Partida
        
        Case 42
             
             M_IncIngExceso.Show 0, Partida
        
        Case 60 '-------> ParametroxCeco
            
            M_fechadespachoCecos.Show 0, Partida
        
        Case 70 '-------> ParametroxProveedor
            
            M_fecha_despachos.Show 0, Partida
        
        Case 80
        
            If Not formAbierto("Masofamsgpgrupodespacho") Then
           
               Dim Masofamsgpgrupodespacho As New M_AsoFamSGPGrupoDespacho
               Masofamsgpgrupodespacho.lc_Aux = "Masofamsgpgrupodespacho"
               Masofamsgpgrupodespacho.Tag = "Masofamsgpgrupodespacho"
               Masofamsgpgrupodespacho.Show 0, Partida
               Set Masofamsgpgrupodespacho = Nothing
        
            End If
        
'        Case 90 '-------> Configurar Despacho Ceco - Proveedor
'
'            M_Calendario_fechas_despachos.Show 0, Partida
    
        Case 90 '-------> consultar convenios pel
        
            If Not formAbierto("ConsultarConveniosPel") Then
           
               Dim ConsultarConveniosPel As New C_ConsultarActualizarConveniosPel
               ConsultarConveniosPel.lc_Aux = "ConsultarConveniosPel"
               ConsultarConveniosPel.Tag = "ConsultarConveniosPel"
               ConsultarConveniosPel.Show 0, Partida
               Set ConsultarConveniosPel = Nothing
        
            End If
    
        Case 94 '-------> mover cero saldo de arrastre
        
            If Not formAbierto("ArrastreDeSaldo") Then
           
               Dim ArrastreDeSaldo As New P_ArrastreDeSaldo
               ArrastreDeSaldo.lc_Aux = "ArrastreDeSaldo"
               ArrastreDeSaldo.Tag = "ArrastreDeSaldo"
               ArrastreDeSaldo.Show 0, Partida
               Set ArrastreDeSaldo = Nothing
        
            End If
    
    End Select

End Sub

Private Sub General_Click(Index As Integer)
    
    vg_OpcM = General.item(Index).HelpContextID
    
    Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
    
    Select Case Index
    
    Case 0 '-------> Familia producto
        
        T_FamPro.Show 0, Partida
    
    Case 10 '-------> Grupo cambio ingrediente
        
        T_GCaIng.Show 0, Partida
    
    Case 20 '------> Habilitar cambio ingrediente
        
        M_CamIRe.Show 0, Partida
    
    Case 30 '-------> Unidad de Medida
        
        T_Unimed.Show 0, Partida
    
    Case 40 '-------> Unidad de Stock
        
        T_Unienv.Show 0, Partida
    
    Case 50 '-------> Unidad de Embalaje
        
        T_Uniemb.Show 0, Partida
    
    Case 60 '-------> Nuriente
        
        T_Nutrie.Show 0, Partida
    
    Case 70 '-------> Impuestos
        
        T_Impues.Show 0, Partida
    
    Case 80 '-------> Cuenta Contable
        
        T_CtaCon.Show 0, Partida
    
    Case 90 '-------> Tipo Documento
        
        T_TipDoc.Show 0, Partida
    
    Case 100 '-------> Categoria Dietetica
        
        T_CatDie.Show 0, Partida
    
    Case 110 '-------> Tipo Plato
        
        T_TipPla.Show 0, Partida
    
    Case 120 '-------> Regimen
        
        If Not formAbierto("MCasino") Then
           
           Dim MCasino As New M_Casino
           MCasino.lc_Aux = "MCasino"
           MCasino.Tag = "MCasino"
           MCasino.Show 0, Partida
           Set MCasino = Nothing
        
        End If
    
    Case 122
        
        If Not formAbierto("MCaspde") Then
           
           Dim MCaspde As New M_Casino
           MCaspde.lc_Aux = "MCaspde"
           MCaspde.Tag = "MCaspde"
           MCaspde.Show 0, Partida
           Set MCaspde = Nothing
        
        End If
    
    Case 123 'Parametro servicios Principal
        
        If Not formAbierto("MCasppr") Then
           
           Dim MCasppr As New M_Casino
           MCasppr.lc_Aux = "MCasppr"
           MCasppr.Tag = "MCasppr"
           MCasppr.Show 0, Partida
           Set MCasppr = Nothing
        
        End If
    
    Case 124 'Parametro Inventario Calendarizado
        
        If Not formAbierto("MCaspic") Then
           
           Dim MCaspic As New M_Casino
           MCaspic.lc_Aux = "MCaspic"
           MCaspic.Tag = "MCaspic"
           MCaspic.Show 0, Partida
           Set MCaspic = Nothing
        
        End If
      
    Case 125 'Parametro Pvta. Cliente Calendarizado
        
        If Not formAbierto("MCaspcc") Then
           
           Dim MCaspcc As New M_Casino
           MCaspcc.lc_Aux = "MCaspcc"
           MCaspcc.Tag = "MCaspcc"
           MCaspcc.Show 0, Partida
           Set MCaspcc = Nothing
        
        End If
      
    Case 126 'Parametro Categoria Dietetica
        
        If Not formAbierto("MCaspcd") Then
           
           Dim MCaspcd As New M_Casino
           MCaspcd.lc_Aux = "MCaspcd"
           MCaspcd.Tag = "MCaspcd"
           MCaspcd.Show 0, Partida
           Set MCaspcd = Nothing
        
        End If
      
    Case 127 'Parametro Costo Merma
        
        If Not formAbierto("MDesPro") Then
           
           Dim MDesPro As New M_Desconche_Produccion
           MDesPro.lc_Aux = "MDesPro"
           MDesPro.Tag = "MDesPro"
           MDesPro.Show 0, Partida
           Set MDesPro = Nothing
        
        End If
      
    Case 130 '-------> Sub Segmento
        
        T_SubSeg.Show 0, Partida
    
    Case 140 '-------> Servicio
        
        T_Servic.Show 0, Partida
    
    Case 141 'Grupo Estructura
        
        If Not formAbierto("GrupoEstructura") Then
           
           Dim GrupoEstructura As New T_GrupoEstructura
           GrupoEstructura.lc_Aux = "GrupoEstructura"
           GrupoEstructura.Tag = "GrupoEstructura"
           GrupoEstructura.Show 0, Partida
           Set GrupoEstructura = Nothing
        
        End If
    
    Case 150 '-------> Regimen
        
        T_Regime.Show 0, Partida
    
    Case 160 '-------> Zona
        
        T_Zona.Show 0, Partida
    
    Case 170 '-------> Tipo Servicio
        
        T_TipSer.Show 0, Partida
    
    Case 180 '-------> Segmento
        
        T_Segmen.Show 0, Partida
    
    Case 190 '-------> Tipo Interfaz
        
        T_TipInt.Show 0, Partida
    
    Case 192 '-------> Tipo Actividad
        
        T_TipAct.Show 0, Partida
    
    Case 194 '-------> Municipio
        
        T_Munici.Show 0, Partida
    
    Case 195 '-------> Región
        
        T_Region.Show 0, Partida
    
    Case 196 '-------> Retención en la Fuente
        
        T_RetFue.Show 0, Partida
    
    Case 198 '-------> Retención ICA
        
        T_RetIca.Show 0, Partida
    
    Case 199 '-------> Ofertas
        
        M_Ofertas.Show 0, Partida
    
    Case 200 '-------> Unidad Receta
        
        M_UnidadRecetas.Show 0, Partida
    
    Case 201 '-------> Días Feriados
        
        M_ClDiaF.Show 0, Partida
    
    Case 202 '-------> Estacionalidad
        
        If Not formAbierto("T_Estacionalidad") Then
           
           Dim T_Estacionalidad As New T_Estacionalidad
           T_Estacionalidad.lc_Aux = "T_Estacionalidad"
           T_Estacionalidad.Tag = "T_Estacionalidad"
           T_Estacionalidad.Show 0, Partida
           Set T_Estacionalidad = Nothing
        
        End If
    
    Case 203 '-------> Estacionalidad
        
        If Not formAbierto("M_AjusteEstacionales") Then
           
           Dim M_AjusteEstacionales As New M_AjusteEstacionales
           M_AjusteEstacionales.lc_Aux = "M_AjusteEstacionales"
           M_AjusteEstacionales.Tag = "M_AjusteEstacionales"
           M_AjusteEstacionales.Show 0, Partida
           Set M_AjusteEstacionales = Nothing
        
        End If
    
    Case 204 '-------> Homologación Food Up
        
        If Not formAbierto("T_HomologaciónFoodUp") Then

           Dim T_HomologacionFoodUp As New T_HomologacionFoodUp
           T_HomologacionFoodUp.lc_Aux = "T_HomologacionFoodUp"
           T_HomologacionFoodUp.Tag = "T_HomologacionFoodUp"
           T_HomologacionFoodUp.Show 0, Partida
           Set T_HomologacionFoodUp = Nothing

        End If
    
    Case 205 '-------> Tipo de Mermas
        
        If Not formAbierto("T_Mermas") Then

           Dim T_Mermas As New T_Mermas
           T_Mermas.lc_Aux = "T_Mermas"
           T_Mermas.Tag = "T_Mermas"
           T_Mermas.Show 0, Partida
           Set T_Mermas = Nothing

        End If
    
    Case 207 '-------> Parámetro Web Service
        
        If Not formAbierto("PWebSe") Then
           
           Dim PWebSe As New M_Parame
           PWebSe.lc_Aux = "PWebSe"
           PWebSe.Tag = "PWebSe"
           PWebSe.Show 0, Partida
           Set PWebSe = Nothing
        
        End If
    
    Case 210 '-------> Parámetro Web Service
        
        If Not formAbierto("ParRec") Then
           
           Dim ParRec As New M_Parame
           ParRec.lc_Aux = "ParRec"
           ParRec.Tag = "ParRec"
           ParRec.Show 0, Partida
           Set ParRec = Nothing
        
        End If
    
    Case 215 '-------> Parámetro codigo barra
        
        If Not formAbierto("ParCodBarra") Then
           
           Dim ParCodBarra As New T_ParCodigoBarra
           ParCodBarra.lc_Aux = "ParCodBarra"
           ParCodBarra.Tag = "ParCodBarra"
           ParCodBarra.Show 0, Partida
           Set ParCodBarra = Nothing
        
        End If
    
    Case 217 '-------> Parámetro sgp local
        
        If Not formAbierto("Parsgplocal") Then
           
           Dim Parsgplocal As New M_Parame
           Parsgplocal.lc_Aux = "Parsgplocal"
           Parsgplocal.Tag = "Parsgplocal"
           Parsgplocal.Show 0, Partida
           Set Parsgplocal = Nothing
        
        End If
    
    Case 220 '-------> Ferfiles de Acceso
        
        M_Perfil.Show 0, Partida
    
    Case 230 '-------> Usuario
        
        M_Usuari.Show 0, Partida
    
    Case 5 '-------> 'MVA - MVI - Carga de archivos
        
        MVI_ImpArchivos.Show 0, Partida
        
    'Case 240 '-------> Cambio de Usuario
    '    Partida.Hide
    '    Partida Unload
    '
    '    V_Acceso.Show 0, Partida
    End Select
End Sub

Private Sub Informe_Click(Index As Integer)

vg_OpcM = Informe.item(Index).HelpContextID
Select Case Index

Case 0
    
    If Not formAbierto("Planif") Then
       
       Dim Planif As New I_Planif
       Planif.lc_Aux = "Planif"
       Planif.Tag = "Planif"
       Planif.Show 0, Partida
       Set Planif = Nothing
    
    End If

Case 2 ' minuta costo
    
    If Not formAbierto("MinCos") Then
       
       Dim MinCos As New I_Planif
       MinCos.lc_Aux = "MinCos"
       MinCos.Tag = "MinCos"
       MinCos.Show 0, Partida
       Set MinCos = Nothing
    
    End If

Case 3 ' Minuta Bloque
    
    If Not formAbierto("PlanifBloque") Then
       
       Dim PlanifBloque As New I_PlanifBloque
       PlanifBloque.lc_Aux = "PlanifBloque"
       PlanifBloque.Tag = "PlanifBloque"
       PlanifBloque.Show 0, Partida
       Set PlanifBloque = Nothing
    
    End If

Case 10 'Consumo Ingrediente Minuta
    
    C_ConIng.Show 0, Partida

Case 20 'Consumo Ingrediente Minuta Bloque
    
    If Not formAbierto("ConIngMinBlo") Then
       
       Dim ConIngMinBlo As New C_ConIngMinBlo
       ConIngMinBlo.lc_Aux = "ConIngMinBlo"
       ConIngMinBlo.Tag = "ConIngMinBlo"
       ConIngMinBlo.Show 0, Partida
       Set ConIngMinBlo = Nothing
    
    End If

Case 30
    
    If Not formAbierto("ExpMinBlo") Then
       
       Dim ExpMinBlo As New I_ExpMinBlo
       ExpMinBlo.lc_Aux = "ExpMinBlo"
       ExpMinBlo.Tag = "ExpMinBlo"
       ExpMinBlo.Show 0, Partida
       Set ExpMinBlo = Nothing
    
    End If

Case 40
    
    If Not formAbierto("I_ExpDetMinBloque") Then
       
       Dim I_ExpDetMinBloque As New I_ExpDetMinBloque
       I_ExpDetMinBloque.lc_Aux = "I_ExpDetMinBloque"
       I_ExpDetMinBloque.Tag = "I_ExpDetMinBloque"
       I_ExpDetMinBloque.Show 0, Partida
       Set I_ExpDetMinBloque = Nothing
    
    End If

Case 50 'Exportar excel detalle minuta II
    
    If Not formAbierto("E_PlanMinuta") Then
       
       Dim E_PlanMinuta As New E_PlanMinuta
       E_PlanMinuta.lc_Aux = "E_PlanMinuta"
       E_PlanMinuta.Tag = "E_PlanMinuta"
       E_PlanMinuta.Show 0, Partida
       Set E_PlanMinuta = Nothing

    End If

Case 52 'Template Minuta Bloque
    
    If Not formAbierto("E_TemplateMinI") Then
       
       Dim E_TemplateMinI As New E_TemplateMinI
       E_TemplateMinI.lc_Aux = "E_TemplateMinI"
       E_TemplateMinI.Tag = "E_TemplateMinI"
       E_TemplateMinI.Show 0, Partida
       Set E_TemplateMinI = Nothing

    End If

Case 60 'Historial Trabajos por Lotes
    
    If Not formAbierto("E_TrabajosPorLotes") Then
       
       Dim E_TrabajosPorLotes As New E_TrabajosPorLotes
       E_TrabajosPorLotes.lc_Aux = "E_TrabajosPorLotes"
       E_TrabajosPorLotes.Tag = "E_TrabajosPorLotes"
       E_TrabajosPorLotes.Show 0, Partida
       Set E_TrabajosPorLotes = Nothing
    
    End If

Case 70 'Exportar excel Q Sitios
    
    If Not formAbierto("E_QSitios") Then
       
       Dim E_QSitios As New E_QSitios
       E_QSitios.lc_Aux = "E_QSitios"
       E_QSitios.Tag = "E_QSitios"
       E_QSitios.Show 0, Partida
       Set E_QSitios = Nothing
    
    End If

Case 90 'Aportes Nutricional Sansis
    
    If Not formAbierto("I_ApoNutSansis") Then
       
       Dim I_ApoNutSansis As New I_ApoNutSansis
       I_ApoNutSansis.lc_Aux = "I_ApoNutSansis"
       I_ApoNutSansis.Tag = "I_ApoNutSansis"
       I_ApoNutSansis.Show 0, Partida
       Set I_ApoNutSansis = Nothing
    
    End If

Case 100 'Planificación Minutas Sansis
    
    If Not formAbierto("I_SetPlaSansis") Then
       
       Dim I_SetPlaSansis As New I_SetPlaSansis
       I_SetPlaSansis.lc_Aux = "I_SetPlaSansis"
       I_SetPlaSansis.Tag = "I_SetPlaSansis"
       I_SetPlaSansis.Show 0, Partida
       Set I_SetPlaSansis = Nothing
    
    End If

Case 110 'Frecuencia De Recetas o Gramos Producto Mensual
    
    If Not formAbierto("I_fregrp") Then
       
       Dim I_FreGrP As New I_FreGrP
       I_FreGrP.lc_Aux = "I_fregrp"
       I_FreGrP.Tag = "I_fregrp"
       I_FreGrP.Show 0, Partida
       Set I_FreGrP = Nothing
    
    End If

Case 120 'Composición Minutas Sansis
    
    If Not formAbierto("E_ComposicionMinutasSansis") Then
       
       Dim E_ComposicionMinutasSansis As New E_ComposicionMinutasSansis
       E_ComposicionMinutasSansis.lc_Aux = "E_ComposicionMinutasSansis"
       E_ComposicionMinutasSansis.Tag = "E_ComposicionMinutasSansis"
       E_ComposicionMinutasSansis.Show 0, Partida
       Set E_ComposicionMinutasSansis = Nothing
    
    End If

Case 122 'Costo Minutas Sansis
    
    If Not formAbierto("I_CostoSansis") Then
       
       Dim I_CostoSansis As New I_CostoSansis
       I_CostoSansis.lc_Aux = "I_CostoSansis"
       I_CostoSansis.Tag = "I_CostoSansis"
       I_CostoSansis.Show 0, Partida
       Set I_CostoSansis = Nothing
    
    End If

Case 140 'Exportar Excel Varios
    
    If Not formAbierto("E_ExcelVarios") Then
       
       Dim E_ExcelVarios As New E_ExcelVarios
       E_ExcelVarios.lc_Aux = "E_ExcelVarios"
       E_ExcelVarios.Tag = "E_ExcelVarios"
       E_ExcelVarios.Show 0, Partida
       Set E_ExcelVarios = Nothing
    
    End If

Case 160 'Exportar Excel Coosto Planificado Teorico - Real - Realizado
    
    If Not formAbierto("I_CostoPlanTeoRealRealizado") Then
       
       Dim I_CostoPlanTeoRealRealizado As New I_CostoPlanTeoRealRealizado
       I_CostoPlanTeoRealRealizado.lc_Aux = "I_CostoPlanTeoRealRealizado"
       I_CostoPlanTeoRealRealizado.Tag = "I_CostoPlanTeoRealRealizado"
       I_CostoPlanTeoRealRealizado.Show 0, Partida
       Set I_CostoPlanTeoRealRealizado = Nothing
    
    End If

Case 170 'Exportar Excel Costo Merma Desconche
    
    If Not formAbierto("E_CostoMermaSitio") Then
       
       Dim E_CostoMermaSitio As New E_CostoMermaSitio
       E_CostoMermaSitio.lc_Aux = "E_CostoMermaSitio"
       E_CostoMermaSitio.Tag = "E_CostoMermaSitio"
       E_CostoMermaSitio.Show 0, Partida
       Set E_CostoMermaSitio = Nothing
    
    End If

Case 180 'Exportar Excel Ingrediente No Poseen Precio Vigente
    
    If Not formAbierto("E_PrecioIngredienteNoVigente") Then
       
       Dim E_PrecioIngredienteNoVigente As New E_PrecioIngredienteNoVigente
       E_PrecioIngredienteNoVigente.lc_Aux = "E_PrecioIngredienteNoVigente"
       E_PrecioIngredienteNoVigente.Tag = "E_PrecioIngredienteNoVigente"
       E_PrecioIngredienteNoVigente.Show 0, Partida
       Set E_PrecioIngredienteNoVigente = Nothing
    
    End If

Case 190 'Exportar Excel So Health
    
    If Not formAbierto("C_SoHealth") Then
       
       Dim C_SoHealth As New C_SoHealth
       C_SoHealth.lc_Aux = "C_SoHealth"
       C_SoHealth.Tag = "C_SoHealth"
       C_SoHealth.Show 0, Partida
       Set C_SoHealth = Nothing
    
    End If

End Select

End Sub

Private Sub Inventario_Click(Index As Integer)
    
    vg_OpcM = Inventario.item(Index).HelpContextID
    
    Select Case Index
        
        Case 0
            
            M_Provee.Show 0, Partida
'        Case 10
'            M_DocPro.Show 0, Partida
'        Case 20
'            M_SalBod.Show 0, Partida
'        Case 30
'            M_DevBod.Show 0, Partida
'        Case 40
'            M_Traspa.Show 0, Partida
'        Case 50
'            M_Mermas.Show 0, Partida
'        Case 60
'            M_VenDir.Show 0, Partida
'        Case 70 '--- Toma de Inventario
'            M_TomInv.Show 0, Partida
'        Case 80 '--- Control de Raciones Administrador
'            M_ConRac.Show 0, Partida
'        Case 85 '---Precio Venta Cliente
'            M_PVtaCl.Show 0, Partida
'        Case 90 '--- Control Facturas Compras
'            I_CtrFCo.Inicio "Control Facturas Compras", "C"
'            I_CtrFCo.Show 0, Partida
'        Case 100 '--- Control traspasos entre casinos
'            I_CtrFCo.Inicio "Control Traspasos Entre Casinos", "T"
'            I_CtrFCo.Show 0, Partida
'        Case 110 '--- Control Fondo Fijo (Fofi)
'            I_CtrFCo.Inicio "Control Fondo Fijo (Fofi)", "F"
'            I_CtrFCo.Show 0, Partida
'        Case 120 '--- Resultado Operacionales Mensual o A13
'            I_A13.Show 0, Partida
    
    End Select

End Sub

Private Sub Main_Click(Index As Integer)

Select Case Index

Case 40
    
    MDIForm_Unload (0)

End Select

End Sub

Private Sub MDIForm_Activate()

vg_opimp = 0

End Sub

Private Sub MDIForm_Load()

Dim i   As Long
Dim RS1 As New ADODB.Recordset

On Error GoTo ManError

TemSeg = 0
IntMin = 10
Me.Caption = "Administrador Base de Datos SGPADM v " & Trim(Str(App.Major)) & "." & Trim(Str(App.Minor)) & "." & Trim(Str(App.Revision))
StatusBar1.Panels(3).text = "Servidor : " & Trim(vg_SqlNSvr) & " "
StatusBar1.Panels(4).text = "Base : " & Trim(vg_SqlBase) & " "
StatusBar1.Panels(5).text = "Usuario : " & Trim(vg_NUsr) & " "
StatusBar1.Panels(6).text = "Tipo Acceso : " & IIf(vg_Indppr = "1", "Real", IIf(vg_Indppr = "2", "Propuesta", "Ambos"))
StatusBar1.Panels(7).text = "Pais : """
StatusBar1.Panels(8).MinWidth = 7000
StatusBar1.Panels(8).text = "Dir. Trabajo : " & Trim(dir_trabajo) & " "

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_s_pais 1, '" & vg_pais & "', ''")
If Not RS1.EOF Then
   
   StatusBar1.Panels(7).text = "Pais : " & Trim(RS1!pai_nombre) & " "

End If
RS1.Close: Set RS1 = Nothing

Dim imgX As ListImage
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
Set imgX = IL1.ListImages.Add(, "A_ExporReceta", LoadResPicture(150, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Buscar", LoadResPicture(151, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ImportarPrecio", LoadResPicture(152, vbResIcon))
Set imgX = IL1.ListImages.Add(, "Calorias", LoadResPicture(153, vbResIcon))
Set imgX = IL1.ListImages.Add(, "Vinculo Ingrediente", LoadResPicture(154, vbResIcon))
Set imgX = IL1.ListImages.Add(, "Ingrediente", LoadResPicture(155, vbResIcon))
Set imgX = IL1.ListImages.Add(, "Proceso", LoadResPicture(156, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_ImportarDatos", LoadResPicture(157, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Agregar", LoadResPicture(158, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Retrocede", LoadResPicture(159, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Avanza", LoadResPicture(160, vbResIcon))
Set imgX = IL1.ListImages.Add(, "A_Calendario", LoadResPicture(161, vbResIcon))

Set imgX = IL1.ListImages.Add(, "A_Reporcesar", LoadResPicture(102, vbResBitmap))
Set imgX = IL1.ListImages.Add(, "A_VerErrores", LoadResPicture(103, vbResBitmap))
Set imgX = IL1.ListImages.Add(, "A_VerConvenio", LoadResPicture(104, vbResBitmap))
Set imgX = IL1.ListImages.Add(, "A_Alzas", LoadResPicture(105, vbResBitmap))

For Each item In Minutas
    
    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In Informe
    
    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In General
    
    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In PedidoWeb
    
    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In GeneracionPedido
    
    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In Diwo

    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In PanLed

    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In ActualizarMinutas
    
    If item.Caption <> "-" Then item.Visible = False

Next

For Each item In ParametroReceta

    If item.Caption <> "-" Then item.Visible = False

Next

AbrirBaseWebPed
Main.item(0).Visible = False
Main.item(20).Visible = False
Main.item(30).Visible = False
Main.item(35).Visible = False
Main.item(37).Visible = False

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS1 = vg_db.Execute("sgpadm_Sel_CargaUsuarioPerfil '" & vg_NUsr & "'")
Do While Not RS1.EOF
    
    Select Case Mid(Trim(Str(RS1!dpe_codopc)), 1, 1)
    
    Case 1
        
        For Each item In Minutas
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
            
                item.Visible = True
                Main.item(0).Visible = True
                Exit For
            
            End If
        
        Next
        
        For Each item In ActualizarMinutas
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
               
               item.Visible = True
               Minutas.item(139).Visible = True
               Exit For
            
            End If
        
        Next
        
        For Each item In GeneracionPedido
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
               
               item.Visible = True
               Minutas.item(190).Visible = True
               Exit For
            
            End If
        
        Next

        For Each item In Diwo
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
               
               item.Visible = True
               Main.item(0).Visible = True
               Minutas.item(199).Visible = True
               Minutas.item(200).Visible = True
               Exit For
            
            End If
        
        Next

        For Each item In PanLed
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
               
               item.Visible = True
               Main.item(0).Visible = True
               Minutas.item(201).Visible = True
               Minutas.item(220).Visible = True
               Exit For
            
            End If
        
        Next
    
    Case 2
        
        For Each item In Informe
        
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
               
               item.Visible = True
               Main.item(20).Visible = True
               Exit For
        
            End If
            
        Next
    
    Case 4
        
        For Each item In General
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
            
               item.Visible = True
               Main.item(30).Visible = True
               Exit For
            
            End If
            
        Next
    
        For Each item In ParametroReceta
            
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then
               
               item.Visible = True
               Main.item(0).Visible = True
               General.item(112).Visible = True
'               Minutas.item(1).Visible = True
               Exit For
            
            End If
        
        Next
    
    Case 5
       
       If vg_estopen And Trim(vg_SqlBaseW) <> "" Then
        
            For Each item In PedidoWeb
                
                If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then item.Visible = True: Main.item(35).Visible = True: Exit For
            
            Next
       
       End If
    
    Case 6
        
        For Each item In ServicioLogistico
        
            If item.HelpContextID = RS1!dpe_codopc And RS1!dpe_deracc = 1 Then item.Visible = True: Main.item(37).Visible = True: Exit For
        
        Next
    
    End Select
    
    RS1.MoveNext

Loop
RS1.Close
Set RS1 = Nothing
'-------> Mover cantidad decimales en precio y cantidad
vg_DPr = 2

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_Parametros 'parpredec'")

If Not RS1.EOF Then vg_DPr = RS1!par_valor
RS1.Close
Set RS1 = Nothing
vg_DCa = 2

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_Parametros 'parcandec'")

If Not RS1.EOF Then vg_DCa = RS1!par_valor
RS1.Close
Set RS1 = Nothing
vg_RDCa = 2

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_Parametros 'parrcandec'")

If Not RS1.EOF Then vg_RDCa = RS1!par_valor
RS1.Close
Set RS1 = Nothing

'-------> Traer parametro calculo digito verificador rut
vg_Dig = "S"
If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
Set RS1 = vg_db.Execute("sgpadm_Sel_Parametros 'parcaldig'")

If Not RS1.EOF Then vg_Dig = RS1!par_valor
RS1.Close
Set RS1 = Nothing

MonitoriarSitioRemoto
Main.item(40).Visible = True
Let VarSitioRemoto = False

Exit Sub
ManError:
If Err.Number = 340 Then Resume Next
MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, Me.Caption

End Sub

Private Sub MDIForm_Resize()

Fondo.Width = Me.Width - 120
Command1.Top = ScaleHeight - 2550
vaSpread1.Top = ScaleHeight - 2000 'IIf(Me.WindowState = 2, 11300, ScaleHeight - 2000)
vaSpread1.Left = Fondo.Left
'vaSpread1.Height = Fondo.Height
vaSpread1.Width = Fondo.Width

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    Dim msg, Response   ' Declara variables.
    If SwSalir = 1 Then Exit Sub
    msg = "¿Esta Seguro Salir?"
    Response = MsgBox(msg, 4 + 32, "Sistema Gestión")
    
    Select Case Response
      
      Case 2  ' No permite cerrar.
        
        Cancel = -1
        msg = "El comando ha sido cancelado."
      
      Case 6
        
        SwSalir = 1
        'vg_db.Close
        Me.Hide
        Unload Me
        End
      
      Case 7
        
        Cancel = -1
    
    End Select

End Sub

Private Sub Minutas_Click(Index As Integer)
    
On Error GoTo Man_Error
    
    vg_OpcM = Minutas.item(Index).HelpContextID
    
    Select Case Index
        
        Case 0 '-------> Maestro producto

            M_Produc.Show 0, Partida
            
        Case 2 '-------> Generar Ing. - Nut. - Rec. - Minuta OPTIMUM
            
            P_GenInfOpt.Show 0, Partida
        
        Case 10 '-------> Asociar productos sac vs sgp
            
            If Not formAbierto("SacSgp") Then
               
               Dim SacSgp As New M_SacSgp
               SacSgp.lc_Aux = "SacSgp"
               SacSgp.Tag = "SacSgp"
               SacSgp.Show 0, Partida
               Set SacSgp = Nothing
            
            End If

        Case 12 '-------> Asociar productos sap vs sgp
            
            If Not formAbierto("SapSgp") Then
               
               Dim SapSgp As New M_SacSgp
               SapSgp.lc_Aux = "SapSgp"
               SapSgp.Tag = "SapSgp"
               SapSgp.Show 0, Partida
               Set SapSgp = Nothing
            
            End If
        
        Case 20 '-------> Maestro proveedor
            
            M_Provee.Show 0, Partida
        
        Case 30 '-------> Generación archivos planos productos
            
            If Not formAbierto("GenPro") Then
               
               Dim GenPro As New M_GenPro
               GenPro.lc_Aux = "SacSgp"
               GenPro.Tag = "SacSgp"
               GenPro.Show 0, Partida
               Set GenPro = Nothing
            
            End If

        Case 32 '-------> generación archivos planos listas precios sac
            
            M_GenLpr.Show 0, Partida

        Case 40 '-------> Lista Precio
            
            M_LisPre.Show 0, Partida
        
        Case 45
            If Not formAbierto("B_CostoIngrediente") Then
               
               Dim CostoIngrediente As New P_CostoIngrediente
               CostoIngrediente.lc_Aux = "CostoIngrediente"
               CostoIngrediente.Tag = "CostoIngrediente"
               CostoIngrediente.Show 0, Partida
               Set CostoIngrediente = Nothing
            
            End If
        
        Case 50 '-------> Importar lista de precio desde sac
            
            M_ILpSac.Show 0, Partida
        
        Case 60 '-------> Importar lista de precio desde excel
            
            M_ImLprE.Show 0, Partida
        
        Case 70 '-------> Asociar lista precio
            
            M_AsoLPr.Show 0, Partida
        
        Case 80 '-------> Actualizar lista percio en planificación
            
            M_ActPrP.Show 0, Partida
        
        Case 90 '-------> Recetas
            
            vg_newcodrec = 0
            vg_newnomrec = ""
            vg_newestrec = False
            vg_modreceta = False
            vg_fecha = ""
            vg_PartePlani = False
            
            M_Receta.Show 0, Partida
        
        Case 100 '-------> Generación de archivos planos recetas

            M_GenRec.Show 0, Partida
        
        Case 110 '-------> Planificación de minutas
            
            VarSitioRemoto = False
            M_Plami1.Partidas "Planificación Minutas", "MINTEO"
            M_Plami1.Show 0, Partida
        
        Case 111 '-------> Planificación de minutas
            
            MsgBox "Esta opción esta deshabilitada", vbCritical + vbOKOnly, "Sistema Administrador SGP"
            'M_Plami3.Show 0, Partida
        
        Case 112 '-------> Minuta Real Casino
            
            C_IMiRCa.Show 0, Partida
        
        Case 120 '-------> Tabla gramaje
            
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
            
            M_TabGra.Show 0, Partida
        
        Case 122 '-------> Copia Minuta Lideeres Seguidores
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
            
            If Not formAbierto("M_Copia_Minuta_Lideres") Then
               
               Dim M_Copia_Minuta_Lideres As New M_Copia_Minuta_Lideres
               M_Copia_Minuta_Lideres.lc_Aux = "M_Copia_Minuta_Lideres"
               M_Copia_Minuta_Lideres.Tag = "M_Copia_Minuta_Lideres"
               M_Copia_Minuta_Lideres.Show 0, Partida
               Set M_Copia_Minuta_Lideres = Nothing
            
            End If
        
        Case 130 '-------> Generación de archivos plano planificaión
            
            M_GenPla.Show 0, Partida
        
        Case 132 '-------> Borrar Sitios Masivos
        
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
            
            If Not formAbierto("M_BorrarCecoMasivo") Then
               
               Dim M_BorrarCecoMasivo As New M_BorrarCecoMasivo
               M_BorrarCecoMasivo.lc_Aux = "M_BorrarCecoMasivo"
               M_BorrarCecoMasivo.Tag = "M_BorrarCecoMasivo"
               M_BorrarCecoMasivo.Show 0, Partida
               Set M_BorrarCecoMasivo = Nothing
            
            End If
        
        Case 133 '-------> copia minuta bloque estandar x ceco
            
            If Not formAbierto("Copia_MinutaBloqueEstandar") Then
               
               Dim Copia_MinutaBloqueEstandar As New M_Copia_minutaBloqueEstandar
               Copia_MinutaBloqueEstandar.lc_Aux = "Copia_MinutaBloqueEstandar"
               Copia_MinutaBloqueEstandar.Tag = "Copia_MinutaBloqueEstandar"
               Copia_MinutaBloqueEstandar.Show 0, Partida
               Set Copia_MinutaBloqueEstandar = Nothing
            
            End If
        
        Case 134 '-------> Costo Bandeja x servicio minuta bloque
            
            If Not formAbierto("MBloqueCostoBandejaxServicios") Then
               
               Dim MBloqueCostoBandejaxServicios As New M_MBloqueCostoBandejaxServicios
               MBloqueCostoBandejaxServicios.lc_Aux = "MBloqueCostoBandejaxServicios"
               MBloqueCostoBandejaxServicios.Tag = "MBloqueCostoBandejaxServicios"
               MBloqueCostoBandejaxServicios.Show 0, Partida
               Set MBloqueCostoBandejaxServicios = Nothing
            
            End If
        
        Case 135 '-------> copia minuta bloque x ceco
            
            If Not formAbierto("Copia_MinutaBloqueCeco") Then
               
               Dim Copia_MinutaBloqueCeco As New M_Copia_MinutaBloqueCeco
               Copia_MinutaBloqueCeco.lc_Aux = "Copia_MinutaBloqueCeco"
               Copia_MinutaBloqueCeco.Tag = "Copia_MinutaBloqueCeco"
               Copia_MinutaBloqueCeco.Show 0, Partida
               Set Copia_MinutaBloqueCeco = Nothing
            
            End If
        
        Case 136 '-------> minuta sitio remoto
            
            Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")
            
            Call M_MinSR1.Show(0, Partida)
        
        Case 137 '-------> Envio minuta sitio remoto
            
            Call M_EnMinRem.Show(0, Partida)
        
        Case 138 '-------> Liveracion minuta sitio remoto
            
            Call M_LibMsr.Show(0, Partida)
        
        Case 141 '-------> Parametrizar n. recetas 5 etapas
            
            M_PNRece.Show 0, Partida
        
        Case 150 '-------> parametrizar costo patron piso 5 etapas
            
            If Not formAbierto("CosCom") Then
               
               Dim CosCom As New M_CostosSitios
               CosCom.lc_Aux = "CosCom"
               CosCom.Tag = "CosCom"
               CosCom.Show 0, Partida
               Set CosCom = Nothing
            
            End If
        
        Case 160 '------> parametrizar costo patron techo 5 etapas
            
            If Not formAbierto("CosTec") Then
               
               Dim CosTec As New M_PCPatr
               CosTec.lc_Aux = "CosTec"
               CosTec.Tag = "CosTec"
               CosTec.Show 0, Partida
               Set CosTec = Nothing
            
            End If
        
        Case 170 '-------> Gramaje familia producto 5 etapas
            M_GraFaP.Show 0, Partida
    
    End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub PanLed_Click(Index As Integer)

On Error GoTo Man_Error

   vg_OpcM = PanLed.item(Index).HelpContextID

   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

    Select Case Index
            
        Case 10 'Parametrización ceco estructura servicio

            If Not formAbierto("M_EstructuraServicioPanLed") Then
           
               Dim M_EstructuraServicioPanLed As New M_EstructuraServicioPanLed
               M_EstructuraServicioPanLed.lc_Aux = "M_estructuraservicioPanLed"
               M_EstructuraServicioPanLed.Tag = "M_estructuraservicioPanLed"
               M_EstructuraServicioPanLed.Show 0, Partida
              Set M_EstructuraServicioPanLed = Nothing
        
           End If

    End Select
    
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub ParametroReceta_Click(Index As Integer)

On Error GoTo Man_Error
    
   vg_OpcM = ParametroReceta.item(Index).HelpContextID

   Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Acceso"), CStr(vg_OpcM), "", "", "")

    Select Case Index
            
            Case 10
            
                If Not formAbierto("Color") Then
                   
                   Dim Color As New T_Color
                   Color.lc_Aux = "Color"
                   Color.Tag = "Color"
                   Color.Show 0, Partida
                   Set Color = Nothing
                
                End If
        
            Case 20
            
                If Not formAbierto("CostoReceta") Then
                   
                   Dim CostoReceta As New T_CostoReceta
                   CostoReceta.lc_Aux = "CostoReceta"
                   CostoReceta.Tag = "CostoReceta"
                   CostoReceta.Show 0, Partida
                   Set CostoReceta = Nothing
                
                End If
        
            Case 30
            
                If Not formAbierto("TipoIngPrincipalReceta") Then
                   
                   Dim TipoIngPrincipalReceta As New T_TipoIngPrincipalReceta
                   TipoIngPrincipalReceta.lc_Aux = "TipoIngPrincipalReceta"
                   TipoIngPrincipalReceta.Tag = "TipoIngPrincipalReceta"
                   TipoIngPrincipalReceta.Show 0, Partida
                   Set TipoIngPrincipalReceta = Nothing
                
                End If
    
            Case 40
            
                If Not formAbierto("MetodococcionReceta") Then
                   
                   Dim MetodococcionReceta As New T_MetodoCoccionReceta
                   MetodococcionReceta.lc_Aux = "MetodococcionReceta"
                   MetodococcionReceta.Tag = "MetodococcionReceta"
                   MetodococcionReceta.Show 0, Partida
                   Set MetodococcionReceta = Nothing
                
                End If
    
            Case 50
            
                If Not formAbierto("CategoriaComplejaReceta") Then
                   
                   Dim CategoriaComplejaReceta As New T_CategoriaComplejaReceta
                   CategoriaComplejaReceta.lc_Aux = "CategoriaComplejaReceta"
                   CategoriaComplejaReceta.Tag = "CategoriaComplejaReceta"
                   CategoriaComplejaReceta.Show 0, Partida
                   Set CategoriaComplejaReceta = Nothing
                
                End If
    
            Case 60

                If Not formAbierto("IngCruceGarnituraReceta") Then

                   Dim IngCruceGarnituraReceta As New T_IngCruceGarnituraReceta
                   IngCruceGarnituraReceta.lc_Aux = "IngCruceGarnituraReceta"
                   IngCruceGarnituraReceta.Tag = "IngCruceGarnituraReceta"
                   IngCruceGarnituraReceta.Show 0, Partida
                   Set IngCruceGarnituraReceta = Nothing

                End If
    
            Case 70

                If Not formAbierto("EstacionalidadReceta") Then

                   Dim EstacionalidadReceta As New T_EstacionalidadReceta
                   EstacionalidadReceta.lc_Aux = "EstacionalidadReceta"
                   EstacionalidadReceta.Tag = "EstacionalidadReceta"
                   EstacionalidadReceta.Show 0, Partida
                   Set EstacionalidadReceta = Nothing

                End If
    
            Case 80
            
                If Not formAbierto("EfectoMeteorizanteReceta") Then
                   
                   Dim EfectoMeteorizanteReceta As New T_EfectoMeteorizanteReceta
                   EfectoMeteorizanteReceta.lc_Aux = "EfectoMeteorizanteReceta"
                   EfectoMeteorizanteReceta.Tag = "EfectoMeteorizanteReceta"
                   EfectoMeteorizanteReceta.Show 0, Partida
                   Set EfectoMeteorizanteReceta = Nothing
                
                End If
    
            Case 90
            
                If Not formAbierto("SellosReceta") Then
                   
                   Dim SellosReceta As New T_SellosReceta
                   SellosReceta.lc_Aux = "SellosReceta"
                   SellosReceta.Tag = "SellosReceta"
                   SellosReceta.Show 0, Partida
                   Set SellosReceta = Nothing
                
                End If
    
            Case 100
            
                If Not formAbierto("TipoNegocioReceta") Then
                   
                   Dim TipoNegocioReceta As New T_TipoNegocioReceta
                   TipoNegocioReceta.lc_Aux = "TipoNegocioReceta"
                   TipoNegocioReceta.Tag = "TipoNegocioReceta"
                   TipoNegocioReceta.Show 0, Partida
                   Set TipoNegocioReceta = Nothing
                
                End If
    
            Case 110
            
                If Not formAbierto("IntoleranciaReceta") Then
                   
                   Dim IntoleranciaReceta As New T_IntoleranciaReceta
                   IntoleranciaReceta.lc_Aux = "IntoleranciaReceta"
                   IntoleranciaReceta.Tag = "IntoleranciaReceta"
                   IntoleranciaReceta.Show 0, Partida
                   Set IntoleranciaReceta = Nothing
                
                End If
    
            Case 120
            
                If Not formAbierto("TiempoHHReceta") Then
                   
                   Dim TiempoHHReceta As New T_TiempoHHReceta
                   TiempoHHReceta.lc_Aux = "TiempoHHReceta"
                   TiempoHHReceta.Tag = "TiempoHHReceta"
                   TiempoHHReceta.Show 0, Partida
                   Set TiempoHHReceta = Nothing
                
                End If
    
            Case 130
            
                If Not formAbierto("TiempoCoccionReceta") Then
                   
                   Dim TiempoCoccionReceta As New T_TiempoCoccionReceta
                   TiempoCoccionReceta.lc_Aux = "TiempoCoccionReceta"
                   TiempoCoccionReceta.Tag = "TiempoCoccionReceta"
                   TiempoCoccionReceta.Show 0, Partida
                   Set TiempoCoccionReceta = Nothing
                
                End If
       
            Case 140
            
                If Not formAbierto("Alergeno") Then
                   
                   Dim Alergeno As New T_Alergeno
                   Alergeno.lc_Aux = "Alergeno"
                   Alergeno.Tag = "Alergeno"
                   Alergeno.Show 0, Partida
                   Set Alergeno = Nothing
                
                End If
       
            Case 150
            
                If Not formAbierto("EstiloAlimentacion") Then
                   
                   Dim EstiloAlimentacion As New T_EstiloAlimentacion
                   EstiloAlimentacion.lc_Aux = "EstiloAlimentacion"
                   EstiloAlimentacion.Tag = "EstiloAlimentacion"
                   EstiloAlimentacion.Show 0, Partida
                   Set EstiloAlimentacion = Nothing
                
                End If
       
            Case 160
            
                If Not formAbierto("ParametroAdicional1") Then
                   
                   Dim ParametroAdicional1 As New T_ParametroAdicional1
                   ParametroAdicional1.lc_Aux = "ParametroAdicional1"
                   ParametroAdicional1.Tag = "ParametroAdicional1"
                   ParametroAdicional1.Show 0, Partida
                   Set ParametroAdicional1 = Nothing
                
                End If
       
            Case 170
            
                If Not formAbierto("ParametroAdicional2") Then
                   
                   Dim ParametroAdicional2 As New T_ParametroAdicional2
                   ParametroAdicional2.lc_Aux = "ParametroAdicional2"
                   ParametroAdicional2.Tag = "ParametroAdicional2"
                   ParametroAdicional2.Show 0, Partida
                   Set ParametroAdicional2 = Nothing
                
                End If
       
            Case 180
            
                If Not formAbierto("EtiquetadoSelloReceta") Then
                   
                   Dim EtiquetadoSelloReceta As New T_EtiquetadoSelloReceta
                   EtiquetadoSelloReceta.lc_Aux = "EtiquetadoSelloReceta"
                   EtiquetadoSelloReceta.Tag = "EtiquetadoSelloReceta"
                   EtiquetadoSelloReceta.Show 0, Partida
                   Set EtiquetadoSelloReceta = Nothing
                
                End If
       
            Case 190 ' Grupo Ingrediente Principal
            
                If Not formAbierto("GrupoIngPrincipal") Then
                   
                   Dim GrupoIngPrincipal As New T_GrupoIngPrincipal
                   GrupoIngPrincipal.lc_Aux = "GrupoIngPrincipal"
                   GrupoIngPrincipal.Tag = "GrupoIngPrincipal"
                   GrupoIngPrincipal.Show 0, Partida
                   Set GrupoIngPrincipal = Nothing
                
                End If
       
            Case 200 ' Parametro Salsa
            
                If Not formAbierto("ParametroSalsa") Then
                   
                   Dim ParametroSalsa As New T_ParametroSalsa
                   ParametroSalsa.lc_Aux = "ParametroSalsa"
                   ParametroSalsa.Tag = "ParametroSalsa"
                   ParametroSalsa.Show 0, Partida
                   Set ParametroSalsa = Nothing
                
                End If
       
            Case 210 ' Equipamiento Cocción
            
                If Not formAbierto("EquipamientoCoccion") Then
                   
                   Dim EquipamientoCoccion As New T_EquipamientoCoccion
                   EquipamientoCoccion.lc_Aux = "EquipamientoCoccion"
                   EquipamientoCoccion.Tag = "EquipamientoCoccion"
                   EquipamientoCoccion.Show 0, Partida
                   Set EquipamientoCoccion = Nothing
                
                End If
    
    End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub PedidoWeb_Click(Index As Integer)

'-------> Abrir base sac
    AbrirBaseWebPed
    
    If Not vg_estopen Then
       
       MsgBox "     No esta autorizado de ingresar a esta opciones de sistema" & Chr(13) & "     Hubo problema con la abertura de la base dato (WEBREPORTING)." & Chr(13) & Chr(13) & "Itentelo en unos minutos más tarde ó bien comuniquese con deparamento Informatica.", vbCritical + vbOKOnly, "Sistema Administrador SGP"
       Exit Sub
    
    End If
    
    vg_OpcM = PedidoWeb.item(Index).HelpContextID
    
    Select Case Index
        
        Case 0 '-------> Maestro Ruta
            
            M_Ruta.Show 0, Partida
        
        Case 10 '-------> Calendarios Días Feriados.
            
            M_ClDiaF.Show 0, Partida
        
        Case 20 '-------> Agregar Cantidad y Lista Cantidad Productos
            
            M_ACProd.Show 0, Partida
        
        Case 30 '-------> Reglas de Negocios
            
            M_RegNeg.Show 0, Partida
        
        Case 40 '-------> Lista de precios
            
            M_ListPreWeb.Show 0, Partida
        
        Case 50 '-------> Definir pedidos
            
            M_DefPed.Show 0, Partida
        
        Case 220 '-------> WebToMDB
            
            M_WebToMDB.Show 0, Partida
        
        Case 230 '-------> SacToWeb
            
            M_SacToWeb.Show 0, Partida
    
    End Select

End Sub

Private Sub ServicioLogistico_Click(Index As Integer)

vg_OpcM = ServicioLogistico.item(Index).HelpContextID

Select Case Index

    Case 10 'Descuentos x Volumen
        
        M_SsllDxv.Show , Partida
    
    Case 20 'Indice Precio de Alimento por Periodo
        
        M_SsllIPA.Show , Partida
    
    Case 30 'Costo Servicio
        
        M_SsllPorCosServ.Show , Partida
    
    Case 40 'Referencia Precio
        
        M_SsllPrecioRef.Show , Partida
    
    Case 50 'Nota Venta
        
        M_ssllNotaVenta.Show , Partida
    
    Case 70 'Consolidado Facturación Clientes
        
        If Not formAbierto("SsllCFacCli") Then
           
           Dim SsllCFacCli As New I_SsLlGen
           SsllCFacCli.lc_Aux = "SsllCFacCli"
           SsllCFacCli.Tag = "SsllCFacCli"
           SsllCFacCli.Show 0, Partida
           Set SsllCFacCli = Nothing
        
        End If

End Select

End Sub

'Private Sub Timer1_Timer()
'' variable estática para acumular la cantidad de segundos
''Static Temp_Seg As Long
'' incrementa
'TemSeg = TemSeg + 1
'' comprueba que los segundos no sea igual a la cantidad de minutos _
'  que queremos , en este caso 3 minutos
'If (TemSeg * 30) >= (IntMin * 30) * 30 Then
'   ' reestablece
'   MonitoriarSitioRemoto
'   TemSeg = 0
'End If
'End Sub

Private Sub MonitoriarSitioRemoto()

Dim RS As New ADODB.Recordset
'-------> Sitio Remoto
'Set RS1 = vg_db.Execute("sgpadm_s_logenviominsr")
'                           " WHERE FecPro >= '" & Now & ".000'"
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Sel_MonitorearSitiosRemotos")
vaSpread1.MaxRows = 0
vaSpread1.Visible = False
Do While Not RS.EOF
   
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   
   vaSpread1.Col = 1
   Let vaSpread1.text = RS!fecpros
   
   vaSpread1.Col = 2
   Let vaSpread1.text = RS!fecrec
   
   vaSpread1.Col = 3
   Let vaSpread1.text = Trim(RS!cencos) & " - " & Trim(RS!Cli_nombre) & " - " & Trim(RS!estado) & " - " & Trim(RS!Mensaje) '& VgLinea
   
   RS.MoveNext

Loop

RS.Close
Set RS = Nothing
vaSpread1.Visible = True

End Sub
