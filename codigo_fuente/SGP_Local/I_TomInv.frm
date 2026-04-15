VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form I_TomInv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Toma de Inventario"
   ClientHeight    =   3345
   ClientLeft      =   2025
   ClientTop       =   2730
   ClientWidth     =   9090
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3345
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2835
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9090
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1305
         Left            =   4605
         TabIndex        =   12
         Top             =   1275
         Width           =   4350
         Begin VB.CheckBox Check1 
            Caption         =   "Incluir productos con Stock Sistema cero"
            Height          =   270
            Index           =   2
            Left            =   135
            TabIndex        =   15
            Top             =   930
            Width           =   4065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Incluir productos con Stock Físico cero"
            Height          =   270
            Index           =   1
            Left            =   135
            TabIndex        =   14
            Top             =   585
            Width           =   4065
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Solo productos con diferencias"
            Height          =   270
            Index           =   0
            Left            =   135
            TabIndex        =   13
            Top             =   240
            Width           =   4065
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         Caption         =   "Tipo Listado"
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   165
         TabIndex        =   3
         Top             =   285
         Width           =   4350
         Begin VB.OptionButton optTIPLIS 
            Caption         =   "Diferencias Físico v/s Sistema - Valorizado"
            Height          =   225
            Index           =   4
            Left            =   165
            TabIndex        =   11
            Top             =   1920
            Width           =   4020
         End
         Begin VB.OptionButton optTIPLIS 
            Caption         =   "Listado de inventario Sistema Valorizado"
            Height          =   225
            Index           =   3
            Left            =   165
            TabIndex        =   10
            Top             =   1530
            Width           =   3915
         End
         Begin VB.OptionButton optTIPLIS 
            Caption         =   "Listado de inventario Físico Valorizado"
            Height          =   225
            Index           =   2
            Left            =   165
            TabIndex        =   9
            Top             =   1140
            Width           =   3915
         End
         Begin VB.OptionButton optTIPLIS 
            Caption         =   "Listado de diferencias Físico v/s Sistema"
            Height          =   225
            Index           =   1
            Left            =   165
            TabIndex        =   8
            Top             =   750
            Width           =   3915
         End
         Begin VB.OptionButton optTIPLIS 
            Caption         =   "Listado para la toma de inventario"
            Height          =   225
            Index           =   0
            Left            =   165
            TabIndex        =   7
            Top             =   360
            Width           =   3915
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Familia Producto"
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   4605
         TabIndex        =   2
         Top             =   285
         Width           =   4350
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Todas"
            Height          =   225
            Index           =   1
            Left            =   3330
            TabIndex        =   6
            Top             =   255
            Width           =   855
         End
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Una Familia"
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   5
            Top             =   255
            Width           =   1500
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   0
            ItemData        =   "I_TomInv.frx":0000
            Left            =   135
            List            =   "I_TomInv.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   540
            Width           =   4035
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "I_TomInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset

Private Sub Form_Load()
Me.Width = 9210
Me.Height = 3645
Msgtitulo = "Imprimir Toma de Inventario"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(0).Clear
RS1.Open "SELECT * FROM a_tipopro ORDER BY tip_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo1(0).AddItem RS1!tip_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!tip_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
optTIPLIS(0).Value = True
optTIPPRO(1).Value = True
'Check1(1).Value = 1
'Check1(2).Value = 1
End Sub

Private Sub optTIPPRO_Click(Index As Integer)
    Combo1(0).Enabled = IIf(Index = 0, True, False)
    Combo1(0).ListIndex = IIf(Index = 0, 0, -1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim Sql As String, v_fecinv  As Variant, v_codbod As Long, sqlTP As String, sqlOT As String, optorden As String
fg_carga ""
v_codbod = fg_codigocbo(M_TomInv.Combo1, 0, 10, 0)
v_fecinv = Format(M_TomInv.Date1(0).text, "yyyymmdd")
Select Case Button.Index
Case 1
    sqlTP = "": optorden = ""
    optorden = IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), " ORDER BY pro.pro_ctacon, pro.pro_nombre", " ORDER BY pro.pro_ctacon, pro.pro_codtip, pro.pro_nombre")
    If optTIPPRO(0).Value = True Then sqlTP = "and pro.pro_codtip=" & Val(fg_codigocbo(Combo1, 0, 10, 0)) & " "
    sqlOT = ""
    If Check1(0).Value = 1 Then sqlOT = sqlOT & "and tin.tin_stosis<>tin.tin_stofis "
    If Check1(1).Value = 0 And Check1(2).Value = 0 Then sqlOT = sqlOT & "and (tin.tin_stofis<>0 or tin.tin_stosis<>0) "
    If Check1(1).Value = 1 And Check1(2).Value = 0 Then sqlOT = sqlOT & "and not (tin.tin_stofis<>0 and tin.tin_stosis=0) "
    If Check1(1).Value = 0 And Check1(2).Value = 1 Then sqlOT = sqlOT & "and not (tin.tin_stofis=0 and tin.tin_stosis<>0) "
    Sql = "SELECT pro.pro_ctacon, cta.cta_nombre, pro.pro_codtip, tin.tin_codpro, pro.pro_nombre, tin.tin_propon, uni.uni_nomcor, tin.tin_stosis, tin.tin_stofis " & _
          "FROM b_tomainv tin, b_productos pro, a_unidad uni, a_ctacontable cta WHERE tin.tin_codpro=pro.pro_codigo " & _
          "AND pro.pro_coduni=uni.uni_codigo AND pro.pro_ctacon=cta.cta_codigo AND tin.tin_fectom=" & v_fecinv & " AND tin.tin_codbod=" & v_codbod & " " & sqlTP & sqlOT & optorden
    If optTIPLIS(0).Value = True Then I_Toma1 Sql, IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
    If optTIPLIS(1).Value = True Then I_Toma2 Sql, IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
    If optTIPLIS(2).Value = True Then I_Toma3 Sql, IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
    If optTIPLIS(3).Value = True Then I_Toma4 Sql, IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
    If optTIPLIS(4).Value = True Then I_Toma5 Sql, IIf(0 = (fg_CambiaChar(GetParametro("opfampro"), ";", "','")), False, True)
Case 3
    Me.Hide
    Unload Me
End Select
fg_descarga
End Sub
