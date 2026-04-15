VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form M_TipUsu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo Usuario"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8805
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Listar Usuarios"
      TabPicture(0)   =   "M_TipUsu.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vaSpread1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Usuarios"
      TabPicture(1)   =   "M_TipUsu.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   971
         Left            =   -72360
         TabIndex        =   2
         Top             =   480
         Width           =   7335
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "M_TipUsu.frx":0038
            Left            =   2010
            List            =   "M_TipUsu.frx":004B
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   2500
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Left            =   2010
            TabIndex        =   4
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
            TabIndex        =   7
            Top             =   645
            Width           =   585
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
            TabIndex        =   6
            Top             =   645
            Width           =   1140
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
            TabIndex        =   5
            Top             =   345
            Width           =   1380
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Acceso Módulo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7575
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   11775
         Begin MSComctlLib.TreeView TvwDir 
            Height          =   7095
            Left            =   240
            TabIndex        =   9
            Top             =   360
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   12515
            _Version        =   393217
            LineStyle       =   1
            Style           =   7
            Checkboxes      =   -1  'True
            Appearance      =   1
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6285
         Left            =   -74040
         TabIndex        =   8
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
         MaxCols         =   2
         MaxRows         =   499
         ProcessTab      =   -1  'True
         SpreadDesigner  =   "M_TipUsu.frx":0077
         ScrollBarTrack  =   3
         ClipboardOptions=   0
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   12510
      _ExtentX        =   22066
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
End
Attribute VB_Name = "M_TipUsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, Msgtitulo As String
Dim rootNode As Node
Dim est As Boolean, OpGr As Boolean
Dim codigo As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 9285
Me.Width = 12615
Msgtitulo = "Tipo Usuarios"
fg_centra Me
SSTab1.Tab = 0
modo = ""
OpGr = False
est = True
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 1, modo
MoverDatosGrilla
CargarDatosTipoOpcion
CargarTipoUsuario
Gl_Ac_Botones Me, 1, IIf(vaSpread1.MaxRows > 0, 1, 2), modo
End Sub

Sub MoverDatosGrilla()
fg_carga ""
OpGr = True
vaSpread1.Visible = False
vaSpread1.MaxRows = 0
vaSpread1.Row = -1
vaSpread1.Col = 1
vaSpread1.Lock = True
Set RS = vg_dbpedweb.Execute("pedweb_s_tipodeusuarios 2, '', ''")
Do While Not RS.EOF
   vaSpread1.MaxRows = vaSpread1.MaxRows + 1
   vaSpread1.Row = vaSpread1.MaxRows
   vaSpread1.Col = 1
   vaSpread1.text = Trim(RS!tipo_usuario)
   vaSpread1.Col = 2
   vaSpread1.text = Trim(IIf(IsNull(RS!descripcion), "", RS!descripcion))
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
vaSpread1.Visible = True
If vaSpread1.MaxRows > 0 Then
   vaSpread1.Row = 1
   vaSpread1.Col = 1
   codigo = ""
   codigo = vaSpread1.text
   vaSpread1.SetActiveCell 1, 1
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
OpGr = False
fg_descarga
End Sub

Sub CargarDatosTipoOpcion()
fg_carga ""
TvwDir.Nodes.Clear
Set rootNode = TvwDir.Nodes.Add(, , "RPEDIDOS     ", "PEDIDOS")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PEDIDOS      " & "H" & "CREAR PEDIDO", "Crear Pedido"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PEDIDOS      " & "H" & "CONSULTAR PEDIDO", "Consultar Pedido"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PEDIDOS      " & "H" & "MODIFICAR PEDIDO", "Modificar Pedido"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PEDIDOS      " & "H" & "PEDIDO ADICIONAL", "Pedido Adicional"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PEDIDOS      " & "H" & "ANULAR PEDIDO", "Anular Pedido"
'TvwDir.Nodes.Add rootNode.Index, tvwChild, "PEDIDOS      " & "H" & "Pedido Adicional", "Pedido Adicional"
Set rootNode = TvwDir.Nodes.Add(, , "RGUIAS      ", "GUIAS")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "GUIAS        " & "H" & "VER CHECKLIST", "Ver Checklist"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "GUIAS        " & "H" & "CONSULTAR GUIAS", "Consultar Guias"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "GUIAS        " & "H" & "RECIBIR GUIA", "Recibir Guias"
Set rootNode = TvwDir.Nodes.Add(, , "RPROD. PRECIOS", "PROD. PRECIOS")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PROD. PRECIOS" & "H" & "LISTA DE PRECIOS", "Lista de Precios"
'TvwDir.Nodes.Add rootNode.Index, tvwChild, "PROD. PRECIOS" & "H" & "Prod. CD anterior", "Prod. CD anterior"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PROD. PRECIOS" & "H" & "PRODUCTO", "Prod. CD anterior"
Set rootNode = TvwDir.Nodes.Add(, , "RRUTA        ", "RUTA")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "DEFINIR RUTA", "Definir Ruta"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "MANTIENE RUTA", "Definir Ruta"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "DEFINE DESPACHOS", "Define Despachos"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "COPIA RUTA", "Copia Ruta"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "COPIA CASINOS/RUTA", "Copia Casinos/Ruta"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "CONSULTA RUTA", "Consulta Ruta"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "RUTA         " & "H" & "FERIADOS", "Feriados"
Set rootNode = TvwDir.Nodes.Add(, , "RPRODUCTOS    ", "PRODUCTOS")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PRODUCTOS    " & "H" & "AGREGA CANTIDADES", "Agrega Cantidades"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PRODUCTOS    " & "H" & "LISTA CANTIDADES", "Lista Cantidades"
Set rootNode = TvwDir.Nodes.Add(, , "RDATOS        ", "DATOS")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "DATOS        " & "H" & "CASINOS", "Casinos"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "DATOS        " & "H" & "UNIDAD MEDIDA", "Unidad Medida"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "DATOS        " & "H" & "FAMILIA DE PROD.", "Familia de Prod."
TvwDir.Nodes.Add rootNode.Index, tvwChild, "DATOS        " & "H" & "PRODUCTOS RAÍZ", "Productos Raíz"
'TvwDir.Nodes.Add rootNode.Index, tvwChild, "DATOS        " & "H" & "PRODUCTO", "Productos"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "DATOS        " & "H" & "PROVEEDORES", "Proveedores"
Set rootNode = TvwDir.Nodes.Add(, , "RINFORMES     ", "INFORMES")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "INFORMES     " & "H" & "PEDIDOS PENDIENTES", "Pedidos Pendientes"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "INFORMES     " & "H" & "PEDIDOS GENERALES", "Pedidos Generales"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "INFORMES     " & "H" & "CONTROL PRODUCTOS", "Control Productos"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "INFORMES     " & "H" & "RESUMEN PRODUCTOS", "Resumen Productos"
'TvwDir.Nodes.Add rootNode.Index, tvwChild, "INFORMES     " & "H" & "Informe Casinos", "Informe Casinos"
Set rootNode = TvwDir.Nodes.Add(, , "RSISTEMA      ", "SISTEMA")
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "TIPOS DE USUARIO", "Tipos de Usuario"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "CREAR USUARIO", "Crea Usuario"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "MODIFICAR USUARIO", "Modificar Usuario"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "CAMBIAR CLAVE", "Cambiar Clave"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "DEFINE PEDIDOS", "Define Pedidos"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "SUBE DATA", "Sube Data"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "BAJA DATA", "Baja Data"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "BLOQUEAR SISTEMA", "Bloquear Sistema"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "REGLAS DE NEGOCIOS", "Reglas de Negocios"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "SISTEMA      " & "H" & "LISTAS DE PRECIOS", "Lista de Precios"
Set rootNode = TvwDir.Nodes.Add(, , "RPRECIOS      ", "PRECIOS")
'TvwDir.Nodes.Add rootNode.Index, tvwChild, "PRECIOS      " & "H" & "Listas de Precios", "Listas de Precios"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PRECIOS      " & "H" & "ASIGNAR CONTRATOS", "Asignar Contratos"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PRECIOS      " & "H" & "MANTIENE LISTAS", "Mantiene Listas"
TvwDir.Nodes.Add rootNode.Index, tvwChild, "PRECIOS      " & "H" & "VER PRECIOS", "Ver Precios"
'TvwDir.Nodes.Item(dest.Key).Selected = True
fg_descarga
End Sub

Sub CargarTipoUsuario()
Dim i As Long, j As Long, Nombre As String
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
codigo = vaSpread1.text
vaSpread1.Col = 2
Nombre = vaSpread1.text
Frame3.Visible = False
TvwDir.Visible = False
Frame3.Caption = "Acceso Módulo "
Frame3.Caption = Frame3.Caption & " " & codigo & " - " & Nombre
Set RS = vg_dbpedweb.Execute("pedweb_s_tipodeusuarios 4, '" & codigo & "', ''")
Do While Not RS.EOF
'    For i = 1 To TvwDir.Nodes.count
'        If TvwDir.Nodes.Item(i).Checked = True And InStr(TvwDir.Nodes.Item(i).Key, "CASINO") <> 0 Then nEst = True: Exit For
'    Next
   For i = 1 To TvwDir.Nodes.count
       If UCase(Mid(Trim(TvwDir.Nodes.item(i).Key), 15, 90)) = Trim(UCase(RS!modulo)) Then
          TvwDir.Nodes.item(i).Checked = True
          TvwDir.Nodes.item(i - 1).Checked = True
          Exit For
       End If
   Next i
   RS.MoveNext
Loop
RS.Close: Set RS = Nothing
Frame3.Visible = True
TvwDir.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Select Case SSTab1.Tab
Case 0
Case 1
    CargarDatosTipoOpcion
    CargarTipoUsuario
End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    SSTab1.Tab = 0
    modo = "A"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.TabEnabled(1) = False
    vaSpread1.MaxRows = vaSpread1.MaxRows + 1
    vaSpread1.Row = vaSpread1.MaxRows
    vaSpread1.Col = 1: vaSpread1.Lock = False
    vaSpread1.Col = 2: vaSpread1.Lock = False
    vaSpread1.SetActiveCell 1, vaSpread1.MaxRows: vaSpread1.SetFocus
Case 3
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
    SSTab1.Tab = 1
Case 5
    If vaSpread1.ActiveRow < 1 Then MsgBox "Debe seleccionar un registro...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    If MsgBox("Elimina registro...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    vaSpread1.Row = vaSpread1.ActiveRow
    vaSpread1.Col = 1: codigo = vaSpread1.text
    vg_dbpedweb.Execute ("pedweb_d_tipousuarios 2, " & codigo & "")
    vaSpread1.DeleteRows vaSpread1.Row, 1
    vaSpread1.MaxRows = vaSpread1.MaxRows - 1
    Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
Case 7
    MoverDatosGrilla
Case 10
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
    If modo = "A" Then
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.DeleteRows vaSpread1.Row, 1
        vaSpread1.MaxRows = vaSpread1.MaxRows - 1
        SSTab1.Tab = 0
        MoverDatosGrilla
    Else
        OpGr = True
        vaSpread1.Row = vaSpread1.ActiveRow
        vaSpread1.Col = 1: codigo = vaSpread1.Value
        Select Case SSTab1.Tab
        Case 0
            MoverDatosGrilla
        Case 1
            CargarDatosTipoOpcion
            CargarTipoUsuario
        End Select
        OpGr = False
    End If
    Gl_Ac_Botones Me, 1, 1, modo
    Combo1.Enabled = True: fpText1.Enabled = True
Case 12
    If SSTab1.Tab = 0 Then
       GrabaRegistro vaSpread1.ActiveRow
    Else
        Dim modulo As String
        OpGr = True
        vaSpread1.Row = vaSpread1.ActiveRow: vaSpread1.Col = 1: codigo = vaSpread1.Value
        vg_dbpedweb.Execute ("pedweb_d_tipousuarios 3, " & codigo & "")
        For i = 1 To TvwDir.Nodes.count
            If (UCase(Mid(Trim(TvwDir.Nodes.item(i).Key), 14, 1)) = "H" Or UCase(Mid(Trim(TvwDir.Nodes.item(i).Key), 1, 1)) = "R") And TvwDir.Nodes.item(i).Checked = True Then
               modulo = IIf(UCase(Mid(Trim(TvwDir.Nodes.item(i).Key), 14, 1)) = "H", UCase(Mid(Trim(TvwDir.Nodes.item(i).Key), 15, 90)), UCase(Mid(Trim(TvwDir.Nodes.item(i).Key), 2, 14)))
               vg_dbpedweb.Execute ("pedweb_i_tipoaccesousuarios " & codigo & ", '" & modulo & "' ")
            End If
        Next i
        CargarDatosTipoOpcion
        CargarTipoUsuario
        modo = "A": Gl_Ac_Botones Me, 1, 1, modo
        OpGr = False
        SSTab1.TabEnabled(0) = True
        SSTab1.TabEnabled(1) = True
        SSTab1.Tab = 0
    End If
Case 15
    If vaSpread1.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'    If SSTab1.Tab = 0 Then
'       I_TipoDeUsuarios
'    Else
'       I_acceso
'    End If
Case 18
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub TvwDir_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim lCheck As Boolean, lCheck1 As Boolean
Dim i As Long
fg_carga ""
If modo = "" Then modo = "M"
TvwDir.Nodes.item(Node.Key).Selected = True
lCheck = TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Checked
lCheck1 = TvwDir.Nodes.item(TvwDir.SelectedItem.Index).Checked

If TvwDir.SelectedItem.Children > 0 Then
   For i = TvwDir.SelectedItem.Index + 1 To TvwDir.Nodes.count
       If TvwDir.Nodes.item(i).Children > 0 Then Exit For
       TvwDir.Nodes.item(i).Checked = lCheck1
   Next i
Else
   For i = TvwDir.SelectedItem.Index + 1 To TvwDir.Nodes.count
       If TvwDir.Nodes.item(i).Children > 0 Then Exit For
       If TvwDir.Nodes.item(i).Checked = True Then lCheck1 = True 'TvwDir.Nodes.Item(i).Checked: Exit For
   Next i
   For i = (TvwDir.SelectedItem.Index - 1) To 1 Step -1
       If TvwDir.Nodes.item(i).Children > 0 Then
          TvwDir.Nodes.item(i).Checked = lCheck1
          Exit For
       ElseIf TvwDir.Nodes.item(i).Checked = True Then
          Exit For
       End If
   Next i
End If
SSTab1.TabEnabled(0) = False: SSTab1.TabEnabled(1) = False: Gl_Ac_Botones Me, 1, 0, modo
fg_descarga
End Sub

Private Sub vaSpread1_EditChange(ByVal Col As Long, ByVal Row As Long)
If modo = "" Then modo = "M"
Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub vaSpread1_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
If Not OpGr And Row <> NewRow And NewRow > 0 And (modo = "A" Or modo = "M") And Toolbar1.Buttons(12).Visible = True Then
   GrabaRegistro Row
End If
End Sub

Private Sub GrabaRegistro(Fila As Long)
Dim codigo As String, Nombre As String
OpGr = True
vaSpread1.Row = Fila
vaSpread1.Col = 1: codigo = Trim(LimpiaDato(vaSpread1.Value))
vaSpread1.Col = 2: Nombre = Trim(LimpiaDato(vaSpread1.Value))
If Trim(Nombre) = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread1.Row = Fila: vaSpread1.Col = 2: vaSpread1.SetActiveCell 2, vaSpread1.MaxRows: vaSpread1.SetFocus: OpGr = False: Exit Sub
If modo = "A" Then
   vg_dbpedweb.Execute ("pedweb_iu_tiposusuarios 'A', '" & codigo & "', '" & Nombre & "'")
   vaSpread1.Col = 1: vaSpread1.text = codigo
Else
   vg_dbpedweb.Execute ("pedweb_iu_tiposusuarios 'M', '" & codigo & "', '" & Nombre & "'")
End If
Label2.Caption = Format(vaSpread1.MaxRows, fg_Pict(7, 0)) & " Registro"
Combo1.Enabled = True: fpText1.Enabled = True
modo = ""
Gl_Ac_Botones Me, 1, 1, modo
OpGr = False
SSTab1.TabEnabled(0) = True
SSTab1.TabEnabled(1) = True
End Sub

Private Sub Cancela()
Dim codigo As Long
OpGr = True
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1: codigo = vaSpread1.text
Set RS1 = vg_dbpedweb.Execute("pedweb_s_tipodeusuarios 3, '" & codigo & "', ''")
DoEvents
If Not RS1.EOF Then
   vaSpread1.Col = 2: vaSpread1.Value = IIf(IsNull(RS1!descripcion), "", Trim(RS1!descripcion))
End If
RS1.Close: Set RS1 = Nothing
OpGr = False
End Sub
