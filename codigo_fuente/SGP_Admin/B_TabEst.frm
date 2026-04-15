VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Begin VB.Form B_TabEst 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   6120
   ClientLeft      =   3645
   ClientTop       =   3165
   ClientWidth     =   6840
   ControlBox      =   0   'False
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
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   6840
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   0
      Width           =   5160
      Begin VB.OptionButton Option1 
         Caption         =   "Propuesta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3360
         TabIndex        =   9
         Top             =   1000
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Ambos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   1000
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Real"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   1000
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   555
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "B_TabEst.frx":0000
         Left            =   1680
         List            =   "B_TabEst.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   90
         TabIndex        =   5
         Top             =   660
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   " Buscar Columna"
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
         Left            =   90
         TabIndex        =   4
         Top             =   345
         Width           =   1440
      End
   End
   Begin FPSpread.vaSpread vaSpread1 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   5175
      _Version        =   393216
      _ExtentX        =   9128
      _ExtentY        =   7646
      _StockProps     =   64
      ButtonDrawMode  =   2
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
      MaxRows         =   1
      OperationMode   =   3
      SelectBlockOptions=   0
      SpreadDesigner  =   "B_TabEst.frx":001E
      ScrollBarTrack  =   3
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   4  'Align Right
      Height          =   6120
      Left            =   6300
      TabIndex        =   6
      Top             =   0
      Width           =   540
      _ExtentX        =   953
      _ExtentY        =   10795
      ButtonWidth     =   1138
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00D9D9FF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   1
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "B_TabEst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim RS1         As New ADODB.Recordset
Dim i           As Long
Dim ibusca      As Long
Dim CategoriDie As Long
Dim TipoPlato   As Long  ' estas variables son solo para traer la categoria dietetica y el tipo de plato desde el formulario M_TabGra cuando llama al proc "LlenaDatos"
Dim Tabla       As String
Dim Suf         As String
Dim Opx         As String
Dim Titulo      As String
Dim SqlText     As String
Dim iCombo      As Integer

Dim BtnX        As Variant
Dim KeyAscii    As Variant
Dim IRow        As Long

Private Sub Combo1_Click()
    
On Error GoTo Man_Error

If iCombo = 0 Then Text1.text = ""

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Man_Error

Select Case KeyCode
    
    Case 27
        Cerrar

End Select

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Form_Activate()
    
On Error GoTo Man_Error

Call fg_descarga

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error
    
    Call fg_centra(Me)
    Me.Width = 5790
    Me.Left = vg_left
    fg_carga ""
    iCombo = 1
    Combo1.ListIndex = 1
    'LlenaDatos
    Toolbar1.ImageList = Partida.IL1
    Toolbar1.Buttons.Clear
    Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = True: BtnX.ToolTipText = "Confirmar "
    Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
    Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
    iCombo = 0
    Call fg_descarga
    Exit Sub

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Text1_Change

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Public Sub Text1_Change()

On Error GoTo Man_Error

Dim RS1       As New ADODB.Recordset
Dim Sql       As String
Dim z         As Long
Dim Activo    As String
Dim Est       As Boolean
Dim AuxIndppr As String

AuxIndppr = vg_Indppr
vg_Indppr = IIf(Option1(0).Value = True, 3, IIf(Option1(1).Value = True, 1, 2))
    
    If LimpiaDato(Trim(Text1.text)) & Chr(KeyAscii) = "" Then
        
        Exit Sub
    
    End If
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient
    
    If Combo1.ItemData(Combo1.ListIndex) = 0 Then
        
        If Opx = "AgregarIngxReceta" Then
        
            vaSpread1.maxcols = 3
            vaSpread1.Row = 0
            vaSpread1.Col = 3
            vaSpread1.text = "Tipo"
            Me.Width = 6930
            Frame1.Left = 600
            vaSpread1.Width = 6150
            Set RS1 = vg_db.Execute("sgpadm_Sel_ListaIngRecetaCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "Gen" Then 'Or Opx = "ProInv") Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo")
        
       ElseIf Opx = "GenUFacIng" Then  ' Unidad Factor Ingrediente
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_UnidadConversionIngredienteCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "homestser" Then
       
           Set RS1 = vg_db.Execute("SELECT DISTINCT id_homologacionestservicio, Descripcion FROM a_homologacionestservicio WHERE UPPER(id_homologacionestservicio) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY id_homologacionestservicio")
        
       ElseIf Opx = "GrpIngPri" Then  ' Grupo Ingrediente Principal
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_BuscarCodigoGrupoIngPrincipalActivo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "TipIngPri" Then  ' Tipo Ingrediente Principal x codigo
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBuscaTipoIngPrincipalRecetaxCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "MetCocc" Then  ' Metodo Cocción x codigo
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaMetodoCoccionRecetaxCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
       
       ElseIf Opx = "IngCruGar" Then  ' Ingrediente Cruce Garnitura x codigo
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaIngCruceGarnituraRecetaxCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "TiempoHH" Then  ' Tiempo HH x codigo
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaTiempoHhRecetaxCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
         
       ElseIf Opx = "Color" Then  ' Color x codigo
        
             Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaColoxCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "TiempoCoccion" Then  ' Tiempo Cocción x codigo
        
             Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaTiempoCoccionRecetaxCodigo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "GrpEstru" Then  ' Grupo Estructura
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBuscarCodigoGrupoEstructura '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "EstSer" Then
          
            Set RS1 = vg_db.Execute("sgpadm_Sel_ListaEstServicioCodigo " & Suf & ", '%" & UCase(LimpiaDato(Text1.text)) & "%', '1'")
        
        ElseIf Opx = "Celo" Then
            
            Sql = ""
            Sql = "sgpadm_Sel_BuscarOrgCompras "
            Sql = Sql & " '" & UCase(LimpiaDato(Text1.text)) & "' "
            Set RS1 = vg_db.Execute(Sql)
        
        ElseIf Opx = "CecoGral" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 44, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Zon" Then 'Or Opx = "ProInv") Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' and zon_activo = '1' ORDER BY " & Suf & "codigo")
        
        ElseIf Opx = "clientesimap" Then 'formato compras sap
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_Cliente 3, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "ProveedorSimap" Then 'Busca los proveedores
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_proveedor 3, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "ForComSap" Then 'formato compras sap
            
            Set RS1 = vg_db.Execute("sgpadm_s_productossap 3, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "CasReg" Then 'tabla regimen casino
            
            Set RS1 = vg_db.Execute("sgpadm_s_casregimen 2, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "CasSer" Then 'tabla servicio casino
            
            Set RS1 = vg_db.Execute("sgpadm_s_casservicio 2, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "CecoPortalElec" Then ' clientes con portal electronico
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 37, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cli5Eta" Then 'Or Opx = "ProInv") Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 3, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "recorrido" Then  'tabla recorrido
            
            Set RS1 = vg_dbpedweb.Execute("SELECT recorrido, descripcion FROM s_Recorrido WHERE recorrido LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY recorrido")
        
        ElseIf Opx = "regneg" Then  'tabla recorrido
            
            Set RS1 = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 2, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "lispreweb" Then  'tabla recorrido
            
            Set RS1 = vg_dbpedweb.Execute("pedweb_s_listaprecios 4, '', '%" & UCase(LimpiaDato(Text1.text)) & "%', ''")
        
        ElseIf (Opx = "Ser" Or Opx = "Sub") Then
               
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre, " & Suf & "indppr FROM " & Tabla & " WHERE " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' = '3') ORDER BY " & Suf & "codigo")
     
        ElseIf Opx = "Reg" Then
               
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre, " & Suf & "indppr FROM " & Tabla & " WHERE " & Suf & "activo = '1' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' = '3') ORDER BY " & Suf & "codigo")
        
        ElseIf Opx = "RegBlo" Then
            
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_Sel_RegimenBloquexCodigo '%" & Sql & "%'")
        
        ElseIf Opx = "SerBlo" Then
            
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioBloquexCodigo '%" & Sql & "%'")
        
        ElseIf Opx = "0" Or Opx = "1" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE (" & Suf & "tiprec = '" & Opx & "' OR '" & Opx & "' = '') AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo")
        
        ElseIf Opx = "Casino" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=0 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Traspaso" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=2 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Cliente" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=1 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "CentCost" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=0 AND " & Suf & "activo=1 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "SsllListProv" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "activo=0 AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "SsllListFormComp" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codsac, " & Suf & "nomsac FROM " & Tabla & " WHERE (" & Suf & "flexec=0 OR (" & Suf & "flexec = -1 AND " & Suf & "vigfin > " & Format(Date, "yyy/mm/dd") & ")) AND " & Suf & "codsac IN (select distinct fcs_codsac from b_formatocomprassgp) AND " & Suf & "codsac LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Proing" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_productosing WHERE " & Suf & "codigo=b_productosing.pri_coding AND b_productosing.pri_codpro='" & vg_codigo & "' AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Ingrec" Then
            
            Set RS1 = vg_db.Execute("SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_receta, b_recetadet WHERE b_receta.rec_codigo=b_recetadet.red_codigo AND " & Suf & "codigo=b_recetadet.red_codpro AND (b_receta.rec_catdie=" & vg_filcatdie & " OR " & vg_filcatdie & "=0) AND (b_receta.rec_tippla=" & vg_filtippla & " OR " & vg_filtippla & "=0) AND " & Suf & "codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo") '
        
        ElseIf Opx = "LisPre" Then
            
            Activo = IIf(Opx = "LisPre", "('0','1')", "('1')")
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_Sel_ListaPrecioxCodigo_V02 '%" & Sql & "%'")
        
        ElseIf Opx = "ProdActi" Then
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_productos 24, '" & vg_auxcod + vg_Indppr & "', '%" & UCase(LimpiaDato(Text1.text)) & "%', '" & vg_NUsr & "'")
        
        ElseIf Opx = "ForCom" Then
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_productos 20, 0, '%" & UCase(LimpiaDato(Text1.text)) & "%', '" & vg_NUsr & "'")
        
        ElseIf Opx = "SacCco" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_saccentrocosto 1, '" & vg_codigo & "', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cliente_SitioRemotoI" Then 'Minuta sitio remoto
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 22, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Regimen_SitioRemotoI" Then 'Minuta sitio remoto
            
            Set RS1 = vg_db.Execute("sgpadm_s_casregimen 5, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Servicio_SitioRemotoI" Then 'Minuta sitio remoto
            
            Set RS1 = vg_db.Execute("sgpadm_s_casservicio 4, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cliente_SitioRemoto" Then 'Minuta sitio remoto
        
            If swEsCopia = True And SeleccTipoMinuta_MVI = "Segmento" And swUp_CECO = True Then
                
               Sql = " SELECT DISTINCT a.sub_codigo, a.sub_nombre"
               Sql = Sql & " FROM a_subsegmento a "
               Sql = Sql & " LEFT outer JOIN b_detlistaprecio b on a.sub_codigo = b.dlp_codigo "
               Sql = Sql & " WHERE sub_activo = 1"
               Sql = Sql & " AND a.sub_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
               Sql = Sql & " ORDER BY a.sub_nombre"
               Set RS1 = vg_db.Execute(Sql)
            
            Else
               
               Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 22, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
            End If
        
        ElseIf Opx = "Cliente_EnvioBloque" Then 'Minuta envio minuta bloque
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 38, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cliente_CopiaMinutaBloque" Then 'Minuta envio minuta bloque
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 40, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Regimen_SitioRemoto" Then 'Minuta sitio remoto
             
             'esta seleccion es selecc. SEGMENTO (TABLA B_...)
            If SeleccTipoMinuta_MVI = "Segmento" Or SeleccTipoMinuta_MVI = "" Then
                
               Set RS1 = vg_db.Execute("sgpadm_s_casregimen 5, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
            Else
                
               Sql = " SELECT DISTINCT"
               Sql = Sql & " reg_codigo, reg_nombre  "
               Sql = Sql & " FROM cas_a_regimen reg With(NoLock)  "
               Sql = Sql & " WHERE  reg_activo = '1'  "
               Sql = Sql & " and upper(reg_codigo) LIKE '%" & Text1.text & "%'"
               Sql = Sql & " and reg_cecori = '" & IIf(swUp_CECO = True, m_copia_min_seg.fpText, m_copia_min_seg.fpText1) & "'"
               Sql = Sql & " ORDER BY reg_nombre"
               Set RS1 = vg_db.Execute(Sql)
            
            End If
        
        'MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
        ElseIf Opx = "Regimen_SitioRemoto_block" Then 'Minuta sitio remoto
            
                Sql = " SELECT DISTINCT"
                Sql = Sql & " reg_codigo, reg_nombre  "
                Sql = Sql & " FROM a_regimen reg With(NoLock)  "
                Sql = Sql & " WHERE  reg_activo = '1'  "
                Sql = Sql & " and    reg_indppr = '1'  "
                Sql = Sql & " and upper(reg_codigo) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
                Sql = Sql & " ORDER BY reg_nombre"
                Set RS1 = vg_db.Execute(Sql)
            
        'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
        ElseIf Opx = "Servicio_SitioRemoto_block" Then 'Minuta sitio remoto
                
                Sql = " SELECT DISTINCT  "
                Sql = Sql & " ser_codigo, ser_nombre   "
                Sql = Sql & " FROM a_servicio With(NoLock)  "
                Sql = Sql & " where  ser_activo = '1' and ser_indppr = '1' "
                Sql = Sql & " and ser_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
                Sql = Sql & " ORDER BY ser_nombre "
                Set RS1 = vg_db.Execute(Sql)
        
        ElseIf Opx = "Servicio_SitioRemoto" Then 'Minuta sitio remoto
            
            'esta seleccion es selecc. SEGMENTO (TABLA B_...)
            If SeleccTipoMinuta_MVI = "Segmento" Then
                
                Set RS1 = vg_db.Execute("sgpadm_s_casservicio 4, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
            Else
                
                Sql = " SELECT DISTINCT  "
                Sql = Sql & " ser_codigo, ser_nombre   "
                Sql = Sql & " FROM cas_a_servicio Reg With(NoLock)  "
                Sql = Sql & " where  ser_activo = '1'  "
                Sql = Sql & " and ser_codigo LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
                Sql = Sql & " and ser_cecori = '" & IIf(swUp_CECO = True, m_copia_min_seg.fpText, m_copia_min_seg.fpText1) & "'"
                Sql = Sql & " ORDER BY ser_nombre "
                Set RS1 = vg_db.Execute(Sql)
            
            End If
        
        ElseIf Opx = "AgregarIng" Then
            
            SqlText = ""
            If VarSitioRemoto = False Then
               
               Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 11, '" & vg_Indppr & "', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
            Else
               
               Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 15, 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
            End If
        
        ElseIf Opx = "AgregarRec" Then
            
            SqlText = ""
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_RecetasActivas '1', '" & vg_Indppr & "', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "IngReal" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 11, '1', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "IngRealCasino" Then
            
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_sel_ingredientexcecoxcodigo '" & LimpiaDato(M_ForComPrexCeCo.fpText) & "','%" & Sql & "%'")
        
        ElseIf Opx = "CliAct" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 44, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        End If
    
    ElseIf Combo1.ItemData(Combo1.ListIndex) = 1 Then
        
        If Opx = "AgregarIngxReceta" Then
        
            vaSpread1.maxcols = 3
            vaSpread1.Row = 0
            vaSpread1.Col = 3
            vaSpread1.text = "Tipo"
            Me.Width = 6930
            Frame1.Left = 600
            vaSpread1.Width = 6150
            Set RS1 = vg_db.Execute("sgpadm_Sel_ListaIngRecetaNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Gen" Then 'Or Opx = "ProInv" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "nombre")
        
       ElseIf Opx = "GenUFacIng" Then  ' Unidad Factor Ingrediente
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_UnidadConversionIngredienteNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "homestser" Then
       
           Set RS1 = vg_db.Execute("SELECT DISTINCT id_homologacionestservicio, Descripcion FROM a_homologacionestservicio WHERE UPPER(Descripcion) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY Descripcion")
        
        
       ElseIf Opx = "GrpIngPri" Then  ' Grupo Ingrediente Principal
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_BuscarNombreGrupoIngPrincipalActivo '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "TipIngPri" Then  ' Tipo Ingrediente Principal x nombre
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBuscaTipoIngPrincipalRecetaxNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
       
       ElseIf Opx = "MetCocc" Then  ' Metodo Cocción x nombre
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaMetodoCoccionRecetaxNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
             
       ElseIf Opx = "IngCruGar" Then  ' Ingrediente Cruce Garnitura x nombre
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaIngCruceGarnituraRecetaxNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
               
       ElseIf Opx = "TiempoHH" Then  ' Tiempo HH x nombre
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaTiempoHhRecetaxNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
   
       ElseIf Opx = "Color" Then  ' Color x nombre
        
             Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaColoxNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "TiempoCoccion" Then  ' Tiempo Cocción x nombre
        
             Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBusquedaTiempoCoccionRecetaxNombre '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
       ElseIf Opx = "GrpEstru" Then  ' Grupo Estructura
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaBuscarNombreGrupoEstructura '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "EstSer" Then
        
            Set RS1 = vg_db.Execute("sgpadm_Sel_ListaEstServicioCodigo " & Suf & ", '%" & UCase(LimpiaDato(Text1.text)) & "%', '2'")
        
        ElseIf Opx = "Celo" Then
            
            Sql = ""
            Sql = "sgpadm_Sel_BuscarOrgCompras "
            Sql = Sql & " '" & UCase(LimpiaDato(Text1.text)) & "'"
            Set RS1 = vg_db.Execute(Sql)
        
        ElseIf Opx = "CecoGral" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 43, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Zon" Then 'Or Opx = "ProInv" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' and zon_activo = '1' ORDER BY " & Suf & "nombre")
        
        ElseIf Opx = "clientesimap" Then 'formato compras sap
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_Cliente 2, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "ProveedorSimap" Then 'Busca los proveedores
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_Proveedor 2, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "ForComSap" Then 'formato compras sap
            
            Set RS1 = vg_db.Execute("sgpadm_s_productossap 2, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "CasReg" Then 'tabla regimen casino
            
            Set RS1 = vg_db.Execute("sgpadm_s_casregimen 3, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "CasSer" Then 'tabla servicio casino
            
            Set RS1 = vg_db.Execute("sgpadm_s_casservicio 3, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "CecoPortalElec" Then ' clientes con portal electronico
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 36, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cli5Eta" Then 'Or Opx = "ProInv") Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 4, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "recorrido" Then  'tabla recorrido
            
            Set RS1 = vg_dbpedweb.Execute("SELECT recorrido, descripcion FROM s_Recorrido WHERE descripcion LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY descripcion")
        
        ElseIf Opx = "regneg" Then ' tabla reglas de negocios
            
            Set RS1 = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 3, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "lispreweb" Then ' tabla reglas de negocios
            
            Set RS1 = vg_dbpedweb.Execute("pedweb_s_listaprecios 5, '', '%" & UCase(LimpiaDato(Text1.text)) & "%', ''")
        
        ElseIf (Opx = "Ser" Or Opx = "Sub") Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre, " & Suf & "indppr FROM " & Tabla & " WHERE UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' = '3') ORDER BY " & Suf & "nombre")
        
        ElseIf Opx = "Reg" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre, " & Suf & "indppr FROM " & Tabla & " WHERE " & Suf & "activo = '1' AND UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' AND (" & Suf & "indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' = '3') ORDER BY " & Suf & "nombre")
        
        ElseIf Opx = "RegBlo" Then
            
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_Sel_RegimenBloquexNombre '%" & Sql & "%'")
        
        ElseIf Opx = "SerBlo" Then
            
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioBloquexNombre '%" & Sql & "%'")
        
        ElseIf Opx = "0" Or Opx = "1" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE (" & Suf & "tiprec='" & Opx & "' OR '" & Opx & "'='') AND UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "codigo")
        
        ElseIf Opx = "Casino" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=0 AND UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Traspaso" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=2 AND UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Cliente" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=1 and UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Proing" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_productosing WHERE " & Suf & "codigo=b_productosing.pri_coding AND b_productosing.pri_codpro='" & vg_codigo & "' AND UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "Ingrec" Then
            
            Set RS1 = vg_db.Execute("SELECT DISTINCT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & ", b_receta, b_recetadet WHERE b_receta.rec_codigo=b_recetadet.red_codigo AND " & Suf & "codigo=b_recetadet.red_codpro AND (b_receta.rec_catdie=" & vg_filcatdie & " OR " & vg_filcatdie & "=0) AND (b_receta.rec_tippla=" & vg_filtippla & " OR " & vg_filtippla & "=0) AND UPPER(" & Suf & "nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ORDER BY " & Suf & "nombre")
        
        ElseIf Opx = "LisPre" Then
            
            Activo = IIf(Opx = "LisPre", "('0','1')", "('1')")
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_Sel_ListaPrecioxNombre_V02 '%" & Sql & "%'")
        
        ElseIf Opx = "ProdActi" Then
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_productos 25, '" & vg_auxcod + vg_Indppr & "', '%" & UCase(LimpiaDato(Text1.text)) & "%', '" & vg_NUsr & "'")
        
        ElseIf Opx = "ForCom" Then
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_productos 19, 0, '%" & UCase(LimpiaDato(Text1.text)) & "%', '" & vg_NUsr & "'")
        
        ElseIf Opx = "SacCco" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_saccentrocosto 2, '" & vg_codigo & "', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cliente_SitioRemotoI" Then 'Minuta sitio remoto
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 23, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Regimen_SitioRemotoI" Then 'Minuta sitio remoto
            
            Set RS1 = vg_db.Execute("sgpadm_s_casregimen 4, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Servicio_SitioRemotoI" Then 'Minuta sitio remoto
            
            Set RS1 = vg_db.Execute("sgpadm_s_casservicio 5, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cliente_SitioRemoto" Then 'Minuta sitio remoto
           
           If swEsCopia = True And SeleccTipoMinuta_MVI = "Segmento" And swUp_CECO = True Then
              
              Sql = " SELECT DISTINCT a.sub_codigo, a.sub_nombre"
              Sql = Sql & " FROM a_subsegmento a "
              Sql = Sql & " LEFT outer JOIN b_detlistaprecio b on a.sub_codigo = b.dlp_codigo "
              Sql = Sql & " WHERE sub_activo = 1"
              Sql = Sql & " AND a.sub_nombre LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
              Sql = Sql & " ORDER BY a.sub_nombre"
              Set RS1 = vg_db.Execute(Sql)
           
           Else
              
              Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 23, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
           
           End If
        
        ElseIf Opx = "Cliente_EnvioBloque" Then 'Minuta envio minuta bloque
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 39, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Cliente_CopiaMinutaBloque" Then 'Minuta envio minuta bloque
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 41, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "Regimen_SitioRemoto" Then 'Minuta sitio remoto
            
             'PANTERA
             'esta seleccion es selecc. SEGMENTO (TABLA B_...)
             If swEsCopia = True And SeleccTipoMinuta_MVI = "Segmento" And swUp_REG = False Then
                    
                Sql = " SELECT DISTINCT"
                Sql = Sql & " reg_codigo, reg_nombre "
                Sql = Sql & " FROM cas_a_regimen Reg With(NoLock)"
                Sql = Sql & " where  reg_activo = '1'"
                Sql = Sql & " and reg_cecori = '" & IIf(swUp_REG = True, Suf, m_copia_min_seg.fpText1) & "'"
                Sql = Sql & " and reg_nombre LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
                Sql = Sql & " ORDER BY reg_nombre"
                Set RS1 = vg_db.Execute(Sql)
             
             ElseIf SeleccTipoMinuta_MVI = "Segmento" Then
                 
                Set RS1 = vg_db.Execute("sgpadm_s_casregimen 4, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
             
             ElseIf swEsCopia = True And SeleccTipoMinuta_MVI = "Segmento" And swUp_REG = False Then
                 
                'por ende Bloque
                Sql = " SELECT DISTINCT "
                Sql = Sql & " reg_codigo, reg_nombre "
                Sql = Sql & " FROM cas_a_regimen  Reg With(NoLock)"
                Sql = Sql & " WHERE UPPER(reg_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
                Sql = Sql & " and reg_cecori = '" & IIf(swUp_CECO = True, m_copia_min_seg.fpText, m_copia_min_seg.fpText1) & "'"
                Sql = Sql & " and reg_activo = '1' ORDER BY reg_nombre"
                Set RS1 = vg_db.Execute(Sql)
             
             End If
             
             'MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
        ElseIf Opx = "Regimen_SitioRemoto_block" Then 'Minuta sitio remoto
                
                Sql = " SELECT DISTINCT "
                Sql = Sql & " reg_codigo, reg_nombre "
                Sql = Sql & " FROM a_regimen With(NoLock)"
                Sql = Sql & " WHERE UPPER(reg_nombre) LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
                Sql = Sql & " and reg_activo = '1' and reg_indppr = '1' ORDER BY reg_nombre"
                Set RS1 = vg_db.Execute(Sql)
       
       ElseIf Opx = "Servicio_SitioRemoto" Then 'Minuta sitio remoto
            
            'esta seleccion es selecc. SEGMENTO (TABLA B_...)
            If SeleccTipoMinuta_MVI = "Segmento" Then
                
               Set RS1 = vg_db.Execute("sgpadm_s_casservicio 5, '" & Suf & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
            Else
                
               Sql = " SELECT DISTINCT  "
               Sql = Sql & " ser_codigo, ser_nombre   "
               Sql = Sql & " FROM cas_a_servicio Reg With(NoLock)  "
               Sql = Sql & " where  ser_activo = '1'  "
               Sql = Sql & " and ser_nombre LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
               Sql = Sql & " and ser_cecori = '" & IIf(swUp_CECO = True, m_copia_min_seg.fpText, m_copia_min_seg.fpText1) & "'"
               Sql = Sql & " ORDER BY ser_nombre "
               Set RS1 = vg_db.Execute(Sql)
            
            End If
       
       ElseIf Opx = "Servicio_SitioRemoto_block" Then 'Minuta sitio remoto
                 
           Sql = " SELECT DISTINCT  "
           Sql = Sql & " ser_codigo, ser_nombre   "
           Sql = Sql & " FROM a_servicio With(NoLock)  "
           Sql = Sql & " where  ser_activo = '1' and ser_indppr = '1' "
           Sql = Sql & " and ser_nombre LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%'"
           Sql = Sql & " ORDER BY ser_nombre "
                 Set RS1 = vg_db.Execute(Sql)
        'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
        
        ElseIf Opx = "CentCost" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "tipo=0 AND " & Suf & "activo=1 AND " & Suf & "nombre LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "SsllListProv" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "activo=0 AND " & Suf & "nombre LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "SsllListFormComp" Then
            
            Set RS1 = vg_db.Execute("SELECT " & Suf & "codsac, " & Suf & "nomsac FROM " & Tabla & " WHERE (" & Suf & "flexec = 0 OR (" & Suf & "flexec = -1 AND " & Suf & "vigfin > " & Format(Date, "yyyy/mm/dd") & ")) AND " & Suf & "nomsac LIKE '%" & UCase(LimpiaDato(Text1.text)) & "%' ")
        
        ElseIf Opx = "AgregarIng" Then
            SqlText = ""
            If VarSitioRemoto = False Then
               
               Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 10, '" & vg_Indppr & "', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
               
            Else
                  
               Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 13, 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
               
            End If

        ElseIf Opx = "AgregarRec" Then
            
            SqlText = ""
            
            Set RS1 = vg_db.Execute("sgpadm_Sel_RecetasActivas '2', '" & vg_Indppr & "', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
            
        ElseIf Opx = "IngReal" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 10, '1', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        ElseIf Opx = "IngRealCasino" Then
            
            Sql = UCase(LimpiaDato(Text1.text))
            Set RS1 = vg_db.Execute("sgpadm_sel_ingredientexcecoxnombre '" & LimpiaDato(M_ForComPrexCeCo.fpText) & "','%" & Sql & "%'")
        
        ElseIf Opx = "CliAct" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 43, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        End If
    
    End If
    
    i = 1
    vaSpread1.MaxRows = 0
    
    If Not RS1.EOF Then
        
        Do While Not RS1.EOF
            
            Est = True
            
            If Opx = "ProInv" Then
                Est = True
            End If
            
            If Est Then
                
                vaSpread1.MaxRows = i: vaSpread1.Row = i: i = i + 1
                vaSpread1.Col = 1: vaSpread1.text = IIf(Opx = "Cliente", fg_PintaRut(RS1(0)), fg_pone_espacio(RS1(0), 13))
                vaSpread1.Col = 2: vaSpread1.text = RS1(1)
                
                If Opx = "ForCom" Or Opx = "ProdActi" Or _
                                                Opx = "AgregarIng" Or Opx = "Ser" Or Opx = "Sub" Or Opx = "Reg" Then
                   vaSpread1.Col = 3
                   vaSpread1.text = ""
                   vaSpread1.text = IIf(Opx = "ProdActi" Or Opx = "ForCom", RS1(2), IIf(RS1(2) = "1", "Real", "Propuesta"))
                
                ElseIf Mid(Opx, 1, 6) = "LisPre" Then
                   
                   vaSpread1.Col = 3
                   vaSpread1.text = IIf(IsNull(RS1(2)) Or RS1(2) = "0", " ", fg_pone_espacio(Mid(RS1(2), 5, 2) & "/" & Mid(RS1(2), 1, 4), 9))
                
                End If
                
                If Opx = "ProdActi" Then
                    
                    vaSpread1.Col = 4
                    vaSpread1.text = ""
                    vaSpread1.text = IIf(Opx = "ProdActi", IIf(IsNull(RS1(4)), "", IIf(RS1(4) = "1", "Real", "Propuesta")), "")
                    vaSpread1.Col = 5: vaSpread1.text = "": vaSpread1.text = Mid(RS1(5), 7, 2) & "/" & Mid(RS1(5), 5, 2) & "/" & Mid(RS1(5), 1, 4)
                    If RS1(6) = "0" Then vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                    vaSpread1.Col = 6: vaSpread1.text = "": vaSpread1.text = IIf(IsNull(RS1(7)), "", Trim(RS1(7)))
                
                ElseIf Opx = "ForCom" And vg_auxcod = "sap" Then
                    
                    vaSpread1.Col = 4: vaSpread1.text = "": vaSpread1.text = IIf(IsNull(RS1(3)), "", Trim(RS1(3)))
                
                End If
            
            End If
            
            RS1.MoveNext
        
        Loop
    
    End If
    RS1.Close: Set RS1 = Nothing
    vaSpread1.SetActiveCell 1, 1
    vg_Indppr = AuxIndppr

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Man_Error

    If KeyCode = 27 Then Cerrar: Exit Sub
    If KeyCode = 40 Or KeyCode = 34 And IRow > 0 Then vaSpread1.SetFocus

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
On Error GoTo Man_Error

Select Case Button.Index
        
    Case 1
            
        MoverDatos
    
    Case 3
            
        vg_codigo = ""
        Cerrar
        
End Select

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub vaSpread1_AfterUserSort(ByVal Col As Long)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
Dim i As Long
Dim valor As String
Select Case Col

Case 3
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = Col
        valor = Trim(vaSpread1.text)
        
        If valor <> "" Then
           
           valor = Mid(valor, 5, 2) & "/" & Mid(valor, 1, 4)
        
        End If
        
        vaSpread1.text = valor
    
    Next

End Select

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub vaSpread1_BeforeUserSort(ByVal Col As Long, ByVal State As Long, DefaultAction As Long)

On Error GoTo Man_Error

Dim valor As String
Dim i As Long
If vaSpread1.MaxRows < 1 Then Exit Sub

Select Case Col

Case 3
    
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        vaSpread1.Col = Col
        valor = Trim(vaSpread1.text)
        
        If valor <> "" Then
           
           valor = Replace(valor, "/", "")
           valor = Mid(valor, 3, 4) & Mid(valor, 1, 2)
           vaSpread1.text = Trim(valor)
        
        Else
           
           vaSpread1.text = " "
        
        End If
    
    Next

End Select

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub vaSpread1_DblClick(ByVal Col As Long, ByVal Row As Long)

On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Or Col = 0 Or Row = 0 Then Exit Sub
Call MoverDatos

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub vaSpread1_KeyPress(KeyAscii As Integer)
    
On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"
MoverDatos

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub vaSpread1_KeyUp(KeyCode As Integer, Shift As Integer)
    
On Error GoTo Man_Error

If KeyCode = 27 Then Cerrar: Exit Sub

If TeclasNoPermitidas(KeyCode) = True Then
        
   Text1.text = IIf(KeyCode = 8, Text1.text, Text1.text & Chr(KeyCode))
   Text1.SetFocus
   Text1.SelStart = Len(Text1.text)
    
End If

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Private Sub MoverDatos()
    
On Error GoTo Man_Error

If vaSpread1.MaxRows < 1 Then Exit Sub
vaSpread1.Row = vaSpread1.ActiveRow
vaSpread1.Col = 1
vg_codigo = Trim(vaSpread1.text)
vaSpread1.Col = 2
vg_nombre = Trim(vaSpread1.text)
vaSpread1.Col = 3
vg_ames = vaSpread1.text

If Opx = "ForCom" And vg_auxcod = "sap" Then
       
   vaSpread1.Col = 4: vg_auxcod = vaSpread1.text

End If
    
Call Cerrar

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Sub Cerrar()
    
On Error GoTo Man_Error

Me.Hide
Unload Me

Exit Sub
Man_Error:
    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub

Sub LlenaDatos(TablaGen As String, SufGen As String, titgen As String, op As String, Optional catdie As Long, Optional tippla As Long)

On Error GoTo Man_Error

Dim RS1 As New ADODB.Recordset
Dim Sql As String
Dim Est As Boolean
Dim z   As Long

    Opx = op
    Tabla = TablaGen
    Suf = SufGen
    Titulo = titgen
    Me.Caption = titgen
    Frame1.Left = 0
    CategoriDie = catdie
    TipoPlato = tippla
    
    If Opx = "AgregarIng" Or Opx = "AgregarIngxReceta" Or Opx = "AgregarRec" Then
       
       Option1(0).Visible = True
       Option1(1).Visible = True
       Option1(2).Visible = True
    
    ElseIf Opx = "Sub" Or Opx = "Reg" Or Opx = "Ser" Or Opx = "ProdActi" Then
       
       Option1(0).Visible = True
       Option1(0).Enabled = IIf(vg_Indppr = "1" Or vg_Indppr = "2", False, True)
       Option1(1).Visible = True
       Option1(1).Value = IIf(vg_Indppr = "1", True, False)
       Option1(1).Enabled = IIf(vg_Indppr = "2", False, True)
       Option1(2).Visible = True
       Option1(2).Value = IIf(vg_Indppr = "2", True, False)
       Option1(2).Enabled = IIf(vg_Indppr = "1", False, True)
    
    End If
    
    vaSpread1.MaxRows = 0
    vaSpread1.maxcols = 2
    vaSpread1.Width = 5160
    
    If RS1.State = 1 Then RS1.Close
    RS1.CursorLocation = adUseClient
    vg_db.CursorLocation = adUseClient

    If Opx = "Gen" Then  ' tabla generica
        
        Set RS1 = vg_db.Execute("SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & "  ORDER BY " & SufGen & "nombre")
        
    ElseIf Opx = "GenUFacIng" Then  ' Unidad Factor Ingrediente
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_UnidadConversionIngrediente")
        
    ElseIf Opx = "homestser" Then
    
        Set RS1 = vg_db.Execute("SELECT DISTINCT id_homologacionestservicio, Descripcion FROM a_homologacionestservicio ORDER BY Descripcion")
       
    ElseIf Opx = "GrpIngPri" Then  ' Grupo Ingrediente Principal
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_GrupoIngPrincipalActivo")

    ElseIf Opx = "TipIngPri" Then  ' Tipo Ingrediente Principal
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_TipoIngPrincipalRecetaActivo")

    ElseIf Opx = "MetCocc" Then  ' Metodo Cocción
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_MetodoCoccionRecetaActivo")
   
    ElseIf Opx = "IngCruGar" Then  ' Ingrediente Cruce Garnitura
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_IngCruceGarnituraRecetaActivo")
   
    ElseIf Opx = "TiempoHH" Then  ' Tiempo HH
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_TiempoHhRecetaActivo")
   
    ElseIf Opx = "Color" Then  ' Color
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_ColorActivo")
   
    ElseIf Opx = "TiempoCoccion" Then  ' Tiempo Cocción
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_TiempoCoccionRecetaActivo")
      
    ElseIf Opx = "GrpEstru" Then  ' Grupo Estructura
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_AyudaGrupoEstructura")
    
    ElseIf Opx = "EstSer" Then
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_ListaEstServicioCodigo " & Suf & ", '', '0'")
        
    ElseIf Opx = "Celo" Then
        
        Sql = ""
        Sql = "sgpadm_Sel_OrgCompras_V02 "
        Sql = Sql & " '' "
        Set RS1 = vg_db.Execute(Sql)
    
    ElseIf Opx = "CecoGral" Then  ' tabla generica
        
        Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 42, '',''")
    
    ElseIf Opx = "Zon" Then  ' tabla generica
        
        Set RS1 = vg_db.Execute("SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & "  where zon_activo = '1' ORDER BY " & SufGen & "nombre")
    
    ElseIf Opx = "clientesimap" Then 'formato compras sap
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_Cliente 1, ''")
    
    ElseIf Opx = "ProveedorSimap" Then 'Selecciona Proveedores
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_Proveedor 1, ''")
    
    ElseIf Opx = "ForComSap" Then 'formato compras sap
        
        Set RS1 = vg_db.Execute("sgpadm_s_productossap 1, '', ''")
    
    ElseIf Opx = "CasReg" Then 'tabla regimen casino
        
        Set RS1 = vg_db.Execute("sgpadm_s_casregimen 1, '" & SufGen & "', 0, ''")
    
    ElseIf Opx = "CasSer" Then 'tabla regimen casino
        
        Set RS1 = vg_db.Execute("sgpadm_s_casservicio 1, '" & SufGen & "', 0, ''")
    
    ElseIf Opx = "Cli5Eta" Then
        
        Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 20, '', ''")
    
    ElseIf Opx = "CecoPortalElec" Then 'Centro costo con portal electronico
        
        Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 35, '', ''")
    
    ElseIf Opx = "recorrido" Then  ' tabla generica
        
        Set RS1 = vg_dbpedweb.Execute("SELECT recorrido, descripcion FROM s_Recorrido  ORDER BY descripcion")
    
    ElseIf Opx = "regneg" Then 'tabla reglas de negocios
        
        Set RS1 = vg_dbpedweb.Execute("pedweb_s_reglasdenegocios 1, '', ''")
    
    ElseIf Opx = "lispreweb" Then 'tabla lista de precios
        
        Set RS1 = vg_dbpedweb.Execute("pedweb_s_listaprecios 2, '', '', ''")
    
    ElseIf Opx = "Cliente_SitioRemotoI" Then 'Minuta sitio remoto
        
        Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 22, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
    
    ElseIf Opx = "Regimen_SitioRemotoI" Then 'Minuta sitio remoto
        
        Set RS1 = vg_db.Execute("sgpadm_s_casregimen 4, '" & SufGen & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
    
    ElseIf Opx = "Servicio_SitioRemotoI" Then 'Minuta sitio remoto
        
        Set RS1 = vg_db.Execute("sgpadm_s_casservicio 4, '" & SufGen & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
    
    ElseIf Opx = "Cliente_SitioRemoto" Then 'Minuta sitio remoto
            
        If swEsCopia = True And SeleccTipoMinuta_MVI = "Bloque" And swUp_CECO = True Then
        'aca carga la query de los sub segmentos
        
            Sql = " "
            Sql = " SELECT DISTINCT cli_codigo, cli_nombre "
            Sql = Sql & " From b_clientes"
            Sql = Sql & " Where cli_activo = 1"
            Sql = Sql & " AND cli_minsre = 1"
            Sql = Sql & " ORDER BY cli_nombre"
       
            Set RS1 = vg_db.Execute(Sql)
            
        ElseIf swEsCopia = True And SeleccTipoMinuta_MVI = "Segmento" And swUp_CECO = True Then
            
            Sql = " SELECT DISTINCT a.sub_codigo, a.sub_nombre"
            Sql = Sql & " FROM a_subsegmento a "
            Sql = Sql & " LEFT outer JOIN b_detlistaprecio b on a.sub_codigo = b.dlp_codigo "
            Sql = Sql & " WHERE sub_activo = 1"
            Sql = Sql & " ORDER BY a.sub_nombre"
            
            Set RS1 = vg_db.Execute(Sql)
        
        Else
            
            Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 22, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        
        End If
    
    ElseIf Opx = "Cliente_EnvioBloque" Then 'Minuta envio minuta bloque
        
        Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 38, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
    
    ElseIf Opx = "Cliente_CopiaMinutaBloque" Then 'Minuta envio minuta bloque
        
        Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 40, '', '%" & UCase(LimpiaDato(Text1.text)) & "%'")
    
    ElseIf Opx = "Regimen_SitioRemoto" Then 'Minuta sitio remoto
    
    'MVA - MVI - COPIA DE MINUTA
        
            'esta seleccion es selecc. SEGMENTO (TABLA B_...)
        If SeleccTipoMinuta_MVI = "Segmento" And swUp_REG = True Then
               
               Set RS1 = vg_db.Execute("sgpadm_s_casregimen 4, '" & SufGen & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
   
        ElseIf swEsCopia = True And swUp_REG = False Then
   
               Sql = " SELECT DISTINCT"
               Sql = Sql & " reg_codigo, reg_nombre "
               Sql = Sql & " FROM cas_a_regimen Reg With(NoLock)"
               Sql = Sql & " where  reg_activo = '1'"
               Sql = Sql & " and reg_cecori = '" & IIf(swUp_REG = True, SufGen, m_copia_min_seg.fpText1) & "'"
               Sql = Sql & " ORDER BY reg_nombre"
        
               Set RS1 = vg_db.Execute(Sql)
  
        Else
               Sql = " SELECT DISTINCT"
               Sql = Sql & " reg_codigo, reg_nombre "
               Sql = Sql & " FROM cas_a_regimen Reg With(NoLock)"
               Sql = Sql & " where  reg_activo = '1'"
               Sql = Sql & " and reg_cecori = '" & IIf(swUp_REG = True, SufGen, m_copia_min_seg.fpText1) & "'"
               Sql = Sql & " ORDER BY reg_nombre"
        
               Set RS1 = vg_db.Execute(Sql)
        End If
   
   'FIN MVA - MVI - COPIA DE MINUTA
   
   
   'MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
    ElseIf Opx = "Regimen_SitioRemoto_block" Then 'Minuta sitio remoto block
                
                Sql = " SELECT DISTINCT"
                Sql = Sql & " reg_codigo, reg_nombre "
                Sql = Sql & " FROM a_regimen Reg With(NoLock)"
                Sql = Sql & " where  reg_activo = '1' and reg_indppr = '1'"
                Sql = Sql & " ORDER BY reg_nombre"
            
                Set RS1 = vg_db.Execute(Sql)
    
    ElseIf Opx = "Servicio_SitioRemoto" Then 'Minuta sitio remoto
        
        'esta seleccion es selecc. SEGMENTO (TABLA B_...)
        If SeleccTipoMinuta_MVI = "Segmento" Then
            
            Set RS1 = vg_db.Execute("sgpadm_s_casservicio 4, '" & SufGen & "', 0, '%" & UCase(LimpiaDato(Text1.text)) & "%'")
        Else
           
           Sql = " SELECT DISTINCT  "
           Sql = Sql & " ser_codigo, ser_nombre   "
           Sql = Sql & " FROM cas_a_servicio Reg With(NoLock)  "
           Sql = Sql & " where  ser_activo = '1'  "
           Sql = Sql & " and ser_cecori = '" & IIf(swUp_CECO = True, m_copia_min_seg.fpText, m_copia_min_seg.fpText1) & "'"
           Sql = Sql & " ORDER BY ser_nombre  "
           
           Set RS1 = vg_db.Execute(Sql)
           
        End If
    ElseIf Opx = "Servicio_SitioRemoto_block" Then 'Servicio sitio remoto block
                
                Sql = " SELECT DISTINCT  "
                Sql = Sql & " ser_codigo, ser_nombre   "
                Sql = Sql & " FROM a_servicio With(NoLock)  "
                Sql = Sql & " where  ser_activo = '1' and ser_indppr = '1' "
                Sql = Sql & " ORDER BY ser_nombre  "
                Set RS1 = vg_db.Execute(Sql)
   'FIN MVA - CREACION MINUTA DESDE CENTRAL - 2013-01-10
    
    ElseIf (Opx = "Ser" Or Opx = "Sub") Then
        
        vaSpread1.maxcols = 3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "Tipo"
        Me.Width = 6930
        Frame1.Left = 600
        vaSpread1.Width = 6150
    
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre, " & SufGen & "indppr FROM " & TablaGen & " WHERE (" & SufGen & "indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' = '3') ORDER BY " & SufGen & "nombre")
    
    ElseIf Opx = "Reg" Then
        
        vaSpread1.maxcols = 3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "Tipo"
        Me.Width = 6930
        Frame1.Left = 600
        vaSpread1.Width = 6150
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre, " & SufGen & "indppr FROM " & TablaGen & " WHERE " & SufGen & "activo = '1' AND (" & SufGen & "indppr = '" & vg_Indppr & "' OR '" & vg_Indppr & "' = '3') ORDER BY " & SufGen & "nombre")
    
    ElseIf Opx = "RegBlo" Then
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_RegimenBloque 0")
    
    ElseIf Opx = "SerBlo" Then
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_ServicioBloque 0")
    
    ElseIf Opx = "0" Or Opx = "1" Then 'Tabla Recetas estado 0
        
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE (" & SufGen & "tiprec='" & op & "' OR '" & op & "'='') ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "Casino" Then
        
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo=0 ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "Traspaso" Then
        
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo=2 ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "Cliente" Then
        
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo=1 ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "CentCost" Then
        
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & " WHERE " & SufGen & "tipo=0 AND " & SufGen & "activo=1 ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "SsllListProv" Then
        
        Set RS1 = vg_db.Execute("SELECT " & Suf & "codigo, " & Suf & "nombre FROM " & Tabla & " WHERE " & Suf & "activo=0 ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "SsllListFormComp" Then
        
        Set RS1 = vg_db.Execute("SELECT " & Suf & "codsac, " & Suf & "nomsac FROM " & Tabla & " WHERE (" & Suf & "flexec = 0 OR (" & Suf & "flexec = -1 AND " & Suf & "vigfin > " & Format(Date, "yyyy/mm/dd") & ")) ORDER BY " & Suf & "codsac")
    
    ElseIf Opx = "Proing" Then
        
        Set RS1 = vg_db.Execute("SELECT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & ", b_productosing WHERE " & SufGen & "codigo=b_productosing.pri_coding AND b_productosing.pri_codpro='" & vg_codigo & "' ORDER BY " & SufGen & "codigo")
    
    ElseIf Opx = "Ingrec" Then
        
        Set RS1 = vg_db.Execute("SELECT DISTINCT " & SufGen & "codigo, " & SufGen & "nombre FROM " & TablaGen & ", b_receta, b_recetadet WHERE b_receta.rec_codigo=b_recetadet.red_codigo AND " & SufGen & "codigo=b_recetadet.red_codpro AND (b_receta.rec_catdie=" & vg_filcatdie & " OR " & vg_filcatdie & "=0) AND (b_receta.rec_tippla=" & vg_filtippla & " OR " & vg_filtippla & "=0) ORDER BY " & SufGen & "nombre")
    
    ElseIf Opx = "SacCco" Then
        
        Set RS1 = vg_db.Execute("sgpadm_s_saccentrocosto 3, '" & vg_codigo & "', ''")
    
    ElseIf Opx = "AgregarIng" Then
        
        vaSpread1.maxcols = 3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "Tipo"
        Me.Width = 6930
        Frame1.Left = 600
        vaSpread1.Width = 6150
        Set RS1 = vg_db.Execute("sgpadm_Sel_ListaIngMinutaxNombre " & Val(M_TabGra.fpLongInteger1(0)) & ", " & Val(M_TabGra.fpLongInteger1(1)) & ", " & Val(M_TabGra.fpLongInteger1(2)) & ", " & CategoriDie & ", " & TipoPlato & ", " & Val(vg_fecha) & ", '%%'")
        
        If VarSitioRemoto = False Then
              
              Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 12, 0, '" & vg_Indppr & "'")
           
        Else
              
              Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 14, 0, '" & vg_Indppr & "'")
        End If
        
    ElseIf Opx = "AgregarIngxReceta" Then
        
        vaSpread1.maxcols = 3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "Tipo"
        Me.Width = 6930
        Frame1.Left = 600
        vaSpread1.Width = 6150
        Set RS1 = vg_db.Execute("sgpadm_Sel_ListaIngRecetaCodigo '%%'")
        
    ElseIf Opx = "AgregarRec" Then
        
        vaSpread1.maxcols = 3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "Tipo"
        Me.Width = 6930
        Frame1.Left = 600
        vaSpread1.Width = 6150
       
        Set RS1 = vg_db.Execute("sgpadm_Sel_RecetasActivas '0' ,'" & vg_Indppr & "', ''")
    
    ElseIf Opx = "IngReal" Then
        
        Set RS1 = vg_db.Execute("sgpadm_s_ingrediente_V02 12, 0, '1'")
    
    ElseIf Opx = "ProdActi" Then
        
        vaSpread1.maxcols = 6 '5 '4 '3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "U.M."
        vaSpread1.Col = 4
        vaSpread1.text = "Tipo Productos"
        vaSpread1.Col = 5
        vaSpread1.text = "Fecha Vigencia"
        vaSpread1.Col = 6
        vaSpread1.text = "Cta. Contable"
        
        Me.Width = 9830 '8930
        Frame1.Left = 1400 '1300
        vaSpread1.Width = 9080 '8150
        Set RS1 = vg_db.Execute("sgpadm_Sel_productos 23, '" & vg_auxcod + vg_Indppr & "', '', '" & vg_NUsr & "'")
    
    ElseIf Opx = "ForCom" Then
        
        vaSpread1.maxcols = IIf(vg_auxcod = "sap", 4, 3)
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "U.M."
        
        If vg_auxcod = "sap" Then
           
           vaSpread1.Col = 4
           vaSpread1.text = "Cta. Con."
           Me.Width = 7930
           Frame1.Left = 1100
           vaSpread1.Width = 7120
        
        Else
           
           Me.Width = 6930
           Frame1.Left = 600
           vaSpread1.Width = 6150
        
        End If
        
        Set RS1 = vg_db.Execute("sgpadm_Sel_productos 18, '" & vg_auxcod & "', '', '" & vg_NUsr & "'")
    
    ElseIf Mid(Opx, 1, 6) = "LisPre" Then
        
        Dim Activo As String
        Activo = IIf(Opx = "LisPre", "('0','1')", "('1')")
        vaSpread1.maxcols = 3
        vaSpread1.Row = 0
        vaSpread1.Col = 3
        vaSpread1.text = "Fecha"
        Me.Width = 6930
        Frame1.Left = 600
        vaSpread1.Width = 6150
        vaSpread1.ColUserSortIndicator(1) = ColUserSortIndicatorAscending
        vaSpread1.ColUserSortIndicator(2) = ColUserSortIndicatorAscending
        vaSpread1.ColUserSortIndicator(3) = ColUserSortIndicatorAscending
        vaSpread1.UserColAction = UserColActionSort
        Set RS1 = vg_db.Execute("sgpadm_Sel_ListaPrecio")
    
    ElseIf Opx = "IngRealCasino" Then ' trae ingrediente reales
        
        Set RS1 = vg_db.Execute("sgpadm_sel_ingredientexceco " & "'" & M_ForComPrexCeCo.fpText & "'")
    
    ElseIf Opx = "CliAct" Then  ' tabla generica
         
         Set RS1 = vg_db.Execute("sgpadm_s_cliente_V02 42, '',''")
    
    End If
    
    If Not RS1.EOF Then
        
        Do While Not RS1.EOF
            
            Est = True
            
            If Opx = "ProInv" Then
                
                Est = True
            
            End If
            
            If Est Then
                
                vaSpread1.MaxRows = vaSpread1.MaxRows + 1
                vaSpread1.Row = vaSpread1.MaxRows
                vaSpread1.Col = 1
                vaSpread1.text = IIf(Opx = "Cliente", fg_PintaRut(RS1(0)), fg_pone_espacio(RS1(0), 10))
                vaSpread1.Col = 2
                vaSpread1.text = Trim(RS1(1))
                If Opx = "ForCom" Or Opx = "ProdActi" Or Opx = "AgregarIng" Or Opx = "Ser" Or Opx = "Sub" Or Opx = "Reg" Then
                   
                   vaSpread1.Col = 3
                   vaSpread1.text = ""
                   vaSpread1.text = IIf(Opx = "LisPre", IIf(IsNull(RS1(2)) Or RS1(2) = "0", " ", fg_pone_espacio(Mid(RS1(2), 5, 2) & "/" & Mid(RS1(2), 1, 4), 9)), IIf(Opx = "ProdActi" Or Opx = "ForCom", RS1(2), IIf(RS1(2) = "1", "Real", "Propuesta")))
                
                ElseIf Mid(Opx, 1, 6) = "LisPre" Then
                   
                   vaSpread1.Col = 3
                   vaSpread1.text = IIf(IsNull(RS1(2)) Or RS1(2) = "0", " ", fg_pone_espacio(Mid(RS1(2), 5, 2) & "/" & Mid(RS1(2), 1, 4), 9))
                
                End If
                
                If Opx = "ProdActi" Then
                   
                   vaSpread1.Col = 4: vaSpread1.text = "": vaSpread1.text = IIf(Opx = "ProdActi", IIf(IsNull(RS1(3)), "", IIf(RS1(3) = "1", "Real", "Propuesta")), "")
                   vaSpread1.Col = 5: vaSpread1.text = "": vaSpread1.text = Mid(RS1(4), 7, 2) & "/" & Mid(RS1(4), 5, 2) & "/" & Mid(RS1(4), 1, 4)
                   If RS1(5) = "0" Then vaSpread1.Col = -1: vaSpread1.BackColor = Shape1(1).FillColor
                   vaSpread1.Col = 6: vaSpread1.text = "": vaSpread1.text = IIf(IsNull(RS1(6)), "", Trim(RS1(6)))
                
                ElseIf Opx = "ForCom" And vg_auxcod = "sap" Then
                   
                   vaSpread1.Col = 4
                   vaSpread1.text = ""
                   vaSpread1.text = IIf(IsNull(RS1(3)), "", Trim(RS1(3)))
                
                End If
            
            End If
            
            RS1.MoveNext
        
        Loop
        
        If RS1(0).Type = adVarWChar Then
            
           vaSpread1.Col = 1
           vaSpread1.Row = -1
           vaSpread1.TypeHAlign = TypeHAlignLeft
        
        Else
            
           vaSpread1.Col = 1
           vaSpread1.Row = -1
           vaSpread1.TypeHAlign = TypeHAlignLeft
        
        End If
    
    End If
    
    RS1.Close: Set RS1 = Nothing

Exit Sub
Man_Error:

    Call MsgBox(Err & ":  " & Error$(Err), vbCritical, Titulo)

End Sub
