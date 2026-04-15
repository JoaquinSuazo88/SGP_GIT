VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_SubBDt 
   Caption         =   "Subir Datos"
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3525
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2970
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   8250
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de productos"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   6
         Top             =   315
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de recetas"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   5
         Top             =   645
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   1530
         Width           =   7215
      End
      Begin VB.TextBox Text1 
         Height          =   585
         Index           =   1
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1515
         Visible         =   0   'False
         Width           =   7665
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de planificación"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   3300
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   240
         Left            =   270
         TabIndex        =   3
         Top             =   2505
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de Origen"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   8
         Top             =   1335
         Width           =   1275
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   7515
         Picture         =   "P_SubBDt.frx":0000
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Archivo en Proceso"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   2190
         Visible         =   0   'False
         Width           =   7650
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   495
      Top             =   2790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   0
      OleObjectBlob   =   "P_SubBDt.frx":030A
      Top             =   0
   End
End
Attribute VB_Name = "P_SubBDt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim Msgtitulo As String, tipopc As String

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 3900
Me.Width = 8460
tipopc = ""
Msgtitulo = "Actualización de Base de Datos"
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "Actualizar", , tbrDefault, "ActuBD"): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "Salir", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
Label1(0).Visible = True: Text1(0).Visible = True: Image1(1).Visible = True
Text1(1).Visible = False
End Sub

Private Sub Image1_Click(Index As Integer)
Cd.Filter = "Todos los archivos (*.*)|*.*"
Cd.DefaultExt = "*.*"
Cd.InitDir = dir_trabajo & "Actualizar"
Cd.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
Cd.ShowOpen
If Cd.Filename = "" Then Text1(0).text = "" Else Text1(0).text = Dir(Cd.Filename)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim oError As Boolean
tipopc = "SGP"
Select Case Button.Key
Case "Actualizar"
    Toolbar1.Enabled = False: Frame1.Enabled = False
    If Trim(Text1(0).text) = "" Then MsgBox "Debe selecionar archivo origen", vbInformation + vbOKOnly, Msgtitulo: Toolbar1.Enabled = True: Frame1.Enabled = True: Exit Sub
    oError = IIf(ActArchivo(Trim(Text1(0).text)), False, True)
    If Not oError And Dir(Cd.Filename) <> "" Then
       Name Trim(Text1(0).text) As Mid(Trim(Text1(0).text), 1, Len(Trim(Text1(0).text)) - 3) & "dwl"
       Text1(0).text = ""
    End If
    If oError Then
        MsgBox "Proceso de Actualización Falló", vbInformation + vbOKOnly, Msgtitulo
    Else
        MsgBox "Proceso de Actualización Finalizado", vbInformation + vbOKOnly, Msgtitulo
    End If
    Label1(1).Visible = False: PB.Visible = False
    Toolbar1.Enabled = True: Frame1.Enabled = True
Case "Salir"
    Me.Hide
    Unload Me
End Select
End Sub

Private Function ActArchivo(ByVal cdbz As String) As Long
Dim fso As New FileSystemObject, cdbi As String, indice As Long, cDBO As String
On Error GoTo Man_Error
ActArchivo = False
cdbi = dir_trabajo & "Actualizar\" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb"
cDBO = dir_trabajo & BaseDeDato
Set dbI = New ADODB.Connection
dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbI.ConnectionTimeout = 3600
dbI.CommandTimeout = 3600
dbI.Open

    PB.Min = 0: PB.Value = 0: PB.max = 23
    Label1(1).Visible = True: PB.Visible = True
    
    Label1(1).Caption = "Importando a_bodega": DoEvents
    RS1.Open "SELECT * FROM a_bodega", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_bodega WHERE bod_codigo = " & RS1!bod_codigo & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    Label1(1).Caption = "Importando b_proveedor": DoEvents
    RS1.Open "SELECT * FROM b_proveedor", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_proveedor WHERE prv_codigo= '" & RS1!prv_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> bodegas
    Label1(1).Caption = "Importando bodegas": DoEvents
    RS1.Open "SELECT * FROM b_bodegas", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_bodegas WHERE bod_codbod=" & RS1!bod_codbod & " AND bod_codpro = '" & RS1!bod_codpro & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> b_totcompras
    Label1(1).Caption = "Importando b_totcompras": DoEvents
    RS1.Open "SELECT * FROM b_totcompras", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_totcompras WHERE toc_rutpro = '" & RS1!toc_rutpro & "' AND toc_tipdoc = '" & RS1!toc_tipdoc & "' AND toc_numdoc = " & RS1!toc_numdoc & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> b_detcompras
    Label1(1).Caption = "Importando b_detcompras": DoEvents
    RS1.Open "SELECT * FROM b_detcompras", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detcompras WHERE dec_rutpro = '" & RS1!dec_rutpro & "' AND dec_tipdoc = '" & RS1!dec_tipdoc & "' AND dec_numdoc = " & RS1!dec_numdoc & " AND dec_numlin = " & RS1!dec_numlin & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipos de Producto
    Label1(1).Caption = "Importando b_detcomprasimp": DoEvents
   
    RS1.Open "SELECT * FROM b_detcomprasimp", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detcomprasimp WHERE imd_rutdoc = '" & RS1!imd_rutdoc & "' AND imd_tipdoc = '" & RS1!imd_tipdoc & "' AND imd_numdoc = " & RS1!imd_numdoc & " AND imd_numlin = " & RS1!imd_numlin & " AND imd_codimp = " & RS1!imd_codimp & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipos documento
    Label1(1).Caption = "Importando b_detpreciocaf": DoEvents
    RS1.Open "SELECT * FROM b_detpreciocaf", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detpreciocaf WHERE dpc_codigo = '" & RS1!dpc_codigo & "' AND dpc_codmer = '" & RS1!dpc_codmer & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Parametro Despacho
    Label1(1).Caption = "Importando b_totventas": DoEvents
    RS1.Open "SELECT * FROM b_totventas", dbI, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_totventas WHERE tov_rutcli = '" & RS1!tov_rutcli & "' AND tov_tipdoc = '" & RS1!tov_tipdoc & "' AND tov_numdoc = " & RS1!tov_numdoc & ""
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Parametro Despacho
    Label1(1).Caption = "Importando b_detventas": DoEvents
    RS1.Open "SELECT * FROM b_detventas", dbI, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detventas WHERE dev_rutcli = '" & RS1!dev_rutcli & "' AND dev_tipdoc = '" & RS1!dev_tipdoc & "' AND dev_numdoc = " & RS1!dev_numdoc & " AND dev_numlin = " & RS1!dev_numlin & ""
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
     
    '-------> Unidades de medida
    Label1(1).Caption = "Importando b_totpreciocaf": DoEvents
    RS1.Open "SELECT * FROM b_totpreciocaf", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_totpreciocaf WHERE tpc_codigo = '" & RS1!tpc_codigo & "' AND tpc_cencos = '" & RS1!tpc_cencos & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    
    '-------> Unidades de medida
    Label1(1).Caption = "Importando b_totventascaf": DoEvents
    RS1.Open "SELECT * FROM b_totventascaf", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_totventascaf WHERE tvc_cencos = '" & RS1!tvc_cencos & "' AND tvc_fecing = '" & RS1!tvc_fecing & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de medida
    Label1(1).Caption = "Importando b_detventascaf": DoEvents
    RS1.Open "SELECT * FROM b_detventascaf", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detventascaf WHERE dvc_cencos = '" & RS1!dvc_cencos & "' AND dvc_fecing = " & RS1!dvc_fecing & " AND dvc_numlin = " & RS1!dvc_numlin & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de stock
    Label1(1).Caption = "Importando b_detventascafpro"
    DoEvents
    RS1.Open "SELECT * FROM b_detventascafpro", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detventascafpro WHERE dvp_cencos = '" & RS1!dvp_cencos & "' AND dvp_fecing = " & RS1!dvp_fecing & " AND dvp_codmer = '" & RS1!dvp_codmer & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de embalaje
    Label1(1).Caption = "Importando b_detventasimp"
    DoEvents
    RS1.Open "SELECT * FROM b_detventasimp", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detventasimp WHERE imd_rutdoc = '" & RS1!emb_codigo & "' AND imd_tipdoc = '" & RS1!imd_tipdoc & "' AND imd_numdoc = " & RS1!imd_numdoc & " AND imd_numlin = " & RS1!imd_numlin & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Cuentas Contables
    Label1(1).Caption = "Importando b_minuta"
    DoEvents
    RS1.Open "SELECT * FROM b_minuta", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minuta WHERE min_codigo = " & RS1!min_codigo & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    Label1(1).Caption = "Importando b_minutadet"
    DoEvents
    RS1.Open "SELECT * FROM b_minutadet", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutadet WHERE mid_codigo = " & RS1!mid_codigo & " AND mid_tipmin = '" & RS1!mid_tipmin & "' and mid_numlin = " & RS1!mid_numlin & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Impuestos
    
    Label1(1).Caption = "Importando b_minutafijadia"
    DoEvents
    RS1.Open "SELECT * FROM b_minutafijadia", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutafijadia WHERE mfd_cencos = '" & RS1!mfd_cencos & "' and mfd_codreg = " & RS1!mfd_codreg & " and  mfd_codser = " & RS1!mfd_codser & " AND mfd_fecha = " & RS1!mfd_fecha & " AND mfd_codpro = '" & RS1!mfd_codpro & "' AND mfd_tipmin = '" & RS1!mfd_tipmin & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Nutrientes
    Label1(1).Caption = "Importando b_minutaraciones"
    DoEvents
    RS1.Open "SELECT * FROM b_minutaraciones", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutaraciones WHERE mir_cencos = '" & RS1!mir_cencos & "' AND mir_codreg = " & RS1!mir_codreg & " AND mir_codser = " & RS1!mir_codser & " AND mir_fecmin = " & RS1!mir_fecmin & " AND mir_rutcli = '" & RS1!MIR_RUTCLI & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Articulos de Stock
    Label1(1).Caption = "Importando b_preciovta": DoEvents
    RS1.Open "SELECT * FROM b_preciovta", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_preciovta WHERE prv_cencos = '" & RS1!prv_cencos & "' AND prv_codreg = " & RS1!prv_codreg & " AND prv_codser = " & RS1!prv_codser & " AND prv_fecvig = " & RS1!prv_fecvig & " AND prv_rutcli = '" & RS1!prv_rutcli & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Impuestos Articulos de Stock
    Label1(1).Caption = "Importando b_productospmpdia": DoEvents
    RS1.Open "SELECT * FROM b_productospmpdia", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productospmpdia WHERE ppd_cencos = '" & RS1!ppd_cencos & "' AND ppd_codpro= '" & RS1!ppd_codpro & "' AND ppd_fecdia = " & RS1!ppd_fecdia & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Ingredientes
    Label1(1).Caption = "Importando b_tomainv": DoEvents
    RS1.Open "SELECT * FROM b_tomainv", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_tomainv WHERE tin_fectom = " & RS1!tin_fectom & " AND tin_codbod = " & RS1!tin_codbod & " AND tin_codpro = '" & RS1!tin_codpro & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Aportes Nutricionales Ingrediente
    Label1(1).Caption = "Importando b_ventacontado": DoEvents
    RS1.Open "SELECT * FROM b_ventacontado", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_ventacontado WHERE vtc_codigo = " & RS1!vtc_codigo & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Proveedores
    Label1(1).Caption = "Importando b_ventacontadodet": DoEvents
    RS1.Open "SELECT * FROM b_ventacontadodet", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_ventacontadodet WHERE vtd_codigo = " & RS1!vtd_codigo & " AND vtd_numlin = " & RS1!vtd_numlin & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    dbI.Close: Set dbI = Nothing
    ActArchivo = True
Exit Function
Man_Error:
If Err = -2147217865 Or Err = 3265 Then
    dbI.Close: Set dbI = Nothing
   Exit Function
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Function
vg_db.RollbackTrans
If Err.Number = -2147467259 Then
    MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Function

Private Function ActRegistro(RSO As Fields, DBO As ADODB.Connection, DBD As ADODB.Connection, cSql As String)
Dim RS2 As New ADODB.Recordset, i As Long, bAdd As Boolean
On Error GoTo ManError
DoEvents
bAdd = False
RS2.Open cSql, DBD, adOpenDynamic, adLockOptimistic
vg_db.BeginTrans
If RS2.EOF Then
    bAdd = True
    RS2.AddNew
End If
For i = 0 To RSO.count - 1
    If bAdd Or (RS2.Fields(i).Name <> "pro_upreco" And RS2.Fields(i).Name <> "pro_fecuco" And RS2.Fields(i).Name <> "pro_propon" And RS2.Fields(i).Name <> "ing_precos" And RS2.Fields(i).Name <> "ing_feccos" And RS2.Fields(i).Name <> "rec_tiprec" And RS2.Fields(i).Name <> "ing_codcom" And RS2.Fields(i).Name <> "ing_codped" And RS2.Fields(i).Name <> "ess_codsec") Then
        Select Case RS2.Fields(i).Type
        Case adChar, adVarChar, adVarWChar
            If TipoDato(RS2.Fields(i).Value, "") <> Trim(TipoDato(RSO.Item(i).Value, "")) Then RS2.Fields(i).Value = IIf(Trim((RSO.Item(i).Value)) = "", " ", Trim(RSO.Item(i).Value))
        Case Else
            If TipoDato(RS2.Fields(i).Value, 0) <> TipoDato(RSO.Item(i).Value, 0) Or RS2.Fields(i).Name = "red_tiprec" Or RS2.Fields(i).Name = "mid_tiprec" Or RS2.Fields(i).Name = "mid_racteo" Or RS2.Fields(i).Name = "mid_racrea" Then
               RS2.Fields(i).Value = RSO.Item(i).Value
            ElseIf RS2.Fields(i).Name = "mid_cosrec" Then
               RS2.Fields(i).Value = fg_CalCtoRecInv(RSO.Item(i - 3).Value, RSO.Item(i + 2).Value, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")))
            ElseIf RS2.Fields(i).Name = "mid_cosdes" Then
               RS2.Fields(i).Value = fg_CalCtoRecInv(RSO.Item(i - 8).Value, RSO.Item(i - 3).Value, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','"))) 'fg_CalCtoRecInv(RSO.Item(i - 3).Value, RSO.Item(i + 2).Value, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")))
            End If
        End Select
    ElseIf RS2.Fields(i).Name = "rec_tiprec" And RSO.Item(i).Value > 0 Then
        RS2.Fields(i).Value = RSO.Item(i).Value
    ElseIf RS2.Fields(i).Name = "ess_codsec" And RSO.Item(i).Value > 0 Then
        RS2.Fields(i).Value = RSO.Item(i).Value
    End If
Next
RS2.Update
RS2.Close: Set RS2 = Nothing
vg_db.CommitTrans
Exit Function
ManError:
If Err.Number = -2147217887 Then Resume Next
RS2.Close: Set RS2 = Nothing
MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, Msgtitulo
End Function

