VERSION 5.00
Object = "{5B7759CE-C04E-4C5D-993B-8297E30D9065}#1.0#0"; "ChilkatFTP.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{830C6AA3-5274-11D4-BD8D-912BC639A87B}#1.0#0"; "activezip.ocx"
Begin VB.Form M_ActuBD 
   Caption         =   "Actualizar Base de Datos"
   ClientHeight    =   3450
   ClientLeft      =   2715
   ClientTop       =   2790
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   8340
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CD 
      Left            =   495
      Top             =   2790
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2970
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   8250
      Begin ACTIVEZIPLib.ActiveZip AZ1 
         Left            =   7920
         Top             =   960
         _Version        =   65536
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   0
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de planificación"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.TextBox Text1 
         Height          =   585
         Index           =   1
         Left            =   270
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1515
         Visible         =   0   'False
         Width           =   7665
      End
      Begin MSComctlLib.ProgressBar PB 
         Height          =   240
         Left            =   270
         TabIndex        =   6
         Top             =   2505
         Visible         =   0   'False
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   1530
         Width           =   7215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de recetas"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   2
         Top             =   645
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Actualizar maestro de productos"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   315
         Visible         =   0   'False
         Width           =   3300
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
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   7515
         Picture         =   "M_ActuBD.frx":0000
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo de Origen"
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   5
         Top             =   1335
         Width           =   1275
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin CHILKATFTPLibCtl.ChilkatFTP oFTP 
      Left            =   4140
      OleObjectBlob   =   "M_ActuBD.frx":030A
      Top             =   3915
   End
End
Attribute VB_Name = "M_ActuBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim RS3 As New ADODB.Recordset
Dim MsgTitulo As String, tipopc As String
Dim estarc As Boolean

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Height = 3900
Me.Width = 8460
tipopc = ""
MsgTitulo = "Actualización de Base de Datos"
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1
Set BtnX = Toolbar1.Buttons.Add(, "Actualizar", , tbrDefault, "ActuBD"): BtnX.Visible = True: BtnX.ToolTipText = "Vista Previa": BtnX.Enabled = IIf(Mid(ValidarUsuario(Me), 4, 1) = "1", True, False)
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "Salir", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"
If GetParametro("metactbd") = 2 Then
    Label1(0).Visible = True: Text1(0).Visible = True: Image1(1).Visible = True
    Text1(1).Visible = False
Else
    Label1(0).Visible = False: Text1(0).Visible = False: Image1(1).Visible = False
    Text1(1).Visible = True
End If
End Sub

Private Sub Image1_Click(Index As Integer)
'If Option1(0).Value Then
''    CD.Filter = "Todos los archivos (MP*.MDB)|MP*.MDB"
''    CD.DefaultExt = "MP*.MDB"
'    CD.Filter = "Todos los archivos (MP*.ZIP)|MP*.ZIP"
'    CD.DefaultExt = "MP*.ZIP"
'ElseIf Option1(1).Value Then
''    CD.Filter = "Todos los archivos (MR*.MDB)|MR*.MDB"
''    CD.DefaultExt = "MR*.MDB"
'    CD.Filter = "Todos los archivos (MR*.ZIP)|MR*.ZIP"
'    CD.DefaultExt = "MR*.ZIP"
'ElseIf Option1(2).Value Then
    CD.Filter = "Todos los archivos (SGP*.ZIP)|SGP*.ZIP"
    CD.DefaultExt = "SGP*.ZIP"
'Else
'    CD.Filter = "Todos los archivos (*.*)|*.*"
'    CD.DefaultExt = "*.*"
'End If
CD.InitDir = dir_trabajo & "Actualizar"
CD.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
CD.ShowOpen
'Text1(0).Text = CD.FileName
If CD.FileName = "" Then Text1(0).text = "" Else Text1(0).text = Dir(CD.FileName)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Dim oError As Boolean
Dim sql1 As String, cHost As String, Cdire As String, cUser As String, cPass As String, Cpuer As Long
'If Option1(0).Value = True Then
'   tipopc = "MP"
'ElseIf Option1(1).Value = True Then
'   tipopc = "MR"
'ElseIf Option1(2).Value = True Then
   tipopc = "SGP"
'End If
Select Case Button.Key
Case "Actualizar"
    
    Toolbar1.Enabled = False: Frame1.Enabled = False
    If GetParametro("metactbd") = 1 Then
       
       '-------> Traer datos FTP
       sql1 = IIf(vg_tipbase = "1", " ucase(par_codigo) ", " upper(par_codigo) ")
       Set RS1 = vg_db.Execute("SELECT par_codigo, par_valor FROM a_param WHERE " & sql1 & " LIKE '%" & LimpiaDato(UCase("ftp")) & "%' AND par_cencos = '" & MuestraCasino(1) & "'")
       If RS1.EOF Then fg_descarga: RS1.Close: Set RS1 = Nothing: MsgBox "No existe Parametrización FTP, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
       Do While Not RS1.EOF
          
          If RS1!par_codigo = "ftpser" Then cHost = fg_Desencripta(TipoDato(RS1!par_valor, ""))
          If RS1!par_codigo = "ftpdir" Then Cdire = fg_Desencripta(TipoDato(RS1!par_valor, ""))
          If RS1!par_codigo = "ftpusu" Then cUser = fg_Desencripta(TipoDato(RS1!par_valor, ""))
          If RS1!par_codigo = "ftppas" Then cPass = fg_Desencripta(TipoDato(RS1!par_valor, ""))
          If RS1!par_codigo = "ftppue" Then Cpuer = fg_Desencripta(TipoDato(RS1!par_valor, ""))
          RS1.MoveNext
       
       Loop
       
       RS1.Close: Set RS1 = Nothing
       If Trim(cHost) = "" Then MsgBox "No existe Parametrización FTP del servidor, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
       If Trim(Cdire) = "" Then MsgBox "No existe Parametrización FTP del nombre directorio, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
       If Trim(cUser) = "" Then MsgBox "No existe Parametrización FTP del usuario, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
       If Trim(cPass) = "" Then MsgBox "No existe Parametrización FTP del password, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
       If Trim(Cpuer) = "" Then MsgBox "No existe Parametrización FTP del puerto, Proceso cancelado, contactese con departamento informactica ", vbCritical, MsgTitulo: Exit Sub
        oError = False
        oFTP.UseIEProxy = False
        oFTP.Port = Cpuer '21
        oFTP.HostName = cHost '"sgp.sodexhochile.cl" '"64.76.146.65" '"64.76.138.76" '"64.76.45.71"  'fg_Desencripta(TipoDato(cHost, ""))
        oFTP.UserName = cUser '"userftp" '"sodexho"   'fg_Desencripta(TipoDato(cUser, ""))
        oFTP.password = cPass '"*sdxo7528*" '"*sdxo123*" '"shx873" 'fg_Desencripta(TipoDato(cPass, ""))
        oFTP.Connect
        If oFTP.IsConnected Then
            Text1(1).text = oFTP.LastErrorText
            lDir = oFTP.GetCurrentDirListing("*.*")
            oFTP.SaveLastError ("aaa.xml")
            Text1(1).text = oFTP.LastErrorText: DoEvents
'            a = oFTP.ChangeRemoteDir("/casinos/bd")
            a = oFTP.ChangeRemoteDir(Cdire)
            oFTP.SaveLastError ("aaa.xml")
            Text1(1).text = oFTP.LastErrorText: DoEvents
'            lDir = oFTP.GetCurrentDirListing("MP" & Trim(GetParametro("casino")) & "*.zip")
            lDir = oFTP.GetCurrentDirListing(tipopc & Trim(GetParametro("casino")) & "*.zip")
            oFTP.SaveLastError ("aaa.xml")
            Text1(1).text = oFTP.LastErrorText: DoEvents
            For i = 0 To oFTP.NumFilesAndDirs - 1
                a = oFTP.GetFile(oFTP.GetFileName(i), App.Path & "\Actualizar\" & oFTP.GetFileName(i))
                oFTP.SaveLastError ("aaa.xml")
                Text1(1).text = oFTP.LastErrorText: DoEvents
                If vg_tipbase = "1" Then
                   If ActArchivoAccess(oFTP.GetFileName(i)) Then
                      a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                      oFTP.SaveLastError ("aaa.xml")
                      Text1(1).text = oFTP.LastErrorText: DoEvents
                      a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                      oFTP.SaveLastError ("aaa.xml")
                      Text1(1).text = oFTP.LastErrorText: DoEvents
                      'Borrar dato
                       a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                       oFTP.SaveLastError ("aaa.xml")
                       Text1(1).text = oFTP.LastErrorText: DoEvents
                      'Borrar dato
                       a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                       oFTP.SaveLastError ("aaa.xml")
                       Text1(1).text = oFTP.LastErrorText: DoEvents
                   Else
                      If estarc Then
                         a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                         a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                         'Borrar dato
                         a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                         'Borrar dato
                         a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                      Else
                         oError = True
                      End If
                   End If
                ElseIf vg_tipbase <> "1" Then
                    If ActArchivoSql(oFTP.GetFileName(i)) Then
                       a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                       oFTP.SaveLastError ("aaa.xml")
                       Text1(1).text = oFTP.LastErrorText: DoEvents
                       a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                       oFTP.SaveLastError ("aaa.xml")
                       Text1(1).text = oFTP.LastErrorText: DoEvents
                       'Borrar dato
                       a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                       oFTP.SaveLastError ("aaa.xml")
                       Text1(1).text = oFTP.LastErrorText: DoEvents
                       'Borrar dato
                       a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                       oFTP.SaveLastError ("aaa.xml")
                       Text1(1).text = oFTP.LastErrorText: DoEvents
                   Else
                      If estarc Then
                         a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                         a = oFTP.RenameRemoteFile(oFTP.GetFileName(i), Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                         'Borrar dato
                         a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                         'Borrar dato
                         a = oFTP.DeleteRemoteFile(Mid(oFTP.GetFileName(i), 1, Len(oFTP.GetFileName(i)) - 3) & "dwl")
                         oFTP.SaveLastError ("aaa.xml")
                         Text1(1).text = oFTP.LastErrorText: DoEvents
                      Else
                         oError = True
                      End If
                   End If
                End If
            Next i
            oFTP.Disconnect
            Text1(1).text = oFTP.LastErrorText: DoEvents
        End If
    Else
        If Trim(Text1(0).text) = "" Then MsgBox "Debe selecionar archivo origen", vbInformation + vbOKOnly, MsgTitulo: Toolbar1.Enabled = True: Frame1.Enabled = True: Exit Sub
        If vg_tipbase = "1" Then
           oError = IIf(ActArchivoAccess(Trim(Text1(0).text)), False, True)
        Else
           oError = IIf(ActArchivoSql(Trim(Text1(0).text)), False, True)
        End If
        If Not oError And Dir(CD.FileName) <> "" Then
           Name Trim(Text1(0).text) As Mid(Trim(Text1(0).text), 1, Len(Trim(Text1(0).text)) - 3) & "dwl"
           Text1(0).text = ""
        End If
    End If
    If oError Then
        MsgBox "Proceso de Actualización Falló", vbInformation + vbOKOnly, MsgTitulo
    Else
        MsgBox "Proceso de Actualización Finalizado", vbInformation + vbOKOnly, MsgTitulo
    End If
    Label1(1).Visible = False: PB.Visible = False
    Toolbar1.Enabled = True: Frame1.Enabled = True

Case "Salir"
    
    Me.Hide
    Unload Me

End Select

Exit Sub
Man_Error:
If Err = -2147467259 Then Resume Next

End Sub

Private Function ActArchivoAccess(ByVal cdbz As String) As Long
Dim fso As New FileSystemObject, cdbi As String, indice As Long, cDBO As String
On Error GoTo Man_Error
ActArchivoAccess = False
estarc = False
If vg_contra <> Mid(cdbz, 4, InStr(cdbz, "-") - 4) Then MsgBox "Base de dato no corresponde centro de costo origen...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
If Not fso.FileExists(dir_trabajo & "Actualizar\" & cdbz) Then MsgBox "No se encuentra el archivo para importar datos...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
cdbi = dir_trabajo & "Actualizar\" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb"
cDBO = dir_trabajo & BaseDeDato
RS1.Open "SELECT * FROM log_actualizacion WHERE archivo = '" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
   If GetParametro("metactbd") <> 1 Then
      MsgBox "Ya fue aplicada la actualización para el archivo" & VgLinea & cdbi & "...", vbExclamation + vbOKOnly, MsgTitulo
   Else
      estarc = True
   End If
   RS1.Close: Set RS1 = Nothing
   Exit Function
End If
RS1.Close: Set RS1 = Nothing
AZ1.OpenZip dir_trabajo & "Actualizar\" & cdbz
AZ1.ExtractFile AZ1.FileName(0), dir_trabajo & "Actualizar\" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb", ""
AZ1.Close
Set dbI = New ADODB.Connection
dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbI.ConnectionTimeout = 3600
dbI.CommandTimeout = 3600
dbI.Open
If UCase(Mid(cdbi, Len(dir_trabajo & "Actualizar\") + 1, 3)) = "SGP" Then
    RS1.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_fecmin IN (SELECT min_fecmin FROM b_minuta IN '" & cdbi & "') AND min_indblo = 1 AND min_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
    If Not RS1.EOF Then
       RS1.Close: Set RS1 = Nothing
       dbI.Close: Set dbI = Nothing
       fso.DeleteFile cdbi: ActArchivoAccess = False
'       Name Trim(Text1(0).Text) As Mid(Trim(Text1(0).Text), 1, Len(Trim(Text1(0).Text)) - 3) & "dwl"
       Name cdbz As Mid(cdbz, 1, Len(cdbz) - 3) & "dwl"
       Text1(0).text = ""
       MsgBox "Planificación minuta esta bloqueada, proceso cancelado...", vbInformation + vbOKOnly, MsgTitulo: ActArchivoAccess = True: Exit Function
    End If
    RS1.Close: Set RS1 = Nothing

    PB.Min = 0: PB.Value = 0: PB.max = 50
    Label1(1).Visible = True: PB.Visible = True
    
    '-------> envío sap
    Label1(1).Caption = "Importando Envío Sap": DoEvents
    RS1.Open "SELECT * FROM a_tipointerfaz", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipointerfaz WHERE tii_codigo = " & RS1!tii_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
'    '-------> par_tipo_vales
'    If validarsiexistetabla(cdbi, "a_par_tipo_vales") Then
'       Label1(1).Caption = "Importando parametro tipo vales": DoEvents
'       RS1.Open "SELECT * FROM a_par_tipo_vales", dbI, adOpenStatic
'       Do While Not RS1.EOF
'          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_par_tipo_vales WHERE ID_Tipo_Vale = '" & RS1!ID_Tipo_Vale & "' and cli_codigo = '" & RS1!cli_codigo & "'"
'          RS1.MoveNext
'       Loop
'       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
'    End If
    
    '-------> par_codigo_barra_cas
    If validarsiexistetabla(cdbi, "a_par_codigo_barra") Then
       Label1(1).Caption = "Importando parametro código barra": DoEvents
       RS1.Open "SELECT * FROM a_par_codigo_barra", dbI, adOpenStatic
       Do While Not RS1.EOF
          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_par_codigo_barra_cas WHERE a_par_id_codigo = " & a_par_id_codigo & " and atr_codigo_barra = '" & RS1!atr_codigo_barra & "' and cli_codigo = '" & RS1!cli_codigo & "'"
          RS1.MoveNext
       Loop
       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    End If
    
    '-------> contrato envío sap
    Label1(1).Caption = "Importando Contrato Envío Sap": DoEvents
    vg_db.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos = '" & MuestraCasino(1) & "'"
    RS1.Open "SELECT * FROM b_casinointerfaz", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_casinointerfaz WHERE cai_cencos = '" & RS1!cai_cencos & "' AND cai_codtii = " & RS1!cai_codtii & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipos de Producto
    Label1(1).Caption = "Importando Tipos de Producto": DoEvents
    dbI.Execute "ALTER TABLE a_tipopro ADD COLUMN tip_activo char(1)"
    dbI.Execute "UPDATE a_tipopro SET tip_activo = 'S'"
    RS1.Open "SELECT * FROM a_tipopro", dbI, adOpenStatic
    If Not RS1.EOF Then
       vg_db.Execute "UPDATE a_tipopro SET tip_activo = 'N'"
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipopro WHERE tip_codigo = " & RS1!tip_codigo
        RS1.MoveNext
    Loop
    End If
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipos documento
    Label1(1).Caption = "Importando Tipos de Documento": DoEvents
    RS1.Open "SELECT * FROM a_tipodocumento", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipodocumento WHERE tdo_codigo = '" & RS1!tdo_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
'    '-------> Parametro Despacho
'    Label1(1).Caption = "Importando Parametro de Despacho": DoEvents
'    RS1.Open "SELECT b.tip_codigo, b.tip_nombre FROM a_tipopro a INNER JOIN a_tipopro AS b ON a.tip_codigo = b.tip_previo WHERE a.tip_previo = 0", vg_db, adOpenStatic
'    If Not RS1.EOF Then
'       Do While Not RS1.EOF
'          RS2.Open "SELECT DISTINCT pad_codigo FROM b_paramdesp WHERE pad_cencos = '" & MuestraCasino(1) & "' AND pad_codigo = " & RS1!tip_codigo & "", vg_db, adOpenStatic
'          If RS2.EOF Then vg_db.Execute "INSERT INTO b_paramdesp VALUES (" & RS1!tip_codigo & ", 'S', '" & MuestraCasino(1) & "', '')"
'          RS2.Close: Set RS2 = Nothing
'          RS1.MoveNext
'       Loop
'    End If
'    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
     
    '-------> Unidades de medida
    Label1(1).Caption = "Importando Unidades de Medida": DoEvents
    RS1.Open "SELECT * FROM a_unidadmed", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidadmed WHERE unm_codigo = " & RS1!unm_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de stock
    Label1(1).Caption = "Importando Unidades de Stock"
    DoEvents
    RS1.Open "SELECT * FROM a_unidad", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidad WHERE uni_codigo = " & RS1!uni_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de embalaje
    Label1(1).Caption = "Importando unidades de embalaje"
    DoEvents
    RS1.Open "SELECT * FROM a_embalaje", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_embalaje WHERE emb_codigo = " & RS1!emb_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Cuentas Contables
    Label1(1).Caption = "Importando Cuentas Contables"
    DoEvents
    RS1.Open "SELECT * FROM a_ctacontable", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_ctacontable WHERE cta_codigo = '" & RS1!cta_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Parametro
    '-------> Agregar campo
    dbI.Execute "alter table a_param add column par_cencos char(10)"
    dbI.Execute "update a_param set par_cencos = '" & MuestraCasino(1) & "'"
    
    Label1(1).Caption = "Importando Parametros"
    DoEvents
    RS1.Open "SELECT * FROM a_param", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_param WHERE par_cencos = '" & RS1!par_cencos & "' AND par_codigo = '" & RS1!par_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Impuestos
    Label1(1).Caption = "Importando Impuestos"
    DoEvents
    RS1.Open "SELECT * FROM a_impuesto", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_impuesto WHERE imp_codigo = " & RS1!imp_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Nutrientes
    Label1(1).Caption = "Importando Nutrientes"
    DoEvents
    RS1.Open "SELECT * FROM a_nutriente", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_nutriente WHERE nut_codigo = " & RS1!nut_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Articulos de Stock
    Label1(1).Caption = "Importando Artículos de Stock": DoEvents
    RS1.Open "SELECT * FROM b_productos", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productos WHERE pro_codigo = '" & RS1!pro_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Impuestos Articulos de Stock
    Label1(1).Caption = "Importando Impuestos Relacionados": DoEvents
    RS1.Open "SELECT * FROM b_productosimp", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosimp WHERE ipr_codpro = '" & RS1!ipr_codpro & "' AND ipr_codimp = " & RS1!ipr_codimp
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Ingredientes
    Label1(1).Caption = "Importando Ingredientes": DoEvents
    RS1.Open "SELECT * FROM b_ingrediente", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_ingrediente WHERE ing_codigo = '" & RS1!ing_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Ingredientes Articulos de Stock
    Label1(1).Caption = "Importando Ingredientes Relacionados": DoEvents
    vg_db.Execute "DELETE FROM b_productosing WHERE pri_codpro IN (SELECT pri_codpro FROM b_productosing IN '" & cdbi & "')"
    RS1.Open "SELECT * FROM b_productosing", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosing WHERE pri_codpro = '" & RS1!pri_codpro & "' AND pri_coding = '" & RS1!pri_coding & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Aportes Nutricionales Ingrediente
    Label1(1).Caption = "Importando Aportes Nutricionales Ingrediente": DoEvents
    RS1.Open "SELECT * FROM b_productonut", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productonut WHERE pnu_codpro = '" & RS1!pnu_codpro & "' AND pnu_codapo = " & RS1!pnu_codapo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Proveedores
    Label1(1).Caption = "Importando Proveedores": DoEvents
    RS1.Open "SELECT * FROM b_proveedor", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_proveedor WHERE prv_codigo = '" & RS1!prv_codigo & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    '-------> Validar si existe actualización posterioes maestro de productos
    RS1.Open "SELECT MAX(prv_fecumo) AS prv_fecumo FROM b_proveedor", vg_db, adOpenStatic
    If Not RS1.EOF And Not IsNull(RS1!prv_fecumo) Then
       '-------> Actualizar tabla actualiza dato
       vg_db.Execute "UPDATE b_actuadatos SET ada_fecumo = '" & RS1!prv_fecumo & "' WHERE ada_nomtab = 'b_proveedor' AND (ada_fecumo < cdate('" & RS1!prv_fecumo & "') OR (ada_fecumo) IS NULL)"
    End If
    RS1.Close: Set RS1 = Nothing
    '-------> Mover zero al stock si es negativo
    vg_db.Execute "UPDATE b_bodegas set bod_canmer = 0 WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 0"
    
    '-------> Categoría de Receta
    Label1(1).Caption = "Importando Categoría de Receta": DoEvents
    RS1.Open "SELECT * FROM a_recetacatdie", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetacatdie WHERE car_codigo = " & RS1!car_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Tipo de Plato
    Label1(1).Caption = "Importando Tipo de Plato": DoEvents
    RS1.Open "SELECT * FROM a_recetatippla", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetatippla WHERE tip_codigo = " & RS1!tip_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Gramos Familia Producto
    Label1(1).Caption = "Importando Gramos Familia Producto": DoEvents
    RS1.Open "SELECT * FROM b_gramofamproducto", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_gramofamproducto WHERE gfp_cencos = '" & RS1!gfp_cencos & "' AND gfp_codreg = " & RS1!gfp_codreg & " AND gfp_catdie = " & RS1!gfp_catdie & " AND gfp_tiprec = " & RS1!gfp_tiprec & " AND gfp_fampro = " & RS1!gfp_fampro & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Recetas
    Label1(1).Caption = "Importando Recetas": DoEvents
    RS1.Open "SELECT * FROM b_receta", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_receta WHERE rec_codigo = " & RS1!rec_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Ingredientes de Recetas
    '-------> Agregar campo
    dbI.Execute "ALTER TABLE b_recetadet ADD COLUMN red_cencos char(10)"
    dbI.Execute "UPDATE b_recetadet SET red_cencos = '0' WHERE red_tiprec = 0"
    dbI.Execute "UPDATE b_recetadet SET red_cencos = '" & MuestraCasino(1) & "' WHERE red_tiprec <> 0"
    Label1(1).Caption = "Importando Ingredientes Recetas": DoEvents
    '-------> Respaldar recetas 5 epatas y patron
    RS1.Open "SELECT DISTINCT * INTO r_recetadet FROM b_recetadet IN '" & cDBO & "' WHERE red_codigo IN (SELECT DISTINCT red_codigo FROM b_recetadet IN '" & cdbi & "' WHERE (red_tiprec=0 OR red_tiprec>=10000) AND ((red_tiprec<>0 AND red_cencos='" & MuestraCasino(1) & "') OR (red_tiprec=0 AND red_cencos='0'))) AND (red_tiprec=0 OR red_tiprec>=10000) AND ((red_tiprec>0 AND red_cencos='" & MuestraCasino(1) & "') OR (red_tiprec=0 AND red_cencos='0'))", dbI, adOpenStatic
    Set RS1 = Nothing
    RS1.Open "SELECT DISTINCT rec_codigo INTO r_receta FROM b_receta IN '" & cDBO & "' WHERE rec_codigo IN (SELECT DISTINCT rec_codigo FROM b_receta IN '" & cdbi & "' )", dbI, adOpenStatic
    Set RS1 = Nothing
    RS1.Open "SELECT DISTINCT rec_codigo, '0' AS rec_tiprec INTO x_receta FROM b_receta IN '" & cDBO & "' WHERE rec_codigo IN (SELECT DISTINCT rec_codigo FROM b_receta IN '" & cdbi & "' )", dbI, adOpenStatic
    Set RS1 = Nothing
    dbI.Execute "ALTER TABLE r_receta ADD Constraint r_receta_pk Primary Key (rec_codigo)"
    dbI.Execute "ALTER TABLE x_receta ADD Constraint x_receta_pk Primary Key (rec_codigo)"
    dbI.Execute "ALTER TABLE b_tablagramaje ADD Constraint b_tablagramaje_pk Primary Key (tgr_codreg,tgr_codrec,tgr_coding)"
    '-------> Insertar recetas 5 etapas
    RS1.Open "SELECT DISTINCT * FROM a_regimen WHERE reg_codigo >= 10000", dbI, adOpenStatic
    Do While Not RS1.EOF
       DoEvents
       '-------> Mover receta desde tabla gramaje
       dbI.Execute "UPDATE x_receta SET x_receta.rec_tiprec = '0'"
       dbI.Execute "UPDATE x_receta INNER JOIN b_tablagramaje ON x_receta.rec_codigo = b_tablagramaje.tgr_codrec SET x_receta.rec_tiprec = '1' " & _
                   "WHERE b_tablagramaje.tgr_codreg = " & RS1!reg_codigo & ""
       dbI.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos) SELECT DISTINCT b.red_codigo, b.red_nroite, b.red_codpro, b.red_canpro, b.red_cospro, b.red_pctapr, b.red_pctcoc, b.red_pctnut, " & RS1!reg_codigo & " AS red_tiprec, '" & MuestraCasino(1) & "' AS red_cencos FROM x_receta a, b_recetadet b WHERE a.rec_codigo = b.red_codigo AND (a.rec_tiprec = '0' and b.red_tiprec = 0)"
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing
    '-------> Leer tabla cencos si sobreescribe receta
    Dim sobrec As String
    sobrec = ""
    Label1(1).Caption = "Actualizando Detalle Recetas 5 Etapas": DoEvents
    RS1.Open "SELECT * FROM a_cencos", dbI, adOpenStatic
    If Not RS1.EOF Then sobrec = RS1!cen_sobrec
    RS1.Close: Set RS1 = Nothing
    If sobrec = "0" Then '-------> Actualiza solamente las recetas que tienen parametro de actualziación
       '-------> Crear tabla temporal de parametro de recetas
       dbI.Execute "SELECT DISTINCT b_ingrediente.ing_codigo INTO b_detparametro " & _
                   "FROM b_clientes, b_ingrediente, b_recetadet, b_receta, b_gramofamproducto, a_regimen, b_productos IN '" & cDBO & "' " & _
                   "WHERE b_recetadet.red_codpro = b_ingrediente.ing_codigo " & _
                   "AND   b_ingrediente.ing_codcom = b_productos.pro_codigo " & _
                   "AND   b_receta.rec_catdie = b_gramofamproducto.gfp_catdie " & _
                   "AND   b_receta.rec_tippla = b_gramofamproducto.gfp_tiprec " & _
                   "AND   b_receta.rec_codigo = b_recetadet.red_codigo " & _
                   "AND   b_productos.pro_codtip = b_gramofamproducto.gfp_fampro " & _
                   "AND   a_regimen.reg_codigo = b_gramofamproducto.gfp_codreg " & _
                   "AND   b_clientes.cli_codigo = b_gramofamproducto.gfp_cencos " & _
                   "AND   b_gramofamproducto.gfp_grafin > 0 AND b_clientes.cli_codigo = '" & MuestraCasino(1) & "'"
       
       '-------> Actualizar recetas 5 etapas
       dbI.Execute "UPDATE (b_recetadet INNER JOIN r_recetadet ON (b_recetadet.red_codigo = r_recetadet.red_codigo) AND (b_recetadet.red_nroite = r_recetadet.red_nroite) AND (b_recetadet.red_codpro = r_recetadet.red_codpro) AND (b_recetadet.red_tiprec = r_recetadet.red_tiprec) AND (b_recetadet.red_cencos = r_recetadet.red_cencos)) INNER JOIN b_detparametro ON r_recetadet.red_codpro = b_detparametro.ing_codigo SET b_recetadet.red_canpro = r_recetadet.red_canpro " & _
                   "WHERE b_recetadet.red_tiprec >= 10000 AND r_recetadet.red_tiprec >= 10000 AND r_recetadet.red_cencos = '" & MuestraCasino(1) & "'"
    ElseIf sobrec = "2" Then '-------> No actualiza ninguna receta
       dbI.Execute "UPDATE b_recetadet INNER JOIN r_recetadet ON (b_recetadet.red_cencos = r_recetadet.red_cencos) AND (b_recetadet.red_tiprec = r_recetadet.red_tiprec) AND (b_recetadet.red_codpro = r_recetadet.red_codpro) AND (b_recetadet.red_nroite = r_recetadet.red_nroite) AND (b_recetadet.red_codigo = r_recetadet.red_codigo) SET b_recetadet.red_canpro = r_recetadet.red_canpro " & _
                   "WHERE b_recetadet.red_tiprec >= 10000 AND r_recetadet.red_tiprec >= 10000 AND r_recetadet.red_cencos = '" & MuestraCasino(1) & "'"
    End If
    PB.Value = PB.Value + 1
    DoEvents
'    vg_db.BeginTrans
    '-------> Borrar detalle receta patron
    vg_db.Execute "DELETE FROM b_recetadet WHERE red_codigo IN (SELECT DISTINCT rec_codigo FROM r_receta IN '" & cdbi & "') AND red_tiprec = 0 AND red_cencos = '0'"
    DoEvents
    '-------> Borrar detalle receta local
    vg_db.Execute "DELETE FROM b_recetadet WHERE red_codigo IN (SELECT DISTINCT red_codigo FROM b_recetadet IN '" & cdbi & "' where red_tiprec = -1) AND red_tiprec = -1 AND red_cencos = '" & MuestraCasino(1) & "'"
'    vg_db.CommitTrans
    DoEvents
'    vg_db.BeginTrans
    vg_db.Execute "DELETE a.* FROM b_recetadet a WHERE a.red_codigo IN (SELECT DISTINCT rec_codigo FROM r_receta IN '" & cdbi & "') AND a.red_tiprec IN (SELECT DISTINCT reg_codigo FROM a_regimen IN '" & cdbi & "' WHERE reg_codigo >= 10000) AND a.red_cencos = '" & MuestraCasino(1) & "' AND a.red_tiprec >= 10000"
'    vg_db.CommitTrans
    DoEvents
'    vg_db.BeginTrans
    vg_db.Execute "INSERT INTO b_recetadet SELECT DISTINCT red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos FROM b_recetadet IN '" & cdbi & "' ORDER BY red_codigo, red_nroite, red_tiprec, red_cencos"
'    vg_db.CommitTrans
    PB.Value = PB.Value + 1
   
    '-------> Importando Regimen
    Label1(1).Caption = "Importando Regimen": DoEvents
    RS1.Open "SELECT * FROM a_regimen", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_regimen WHERE reg_codigo=" & RS1!reg_codigo
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '------> Importando Servicio
    Label1(1).Caption = "Importando Servicio": DoEvents
    RS1.Open "SELECT * FROM a_servicio", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_servicio WHERE ser_codigo=" & RS1!ser_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '------> Importando Estructura Servicio
    Label1(1).Caption = "Importando Estructura Servicio": DoEvents
    dbI.Execute "ALTER TABLE a_estservicio ADD COLUMN ess_cencos char(10)"
    dbI.Execute "UPDATE a_estservicio SET ess_cencos='" & MuestraCasino(1) & "'"
    RS1.Open "SELECT * FROM a_estservicio", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_estservicio WHERE ess_codser=" & RS1!ess_codser & " AND ess_codigo=" & RS1!ess_codigo & " AND ess_cencos='" & RS1!ess_cencos & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Validar si existe planificación minutas
    vg_db.BeginTrans
    indice = 0
    Label1(1).Caption = "Validar Planificación Minutas": DoEvents
    RS1.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha, min_codreg, reg_nombre, min_codser, ser_nombre FROM b_minuta a, a_regimen b, a_servicio c WHERE a.min_cencos='" & MuestraCasino(1) & "' AND a.min_codreg=b.reg_codigo AND a.min_codser=c.ser_codigo", dbI, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          DoEvents
          RS2.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_cencos='" & MuestraCasino(1) & "' AND VAL(MID(min_fecmin,1,6))=" & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & "", vg_db, adOpenStatic
          If Not RS2.EOF Then
             If MsgBox("Existe planificación minuta, desea borrar la información existente... " & VgLinea & VgLinea & "Regimen  : " & RS1!min_codreg & " " & Trim(RS1!reg_nombre) & VgLinea & "Servicio   :  " & RS1!min_codser & " " & Trim(RS1!ser_nombre), vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
                '-------> Borrar planificación contrato
                vg_db.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo WHERE b_minuta.min_cencos='" & MuestraCasino(1) & "' AND VAL(MID(b_minuta.min_fecmin,1,6))=" & RS1!Fecha & " AND b_minuta.min_codreg=" & RS1!min_codreg & " AND b_minuta.min_codser=" & RS1!min_codser & ""
                vg_db.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos='" & MuestraCasino(1) & "' AND VAL(MID(min_fecmin,1,6))=" & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & ""
             Else
                '-------> Borrar planificación de la base carga
                dbI.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo WHERE b_minuta.min_cencos='" & MuestraCasino(1) & "'AND VAL(MID(b_minuta.min_fecmin,1,6))=" & RS1!Fecha & " AND b_minuta.min_codreg=" & RS1!min_codreg & " AND b_minuta.min_codser=" & RS1!min_codser & ""
                dbI.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos='" & MuestraCasino(1) & "' AND VAL(MID(min_fecmin,1,6))=" & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser=" & RS1!min_codser & ""
             End If
          End If
          RS2.Close: Set RS2 = Nothing
          '-------> Traer ultimo correlativo
          If indice = 0 Then
             RS2.Open "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC", vg_db, adOpenStatic
             If Not RS2.EOF Then RS2.MoveFirst: indice = RS2!min_codigo + 1 Else indice = 1
             RS2.Close: Set RS2 = Nothing
          End If
          '-------> actualizar correlativo planificación base externa
          RS2.Open "SELECT DISTINCT min_codigo, min_codreg FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & "", dbI, adOpenStatic
          If Not RS2.EOF Then
             Do While Not RS2.EOF
                dbI.Execute "UPDATE b_minutadet SET mid_codigo = " & indice & ", mid_tiprec = " & RS2!min_codreg & " WHERE mid_codigo = " & RS2!min_codigo & ""
                dbI.Execute "UPDATE b_minuta SET min_codigo = " & indice & " WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codigo = " & RS2!min_codigo & ""
                RS2.MoveNext: indice = indice + 1
             Loop
          End If
          RS2.Close: Set RS2 = Nothing
          '-------> actualizar detalle planificación el campo mid_tiprec
          '-------> actualizar nro. raciones totales
          RS2.Open "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos = '" & MuestraCasino(1) & "' AND sra_codser = " & RS1!min_codser & " GROUP BY sra_serdia ORDER BY sra_serdia", vg_db, adOpenStatic
          If Not RS2.EOF Then
             Do While Not RS2.EOF
                dbI.Execute "UPDATE b_minuta SET min_racteo=" & RS2!raciones & " WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & " AND (min_racteo = 0 OR (min_racteo) is null) AND IIF(datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)) = 1,7,IIf(Val(Mid(min_fecmin, 5, 4)) = 229,datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)),datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)- 1))) = " & RS2!sra_serdia & ""
                RS2.MoveNext
             Loop
          End If
          RS2.Close: Set RS2 = Nothing
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing
    vg_db.CommitTrans

    '-------> Encabezado Planificación
    Label1(1).Caption = "Importando Planificación Encabezado": DoEvents
    RS1.Open "SELECT * FROM b_minuta", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minuta WHERE min_codigo = " & RS1!min_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Detalle Planificación
    Label1(1).Caption = "Importando Planificación Detalle": DoEvents
    RS1.Open "SELECT * FROM b_minutadet", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutadet WHERE mid_codigo = " & RS1!mid_codigo & " AND mid_tipmin = '" & RS1!mid_tipmin & "' AND mid_numlin = " & RS1!mid_numlin & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Costo Patron
    Label1(1).Caption = "Importando Costo Patron": DoEvents
    RS1.Open "SELECT * FROM b_costopatron", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_costopatron WHERE cpa_cencos = '" & RS1!cpa_cencos & "' AND cpa_codreg = " & RS1!cpa_codreg & " AND cpa_codser = " & RS1!cpa_codser & " AND cpa_anomes = " & RS1!cpa_anomes & " AND cpa_descripcion = '" & RS1!cpa_descripcion & "'"
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Tipo de Servicio
    Label1(1).Caption = "Importando tipo de servicio"
    DoEvents
    RS1.Open "SELECT * FROM a_tiposervicio", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tiposervicio WHERE tis_codigo = " & RS1!tis_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Segmento
    Label1(1).Caption = "Importando segmento"
    DoEvents
    RS1.Open "SELECT * FROM a_segmento", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_segmento WHERE seg_codigo = " & RS1!seg_codigo
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Actualizar tabla contrato moviendo tipo de servicio y segmento
    vg_db.BeginTrans
    vg_db.Execute "UPDATE b_clientes, a_tiposervicio, a_segmento SET b_clientes.cli_codtis = a_tiposervicio.tis_codigo, b_clientes.cli_codseg = a_segmento.seg_codigo WHERE b_clientes.cli_codigo = '" & MuestraCasino(1) & "'"
    vg_db.CommitTrans
    
    '-------> Actualizar Clientes
    Label1(1).Caption = "Importando sociedad sap"
    DoEvents
    vg_db.BeginTrans
    RS1.Open "SELECT * FROM a_cencos", dbI, adOpenStatic
    Do While Not RS1.EOF
       vg_db.Execute "UPDATE b_clientes SET cli_ccisac = " & RS1!cen_ccisac & ", cli_cecsac = '" & RS1!cen_cecsac & "', cli_socsap = '" & RS1!cen_socsap & "', cli_sobrec = '" & RS1!cen_sobrec & "', cli_codmun = " & RS1!cen_codmun & ", cli_codreg = " & RS1!cen_codreg & " WHERE cli_codigo = '" & RS1!cen_codigo & "'"
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    vg_db.CommitTrans
    
    '-------> Actualizar tabla lista producto y lista ingrediente
    Label1(1).Caption = "Importando producto & Ingredientes"
    DoEvents
'    On Error Resume Next
    vg_db.BeginTrans
    vg_db.Execute "SELECT DISTINCT * INTO " & vg_NUsr & "_Subeproductospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia = " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ""
    vg_db.Execute "INSERT INTO b_productospmpdia (ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo, ppd_upreco, ppd_fecuco) SELECT DISTINCT '" & MuestraCasino(1) & "', a.pro_codigo, " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", 0, 0, 0, '' FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_codigo NOT IN (SELECT DISTINCT ppd_codpro FROM " & vg_NUsr & "_Subeproductospmpdia)"
    vg_db.Execute "DROP TABLE " & vg_NUsr & "_Subeproductospmpdia"
    vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & MuestraCasino(1) & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente WHERE ing_codigo NOT IN (SELECT DISTINCT cpi_coding FROM b_contlistpreing WHERE cpi_cencos = '" & MuestraCasino(1) & "')"
    PB.Value = PB.Value + 1
    vg_db.CommitTrans
    
    '-------> Encabezado formato compras
    Label1(1).Caption = "Importando Encabezado formato compras": DoEvents
    RS1.Open "SELECT * FROM b_formatocompras", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_formatocompras WHERE foc_codsac = '" & RS1!foc_codsac & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Detalle formato compras
    Label1(1).Caption = "Importando Detalle formato compras": DoEvents
    RS1.Open "SELECT * FROM b_formatocomprassgp", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_formatocomprassgp WHERE fcs_codsac = '" & RS1!fcs_codsac & "' AND fcs_codsgp = '" & RS1!fcs_codsgp & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Lista precio sac
    Label1(1).Caption = "Importando Lista Precio Sac": DoEvents
    RS1.Open "SELECT * FROM b_sac_listaprecio", dbI, adOpenStatic
    Do While Not RS1.EOF
'        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_sac_listaprecio WHERE lps_cencos = '" & RS1!lps_cencos & "' AND lps_fecini = cdate('" & RS1!lps_fecini & "') AND lps_fecfin = cdate('" & RS1!lps_fecfin & "') AND lps_periodo = '" & RS1!lps_periodo & "' AND lps_codsac = '" & RS1!lps_codsac & "'"
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_sac_listaprecio WHERE lps_cencos = '" & RS1!lps_cencos & "' AND lps_periodo = '" & RS1!lps_periodo & "' AND lps_codsac = '" & RS1!lps_codsac & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Parametro Despachos
    Label1(1).Caption = "Importando Parametro de Despachos": DoEvents
    RS1.Open "SELECT * FROM b_paramdesp", dbI, adOpenStatic
    If vg_pais = "CO" And Not RS1.EOF Then
       vg_db.Execute "DELETE b_paramdesp FROM b_paramdesp"
    End If
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT pad_cencos, pad_codigo AS pad_codigo, pad_tipo, pad_diaseg, pad_diario  FROM b_paramdesp WHERE pad_cencos = '" & RS1!pad_cencos & "' AND pad_codigo = " & RS1!pad_codtip & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Fechas Inhabiles
    Label1(1).Caption = "Importando Fechas Inhabiles": DoEvents
    RS1.Open "SELECT * FROM b_Fecha_Inhabiles", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & RS1!CFI_CeCo & "' AND cdate(CFI_Fecha) = '" & RS1!CFI_Fecha & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipo Actividad
    Label1(1).Caption = "Importando Tipo Actividad": DoEvents
    RS1.Open "SELECT * FROM a_tipoactividad", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipoactividad WHERE tia_codigo = " & RS1!tia_codigo & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Casino Tipo Actividades
    Label1(1).Caption = "Importando Casino Tipo Actividades": DoEvents
    RS1.Open "SELECT * FROM b_casinotipoactividades", dbI, adOpenStatic
'    If Not RS1.EOF Then
       vg_db.Execute "DELETE b_casinotipoactividades FROM b_casinotipoactividades WHERE cta_cencos = '" & MuestraCasino(1) & "'"
'    End If
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_casinotipoactividades WHERE cta_cencos = '" & RS1!cta_cencos & "' AND cta_tipact = " & RS1!cta_tipact & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Casino Parametro Stock
    Label1(1).Caption = "Importando Casino Parametro Stock": DoEvents
    RS1.Open "SELECT * FROM b_casinoparametrostock", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_casinoparametrostock WHERE cps_cencos = '" & RS1!cps_cencos & "'"
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Municipio
    Label1(1).Caption = "Importando Municipio": DoEvents
    RS1.Open "SELECT * FROM a_municipio", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_municipio WHERE mun_codigo = " & RS1!mun_codigo & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Region
    Label1(1).Caption = "Importando Region": DoEvents
    RS1.Open "SELECT * FROM a_region", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_region WHERE reg_codigo = " & RS1!reg_codigo & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Clase Documento SAP
    Label1(1).Caption = "Importando Clase Documento SAP": DoEvents
    RS1.Open "SELECT * FROM a_clasedocsap", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_clasedocsap WHERE cds_coddoc = '" & RS1!cds_coddoc & "' AND cds_codreg = " & RS1!cds_codreg & ""
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Pais
    Label1(1).Caption = "Importando Pais": DoEvents
    RS1.Open "SELECT * FROM a_pais", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_pais WHERE pai_codigo = '" & RS1!pai_codigo & "'"
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Retención en la Fuente
    Label1(1).Caption = "Importando Retención en la Fuente": DoEvents
    RS1.Open "SELECT * FROM b_retencionfuente", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_retencionfuente WHERE ref_codigo = " & RS1!ref_codigo & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Retención Ica
    Label1(1).Caption = "Importando Retención Ica": DoEvents
    RS1.Open "SELECT * FROM b_retencionica", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_retencionica WHERE rei_codigo = " & RS1!rei_codigo & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Detalle Retención Ica
    Label1(1).Caption = "Importando Detalle Retención Ica": DoEvents
    RS1.Open "SELECT * FROM b_detretencionica", dbI, adOpenStatic
    Do While Not RS1.EOF
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detretencionica WHERE dri_codigo = " & RS1!dri_codigo & " AND dri_codmun = " & RS1!dri_codmun & ""
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    dbI.Close: Set dbI = Nothing
    vg_db.BeginTrans
    fso.DeleteFile cdbi
    vg_db.Execute "INSERT INTO log_actualizacion VALUES ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', cdate('" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm") & "'))"
    ActArchivoAccess = True
    vg_db.CommitTrans
    '-------> Actualizar producto vigente
    ValidarProductoVigente
End If

Exit Function
Man_Error:
If Err = -2147217865 Or Err = 3265 Or Err = -2147467259 Then
    If Err = 3265 Or Err = -2147467259 Then
'       vg_db.RollbackTrans
       Resume Next
    End If
    dbI.Close: Set dbI = Nothing
    '-------> Actualizar tabla lista producto y lista ingrediente
    vg_db.BeginTrans
    vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & MuestraCasino(1) & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente WHERE ing_codigo NOT IN (SELECT DISTINCT cpi_coding FROM b_contlistpreing WHERE cpi_cencos='" & MuestraCasino(1) & "')"
    vg_db.CommitTrans
    
    vg_db.BeginTrans
    fso.DeleteFile cdbi
    vg_db.Execute "INSERT INTO log_actualizacion VALUES ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', cdate('" & Format(Date, "dd/MM/yyyy") & " " & Format(Time, "HH:mm") & "'))"
    ActArchivoAccess = True
    vg_db.CommitTrans
    '-------> Actualizar producto vigente
    ValidarProductoVigente
   Exit Function
End If
If Err = 3034 Then vg_db.RollbackTrans: Exit Function
If Err = 53 Then ActArchivoAccess = True: Exit Function
If Err = -2147168242 Then ActArchivoAccess = True: Exit Function
vg_db.RollbackTrans
If Err.Number = -2147467259 Then
    MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, MsgTitulo
Else
    MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End If
End Function

Private Function ActArchivoSql(ByVal cdbz As String) As Long

Dim fso    As New FileSystemObject
Dim cdbi   As String
Dim indice As Long
Dim cDBO   As String
Dim DBO    As String
Dim spid   As Long
Dim RS1    As New ADODB.Recordset
Dim RS     As New ADODB.Recordset

On Error GoTo Man_Error

estarc = False
ActArchivoSql = False
If vg_contra <> Mid(cdbz, 4, InStr(cdbz, "-") - 4) Then MsgBox "Base de dato no corresponde centro de costo origen...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
If Not fso.FileExists(dir_trabajo & "Actualizar\" & cdbz) Then MsgBox "No se encuentra el archivo para importar datos...", vbExclamation + vbOKOnly, MsgTitulo: Exit Function
cdbi = dir_trabajo & "Actualizar\" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb"
cDBO = dir_trabajo & BaseDeDato
DBO = "'' [ODBC;PROVIDER=MSDASQL;driver={SQL Server};server=" + vg_SqlNSvr + ";uid=" + vg_SqlNUsr + ";pwd=" + vg_SqlPass + ";database=" + vg_SqlBase + ";]"
RS1.Open "SELECT * FROM log_actualizacion WHERE archivo = '" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "'", vg_db, adOpenStatic

If Not RS1.EOF Then
   
   If GetParametro("metactbd") <> 1 Then
      
      MsgBox "Ya fue aplicada la actualización para el archivo" & VgLinea & cdbi & "...", vbExclamation + vbOKOnly, MsgTitulo
   
   Else
      
      estarc = True
   
   End If
   RS1.Close: Set RS1 = Nothing
   Exit Function

End If
'If Not RS1.EOF Then MsgBox "Ya fue aplicada la actualización para el archivo" & VgLinea & cDBI & "...", vbExclamation + vbOKOnly, Msgtitulo: RS1.Close: Set RS1 = Nothing: Exit Function
RS1.Close: Set RS1 = Nothing

AZ1.OpenZip dir_trabajo & "Actualizar\" & cdbz
AZ1.ExtractFile AZ1.FileName(0), dir_trabajo & "Actualizar\" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb", ""
AZ1.Close

Set dbI = New ADODB.Connection
dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbI.ConnectionTimeout = 3600
dbI.CommandTimeout = 3600
dbI.Open

If UCase(Mid(cdbi, Len(dir_trabajo & "Actualizar\") + 1, 3)) = "SGP" Then
    
    RS1.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha FROM b_minuta WHERE min_fecmin IN (SELECT DISTINCT min_fecmin FROM b_minuta IN " & DBO & " WHERE min_indblo = 1 AND min_cencos = '" & MuestraCasino(1) & "')", dbI, adOpenStatic
    If Not RS1.EOF Then
       
       RS1.Close: Set RS1 = Nothing:
       dbI.Close: Set dbI = Nothing
       fso.DeleteFile cdbi: ActArchivoSql = False
'       Name Trim(Text1(0).Text) As Mid(Trim(Text1(0).Text), 1, Len(Trim(Text1(0).Text)) - 3) & "dwl"
       Name cdbz As Mid(cdbz, 1, Len(cdbz) - 3) & "dwl"
       Text1(0).text = ""
       MsgBox "Planificación minuta esta bloqueada, proceso cancelado...", vbInformation + vbOKOnly, MsgTitulo: ActArchivoSql = True: Exit Function
    
    End If
    
    RS1.Close: Set RS1 = Nothing
    
    '-------> Validar que minutas simap no suban
    RS1.Open "select top 1 min_cencos from b_minuta a where a.min_cencos in (SELECT DISTINCT cli_codigo FROM b_clientes IN " & DBO & " WHERE cli_codigo = '" & MuestraCasino(1) & "' and cli_tipo = 0 and cli_tipominuta = 3)", dbI, adOpenStatic
    If Not RS1.EOF Then
       
       RS1.Close: Set RS1 = Nothing:
       dbI.Close: Set dbI = Nothing
       fso.DeleteFile cdbi: ActArchivoSql = False
'       Name Trim(Text1(0).Text) As Mid(Trim(Text1(0).Text), 1, Len(Trim(Text1(0).Text)) - 3) & "dwl"
       Name cdbz As Mid(cdbz, 1, Len(cdbz) - 3) & "dwl"
       Text1(0).text = ""
       MsgBox "Planificación minuta de tipo Simap no puede actualizar, proceso cancelado...", vbInformation + vbOKOnly, MsgTitulo: ActArchivoSql = True: Exit Function
    
    End If
    RS1.Close: Set RS1 = Nothing
    
    '-------> Borrar tabla paso receta -
    vg_db.Execute "DELETE paso_receta WHERE rec_spid = @@spid AND rec_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_recetadet WHERE red_spid = @@spid AND red_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_productosing WHERE pri_spid = @@spid AND pri_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_regimen WHERE reg_spid = @@spid AND reg_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_productospmpdia WHERE ppd_spid = @@spid AND ppd_usuario = '" & vg_NUsr & "'"
    
    '-------> Buscar spid
    Set RS = vg_db.Execute("SELECT @@spid spid")
    If Not RS.EOF Then spid = RS!spid
    RS.Close: Set RS = Nothing

    PB.Min = 0: PB.Value = 0: PB.max = 53
    Label1(1).Visible = True: PB.Visible = True
    
    '-------> envío sap
    Label1(1).Caption = "Importando Envío Sap": DoEvents
    RS1.Open "SELECT * FROM a_tipointerfaz", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipointerfaz WHERE tii_codigo = " & RS1!tii_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
'    '-------> par_tipo_vales
'    If validarsiexistetabla(cdbi, "a_par_tipo_vales") Then
'       Label1(1).Caption = "Importando parametro tipo vales": DoEvents
'       RS1.Open "SELECT * FROM a_par_tipo_vales", dbI, adOpenStatic
'       Do While Not RS1.EOF
'          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_par_tipo_vales WHERE ID_Tipo_Vale = '" & RS1!ID_Tipo_Vale & "' and cli_codigo = '" & RS1!cli_codigo & "'"
'          RS1.MoveNext
'       Loop
'       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
'    End If
    
    If validarsiexistetabla(cdbi, "a_par_codigo_barra") Then
       
       '-------> par_codigo_barra_cas
       Label1(1).Caption = "Importando parametro código barra": DoEvents
       RS1.Open "SELECT * FROM a_par_codigo_barra", dbI, adOpenStatic
       Do While Not RS1.EOF
          
          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_par_codigo_barra_cas WHERE a_par_id_codigo = " & RS1!a_par_id_codigo & " and atr_codigo_barra = '" & RS1!atr_codigo_barra & "' and cli_codigo = '" & RS1!cli_codigo & "'"
          RS1.MoveNext
       
       Loop
       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    End If
    
    If validarsiexistetabla(cdbi, "Cuentas_Sap_AX") Then
       
       '-------> Cuentas_Sap_AX
       Label1(1).Caption = "Importando parametro cuenta contable OPTIMUM": DoEvents
       RS1.Open "SELECT * FROM Cuentas_Sap_AX", dbI, adOpenStatic
       Do While Not RS1.EOF

          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM Cuentas_Sap_AX WHERE Cuentas_Sap = " & RS1!Cuentas_Sap & ""
          RS1.MoveNext
       
       Loop
       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    End If
    
    If validarsiexistetabla(cdbi, "Sociedad_Sap_AX") Then
       
       '-------> Sociedad_Sap_AX
       Label1(1).Caption = "Importando parametro Sociedad contable OPTIMUM": DoEvents
       RS1.Open "SELECT * FROM Sociedad_Sap_AX", dbI, adOpenStatic
       Do While Not RS1.EOF

          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM Sociedad_Sap_AX WHERE Sociedad_Sap = '" & RS1!Sociedad_Sap & "'"
          RS1.MoveNext
       
       Loop
       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    End If
    
    If validarsiexistetabla(cdbi, "Cecos_Sap_AX") Then
       
       '-------> Ceco_Sap_AX
       Label1(1).Caption = "Importando parametro Cecos_Sap_Ax contable OPTIMUM": DoEvents
       RS1.Open "SELECT * FROM Cecos_Sap_AX", dbI, adOpenStatic
       Do While Not RS1.EOF

          ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM Cecos_Sap_AX WHERE Cecos_Sap = '" & RS1!Cecos_Sap & "' and Cecos_AX = '" & RS1!Cecos_AX & "'"
          RS1.MoveNext
       
       Loop
       RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    End If
    
    '-------> contrato envío sap
    Label1(1).Caption = "Importando Contrato Envío Sap": DoEvents
    vg_db.Execute "DELETE b_casinointerfaz FROM b_casinointerfaz WHERE cai_cencos = '" & MuestraCasino(1) & "'"
    RS1.Open "SELECT * FROM b_casinointerfaz", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_casinointerfaz WHERE cai_cencos = '" & RS1!cai_cencos & "' AND cai_codtii = " & RS1!cai_codtii & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipos de Producto
    Label1(1).Caption = "Importando Tipos de Producto": DoEvents
    dbI.Execute "ALTER TABLE a_tipopro ADD COLUMN tip_activo char(1)"
    dbI.Execute "UPDATE a_tipopro SET tip_activo = '1'"
    RS1.Open "SELECT * FROM a_tipopro", dbI, adOpenStatic
    If Not RS1.EOF Then
       vg_db.Execute "UPDATE a_tipopro SET tip_activo = 'N'"
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipopro WHERE tip_codigo = " & RS1!tip_codigo
        RS1.MoveNext
    
    Loop
    End If
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipos documento
    Label1(1).Caption = "Importando Tipos de Documento": DoEvents
    RS1.Open "SELECT * FROM a_tipodocumento", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipodocumento WHERE tdo_codigo = '" & RS1!tdo_codigo & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
'    '-------> Parametro Despacho
'    Label1(1).Caption = "Importando Parametro de Despacho": DoEvents
'    RS1.Open "SELECT b.tip_codigo, b.tip_nombre FROM a_tipopro a INNER JOIN a_tipopro AS b ON a.tip_codigo = b.tip_previo WHERE a.tip_previo = 0", vg_db, adOpenStatic
'    If Not RS1.EOF Then
'       Do While Not RS1.EOF
'          RS2.Open "SELECT DISTINCT pad_codigo FROM b_paramdesp WHERE pad_cencos = '" & MuestraCasino(1) & "' AND pad_codigo = " & RS1!tip_codigo & "", vg_db, adOpenStatic
'          If RS2.EOF Then vg_db.Execute "INSERT INTO b_paramdesp VALUES (" & RS1!tip_codigo & ", 'S', '" & MuestraCasino(1) & "', '')"
'          RS2.Close: Set RS2 = Nothing
'          RS1.MoveNext
'       Loop
'    End If
'    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
     
    '-------> Unidades de medida
    Label1(1).Caption = "Importando Unidades de Medida": DoEvents
    RS1.Open "SELECT * FROM a_unidadmed", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidadmed WHERE unm_codigo = " & RS1!unm_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de stock
    Label1(1).Caption = "Importando Unidades de Stock"
    DoEvents
    RS1.Open "SELECT * FROM a_unidad", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_unidad WHERE uni_codigo = " & RS1!uni_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Unidades de embalaje
    Label1(1).Caption = "Importando unidades de embalaje"
    DoEvents
    RS1.Open "SELECT * FROM a_embalaje", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_embalaje WHERE emb_codigo = " & RS1!emb_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Cuentas Contables
    Label1(1).Caption = "Importando Cuentas Contables"
    DoEvents
    RS1.Open "SELECT * FROM a_ctacontable", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_ctacontable WHERE cta_codigo = '" & RS1!cta_codigo & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Parametro
    '-------> Agregar campo
    dbI.Execute "alter table a_param add column par_cencos char(10)"
    dbI.Execute "update a_param set par_cencos = '" & MuestraCasino(1) & "'"
    
    Label1(1).Caption = "Importando Parametros"
    DoEvents
    RS1.Open "SELECT * FROM a_param", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_param WHERE par_cencos = '" & RS1!par_cencos & "' AND par_codigo = '" & RS1!par_codigo & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Impuestos
    Label1(1).Caption = "Importando Impuestos"
    DoEvents
    RS1.Open "SELECT * FROM a_impuesto", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_impuesto WHERE imp_codigo = " & RS1!imp_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Nutrientes
    Label1(1).Caption = "Importando Nutrientes"
    DoEvents
    RS1.Open "SELECT * FROM a_nutriente", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_nutriente WHERE nut_codigo = " & RS1!nut_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Articulos de Stock
    Label1(1).Caption = "Importando Artículos de Stock": DoEvents
    RS1.Open "SELECT * FROM b_productos", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productos WHERE pro_codigo = '" & RS1!pro_codigo & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Impuestos Articulos de Stock
    Label1(1).Caption = "Importando Impuestos Relacionados": DoEvents
    RS1.Open "SELECT * FROM b_productosimp", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosimp WHERE ipr_codpro = '" & RS1!ipr_codpro & "' AND ipr_codimp = " & RS1!ipr_codimp
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Ingredientes
    Label1(1).Caption = "Importando Ingredientes": DoEvents
    RS1.Open "SELECT * FROM b_ingrediente", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_ingrediente WHERE ing_codigo = '" & RS1!ing_codigo & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Ingredientes Articulos de Stock
    Label1(1).Caption = "Importando Ingredientes Relacionados": DoEvents
'    vg_db.Execute "DELETE FROM b_productosing WHERE pri_codpro IN (SELECT pri_codpro FROM b_productosing IN '" & cDBI & "')"
    RS1.Open "SELECT * FROM b_productosing", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       vg_db.Execute "INSERT INTO paso_productosing VALUES (" & spid & ", '" & vg_NUsr & "', '" & RS1!pri_codpro & "')"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    
    vg_db.Execute "DELETE FROM b_productosing WHERE pri_codpro IN (SELECT pri_codpro FROM paso_productosing WHERE pri_spid = " & spid & " AND pri_usr = '" & vg_NUsr & "')"
    RS1.Open "SELECT * FROM b_productosing", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productosing WHERE pri_codpro = '" & RS1!pri_codpro & "' AND pri_coding = '" & RS1!pri_coding & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Aportes Nutricionales Ingrediente
    Label1(1).Caption = "Importando Aportes Nutricionales Ingrediente": DoEvents
    RS1.Open "SELECT * FROM b_productonut", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_productonut WHERE pnu_codpro = '" & RS1!pnu_codpro & "' AND pnu_codapo = " & RS1!pnu_codapo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Proveedores
    Label1(1).Caption = "Importando Proveedores": DoEvents
    RS1.Open "SELECT * FROM b_proveedor", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_proveedor WHERE prv_codigo = '" & RS1!prv_codigo & "'"
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    '-------> Validar si existe actualización posterioes maestro de productos
    RS1.Open "SELECT MAX(prv_fecumo) AS prv_fecumo FROM b_proveedor", vg_db, adOpenStatic
    If Not RS1.EOF And Not IsNull(RS1!prv_fecumo) Then
       '-------> Actualizar tabla actualiza dato
       
       vg_db.Execute "UPDATE b_actuadatos SET ada_fecumo = '" & Format(RS1!prv_fecumo, "yyyymmdd") & "' WHERE ada_nomtab = 'b_proveedor' AND (ada_fecumo < '" & Format(RS1!prv_fecumo, "yyyymmdd") & "' OR (ada_fecumo) IS NULL)"
    
    End If
    RS1.Close: Set RS1 = Nothing
    '-------> Mover zero al stock si es negativo
    vg_db.Execute "UPDATE b_bodegas set bod_canmer = 0 WHERE bod_codbod = " & vg_codbod & " AND bod_canmer < 0"
    
    '-------> Categoría de Receta
    Label1(1).Caption = "Importando Categoría de Receta": DoEvents
    RS1.Open "SELECT * FROM a_recetacatdie", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetacatdie WHERE car_codigo = " & RS1!car_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Tipo de Plato
    Label1(1).Caption = "Importando Tipo de Plato": DoEvents
    RS1.Open "SELECT * FROM a_recetatippla", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_recetatippla WHERE tip_codigo = " & RS1!tip_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Gramos Familia Producto
    Label1(1).Caption = "Importando Gramos Familia Producto": DoEvents
    RS1.Open "SELECT * FROM b_gramofamproducto", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_gramofamproducto WHERE gfp_cencos = '" & RS1!gfp_cencos & "' AND gfp_codreg = " & RS1!gfp_codreg & " AND gfp_catdie = " & RS1!gfp_catdie & " AND gfp_tiprec = " & RS1!gfp_tiprec & " AND gfp_fampro = " & RS1!gfp_fampro & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Recetas
    Label1(1).Caption = "Importando Recetas": DoEvents
    RS1.Open "SELECT * FROM b_receta", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_receta WHERE rec_codigo = " & RS1!rec_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Ingredientes de Recetas
    '-------> Agregar campo
    dbI.Execute "ALTER TABLE b_recetadet ADD COLUMN red_cencos char(10)"
    dbI.Execute "UPDATE b_recetadet SET red_cencos = '0' WHERE red_tiprec = 0"
    dbI.Execute "UPDATE b_recetadet SET red_cencos = '" & MuestraCasino(1) & "' WHERE red_tiprec <> 0"
    Label1(1).Caption = "Importando Ingredientes Recetas": DoEvents
    '-------> Respaldar recetas 5 epatas y patron
    dbI.Execute "ALTER TABLE b_receta ADD Constraint b_receta_pk Primary Key (rec_codigo)"
    dbI.Execute "ALTER TABLE b_recetadet ADD Constraint b_recetadet_pk Primary Key (red_codigo, red_nroite, red_tiprec, red_cencos)"
    RS1.Open "SELECT DISTINCT * INTO r_recetadet FROM b_recetadet IN " & DBO & " WHERE red_codigo IN (SELECT DISTINCT red_codigo FROM b_recetadet IN '" & cdbi & "' WHERE (red_tiprec=0 OR red_tiprec>=10000) AND ((red_tiprec<>0 AND red_cencos='" & MuestraCasino(1) & "') OR (red_tiprec=0 AND red_cencos='0'))) AND (red_tiprec=0 OR red_tiprec>=10000) AND ((red_tiprec>0 AND red_cencos='" & MuestraCasino(1) & "') OR (red_tiprec=0 AND red_cencos='0'))", dbI, adOpenStatic
    Set RS1 = Nothing
    RS1.Open "SELECT DISTINCT rec_codigo INTO r_receta FROM b_receta IN " & DBO & " WHERE rec_codigo IN (SELECT DISTINCT rec_codigo FROM b_receta IN '" & cdbi & "' )", dbI, adOpenStatic
    Set RS1 = Nothing
    RS1.Open "SELECT DISTINCT rec_codigo, '0' AS rec_tiprec INTO x_receta FROM b_receta IN " & DBO & " WHERE rec_codigo IN (SELECT DISTINCT rec_codigo FROM b_receta IN '" & cdbi & "' )", dbI, adOpenStatic
    Set RS1 = Nothing
    dbI.Execute "ALTER TABLE r_receta ADD Constraint r_receta_pk Primary Key (rec_codigo)"
    dbI.Execute "ALTER TABLE x_receta ADD Constraint x_receta_pk Primary Key (rec_codigo)"
    dbI.Execute "ALTER TABLE b_tablagramaje ADD Constraint b_tablagramaje_pk Primary Key (tgr_codreg,tgr_codrec,tgr_coding)"
    '-------> Insertar recetas 5 etapas
    RS1.Open "SELECT DISTINCT * FROM a_regimen WHERE reg_codigo >= 10000", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       DoEvents
       '-------> Mover receta desde tabla gramaje
       dbI.Execute "UPDATE x_receta SET x_receta.rec_tiprec = '0'"
       dbI.Execute "UPDATE x_receta INNER JOIN b_tablagramaje ON x_receta.rec_codigo = b_tablagramaje.tgr_codrec SET x_receta.rec_tiprec = '1' " & _
                   "WHERE b_tablagramaje.tgr_codreg = " & RS1!reg_codigo & ""
       dbI.Execute "INSERT INTO b_recetadet (red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos) SELECT DISTINCT b.red_codigo, b.red_nroite, b.red_codpro, b.red_canpro, b.red_cospro, b.red_pctapr, b.red_pctcoc, b.red_pctnut, " & RS1!reg_codigo & " AS red_tiprec, '" & MuestraCasino(1) & "' AS red_cencos FROM x_receta a, b_recetadet b WHERE a.rec_codigo = b.red_codigo AND (a.rec_tiprec = '0' and b.red_tiprec = 0)"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    '-------> Leer tabla cencos si sobreescribe receta
    Dim sobrec As String
    sobrec = ""
    Label1(1).Caption = "Actualizando Detalle Recetas 5 Etapas": DoEvents
    RS1.Open "SELECT * FROM a_cencos", dbI, adOpenStatic
    If Not RS1.EOF Then sobrec = RS1!cen_sobrec
    RS1.Close: Set RS1 = Nothing
    
    If sobrec = "0" Then '-------> Actualiza solamente las recetas que tienen parametro de actualziación
       
       '-------> Crear tabla temporal de parametro de recetas
       dbI.Execute "SELECT DISTINCT b_ingrediente.ing_codigo INTO b_detparametro " & _
                   "FROM b_clientes, b_ingrediente, b_recetadet, b_receta, b_gramofamproducto, a_regimen, b_productos IN " & DBO & " " & _
                   "WHERE b_recetadet.red_codpro = b_ingrediente.ing_codigo " & _
                   "AND   b_ingrediente.ing_codcom = b_productos.pro_codigo " & _
                   "AND   b_receta.rec_catdie = b_gramofamproducto.gfp_catdie " & _
                   "AND   b_receta.rec_tippla = b_gramofamproducto.gfp_tiprec " & _
                   "AND   b_receta.rec_codigo = b_recetadet.red_codigo " & _
                   "AND   b_productos.pro_codtip = b_gramofamproducto.gfp_fampro " & _
                   "AND   a_regimen.reg_codigo = b_gramofamproducto.gfp_codreg " & _
                   "AND   b_clientes.cli_codigo = b_gramofamproducto.gfp_cencos " & _
                   "AND   b_gramofamproducto.gfp_grafin > 0 AND b_clientes.cli_codigo = '" & MuestraCasino(1) & "'"
       
       '-------> Actualizar recetas 5 etapas
       dbI.Execute "UPDATE (b_recetadet INNER JOIN r_recetadet ON (b_recetadet.red_codigo = r_recetadet.red_codigo) AND (b_recetadet.red_nroite = r_recetadet.red_nroite) AND (b_recetadet.red_codpro = r_recetadet.red_codpro) AND (b_recetadet.red_tiprec = r_recetadet.red_tiprec) AND (b_recetadet.red_cencos = r_recetadet.red_cencos)) INNER JOIN b_detparametro ON r_recetadet.red_codpro = b_detparametro.ing_codigo SET b_recetadet.red_canpro = r_recetadet.red_canpro " & _
                   "WHERE b_recetadet.red_tiprec >= 10000 AND r_recetadet.red_tiprec >= 10000 AND r_recetadet.red_cencos = '" & MuestraCasino(1) & "'"
    ElseIf sobrec = "2" Then '-------> No actualiza ninguna receta
       
       dbI.Execute "UPDATE b_recetadet INNER JOIN r_recetadet ON (b_recetadet.red_cencos = r_recetadet.red_cencos) AND (b_recetadet.red_tiprec = r_recetadet.red_tiprec) AND (b_recetadet.red_codpro = r_recetadet.red_codpro) AND (b_recetadet.red_nroite = r_recetadet.red_nroite) AND (b_recetadet.red_codigo = r_recetadet.red_codigo) SET b_recetadet.red_canpro = r_recetadet.red_canpro " & _
                   "WHERE b_recetadet.red_tiprec >= 10000 AND r_recetadet.red_tiprec >= 10000 AND r_recetadet.red_cencos = '" & MuestraCasino(1) & "'"
    
    End If
    PB.Value = PB.Value + 1
    DoEvents
    '-------> Mover datos tabla receta paso y regimen
    RS1.Open "SELECT DISTINCT rec_codigo FROM r_receta", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       vg_db.Execute "INSERT INTO paso_receta VALUES (" & spid & ", '" & vg_NUsr & "', " & RS1!rec_codigo & ")"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    RS1.Open "SELECT DISTINCT reg_codigo FROM a_regimen", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       vg_db.Execute "INSERT INTO paso_regimen VALUES (" & spid & ", '" & vg_NUsr & "', " & RS1!reg_codigo & ")"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    '-------> borrar detalle receta patron
    vg_db.Execute "DELETE FROM b_recetadet WHERE red_codigo IN (SELECT DISTINCT rec_codigo FROM paso_receta WHERE rec_spid = " & spid & " AND rec_usr = '" & vg_NUsr & "') AND red_tiprec = 0 AND red_cencos = '0'"

    DoEvents
    vg_db.Execute "DELETE b_recetadet FROM b_recetadet a WHERE a.red_codigo IN (SELECT DISTINCT rec_codigo FROM paso_receta WHERE rec_spid = " & spid & " AND rec_usr = '" & vg_NUsr & "') AND a.red_tiprec IN (SELECT DISTINCT reg_codigo FROM paso_regimen WHERE reg_codigo >= 10000 AND reg_spid = " & spid & " AND reg_usr = '" & vg_NUsr & "') AND a.red_cencos = '" & MuestraCasino(1) & "' AND a.red_tiprec >= 10000"
    DoEvents
    
    RS1.Open "SELECT DISTINCT red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos FROM b_recetadet", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       vg_db.Execute "INSERT INTO paso_recetadet VALUES (" & spid & ", '" & vg_NUsr & "', " & RS1!red_codigo & ", " & RS1!red_nroite & ", '" & RS1!red_codpro & "', " & RS1!red_canpro & ", " & RS1!red_cospro & ", " & RS1!red_pctapr & ", " & RS1!red_pctcoc & ", " & RS1!red_pctnut & ", " & RS1!red_tiprec & ", '" & RS1!red_cencos & "')"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing
    '-------> borrar detalle receta local
    DoEvents
    vg_db.Execute "DELETE FROM b_recetadet WHERE red_codigo IN (SELECT DISTINCT red_codigo FROM paso_recetadet WHERE red_spid = " & spid & " AND red_usr = '" & vg_NUsr & "' AND red_tiprec = -1) AND red_tiprec = -1 AND red_cencos = '" & MuestraCasino(1) & "'"

    vg_db.Execute "INSERT INTO b_recetadet SELECT DISTINCT red_codigo, red_nroite, red_codpro, red_canpro, red_cospro, red_pctapr, red_pctcoc, red_pctnut, red_tiprec, red_cencos FROM paso_recetadet  WHERE red_usr = '" & vg_NUsr & "' AND red_spid = " & spid & " ORDER BY red_codigo, red_nroite, red_tiprec, red_cencos"
    PB.Value = PB.Value + 1
   
    '-------> Importando Regimen
    Label1(1).Caption = "Importando Regimen": DoEvents
    RS1.Open "SELECT * FROM a_regimen", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_regimen WHERE reg_codigo=" & RS1!reg_codigo
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '------> Importando Servicio
    Label1(1).Caption = "Importando Servicio": DoEvents
    RS1.Open "SELECT * FROM a_servicio", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_servicio WHERE ser_codigo=" & RS1!ser_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '------> Importando Estructura Servicio
    Label1(1).Caption = "Importando Estructura Servicio": DoEvents
    dbI.Execute "ALTER TABLE a_estservicio ADD COLUMN ess_cencos char(10)"
    dbI.Execute "UPDATE a_estservicio SET ess_cencos='" & MuestraCasino(1) & "'"
    RS1.Open "SELECT * FROM a_estservicio", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_estservicio WHERE ess_codser=" & RS1!ess_codser & " AND ess_codigo=" & RS1!ess_codigo & " AND ess_cencos='" & RS1!ess_cencos & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Validar si existe planificación minutas
    indice = 0
    Label1(1).Caption = "Validar Planificación Minutas": DoEvents
    RS1.Open "SELECT DISTINCT VAL(MID(min_fecmin,1,6)) AS fecha, min_codreg, reg_nombre, min_codser, ser_nombre FROM b_minuta a, a_regimen b, a_servicio c WHERE a.min_cencos='" & MuestraCasino(1) & "' AND a.min_codreg=b.reg_codigo AND a.min_codser=c.ser_codigo", dbI, adOpenStatic
    If Not RS1.EOF Then
       Do While Not RS1.EOF
          DoEvents
          RS2.Open "SELECT DISTINCT substring(convert(varchar(8),min_fecmin),1,6) AS fecha FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND substring(convert(varchar(8),min_fecmin),1,6) = " & RS1!Fecha & " AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & "", vg_db, adOpenStatic
          If Not RS2.EOF Then
             If MsgBox("Existe planificación minuta, desea borrar la información existente... " & VgLinea & VgLinea & "Regimen  : " & RS1!min_codreg & " " & Trim(RS1!reg_nombre) & VgLinea & "Servicio   :  " & RS1!min_codser & " " & Trim(RS1!ser_nombre), vbQuestion + vbYesNo, MsgTitulo) = vbYes Then
                '-------> Borrar planificación contrato
                vg_db.Execute "DELETE b_minutadet FROM b_minuta, b_minutadet WHERE b_minuta.min_codigo = b_minutadet.mid_codigo AND b_minuta.min_cencos = '" & MuestraCasino(1) & "' AND substring(convert(varchar(8),b_minuta.min_fecmin),1,6) = " & RS1!Fecha & " AND b_minuta.min_codreg = " & RS1!min_codreg & " AND b_minuta.min_codser = " & RS1!min_codser & ""
                vg_db.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND substring(convert(varchar(8),min_fecmin),1,6) = " & RS1!Fecha & " AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & ""
             Else
                '-------> Borrar planificación de la base carga
                dbI.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo = b_minutadet.mid_codigo WHERE b_minuta.min_cencos = '" & MuestraCasino(1) & "'AND VAL(MID(b_minuta.min_fecmin,1,6)) = " & RS1!Fecha & " AND b_minuta.min_codreg = " & RS1!min_codreg & " AND b_minuta.min_codser = " & RS1!min_codser & ""
                dbI.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND VAL(MID(min_fecmin,1,6)) = " & RS1!Fecha & " AND min_codreg=" & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & ""
             End If
          End If
          RS2.Close: Set RS2 = Nothing
          '-------> Traer ultimo correlativo
          If indice = 0 Then
             RS2.Open "SELECT min_codigo FROM b_minuta ORDER BY min_codigo DESC", vg_db, adOpenStatic
             If Not RS2.EOF Then RS2.MoveFirst: indice = RS2!min_codigo + 1 Else indice = 1
             RS2.Close: Set RS2 = Nothing
          End If
          '-------> actualizar correlativo planificación base externa
          RS2.Open "SELECT DISTINCT min_codigo, min_codreg FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & "", dbI, adOpenStatic
          If Not RS2.EOF Then
             Do While Not RS2.EOF
                dbI.Execute "UPDATE b_minutadet SET mid_codigo = " & indice & ", mid_tiprec = " & RS2!min_codreg & " WHERE mid_codigo = " & RS2!min_codigo & ""
                dbI.Execute "UPDATE b_minuta SET min_codigo = " & indice & " WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codigo = " & RS2!min_codigo & ""
                RS2.MoveNext: indice = indice + 1
             Loop
          End If
          RS2.Close: Set RS2 = Nothing
          '-------> actualizar detalle planificación el campo mid_tiprec
          '-------> actualizar nro. raciones totales
          RS2.Open "SELECT sra_serdia, SUM(sra_raciones) AS raciones FROM a_serviciorac WHERE sra_cencos = '" & MuestraCasino(1) & "' AND sra_codser = " & RS1!min_codser & " GROUP BY sra_serdia ORDER BY sra_serdia", vg_db, adOpenStatic
          If Not RS2.EOF Then
             Do While Not RS2.EOF
                dbI.Execute "UPDATE b_minuta SET min_racteo=" & RS2!raciones & " WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_codreg = " & RS1!min_codreg & " AND min_codser = " & RS1!min_codser & " AND (min_racteo = 0 OR (min_racteo) is null) AND IIF(datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)) = 1,7,IIf(Val(Mid(min_fecmin, 5, 4)) = 229,datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)),datepart('w',Mid(min_fecmin, 7, 2) & '/' & Mid(min_fecmin, 5, 2) & '/' & Mid(min_fecmin, 1, 4)- 1))) = " & RS2!sra_serdia & ""
                RS2.MoveNext
             Loop
          End If
          RS2.Close: Set RS2 = Nothing
          RS1.MoveNext
       Loop
    End If
    RS1.Close: Set RS1 = Nothing

    '-------> Encabezado Planificación
    Label1(1).Caption = "Importando Planificación Encabezado": DoEvents
    RS1.Open "SELECT * FROM b_minuta", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minuta WHERE min_codigo = " & RS1!min_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Detalle Planificación
    Label1(1).Caption = "Importando Planificación Detalle": DoEvents
    RS1.Open "SELECT * FROM b_minutadet", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_minutadet WHERE mid_codigo = " & RS1!mid_codigo & " AND mid_tipmin = '" & RS1!mid_tipmin & "' AND mid_numlin = " & RS1!mid_numlin & ""
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Costo Patron
    Label1(1).Caption = "Importando Costo Patron": DoEvents
    RS1.Open "SELECT * FROM b_costopatron", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_costopatron WHERE cpa_cencos = '" & RS1!cpa_cencos & "' AND cpa_codreg = " & RS1!cpa_codreg & " AND cpa_codser = " & RS1!cpa_codser & " AND cpa_anomes = " & RS1!cpa_anomes & " AND cpa_descripcion = '" & RS1!cpa_descripcion & "'"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1

    '-------> Tipo de Servicio
    Label1(1).Caption = "Importando tipo de servicio"
    DoEvents
    RS1.Open "SELECT * FROM a_tiposervicio", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tiposervicio WHERE tis_codigo = " & RS1!tis_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Segmento
    Label1(1).Caption = "Importando segmento"
    DoEvents
    RS1.Open "SELECT * FROM a_segmento", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_segmento WHERE seg_codigo = " & RS1!seg_codigo
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Actualizar tabla contrato moviendo tipo de servicio y segmento
    vg_db.Execute "UPDATE b_clientes SET b_clientes.cli_codtis = a_tiposervicio.tis_codigo, b_clientes.cli_codseg = a_segmento.seg_codigo FROM b_clientes, a_tiposervicio, a_segmento WHERE b_clientes.cli_codigo = '" & MuestraCasino(1) & "'"
    
    '-------> Actualizar Clientes
    Label1(1).Caption = "Importando sociedad sap"
    DoEvents
    RS1.Open "SELECT * FROM a_cencos", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       vg_db.Execute "UPDATE b_clientes SET cli_ccisac = " & RS1!cen_ccisac & ", cli_cecsac = '" & RS1!cen_cecsac & "', cli_socsap = '" & RS1!cen_socsap & "', cli_sobrec = '" & RS1!cen_sobrec & "', cli_codmun = " & RS1!cen_codmun & ", cli_codreg = " & RS1!cen_codreg & " WHERE cli_codigo = '" & RS1!cen_codigo & "'"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Actualizar ceco OPTIMUM
    Label1(1).Caption = "Importando ceco OPTIMUM"
    DoEvents
    RS1.Open "SELECT * FROM a_cencos", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       Set RS = vg_db.Execute("select * from Cecos_Sap_AX where Cecos_Sap = '" & RS1!cen_codigo & "'")
       If Not RS.EOF Then
          vg_db.Execute ("UPDATE Cecos_Sap_AX SET Cecos_AX = '" & RS1!cen_codopt & "',  Sociedad_Sap  =  '" & RS1!cen_socsap & "' WHERE Cecos_Sap = '" & RS1!cen_codigo & "'")
       Else
          vg_db.Execute ("insert into Cecos_Sap_AX (Cecos_Sap, Cecos_AX, Sociedad_Sap) values ('" & RS1!cen_codigo & "', '" & RS1!cen_codopt & "', '" & RS1!cen_socsap & "')")
       End If
       RS.Close: Set RS = Nothing
       
       RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Actualizar tabla lista producto y lista ingrediente
    Label1(1).Caption = "Importando producto & Ingredientes"
    DoEvents
    vg_db.Execute ("sgp_s_pasoproductospmpdia " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", " & Format(CDate(vg_ciedia), "yyyymmdd") & ", '" & MuestraCasino(1) & "', '" & vg_NUsr & "', " & spid & "")
    vg_db.Execute "INSERT INTO b_productospmpdia (ppd_cencos, ppd_codpro, ppd_fecdia, ppd_propon, ppd_saldo, ppd_upreco, ppd_fecuco) SELECT DISTINCT '" & MuestraCasino(1) & "', a.pro_codigo, " & Format(CDate(vg_ciedia) - 1, "yyyymmdd") & ", 0, 0, 0, '' FROM b_productos a, a_tiposervicio b, b_clientes c WHERE (b.tis_codigo = c.cli_codtis OR a.pro_maepro < 1) AND c.cli_codigo = '" & MuestraCasino(1) & "' AND (b.tis_codigo = a.pro_maepro OR a.pro_maepro < 1) AND a.pro_codigo NOT IN (SELECT DISTINCT ppd_codpro FROM paso_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_usuario = '" & vg_NUsr & "' AND ppd_spid = " & spid & ")"
    vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & MuestraCasino(1) & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente WHERE ing_codigo NOT IN (SELECT DISTINCT cpi_coding FROM b_contlistpreing WHERE cpi_cencos = '" & MuestraCasino(1) & "')"
    PB.Value = PB.Value + 1
    
    '-------> Encabezado formato compras
    Label1(1).Caption = "Importando Encabezado formato compras": DoEvents
    RS1.Open "SELECT * FROM b_formatocompras", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_formatocompras WHERE foc_codsac = '" & RS1!foc_codsac & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Detalle formato compras
    Label1(1).Caption = "Importando Detalle formato compras": DoEvents
    RS1.Open "SELECT * FROM b_formatocomprassgp", dbI, adOpenStatic
    Do While Not RS1.EOF
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_formatocomprassgp WHERE fcs_codsac = '" & RS1!fcs_codsac & "' AND fcs_codsgp = '" & RS1!fcs_codsgp & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Lista precio sac
    Label1(1).Caption = "Importando Lista Precio Sac": DoEvents
    RS1.Open "SELECT * FROM b_sac_listaprecio", dbI, adOpenStatic
    Do While Not RS1.EOF
'        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_sac_listaprecio WHERE lps_cencos = '" & RS1!lps_cencos & "' AND lps_fecini = cdate('" & RS1!lps_fecini & "') AND lps_fecfin = cdate('" & RS1!lps_fecfin & "') AND lps_periodo = '" & RS1!lps_periodo & "' AND lps_codsac = '" & RS1!lps_codsac & "'"
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_sac_listaprecio WHERE lps_cencos = '" & RS1!lps_cencos & "' AND lps_periodo = '" & RS1!lps_periodo & "' AND lps_codsac = '" & RS1!lps_codsac & "'"
        RS1.MoveNext
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Parametro Despachos
    Label1(1).Caption = "Importando Parametro de Despachos": DoEvents
    RS1.Open "SELECT * FROM b_paramdesp", dbI, adOpenStatic
    If vg_pais = "CO" And Not RS1.EOF Then
       vg_db.Execute "DELETE b_paramdesp FROM b_paramdesp"
    End If
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT pad_cencos, pad_codigo AS pad_codigo, pad_tipo, pad_diaseg, pad_diario  FROM b_paramdesp WHERE pad_cencos = '" & RS1!pad_cencos & "' AND pad_codigo = " & RS1!pad_codtip & ""
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Fechas Inhabiles
    Label1(1).Caption = "Importando Fechas Inhabiles": DoEvents
    RS1.Open "SELECT CFI_CeCo, CFI_Fecha AS CFI_Fecha, CFI_Glosa FROM b_Fecha_Inhabiles", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT CFI_CeCo, CFI_Fecha, CFI_Glosa FROM b_Fecha_Inhabiles WHERE CFI_CeCo = '" & RS1!CFI_CeCo & "' AND convert(varchar(10),CFI_Fecha,103) = '" & RS1!CFI_Fecha & "'"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Tipo Actividad
    Label1(1).Caption = "Importando Tipo Actividad": DoEvents
    RS1.Open "SELECT * FROM a_tipoactividad", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_tipoactividad WHERE tia_codigo = " & RS1!tia_codigo & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Casino Tipo Actividades
    Label1(1).Caption = "Importando Casino Tipo Actividades": DoEvents
    RS1.Open "SELECT * FROM b_casinotipoactividades", dbI, adOpenStatic
'    If Not RS1.EOF Then
       vg_db.Execute "DELETE b_casinotipoactividades FROM b_casinotipoactividades WHERE cta_cencos = '" & MuestraCasino(1) & "'"
'    End If
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_casinotipoactividades WHERE cta_cencos = '" & RS1!cta_cencos & "' AND cta_tipact = " & RS1!cta_tipact & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Casino Parametro Stock
    Label1(1).Caption = "Importando Casino Parametro Stock": DoEvents
    RS1.Open "SELECT * FROM b_casinoparametrostock", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_casinoparametrostock WHERE cps_cencos = '" & RS1!cps_cencos & "'"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Pais
    Label1(1).Caption = "Importando Pais": DoEvents
    RS1.Open "SELECT * FROM a_pais", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_pais WHERE pai_codigo = '" & RS1!pai_codigo & "'"
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Municipio
    Label1(1).Caption = "Importando Municipio": DoEvents
    RS1.Open "SELECT * FROM a_municipio", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_municipio WHERE mun_codigo = " & RS1!mun_codigo & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Region
    Label1(1).Caption = "Importando Region": DoEvents
    RS1.Open "SELECT * FROM a_region", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_region WHERE reg_codigo = " & RS1!reg_codigo & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Clase Documento SAP
    Label1(1).Caption = "Importando Clase Documento SAP": DoEvents
    RS1.Open "SELECT * FROM a_clasedocsap", dbI, adOpenStatic
    Do While Not RS1.EOF
        
        ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM a_clasedocsap WHERE cds_coddoc = '" & RS1!cds_coddoc & "' AND cds_codreg = " & RS1!cds_codreg & ""
        RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Retención en la Fuente
    Label1(1).Caption = "Importando Retención en la Fuente": DoEvents
    RS1.Open "SELECT * FROM b_retencionfuente", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_retencionfuente WHERE ref_codigo = " & RS1!ref_codigo & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Retención Ica
    Label1(1).Caption = "Importando Retención Ica": DoEvents
    RS1.Open "SELECT * FROM b_retencionica", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_retencionica WHERE rei_codigo = " & RS1!rei_codigo & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Detalle Retención Ica
    Label1(1).Caption = "Importando Detalle Retención Ica": DoEvents
    RS1.Open "SELECT * FROM b_detretencionica", dbI, adOpenStatic
    Do While Not RS1.EOF
       
       ActRegistro RS1.Fields, dbI, vg_db, "SELECT * FROM b_detretencionica WHERE dri_codigo = " & RS1!dri_codigo & " AND dri_codmun = " & RS1!dri_codmun & ""
       RS1.MoveNext
    
    Loop
    RS1.Close: Set RS1 = Nothing: PB.Value = PB.Value + 1
    
    '-------> Borrar tabla paso receta -
    vg_db.Execute "DELETE paso_receta WHERE rec_spid = " & spid & " AND rec_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_recetadet WHERE red_spid = " & spid & " AND red_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_productosing WHERE pri_spid = " & spid & " AND pri_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_regimen WHERE reg_spid = " & spid & " AND reg_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_productospmpdia WHERE ppd_spid = " & spid & " AND ppd_usuario = '" & vg_NUsr & "'"
    
    dbI.Close: Set dbI = Nothing
    fso.DeleteFile cdbi
    vg_db.Execute "INSERT INTO log_actualizacion VALUES ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', '" & Format(Date, "yyyymmdd") & " " & Format(Time, "HH:mm") & "')"
    ActArchivoSql = True
    
    '-------> Actualizar producto vigente
    ValidarProductoVigente
    
End If

Exit Function
Man_Error:
If Err = -2147217865 Or Err = 3265 Or Err = -2147467259 Then
    If Err = 3265 Or Err = -2147467259 Then
       Resume Next
    End If
    dbI.Close: Set dbI = Nothing
    '-------> Actualizar tabla lista producto y lista ingrediente
    '-------> Borrar tabla paso receta -
    vg_db.Execute "DELETE paso_receta WHERE rec_spid = " & spid & " AND rec_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_recetadet WHERE red_spid = " & spid & " AND red_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_productosing WHERE pri_spid = " & spid & " AND pri_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_regimen WHERE reg_spid = " & spid & " AND reg_usr = '" & vg_NUsr & "'"
    vg_db.Execute "DELETE paso_productospmpdia WHERE ppd_spid = " & spid & " AND ppd_usuario = '" & vg_NUsr & "'"
    
    vg_db.Execute "INSERT INTO b_contlistpreing (cpi_cencos, cpi_coding, cpi_precos, cpi_feccos, cpi_codcom, cpi_codped) SELECT '" & MuestraCasino(1) & "', ing_codigo, 0, 0, ing_codcom, ing_codped FROM b_ingrediente WHERE ing_codigo NOT IN (SELECT DISTINCT cpi_coding FROM b_contlistpreing WHERE cpi_cencos='" & MuestraCasino(1) & "')"
    
    fso.DeleteFile cdbi
    vg_db.Execute "INSERT INTO log_actualizacion VALUES ('" & Mid(cdbz, 1, Len(cdbz) - 3) & "mdb" & "', '" & Format(Date, "yyyy/mm/dd") & " " & Format(Time, "HH:mm") & "')"
    ActArchivoSql = True
    '-------> Actualizar producto vigente
    ValidarProductoVigente
   Exit Function
End If
If Err = 53 Then ActArchivoSql = True: Exit Function
If Err = -2147168242 Then ActArchivoSql = True: Exit Function
If Err = 3034 Then
   Exit Function
End If

If Err.Number = -2147467259 Then
    MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, MsgTitulo
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

If RS2.EOF Then
    
    bAdd = True
    RS2.AddNew

End If

For i = 0 To RSO.count - 1
'    If bAdd Or (RS2.Fields(i).Name <> "pro_upreco" And RS2.Fields(i).Name <> "pro_fecven" And RS2.Fields(i).Name <> "pro_fecuco" And RS2.Fields(i).Name <> "pro_maepro" And RS2.Fields(i).Name <> "pro_propon" And RS2.Fields(i).Name <> "ing_precos" And RS2.Fields(i).Name <> "ing_feccos" And RS2.Fields(i).Name <> "rec_tiprec" And RS2.Fields(i).Name <> "ing_codcom" And RS2.Fields(i).Name <> "ing_codped" And RS2.Fields(i).Name <> "ess_codsec") Then
    
    If bAdd Or (RS2.Fields(i).Name <> "pro_upreco" And RS2.Fields(i).Name <> "pro_fecven" And RS2.Fields(i).Name <> "pro_fecuco" And RS2.Fields(i).Name <> "pro_propon" And RS2.Fields(i).Name <> "ing_precos" And RS2.Fields(i).Name <> "ing_feccos" And RS2.Fields(i).Name <> "rec_tiprec" And RS2.Fields(i).Name <> "ing_codcom" And RS2.Fields(i).Name <> "ing_codped" And RS2.Fields(i).Name <> "ess_codsec") Then
        
        Select Case RS2.Fields(i).Type
        
        Case adChar, adVarChar, adVarWChar


            If TipoDato(RS2.Fields(i).Value, "") <> Trim(TipoDato(RSO.Item(i).Value, "")) Then
               
               If Trim(RSO.Item(i).Value) = "" Then
                  
                  RS2.Fields(i).Value = " "
               
               Else
                  
                  RS2.Fields(i).Value = Trim(RSO.Item(i).Value)
               
               End If
                  
                  'IIf(Trim((RSO.Item(i).Value)) = "", " ", Trim(RSO.Item(i).Value))
             
             End If
        
        Case Else
            
            If TipoDato(RS2.Fields(i).Value, 0) <> TipoDato(RSO.Item(i).Value, 0) Or RS2.Fields(i).Name = "red_tiprec" Or RS2.Fields(i).Name = "mid_tiprec" Or RS2.Fields(i).Name = "mid_racteo" Or RS2.Fields(i).Name = "mid_racrea" Or RS2.Fields(i).Name = "pro_fecven" Or RS2.Fields(i).Name = "tip_previo" Or RS2.Fields(i).Name = "car_previo" Or RS2.Fields(i).Name = "pro_ctrsto" Or RS2.Fields(i).Name = "pro_maepro" Then
               
               RS2.Fields(i).Value = RSO.Item(i).Value
            
            ElseIf RS2.Fields(i).Name = "mid_cosrec" Then
               
               RS2.Fields(i).Value = fg_CalCtoRecInv(RSO.Item(i - 3).Value, RSO.Item(i + 2).Value, (fg_CambiaChar(GetParametro("ctainsumo"), ";", "','")))
            
            ElseIf RS2.Fields(i).Name = "mid_cosdes" Then
               
               RS2.Fields(i).Value = fg_CalCtoRecInv(RSO.Item(i - 8).Value, RSO.Item(i - 3).Value, (fg_CambiaChar(GetParametro("ctalimdes"), ";", "','")))
            
            ElseIf RS2.Fields(i).Name = "mid_rec5eta" Then
               
               RS2.Fields(i).Value = RSO.Item(i).Value
            
            End If
        
        End Select
    
    ElseIf RS2.Fields(i).Name = "rec_tiprec" And RSO.Item(i).Value > 0 Then
        
        RS2.Fields(i).Value = RSO.Item(i).Value
    
    ElseIf RS2.Fields(i).Name = "ess_codsec" And RSO.Item(i).Value > 0 Then
        
        RS2.Fields(i).Value = RSO.Item(i).Value
    
    ElseIf RS2.Fields(i).Name = "pro_fecven" And RSO.Item(i).Value > 0 Then
        
        RS2.Fields(i).Value = RSO.Item(i).Value
    
    End If

Next

RS2.Update
RS2.Close: Set RS2 = Nothing

Exit Function
ManError:
If Err.Number = -2147217887 Then Resume Next
MsgBox "Error " & Err.Number & ": " & Trim(Err.Description), vbCritical + vbOKOnly, MsgTitulo

End Function


Function validarsiexistetabla(cdbi As String, nombretabla As String) As Boolean

validarsiexistetabla = False
Dim RS As New ADODB.Recordset
Set dbI = New ADODB.Connection
dbI.ConnectionString = "Provider='" & LTrim(RTrim(Provider)) & "';Data Source= '" & cdbi & "' ;Persist Security Info=False"
dbI.ConnectionTimeout = 3600
dbI.CommandTimeout = 3600
dbI.Open

Set RS = dbI.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
Do While Not RS.EOF
   
   If RS!table_name = nombretabla Then
      
      validarsiexistetabla = True
      Exit Do
   
   End If
   RS.MoveNext

Loop
RS.Close: Set RS = Nothing

End Function
