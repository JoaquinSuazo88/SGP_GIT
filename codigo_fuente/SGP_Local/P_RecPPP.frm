VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form P_RecPPP 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de Recalculo Precio Prom. Ponderado"
   ClientHeight    =   2730
   ClientLeft      =   3750
   ClientTop       =   2085
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
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
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   30
      TabIndex        =   0
      Top             =   450
      Width           =   7845
      Begin VB.Frame Frame2 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   90
         TabIndex        =   4
         Top             =   330
         Width           =   7665
         Begin VB.OptionButton Option1 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   1
            Left            =   3285
            TabIndex        =   6
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Uno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   300
            Width           =   1395
         End
         Begin EditLib.fpText fpText1 
            Height          =   315
            Index           =   0
            Left            =   1005
            TabIndex        =   7
            Top             =   690
            Width           =   1395
            _Version        =   196608
            _ExtentX        =   2461
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
         Begin VB.Label fpayuda 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Index           =   0
            Left            =   2895
            TabIndex        =   9
            Top             =   690
            Width           =   4545
         End
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   480
            Index           =   0
            Left            =   2370
            Picture         =   "P_RecPPP.frx":0000
            Top             =   585
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Producto"
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
            Left            =   165
            TabIndex        =   8
            Top             =   750
            Width           =   780
         End
         Begin VB.Label label 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   2940
            TabIndex        =   10
            Top             =   720
            Width           =   4530
         End
      End
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   165
         Index           =   0
         Left            =   1620
         TabIndex        =   2
         Top             =   1950
         Visible         =   0   'False
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   291
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Un Momento, Procesando Información"
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
         Left            =   1875
         TabIndex        =   3
         Top             =   1650
         Visible         =   0   'False
         Width           =   3285
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "P_RecPPP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim MsgTitulo As String

Private Sub Form_Load()

On Error GoTo Man_Error

Me.Width = 7995
Me.Height = 3165
MsgTitulo = "Proceso de Recalculo Precio Prom. Ponderado"
fg_centra Me
Me.HelpContextID = vg_OpcM
Toolbar1.ImageList = Partida.IL1

Set BtnX = Toolbar1.Buttons.Add(, "A_Confirmar ", , tbrDefault, "A_Confirmar "): BtnX.Visible = IIf(Mid(ValidarUsuario(Me), 2, 1) = "1", True, False): BtnX.ToolTipText = "Confirmar "
Set BtnX = Toolbar1.Buttons.Add(, "I_Confirmar ", , tbrDefault, "I_Confirmar "): BtnX.Visible = IIf(Mid(ValidarUsuario(Me), 2, 1) = "0", True, False): BtnX.ToolTipText = ""
Set BtnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): BtnX.Enabled = False
Set BtnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): BtnX.Visible = True: BtnX.ToolTipText = "Salir"

'------- Traer periodo abierto

If RS1.State = 1 Then RS1.Close
RS1.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT cie_periodo, cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos='" & MuestraCasino(1) & "' AND cie_estado=1", vg_db, adOpenStatic
If Not RS1.EOF Then Frame1.Caption = "Periodo " & Mid(RS1!cie_periodo, 5, 2) & "/" & Mid(RS1!cie_periodo, 1, 4)
RS1.Close
Set RS1 = Nothing
'------- Fin traer periodo abierto

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

If KeyAscii <> 13 Then Exit Sub
SendKeys "{Tab}"

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub fpText1_LostFocus(Index As Integer)

On Error GoTo Man_Error

If Trim(fpText1(0).text) = "" Then fpayuda(0).Caption = "": Exit Sub

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

RS1.Open "SELECT pro_nombre FROM b_productos WHERE pro_codigo='" & fpText1(0).text & "'", vg_db, adOpenStatic
If Not RS1.EOF Then
   
   Do While Not RS1.EOF
      
      fpayuda(0).Caption = RS1!pro_nombre
      RS1.MoveNext
   
   Loop

Else
   
   fpText1(0).text = ""
   fpayuda(0).Caption = ""
   MsgBox "Producto no existe...", vbExclamation + vbOKOnly, MsgTitulo

End If
RS1.Close
Set RS1 = Nothing

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Image1_Click(Index As Integer)

On Error GoTo Man_Error

vg_codigo = ""
vg_nombre = ""
vg_left = fpayuda(0).Left + 4800
B_TabEst.LlenaDatos "b_productos", "pro_", "Productos", "Gen"
B_TabEst.Show 1
Me.Refresh
If Trim(vg_codigo) = "" Or Trim(vg_codigo) = "0" Then Exit Sub
fpText1(Index) = Trim(vg_codigo)
fpayuda(Index).Caption = vg_nombre
fpText1_LostFocus 0

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Option1_Click(Index As Integer)

On Error GoTo Man_Error

Select Case Index
    
    Case 0
        
        Image1(0).Enabled = True
        fpText1(0).Enabled = True
    
    Case 1
        
        Image1(0).Enabled = False
        fpText1(0).text = "": fpText1(0).Enabled = False
        fpayuda(0).Caption = ""

End Select

Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

On Error GoTo Man_Error

Select Case Button.Index
    
    Case 1
        
        If Option1(0).Value = True Then
           
           If RS1.State = 1 Then RS1.Close
           RS1.CursorLocation = adUseClient
           vg_db.CursorLocation = adUseClient

           RS1.Open "SELECT * FROM b_productos WHERE pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "'", vg_db, adOpenStatic
           If RS1.EOF Then RS.Close: Set RS = Nothing: fpText1(0).text = "": fpayuda(0).Caption = "": MsgBox "No existe producto", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
           RS1.Close
           Set RS1 = Nothing
        
        End If
        
        vg_db.BeginTrans
        Dim fecper As Long, fecini As Long, fecfin As Long
        fecper = 0
        fecini = 0
        fecfin = 0
        
        '------- Traer periodo abierto
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient
        
        RS1.Open "SELECT cie_periodo, cie_fecini, cie_fecter FROM b_cierreperiodo WHERE cie_cencos='" & MuestraCasino(1) & "' AND cie_estado=1", vg_db, adOpenStatic
        If Not RS1.EOF Then fecper = RS1!cie_periodo: fecper = fecper - 1: fecini = RS1!cie_fecini: fecfin = RS1!cie_fecter
        RS1.Close
        Set RS1 = Nothing
        '------- Fin traer periodo abierto
        
        Dim aAp As String
        aAp = Trim(vg_NUsr) & "_tmp_Recppp"
        '------- Creo tabla temporal y chequeo si existe antes
        fg_CheckTmp aAp
            
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        RS1.Open "SELECT DISTINCT tin_fectom as fecpro, tin_codpro as codpro, tin_stofis as cansto, tin_propon as propon, 'E+' as tipmov , 0 as numdoc, 'E' as tipdoc, 'E' as rutcli INTO " & aAp & " FROM b_tomainv " & _
                 "WHERE tin_ciemes=" & fecper & " AND tin_stofis>0 AND tin_propon>0 AND (tin_codpro='" & LimpiaDato(Trim(fpText1(0).text)) & "' OR '" & Trim(fpText1(0).text) & "'='') ORDER BY tin_fectom, tin_codpro", vg_db, adOpenStatic
        Set RS1 = Nothing
        If fecper = 0 Then MsgBox "No existe Información, para procesar....", vbExclamation + vbOKOnly, MsgTitulo: Exit Sub
        Label1(0).Visible = True
        Label1(0).Caption = "Un Momento, Procesando Información": DoEvents
        Bar1(0).Min = 0
        Bar1(0).Value = 0
        Bar1(0).max = 8
        Bar1(0).Visible = True
        
        Bar1(0).Value = Bar1(0).Value + 2
        '------- Fin traer Inventario primer inventario
    
        '------- Traer ajuste inventario
        vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecemi, 'yyyymmdd') as fecpro, pro.pro_codigo as codpro, dev.dev_canmin as cansto, dev.dev_precos as propon, iif(aju.aju_tipo = 'A','E+','E-') as tipmov, tov.tov_numdoc as numdoc, tov.tov_tipdoc as tipdoc, tov.tov_rutcli as rutcli " & _
                      "FROM b_totventas tov, b_detventas dev, b_productos pro, a_tipoajuste aju " & _
                      "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
                      "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
                      "AND   tov.tov_numdoc=dev.dev_numdoc " & _
                      "AND   dev.dev_codmer=pro.pro_codigo " & _
                      "AND   tov.tov_codser=aju.aju_codigo " & _
                      "AND   format(tov.tov_fecemi, 'yyyymmdd')>=" & fecini & " " & _
                      "AND   format(tov.tov_fecemi, 'yyyymmdd')<=" & fecfin & " " & _
                      "AND   tov.tov_codbod=" & vg_codbod & " AND tov.tov_tipdoc='AI' AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' AND pro.pro_ctrsto=1 AND (pro.pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "' OR '" & Trim(fpText1(0).text) & "'='') ORDER BY tov.tov_fecemi, pro.pro_codigo"
        Bar1(0).Value = Bar1(0).Value + 1
        '------- Fin traer ajuste inventario
        
        '------- Traer salida y devolución produción
        vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecpro, 'yyyymmdd') as fecpro, pro.pro_codigo as codpro, dev.dev_canmer as cansto, dev.dev_precos as propon, 'S' as tipmov, tov.tov_numdoc as numdoc, tov.tov_tipdoc as tipdoc, tov.tov_rutcli as rutcli " & _
                      "FROM b_totventas tov, b_detventas dev, b_productos pro " & _
                      "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
                      "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
                      "AND   tov.tov_numdoc=dev.dev_numdoc " & _
                      "AND   dev.dev_codmer=pro.pro_codigo " & _
                      "AND   pro.pro_ctrsto=1 " & _
                      "AND  (tov.tov_tipdoc='SP' OR tov.tov_tipdoc='DP') " & _
                      "AND   dev.dev_canmer<>0 AND tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' AND tov.tov_codbod=" & vg_codbod & " " & _
                      "AND   tov.tov_fecpro>=cdate('" & fg_Ctod1(fecini) & "')" & _
                      "AND   tov.tov_fecpro<=cdate('" & fg_Ctod1(fecfin) & "')" & _
                      "AND   (pro.pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "' OR '" & Trim(fpText1(0).text) & "'='')"
        Bar1(0).Value = Bar1(0).Value + 1
        '------- Fin traer salida y devolución produción
        
        '------- Traer mermas
        vg_db.Execute "INSERT INTO " & aAp & " SELECT format(tov.tov_fecemi, 'yyyymmdd') AS fecpro, pro.pro_codigo AS codpro, dev.dev_canmer AS cansto, dev.dev_precos AS propon, 'S' AS tipmov, tov.tov_numdoc AS numdoc, tov.tov_tipdoc AS tipdoc, tov.tov_rutcli AS rutcli " & _
                      "FROM  b_totventas tov, b_detventas dev, b_productos pro " & _
                      "WHERE tov.tov_rutcli=dev.dev_rutcli " & _
                      "AND   tov.tov_tipdoc=dev.dev_tipdoc " & _
                      "AND   tov.tov_numdoc=dev.dev_numdoc " & _
                      "AND   dev.dev_codmer=pro.pro_codigo " & _
                      "AND   tov.tov_codbod=" & vg_codbod & " AND tov.tov_tipdoc='ME' and tov.tov_estdoc<>'A' AND tov.tov_estdoc<>'P' " & _
                      "AND   format(tov.tov_fecemi,'yyyymmdd')>=" & fecini & "" & _
                      "AND   format(tov.tov_fecemi,'yyyymmdd')>=" & fecfin & "" & _
                      "AND   (pro.pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "' OR '" & Trim(fpText1(0).text) & "'='')"
        Bar1(0).Value = Bar1(0).Value + 1
        '------- Fin traer mermas
        
        '------- Traer documento traspaso entrada
        vg_db.Execute "insert into " & aAp & " select format(tov.tov_fecemi, 'yyyymmdd') as fecpro, pro.pro_codigo as codpro, dev.dev_canmer as cansto, dev.dev_precos as propon, iif(tov.tov_codreg=0,'E+','S') as tipmov, tov.tov_numdoc as numdoc, tov.tov_tipdoc as tipdoc, tov.tov_rutcli as rutcli " & _
                      "from   b_totventas tov, b_detventas dev, b_productos pro " & _
                      "where  tov.tov_rutcli=dev.dev_rutcli " & _
                      "and    tov.tov_tipdoc=dev.dev_tipdoc " & _
                      "and    tov.tov_numdoc=dev.dev_numdoc " & _
                      "and    dev.dev_codmer=pro.pro_codigo " & _
                      "and    pro.pro_ctrsto=1 " & _
                      "and    tov.tov_codbod=" & vg_codbod & " AND tov.tov_tipdoc='TR' and tov.tov_estdoc<>'A' and tov.tov_estdoc<>'P' " & _
                      "and    format(tov.tov_fecemi,'yyyymmdd')>=" & fecini & " and format(tov.tov_fecemi,'yyyymmdd')<=" & fecfin & " and dev.dev_canmer>0 and (pro.pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "' or '" & Trim(fpText1(0).text) & "'='') order by tov.tov_fecemi, pro.pro_codigo"
        Bar1(0).Value = Bar1(0).Value + 1
    '                  "and    tov.tov_tipdoc='TR' and tov.tov_estdoc<>'A' and tov.tov_estdoc<>'P' and tov.tov_codreg=0 " & _
        '------- Fin traer documento traspaso entrada
        
        
        '------- Traer Documento Proveedor
        Dim pctimp As Double, pctdes  As Double, precio As Double
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

'        RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, de.dec_numdoc, " & _
'                 "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
'                 "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
'                 "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
'                 "WHERE toc.toc_rutpro=de.dec_rutpro " & _
'                 "AND   toc.toc_tipdoc=de.dec_tipdoc " & _
'                 "AND   toc.toc_numdoc=de.dec_numdoc " & _
'                 "AND   de.dec_codmer=pro.pro_codigo " & _
'                 "AND   de.dec_mueinv='S' and toc.toc_tipdoc<>'SN' " & _
'                 "AND   de.dec_canrec>0 " & _
'                 "AND  format(toc.toc_fecrem, 'yyyymmdd')>='" & fecini & "' " & _
'                 "AND  format(toc.toc_fecrem, 'yyyymmdd')<='" & fecfin & "' " & _
'                 "AND   toc.toc_codbod=" & vg_codbod & " " & _
'                 "AND  (pro.pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "' OR '" & Trim(fpText1(0).text) & "'='') " & _
'                 "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic
        
        RS1.Open "SELECT toc.toc_fecrem, pro.pro_codigo, de.dec_numdoc, " & _
                 "de.dec_codmer, de.dec_canmer, de.dec_precom, de.dec_pctdes, de.dec_numlin, " & _
                 "de.dec_canrec, de.dec_prerec, de.dec_prefle, de.dec_rutpro, de.dec_tipdoc " & _
                 "FROM b_totcompras toc, b_detcompras de, b_productos pro " & _
                 "WHERE toc.toc_rutpro=de.dec_rutpro " & _
                 "AND   toc.toc_tipdoc=de.dec_tipdoc " & _
                 "AND   toc.toc_numdoc=de.dec_numdoc " & _
                 "AND   de.dec_codmer=pro.pro_codigo " & _
                 "AND   de.dec_mueinv='S' and toc.toc_tipdoc in (select tdo_codigo from a_tipodocumento where tdo_IdCodigo <> 'SN') " & _
                 "AND   de.dec_canrec>0 " & _
                 "AND  format(toc.toc_fecrem, 'yyyymmdd')>='" & fecini & "' " & _
                 "AND  format(toc.toc_fecrem, 'yyyymmdd')<='" & fecfin & "' " & _
                 "AND   toc.toc_codbod=" & vg_codbod & " " & _
                 "AND  (pro.pro_codigo='" & LimpiaDato(Trim(fpText1(0).text)) & "' OR '" & Trim(fpText1(0).text) & "'='') " & _
                 "ORDER BY toc.toc_fecrem, pro.pro_nombre", vg_db, adOpenStatic
        
        If Not RS1.EOF Then
           
           Do While Not RS1.EOF
              
              pctimp = 0
              precio = 0
              
              If RS2.State = 1 Then RS2.Close
              RS2.CursorLocation = adUseClient
              vg_db.CursorLocation = adUseClient
              
              RS2.Open "select b_detcomprasimp.imd_monimp, a_impuesto.imp_pctimp, a_impuesto.imp_inccos " & _
                       "from  b_detcomprasimp, a_impuesto " & _
                       "where b_detcomprasimp.imd_rutdoc='" & RS1!dec_rutpro & "' " & _
                       "and   b_detcomprasimp.imd_tipdoc='" & RS1!dec_tipdoc & "' " & _
                       "and   b_detcomprasimp.imd_numdoc=" & RS1!dec_numdoc & " " & _
                       "and   b_detcomprasimp.imd_numlin=" & RS1!dec_numlin & " " & _
                       "and   b_detcomprasimp.imd_codpro='" & RS1!pro_codigo & "' " & _
                       "and   b_detcomprasimp.imd_codimp=a_impuesto.imp_codigo " & _
                       "and   a_impuesto.imp_inccos=1", vg_db, adOpenStatic
              
              If RS2.EOF Then RS2.Close: Set RS2 = Nothing Else pctimp = RS2!imd_monimp: RS2.Close: Set RS2 = Nothing
              
              pctdes = 0
              tanali = 0
              If RS1!dec_pctdes > 0 Then pctdes = RS1!dec_pctdes
              
              If RS1!dec_prefle > 0 Then
                 
                 precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec)) + (RS1!dec_prefle / RS1!dec_canrec)
              
              Else
                 
                 precio = ((RS1!dec_prerec) - (RS1!dec_prerec * (pctdes / 100)) + (pctimp / RS1!dec_canrec))
              
              End If
              vg_db.Execute "insert into " & aAp & " values (" & Val(Format(RS1!toc_fecrem, "yyyymmdd")) & ", '" & Trim(RS1!pro_codigo) & "', " & RS1!dec_canrec & ", " & precio & ", '" & "E+" & "', " & RS1!dec_numdoc & ", '" & Trim(RS1!dec_tipdoc) & "', '" & Trim(RS1!dec_rutpro) & "')"
              RS1.MoveNext
              
           Loop
           
        End If
        Bar1(0).Value = Bar1(0).Value + 1
        RS1.Close
        Set RS1 = Nothing
        '------- Fin traer Documento Proveedor
       
        '------- Procesar Información Precio Promedio Ponderado
        Dim canbod As Double, auxpmp As Double, pmp As Long
        Dim codpro As String
        
        If RS1.State = 1 Then RS1.Close
        RS1.CursorLocation = adUseClient
        vg_db.CursorLocation = adUseClient

        RS1.Open "select * from " & aAp & " where (codpro='" & LimpiaDato(Trim(fpText1(0).text)) & "' or '" & Trim(fpText1(0).text) & "'='') order by codpro, fecpro, tipmov", vg_db, adOpenStatic
        If Not RS1.EOF Then
           
           canbod = 0
           auxpmp = 0
           pmp = 0
           
           Do While Not RS1.EOF
              
              '------- Actualizar documentos salidas produccción y Mermmas
              If RS1!codpro <> codpro Then
    '*             If Trim(CodPro) <> "" Then
    '*                vg_db.Execute "UPDATE b_contlistprepro SET cpp_propon=" & PMP & " WHERE cpp_cencos='" & MuestraCasino(1) & "' AND cpp_codpro='" & CodPro & "'"
    '*                vg_db.Execute "UPDATE (b_contlistpreing INNER JOIN (b_productos INNER JOIN b_productosing ON b_productos.pro_codigo=b_productosing.pri_codpro) ON b_contlistpreing.cpi_coding=b_productosing.pri_coding) INNER JOIN b_contlistprepro ON b_productos.pro_codigo=b_contlistprepro.cpp_codpro SET b_contlistpreing.cpi_precos=" & Format(Date, "yyyymmdd") & ", b_contlistpreing.cpi_feccos=(" & PMP & "/b_productos.pro_facing) " & _
    '*                              "WHERE b_productosing.pri_codpro='" & CodPro & "' AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "' AND b_contlistprepro.cpp_cencos='" & MuestraCasino(1) & "'"
    '*             End If
                 codpro = RS1!codpro
                 canbod = 0
                 auxpmp = 0
                 pmp = 0
                 
              End If
              
              If RS1!tipmov = "S" And codpro = RS1!codpro Then
                 
                 If pmp > 0 Then
                    
                    '------- Actualizar encabezado y detalle ventas
                    vg_db.Execute "UPDATE b_totventas INNER JOIN b_detventas ON (b_totventas.tov_numdoc=b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc=b_detventas.dev_tipdoc) and (b_totventas.tov_rutcli=b_detventas.dev_rutcli) SET b_detventas.dev_precos= " & pmp & ", b_detventas.dev_predoc=" & pmp & ", b_detventas.dev_ptotal=(" & pmp & " * " & RS1!cansto & ") " & _
                                  "WHERE b_totventas.tov_numdoc=" & RS1!numdoc & " AND b_totventas.tov_tipdoc='" & RS1!tipdoc & "' AND b_totventas.tov_rutcli='" & RS1!rutcli & "' AND b_detventas.dev_codmer='" & RS1!codpro & "' AND b_totventas.tov_estdoc<>'A' AND b_totventas.tov_estdoc<>'P' AND b_totventas.tov_codbod=" & vg_codbod & ""
                 
                    If RS2.State = 1 Then RS2.Close
                    RS2.CursorLocation = adUseClient
                    vg_db.CursorLocation = adUseClient
                    
                    RS2.Open "SELECT SUM(dev_ptotal) AS ptotal FROM b_detventas WHERE dev_rutcli='" & RS1!rutcli & "' AND dev_tipdoc='" & RS1!tipdoc & "' AND dev_numdoc=" & RS1!numdoc & " GROUP BY dev_rutcli", vg_db, adOpenStatic
                    
                    If Not RS2.EOF And RS2!ptotal > 0 And Not IsNull(RS2!ptotal) Then
                       
                       vg_db.Execute "UPDATE b_totventas SET b_totventas.tov_totdoc=" & RS2!ptotal & " " & _
                                     "WHERE b_totventas.tov_estdoc<>'A' AND b_totventas.tov_estdoc<>'P' AND b_totventas.tov_rutcli='" & RS1!rutcli & "' AND b_totventas.tov_tipdoc='" & RS1!tipdoc & "' AND b_totventas.tov_numdoc=" & RS1!numdoc & " b_totventas.tov_codbod=" & vg_codbod & ""
                    
                    End If
                    RS2.Close
                    Set RS2 = Nothing
                    '------- Fin actualizar encabezado y detalle ventas
                 
                 Else
                    
                    pmp = RS1!propon
                 
                 End If
                 
                 pmp = ((auxpmp * IIf(canbod < 0, (canbod * -1), canbod)) + (pmp * RS1!cansto)) / (IIf(canbod < 0, (canbod * -1), canbod) + RS1!cansto)
                 auxpmp = pmp: canbod = canbod - RS1!cansto
              
              Else
                 
                 pmp = ((auxpmp * IIf(canbod < 0, (canbod * -1), canbod)) + (RS1!propon * RS1!cansto)) / (IIf(canbod < 0, (canbod * -1), canbod) + RS1!cansto)
                 auxpmp = pmp: canbod = IIf(Trim(RS1!tipmov) = "E+", (canbod + RS1!cansto), (canbod - RS1!cansto))
              
              End If
              '------- Fin actualizar documentos salidas produccción y Mermmas
              RS1.MoveNext
           
           Loop
    '*                vg_db.Execute "UPDATE b_contlistprepro SET cpp_propon=" & PMP & " WHERE cpp_cencos='" & MuestraCasino(1) & "' AND cpp_codpro='" & CodPro & "'"
    '*                vg_db.Execute "UPDATE (b_contlistpreing INNER JOIN (b_productos INNER JOIN b_productosing ON b_productos.pro_codigo=b_productosing.pri_codpro) ON b_contlistpreing.cpi_coding=b_productosing.pri_coding) INNER JOIN b_contlistprepro ON b_productos.pro_codigo=b_contlistprepro.cpp_codpro SET b_contlistpreing.cpi_precos=" & Format(Date, "yyyymmdd") & ", b_contlistpreing.cpi_feccos=(" & PMP & "/b_productos.pro_facing) " & _
    '*                              "WHERE b_productosing.pri_codpro='" & CodPro & "' AND b_contlistpreing.cpi_cencos='" & MuestraCasino(1) & "' AND b_contlistprepro.cpp_cencos='" & MuestraCasino(1) & "'"
           Bar1(0).Value = Bar1(0).Value + 1
        End If
        RS1.Close
        Set RS1 = Nothing
        
        vg_db.CommitTrans
        
        Label1(0).Visible = False: Bar1(0).Visible = False
        MsgBox "Proceso de Actualización Finalizado", vbInformation + vbOKOnly, MsgTitulo
        '------- Fin procesar información precio promedio ponderado
    
    Case 4
        
        Me.Hide
        Unload Me
        
End Select

Exit Sub
Man_Error:
If Err = 3034 Then vg_db.RollbackTrans: Exit Sub
vg_db.RollbackTrans
fg_descarga
MsgBox Err & ":  " & error$(Err), vbCritical, "Boton_Click"
ins_log_error Date & Time & Err & ":  " & error$(Err)

End Sub
