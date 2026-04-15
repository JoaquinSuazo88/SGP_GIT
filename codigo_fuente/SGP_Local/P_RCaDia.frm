VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form P_RCaDia 
   Caption         =   "Proceso Reecalcular Día"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2535
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Procesar"
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
      Left            =   3960
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   1
         Left            =   2310
         TabIndex        =   1
         Top             =   960
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2990
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
         MaxLength       =   255
         MultiLine       =   0   'False
         PasswordChar    =   "*"
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
      Begin EditLib.fpText Nombre 
         Height          =   315
         Index           =   0
         Left            =   2310
         TabIndex        =   0
         Top             =   600
         Width           =   1695
         _Version        =   196608
         _ExtentX        =   2990
         _ExtentY        =   556
         Enabled         =   -1  'True
         MousePointer    =   0
         Object.TabStop         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         AlignTextV      =   0
         AllowNull       =   0   'False
         NoSpecialKeys   =   0
         AutoAdvance     =   0   'False
         AutoBeep        =   0   'False
         AutoCase        =   1
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
         MaxLength       =   20
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
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   2400
         TabIndex        =   7
         Top             =   600
         Width           =   1395
         _Version        =   196608
         _ExtentX        =   2461
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
         ButtonStyle     =   3
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   -1  'True
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   0
         AlignTextV      =   0
         AllowNull       =   -1  'True
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
         Text            =   ""
         DateCalcMethod  =   0
         DateTimeFormat  =   5
         UserDefinedFormat=   "dd/mm/yyyy"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Password"
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
         Left            =   1320
         TabIndex        =   4
         Top             =   1020
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Login"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   645
         Width           =   585
      End
   End
End
Attribute VB_Name = "P_RCaDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Form_Load()
fg_centra Me
Msgtitulo = "Proceso Recalculo Día"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Command1(0).Caption = "&Aceptar"
Command1(1).Caption = "&Salir"
fpDateTime1(0).Visible = False
End Sub

Private Sub Command1_Click(Index As Integer)
Dim RS As New ADODB.Recordset
Dim fecha As Date
Select Case Index
Case 0 '-------> Procesar
    If Command1(0).Caption = "&Aceptar" Then
       '-------> Validar usuario
       RS.Open "SELECT * FROM a_param WHERE par_valor = '" & LimpiaDato(Trim(Nombre(0).text)) & "' AND par_codigo = 'usulimbas' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
       If RS.EOF Then MsgBox "Usuario no existe...": RS.Close: Set RS = Nothing: Nombre(0).text = "": Nombre(0).SetFocus: Exit Sub
       RS.Close: Set RS = Nothing
       RS.Open "SELECT * FROM a_param WHERE par_codigo = 'paslimbas' AND par_cencos = '" & MuestraCasino(1) & "'", vg_db, adOpenStatic
       If Not RS.EOF And UCase(Nombre(1).text) <> UCase(fg_Desencripta(TipoDato(RS!par_valor, ""))) Then MsgBox "La clave no corresponde al usuario...": RS.Close: Set RS = Nothing: Nombre(0).text = "": Nombre(0).SetFocus: Exit Sub
       RS.Close: Set RS = Nothing
       Command1(0).Caption = "&Procesar"
       Command1(1).Caption = "&Cancelar"
       Label1(0).Caption = "Fecha Proceso"
       Label1(0).Left = 1020
       Nombre(0).Visible = False
       Nombre(1).Visible = False
       Label1(1).Visible = False
       fpDateTime1(0).Visible = True
    Else
       If fpDateTime1(0).text = "" Then MsgBox "Fecha no puede ser vacia", vbExclamation + vbOKOnly, "Categoría Dietética": Exit Sub
'       Label1(1).Caption = "Procesando : "
'       Label1(0).Caption = ""
        fecha = fpDateTime1(0).text 'CDate(fg_Ctod1(fecinv)) + 1
        If CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1 < fecha Then
           vg_ciedia = fecha - 1
            V_Acceso.Label1(1).Visible = True
            V_Acceso.Label1(1).Caption = "Procesando : "
            V_Acceso.Label1(0).Caption = Mid(fecha - 1, 1, 2) & "/" & Mid(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1, 1, 2)
            V_Acceso.Combo1(0).Clear
            V_Acceso.Combo1(0).AddItem "Proc. Día : " & fecha - 1
            V_Acceso.Combo1(0).ListIndex = 0
            If vg_tipbase = "1" Then
               CalcularPMPDiaAccess V_Acceso, False, True
            Else
               CalcularPMPDiaSql V_Acceso, False, True
            End If
        End If
        Do While fecha <= CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1
           V_Acceso.Label1(1).Visible = True
           V_Acceso.Label1(1).Caption = "Procesando : "
           V_Acceso.Label1(0).Caption = Mid(fecha, 1, 2) & "/" & Mid(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))) - 1, 1, 2)
           V_Acceso.Combo1(0).Clear
           V_Acceso.Combo1(0).AddItem "Proc. Día : " & fecha
           V_Acceso.Combo1(0).ListIndex = 0
           vg_ciedia = fecha
           If vg_tipbase = "1" Then
              CalcularPMPDiaAccess V_Acceso, False, True
           Else
              CalcularPMPDiaSql V_Acceso, False, True
           End If
           '-------> Actualizar b_productospmpdia
           If vg_tipbase = "1" Then
              vg_db.Execute "UPDATE b_productospmpdia INNER JOIN b_tomainv ON b_productospmpdia.ppd_codpro = b_tomainv.tin_codpro SET b_productospmpdia.ppd_saldo = b_tomainv.tin_stofis " & _
                            "WHERE b_productospmpdia.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND b_productospmpdia.ppd_cencos = '" & vg_contra & "' AND b_tomainv.tin_codbod = " & vg_codbod & " AND b_tomainv.tin_fectom = " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND b_tomainv.tin_ciemes = 0"
           Else
              vg_db.Execute "UPDATE b_productospmpdia SET b_productospmpdia.ppd_saldo = b.tin_stofis FROM b_productospmpdia a, b_tomainv b WHERE a.ppd_codpro = b.tin_codpro " & _
                            "AND a.ppd_fecdia = " & Format(CDate(vg_ciedia), "yyyymmdd") & " AND a.ppd_cencos = '" & vg_contra & "' AND b.tin_codbod = " & vg_codbod & " AND b.tin_fectom = " & Format(CDate(vg_ciedia), "yyyymmdd") & "" 'AND b.tin_ciemes = 0"
           End If
           If vg_tipbase = "1" Then
              '-------> Actualizar precio toma inventario
              vg_db.Execute "UPDATE b_tomainv INNER JOIN b_productospmpdia ON (b_tomainv.tin_fectom = b_productospmpdia.ppd_fecdia) AND (b_tomainv.tin_codpro = b_productospmpdia.ppd_codpro) SET b_tomainv.tin_propon = b_productospmpdia.ppd_propon " & _
                            "WHERE b_tomainv.tin_ciemes = 0 AND Mid(b_tomainv.tin_fectom, 1, 6) = " & Format(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))), "yyyymm") & " And b_tomainv.tin_codbod = " & vg_codbod & " AND b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_productospmpdia.ppd_fecdia = " & Format(CDate(fecha), "yyyymmdd") & ""
              '-------> Actualizar precio en ajuste inventario
              vg_db.Execute "UPDATE (b_totventas INNER JOIN b_detventas ON (b_totventas.tov_numdoc = b_detventas.dev_numdoc) AND (b_totventas.tov_tipdoc = b_detventas.dev_tipdoc) AND (b_totventas.tov_rutcli = b_detventas.dev_rutcli)) INNER JOIN b_productospmpdia ON (b_detventas.dev_codmer = b_productospmpdia.ppd_codpro) AND (b_totventas.tov_rutcli = b_productospmpdia.ppd_cencos) SET b_detventas.dev_precos = b_productospmpdia.ppd_propon, b_detventas.dev_predoc = b_productospmpdia.ppd_propon, b_detventas.dev_ptotal = (b_detventas.dev_canmer*b_productospmpdia.ppd_propon) " & _
                            "WHERE b_productospmpdia.ppd_cencos = '" & MuestraCasino(1) & "' AND b_productospmpdia.ppd_fecdia = " & Format(CDate(fecha), "yyyymmdd") & " AND b_totventas.tov_codbod= " & vg_codbod & "  AND b_totventas.tov_fecemi = CDate('" & fecha & "') AND b_totventas.tov_estdoc <> 'A' AND b_totventas.tov_tipdoc = 'AI'"
           Else
              '-------> Actualizar precio toma inventario
              vg_db.Execute "UPDATE b_tomainv SET b_tomainv.tin_propon = b.ppd_propon FROM b_tomainv a, b_productospmpdia b WHERE a.tin_fectom = b.ppd_fecdia AND a.tin_codpro = b.ppd_codpro  " & _
                            "AND a.tin_ciemes = 0 AND convert(int,substring(convert(varchar(8),a.tin_fectom), 1, 6)) = " & Format(CDate(fg_Desencripta(TipoDato((RS1!par_valor), ""))), "yyyymm") & " And a.tin_codbod = " & vg_codbod & " AND b.ppd_cencos = '" & MuestraCasino(1) & "' AND b.ppd_fecdia = " & Format(CDate(fecha), "yyyymmdd") & ""
              '-------> Actualizar precio en ajuste inventario
              vg_db.Execute "UPDATE b_detventas SET b_detventas.dev_precos = c.ppd_propon, b_detventas.dev_predoc = c.ppd_propon, b_detventas.dev_ptotal = (b.dev_canmer*c.ppd_propon) FROM b_totventas a,  b_detventas b, b_productospmpdia c WHERE a.tov_numdoc = b.dev_numdoc AND a.tov_tipdoc = b.dev_tipdoc AND a.tov_rutcli = b.dev_rutcli AND b.dev_codmer = c.ppd_codpro AND a.tov_rutcli = c.ppd_cencos  " & _
                            "AND c.ppd_cencos = '" & MuestraCasino(1) & "' AND c.ppd_fecdia = " & Format(CDate(fecha), "yyyymmdd") & " AND a.tov_codbod= " & vg_codbod & "  AND a.tov_fecemi = '" & Format(fecha, "yyyymmdd") & "' AND a.tov_estdoc <> 'A' AND a.tov_tipdoc = 'AI'"
           End If
           If vg_tipbase = "1" Then
              RS.Open "SELECT a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi, Sum(b.dev_ptotal) AS dev_ptotal " & _
                      "FROM b_totventas a, b_detventas b " & _
                      "WHERE a.tov_numdoc = b.dev_numdoc " & _
                      "AND   a.tov_tipdoc = b.dev_tipdoc " & _
                      "AND   a.tov_rutcli =  b.dev_rutcli " & _
                      "AND   a.tov_fecemi = CDate('" & fecha & "') " & _
                      "AND   a.tov_tipdoc = 'AI' " & _
                      "AND   a.tov_estdoc <> 'A' " & _
                      "AND   a.tov_codbod = " & vg_codbod & " " & _
                      "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
                      "GROUP BY  a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi", vg_db, adOpenStatic
           Else
              RS.Open "SELECT a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi, Sum(b.dev_ptotal) AS dev_ptotal " & _
                      "FROM b_totventas a, b_detventas b " & _
                      "WHERE a.tov_numdoc = b.dev_numdoc " & _
                      "AND   a.tov_tipdoc = b.dev_tipdoc " & _
                      "AND   a.tov_rutcli =  b.dev_rutcli " & _
                      "AND   a.tov_fecemi = '" & Format(fecha, "yyyymmdd") & "' " & _
                      "AND   a.tov_tipdoc = 'AI' " & _
                      "AND   a.tov_estdoc <> 'A' " & _
                      "AND   a.tov_codbod = " & vg_codbod & " " & _
                      "AND   a.tov_rutcli = '" & MuestraCasino(1) & "' " & _
                      "GROUP BY  a.tov_numdoc, a.tov_tipdoc, a.tov_rutcli, a.tov_fecemi", vg_db, adOpenStatic
           End If
           If Not RS.EOF Then
              Do While Not RS.EOF
                 vg_db.Execute "UPDATE b_totventas SET tov_totdoc = " & RS!dev_ptotal & " WHERE tov_numdoc = " & RS!tov_numdoc & " AND tov_tipdoc = '" & RS!tov_tipdoc & "' AND tov_rutcli = '" & RS!tov_rutcli & "' AND tov_fecemi = " & RS!tov_fecemi & " AND tov_estdoc <> 'A' AND tov_codbod = " & vg_codbod & ""
                 RS.MoveNext
              Loop
           End If
           RS.Close: Set RS = Nothing
           fecha = fecha + 1
        Loop
    End If
Case 1 '-------> Salir
    Me.Hide
    Unload Me
End Select
End Sub

