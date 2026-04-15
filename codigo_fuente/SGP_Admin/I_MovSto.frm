VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form I_MovSto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Stock"
   ClientHeight    =   2625
   ClientLeft      =   1770
   ClientTop       =   3195
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
   ScaleHeight     =   2625
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2160
      Left            =   0
      TabIndex        =   1
      Top             =   360
      Width           =   9090
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         Caption         =   "Cuenta Contable"
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   150
         TabIndex        =   9
         Top             =   870
         Width           =   4350
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   2
            ItemData        =   "I_MovSto.frx":0000
            Left            =   135
            List            =   "I_MovSto.frx":0002
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   540
            Width           =   4035
         End
         Begin VB.OptionButton optCUENTA 
            Caption         =   "Una cuenta"
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   11
            Top             =   255
            Width           =   1425
         End
         Begin VB.OptionButton optCUENTA 
            Caption         =   "Todas"
            Height          =   225
            Index           =   1
            Left            =   3330
            TabIndex        =   10
            Top             =   255
            Width           =   855
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   690
         Left            =   150
         TabIndex        =   6
         Top             =   120
         Width           =   8775
         Begin VB.ComboBox Combo1 
            Height          =   315
            Index           =   0
            Left            =   1455
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   195
            Width           =   2550
         End
         Begin EditLib.fpDateTime Date1 
            Height          =   330
            Index           =   0
            Left            =   1455
            TabIndex        =   7
            Top             =   195
            Visible         =   0   'False
            Width           =   1470
            _Version        =   196608
            _ExtentX        =   2593
            _ExtentY        =   582
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
            ButtonStyle     =   3
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
            AutoAdvance     =   -1  'True
            AutoBeep        =   0   'False
            CaretInsert     =   0
            CaretOverWrite  =   3
            UserEntry       =   0
            HideSelection   =   -1  'True
            InvalidColor    =   -2147483637
            InvalidOption   =   0
            MarginLeft      =   3
            MarginTop       =   3
            MarginRight     =   3
            MarginBottom    =   3
            NullColor       =   -2147483637
            OnFocusAlignH   =   0
            OnFocusAlignV   =   0
            OnFocusNoSelect =   0   'False
            OnFocusPosition =   0
            ControlType     =   0
            Text            =   "04/09/2004"
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
            ThreeDFrameColor=   -2147483633
            Appearance      =   0
            BorderDropShadow=   1
            BorderDropShadowColor=   -2147483632
            BorderDropShadowWidth=   3
            PopUpType       =   0
            DateCalcY2KSplit=   60
            CaretPosition   =   0
            IncYear         =   1
            IncMonth        =   1
            IncDay          =   1
            IncHour         =   1
            IncMinute       =   1
            IncSecond       =   1
            ButtonColor     =   -2147483633
            AutoMenu        =   -1  'True
            StartMonth      =   4
            ButtonAlign     =   0
            BoundDataType   =   0
            OLEDropMode     =   0
            OLEDragMode     =   0
         End
         Begin VB.Label sombra 
            Appearance      =   0  'Flat
            BackColor       =   &H80000010&
            ForeColor       =   &H80000008&
            Height          =   315
            Index           =   0
            Left            =   1500
            TabIndex        =   15
            Top             =   240
            Width           =   2550
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Bodega"
            Height          =   195
            Index           =   0
            Left            =   540
            TabIndex        =   13
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Stock"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "Familia Producto"
         ForeColor       =   &H80000008&
         Height          =   990
         Left            =   4575
         TabIndex        =   2
         Top             =   870
         Width           =   4350
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Todos"
            Height          =   225
            Index           =   1
            Left            =   3330
            TabIndex        =   5
            Top             =   255
            Width           =   855
         End
         Begin VB.OptionButton optTIPPRO 
            Caption         =   "Un Tipo"
            Height          =   225
            Index           =   0
            Left            =   135
            TabIndex        =   4
            Top             =   255
            Width           =   1005
         End
         Begin VB.ComboBox Combo1 
            Appearance      =   0  'Flat
            Height          =   315
            Index           =   1
            ItemData        =   "I_MovSto.frx":0004
            Left            =   135
            List            =   "I_MovSto.frx":0006
            Style           =   2  'Dropdown List
            TabIndex        =   3
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
Attribute VB_Name = "I_MovSto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS1 As New ADODB.Recordset

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.Width = 9210
Me.Height = 3030
Msgtitulo = "Imprimir Toma de Inventario"
fg_centra Me
Toolbar1.ImageList = Partida.IL1
Set btnX = Toolbar1.Buttons.Add(, "A_Previa   ", , tbrDefault, "A_Previa   "): btnX.Visible = True: btnX.ToolTipText = "Vista Previa"
Set btnX = Toolbar1.Buttons.Add(, , "", tbrDefault, 0): btnX.Enabled = False
Set btnX = Toolbar1.Buttons.Add(, "A_Salir    ", , tbrDefault, "A_Salir    "): btnX.Visible = True: btnX.ToolTipText = "Salir"
Combo1(0).Clear
RS1.Open "select * from a_bodega order by bod_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo1(0).AddItem RS1!bod_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!bod_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
Combo1(0).ListIndex = 0
Combo1(1).Clear
RS1.Open "select * from a_tipopro order by tip_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo1(1).AddItem RS1!tip_nombre & Space(150) & "(" & fg_pone_cero(Str(RS1!tip_codigo), 10) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
Combo1(2).Clear
RS1.Open "select * from a_ctacontable order by cta_nombre", vg_db, adOpenStatic
Do While Not RS1.EOF
    Combo1(2).AddItem RS1!cta_nombre & Space(150) & "(" & Space(10 - Len(Trim(RS1!cta_codigo))) & Trim(RS1!cta_codigo) & ")"
    RS1.MoveNext
Loop
RS1.Close: Set RS1 = Nothing
optTIPPRO(1).Value = True
optCUENTA(1).Value = True
End Sub

Private Sub optTIPPRO_Click(Index As Integer)
    Combo1(1).Enabled = IIf(Index = 0, True, False)
    Combo1(1).ListIndex = IIf(Index = 0, 0, -1)
End Sub

Private Sub optCUENTA_Click(Index As Integer)
    Combo1(2).Enabled = IIf(Index = 0, True, False)
    Combo1(2).ListIndex = IIf(Index = 0, 0, -1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim sql1 As String, sql2 As String, v_fecini As String, v_fecter As String
'Dim sqlTP As String, sqlCU As String, aAp As String, aApPrin As String, v_codbod As Long
Select Case Button.Index
Case 1
'    v_codbod = fg_codigocbo(Combo1, 0, 10, 0)
'    sqlTP = ""
'    If optTIPPRO(0).Value = True Then sqlTP = "and pro.pro_codtip=" & Val(fg_codigocbo(Combo1, 1, 10, 0)) & " "
'    sqlCU = ""
'    If optCUENTA(0).Value = True Then sqlCU = "and pro.pro_ctacon='" & Trim(Mid(Trim(Combo1(2).List(Combo1(2).ListIndex)), Len(Trim(Combo1(2).List(Combo1(2).ListIndex))) - 10, 10)) & "' "
'    vg_db.BeginTrans
'    aApPrin = Trim(vg_NUsr) & "_tmp_Stock"
'    fg_CheckTmp aApPrin
'    vg_db.Execute "select pro.pro_ctacon, pro.pro_codtip, pro.pro_codigo, pro.pro_nombre, uni.uni_nomcor, bod.bod_canmer, pro.pro_propon " & _
'                  "into " & aApPrin & " from b_productos pro, a_unidad uni, b_bodegas as bod " & _
'                  "where pro.pro_codigo=bod.bod_codpro and uni.uni_codigo=pro.pro_coduni " & sqlCU & sqlTP & _
'                  "and bod.bod_canmer>0 and bod.bod_codbod=" & v_codbod
'    vg_db.Execute "update " & aApPrin & " a inner join b_bodegas b on a.pro_codigo=b.bod_codpro " & _
'                  "set a.bod_canmer=b.bod_canmer where b.bod_codbod=" & v_codbod
'    '--------------------------COMPRAS-------------------------
'    vg_db.Execute "update " & aApPrin & " as tmp, b_totcompras as toc, b_detcompras as [dec] " & _
'                  "set tmp.bod_canmer=tmp.bod_canmer+IIf(dec.dec_tipdoc='NC',dec.dec_canmer,-dec.dec_canmer) " & _
'                  "where toc.toc_rutpro=dec.dec_rutpro and toc.toc_tipdoc=dec.dec_tipdoc and toc.toc_numdoc=dec.dec_numdoc " & _
'                  "and dec.dec_codmer=tmp.pro_codigo and dec.dec_mueinv='S' and toc.toc_codbod=" & v_codbod & " " & _
'                  "and toc.toc_fecemi>cdate('" & Trim(Date1(0).Text) & "')"
'    '--------------------------VENTAS--------------------------
'    vg_db.Execute "update " & aApPrin & " tmp, b_totventas tov, b_detventas dev " & _
'                  "set tmp.bod_canmer=tmp.bod_canmer+IIf(dev.dev_tipdoc='SP' or dev.dev_tipdoc='ME' or dev.dev_tipdoc='FA' " & _
'                  "or dev.dev_tipdoc='GD' or (dev.dev_tipdoc='AI' and tov.tov_codreg=0) " & _
'                  "or (dev.dev_tipdoc='TR' and tov.tov_codser=0),dev.dev_canmer,-dev.dev_canmer) " & _
'                  "where tov.tov_rutcli=dev.dev_rutcli and tov.tov_tipdoc=dev.dev_tipdoc and tov.tov_numdoc=dev.dev_numdoc " & _
'                  "and dev.dev_codmer=tmp.pro_codigo and dev.dev_mueinv='S' and tov.tov_codbod=" & v_codbod & " " & _
'                  "and tov.tov_fecemi>cdate('" & Trim(Date1(0).Text) & "')"
'    '----------------------------------------------------------
'    vg_db.CommitTrans
    I_StockxFecha Me  'aApPrin
Case 3
    Me.Hide
    Unload Me
End Select
End Sub

