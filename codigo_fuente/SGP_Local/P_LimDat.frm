VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form P_LimDat 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Limpiar Base de Datos"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9540
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Left            =   6840
      TabIndex        =   6
      Top             =   3720
      Width           =   1215
   End
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
      Left            =   8160
      TabIndex        =   7
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   70
      TabIndex        =   3
      Top             =   120
      Width           =   9375
      Begin VB.Frame Frame2 
         Height          =   1695
         Left            =   180
         TabIndex        =   11
         Top             =   2040
         Width           =   9015
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
            Index           =   2
            Left            =   1320
            TabIndex        =   13
            Top             =   645
            Width           =   585
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
            Index           =   3
            Left            =   1320
            TabIndex        =   12
            Top             =   1020
            Width           =   930
         End
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   1800
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Width           =   6855
      End
      Begin EditLib.fpDateTime fpDateTime1 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   1200
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
      Begin MSComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Visible         =   0   'False
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Parametros"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo Limpieza (Hasta)"
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
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Base de Datos"
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
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   1455
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   8760
         Picture         =   "P_LimDat.frx":0000
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "P_LimDat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String
Dim RS As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
Dim sql1 As String
Select Case Index
Case 0 '-------> Proceso Limpiar base de dato
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
       Frame2.Visible = False
       fpDateTime1(0).Visible = True
    Else
        If Trim(Text1(0).text) = "" Then MsgBox "Debe base de datos...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
        If Trim(fpDateTime1(0).text) = "" Then MsgBox "Debe selecionar fecha...", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
'        If Dir(dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 3)) & "ldb") <> "" Then
'           MsgBox "El sistema necesita respaldar la base de dato que esta abierta." & Chr(13) & Chr(13) & "No se ejecutara hasta cerrar la Base o los programas relacionados", vbExclamation + vbOKOnly, Msgtitulo
'        End If
        Command1(0).Enabled = False
        PB.Min = 0: PB.Value = 0: PB.max = 20: Label2.Visible = True: PB.Visible = True
        '-------> Borrar tablas relacionada Minutas
        '-------> Borrar tabla b_minutacambios
        Label2.Caption = "Borrando Tabla b_minutacambios": DoEvents
        vg_db.Execute "DELETE b_minutacambios FROM b_minutacambios WHERE cam_fecmin IN (SELECT DISTINCT min_fecmin FROM b_minuta WHERE min_codigo = cam_codmin AND min_cencos = '" & MuestraCasino(1) & "' AND min_fecmin < " & Format(fpDateTime1(0).text, "yyyymmdd") & ")"
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_minutacosto
        Label2.Caption = "Borrando Tabla b_minutacosto": DoEvents
        vg_db.Execute "DELETE b_minutacosto FROM b_minutacosto WHERE mic_cencos='" & MuestraCasino(1) & "' AND mic_fecval IN (SELECT a.mid_fecval FROM b_minutadet a, b_minuta b WHERE a.mid_codigo=b.min_codigo AND b.min_cencos='" & MuestraCasino(1) & "' AND b.min_fecmin<" & Format(fpDateTime1(0).text, "yyyymmdd") & ")"
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_minutafijadia
        Label2.Caption = "Borrando Tabla b_minutafijadia": DoEvents
        vg_db.Execute "DELETE b_minutafijadia FROM b_minutafijadia WHERE mfd_cencos='" & MuestraCasino(1) & "' AND mfd_fecha<" & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_minutapedido
        Label2.Caption = "Borrando Tabla b_minutapedido": DoEvents
        vg_db.Execute "DELETE b_minutapedidos FROM b_minutapedidos WHERE ped_codcas='" & MuestraCasino(1) & "' AND ped_anomes<" & Format(fpDateTime1(0).text, "yyyymm") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_minutaraciones
        Label2.Caption = "Borrando Tabla b_minutaraciones": DoEvents
        vg_db.Execute "DELETE b_minutaraciones FROM b_minutaraciones WHERE mir_cencos='" & MuestraCasino(1) & "' AND mir_fecmin<" & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_minutadet
        Label2.Caption = "Borrando Tabla b_minutadet": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_minutadet.* FROM b_minuta INNER JOIN b_minutadet ON b_minuta.min_codigo=b_minutadet.mid_codigo " & _
                         "WHERE b_minuta.min_cencos='" & MuestraCasino(1) & "' AND b_minuta.min_fecmin<" & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        Else
           vg_db.Execute "DELETE b_minutadet FROM b_minuta, b_minutadet WHERE b_minuta.min_codigo = b_minutadet.mid_codigo " & _
                         "AND b_minuta.min_cencos = '" & MuestraCasino(1) & "' AND b_minuta.min_fecmin < " & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_minuta
        Label2.Caption = "Borrando Tabla b_minuta": DoEvents
        vg_db.Execute "DELETE b_minuta FROM b_minuta WHERE min_cencos = '" & MuestraCasino(1) & "' AND min_fecmin < " & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_detcomprasimp
        Label2.Caption = "Borrando Tabla b_detcomprasimp": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_detcomprasimp.* FROM b_detcomprasimp INNER JOIN b_totcompras ON (b_detcomprasimp.imd_numdoc=b_totcompras.toc_numdoc) AND (b_detcomprasimp.imd_tipdoc=b_totcompras.toc_tipdoc) AND (b_detcomprasimp.imd_rutdoc=b_totcompras.toc_rutpro) " & _
                         "WHERE b_totcompras.toc_fecrem<CDate('" & fpDateTime1(0).text & "') AND b_totcompras.toc_codbod=" & vg_codbod & ""
        Else
           vg_db.Execute "DELETE b_detcomprasimp FROM b_detcomprasimp, b_totcompras WHERE b_detcomprasimp.imd_numdoc = b_totcompras.toc_numdoc  AND b_detcomprasimp.imd_tipdoc = b_totcompras.toc_tipdoc AND b_detcomprasimp.imd_rutdoc = b_totcompras.toc_rutpro " & _
                         "AND b_totcompras.toc_fecrem < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND b_totcompras.toc_codbod = " & vg_codbod & ""
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_detcompras
        Label2.Caption = "Borrando Tabla b_detcompras": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_detcompras.* FROM b_totcompras INNER JOIN b_detcompras ON (b_totcompras.toc_numdoc=b_detcompras.dec_numdoc) AND (b_totcompras.toc_tipdoc=b_detcompras.dec_tipdoc) AND (b_totcompras.toc_rutpro=b_detcompras.dec_rutpro) " & _
                         "WHERE b_totcompras.toc_fecrem<CDate('" & fpDateTime1(0).text & "') AND b_totcompras.toc_codbod=" & vg_codbod & ""
        Else
           vg_db.Execute "DELETE b_detcompras FROM b_totcompras, b_detcompras WHERE b_totcompras.toc_numdoc = b_detcompras.dec_numdoc AND b_totcompras.toc_tipdoc = b_detcompras.dec_tipdoc AND b_totcompras.toc_rutpro = b_detcompras.dec_rutpro " & _
                         "AND b_totcompras.toc_fecrem < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND b_totcompras.toc_codbod = " & vg_codbod & ""
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_compras
        Label2.Caption = "Borrando Tabla b_compras": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_totcompras FROM b_totcompras WHERE toc_fecrem<CDate('" & fpDateTime1(0).text & "') AND toc_codbod=" & vg_codbod & ""
        Else
           vg_db.Execute "DELETE b_totcompras FROM b_totcompras WHERE toc_fecrem < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND toc_codbod = " & vg_codbod & ""
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_detventasimp
        Label2.Caption = "Borrando Tabla b_detventasimp": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_detventasimp.* FROM b_totventas INNER JOIN b_detventasimp ON (b_totventas.tov_rutcli=b_detventasimp.imd_rutdoc) AND (b_totventas.tov_tipdoc=b_detventasimp.imd_tipdoc) AND (b_totventas.tov_numdoc=b_detventasimp.imd_numdoc) " & _
                         "WHERE IIF(b_totventas.tov_tipdoc IN ('AI','ME','TR'), b_totventas.tov_fecemi<CDate('" & fpDateTime1(0).text & "'), b_totventas.tov_fecpro<CDate('" & fpDateTime1(0).text & "')) AND b_totventas.tov_codbod=" & vg_codbod & ""
        Else
           vg_db.Execute "DELETE b_detventasimp FROM b_totventas, b_detventasimp WHERE b_totventas.tov_rutcli = b_detventasimp.imd_rutdoc AND b_totventas.tov_tipdoc = b_detventasimp.imd_tipdoc AND b_totventas.tov_numdoc = b_detventasimp.imd_numdoc " & _
                         "AND b_totventas.tov_tipdoc IN ('AI','ME','TR') AND b_totventas.tov_fecemi < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND b_totventas.tov_codbod = " & vg_codbod & ""
           
           vg_db.Execute "DELETE b_detventasimp FROM b_totventas, b_detventasimp WHERE b_totventas.tov_rutcli = b_detventasimp.imd_rutdoc AND b_totventas.tov_tipdoc = b_detventasimp.imd_tipdoc AND b_totventas.tov_numdoc = b_detventasimp.imd_numdoc " & _
                         "AND b_totventas.tov_tipdoc NOT IN ('AI','ME','TR') AND b_totventas.tov_fecpro < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND b_totventas.tov_codbod = " & vg_codbod & ""
        
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_detventas
        Label2.Caption = "Borrando Tabla b_detventas": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_detventas.* FROM b_totventas INNER JOIN b_detventas ON (b_totventas.tov_rutcli=b_detventas.dev_rutcli) AND (b_totventas.tov_tipdoc=b_detventas.dev_tipdoc) AND (b_totventas.tov_numdoc=b_detventas.dev_numdoc) " & _
                         "WHERE IIF(b_totventas.tov_tipdoc IN ('AI','ME','TR'), b_totventas.tov_fecemi<CDate('" & fpDateTime1(0).text & "'), b_totventas.tov_fecpro<CDate('" & fpDateTime1(0).text & "')) AND b_totventas.tov_codbod=" & vg_codbod & ""
        Else
           vg_db.Execute "DELETE b_detventas FROM b_totventas, b_detventas WHERE b_totventas.tov_rutcli = b_detventas.dev_rutcli AND b_totventas.tov_tipdoc = b_detventas.dev_tipdoc AND b_totventas.tov_numdoc = b_detventas.dev_numdoc " & _
                         "AND b_totventas.tov_tipdoc IN ('AI','ME','TR') AND b_totventas.tov_fecemi < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND b_totventas.tov_codbod=" & vg_codbod & ""
        
           vg_db.Execute "DELETE b_detventas FROM b_totventas, b_detventas WHERE b_totventas.tov_rutcli = b_detventas.dev_rutcli AND b_totventas.tov_tipdoc = b_detventas.dev_tipdoc AND b_totventas.tov_numdoc = b_detventas.dev_numdoc " & _
                         "AND b_totventas.tov_tipdoc NOT IN ('AI','ME','TR') AND b_totventas.tov_fecpro < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND b_totventas.tov_codbod = " & vg_codbod & ""
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_compras
        Label2.Caption = "Borrando Tabla b_compras": DoEvents
        If vg_tipbase = "1" Then
           vg_db.Execute "DELETE b_totventas FROM b_totventas WHERE IIF(b_totventas.tov_tipdoc IN ('AI','ME','TR'), b_totventas.tov_fecemi<CDate('" & fpDateTime1(0).text & "'), b_totventas.tov_fecpro<CDate('" & fpDateTime1(0).text & "')) AND tov_codbod=" & vg_codbod & ""
        Else
           vg_db.Execute "DELETE b_totventas FROM b_totventas WHERE b_totventas.tov_tipdoc IN ('AI','ME','TR') AND b_totventas.tov_fecemi < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND tov_codbod = " & vg_codbod & ""
           
           vg_db.Execute "DELETE b_totventas FROM b_totventas WHERE b_totventas.tov_tipdoc NOT IN ('AI','ME','TR') AND b_totventas.tov_fecpro < '" & Format(fpDateTime1(0).text, "yyyymmdd") & "' AND tov_codbod=" & vg_codbod & ""
        End If
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_tomainv
        Label2.Caption = "Borrando Tabla b_tomainv": DoEvents
        vg_db.Execute "DELETE b_tomainv FROM b_tomainv WHERE tin_fectom<" & Format(fpDateTime1(0).text, "yyyymmdd") & " AND tin_codbod=" & vg_codbod & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_preciovta
'        Label2.Caption = "Borrando Tabla b_preciovta": DoEvents
'        vg_db.Execute "DELETE b_preciovta FROM b_preciovta WHERE prv_cencos='" & MuestraCasino(1) & "' AND prv_fecvig<" & Format(fpDateTime1(0).text, "yyyymmdd") & ""
'        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_ventacontadodet
        Label2.Caption = "Borrando Tabla b_ventacontadodet": DoEvents
        vg_db.Execute "DELETE b_ventacontadodet FROM b_ventacontadodet WHERE vtd_codigo IN (SELECT DISTINCT vtc_codigo FROM b_ventacontado WHERE vtc_cencos='" & MuestraCasino(1) & "' AND vtc_fecvta<" & Format(fpDateTime1(0).text, "yyyymmdd") & ")"
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_ventacontado
        Label2.Caption = "Borrando Tabla b_ventacontado": DoEvents
        vg_db.Execute "DELETE b_ventacontado FROM b_ventacontado WHERE vtc_cencos='" & MuestraCasino(1) & "' AND vtc_fecvta<" & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_presupuestoproyeccion
        Label2.Caption = "Borrando Tabla b_presupuestoproyeccion": DoEvents
        vg_db.Execute "DELETE b_presupuestoproyeccion FROM b_presupuestoproyeccion WHERE ppr_cencos='" & MuestraCasino(1) & "' AND ppr_anomes<" & Format(fpDateTime1(0).text, "yyyymm") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_gastosa13
        Label2.Caption = "Borrando Tabla b_gastosa13": DoEvents
        vg_db.Execute "DELETE b_gastosa13 FROM b_gastosa13 WHERE gas_cencos='" & MuestraCasino(1) & "' AND gas_anomes<" & Format(fpDateTime1(0).text, "yyyymm") & ""
        PB.Value = PB.Value + 1

        
        '-------> Borrar tabla b_costopatron
        Label2.Caption = "Borrando Tabla b_costopatron": DoEvents
        vg_db.Execute "DELETE b_costopatron FROM b_costopatron WHERE cpa_cencos='" & MuestraCasino(1) & "' AND cpa_anomes<" & Format(fpDateTime1(0).text, "yyyymm") & ""
        PB.Value = PB.Value + 1
        
        '-------> Borrar tabla b_productospmpdia
        Label2.Caption = "Borrando Tabla b_productospmpdia": DoEvents
        vg_db.Execute "DELETE b_productospmpdia FROM b_productospmpdia WHERE ppd_cencos = '" & MuestraCasino(1) & "' AND ppd_fecdia < " & Format(fpDateTime1(0).text, "yyyymmdd") & ""
        PB.Value = PB.Value + 1
        
        PB.Min = 0: PB.Value = 0: Label2.Visible = False: PB.Visible = False
        MsgBox "Proceso de Actualización Finalizado", vbInformation + vbOKOnly, Msgtitulo
        Command1(0).Enabled = True
        

'        '-------> Comprimir base de dato
'        vg_db.Close
'        Dim fso
'        Set fso = CreateObject("Scripting.FileSystemObject")
'        If Dir(dir_trabajo & Mid(BaseDeDato, 1, (Len(BaseDeDato) - 3)) & "ldb") <> "" Then
'           MsgBox "El sistema necesita respaldar la base de dato que esta abierta." & Chr(13) & Chr(13) & "No se ejecutara hasta cerrar la Base o los programas relacionados", vbExclamation + vbOKOnly, Msgtitulo
'        End If
'
'        DBEngine.CompactDatabase dir_trabajo & BaseDeDato, dir_trabajo & "respaldo.mdb", dbLangGeneral
'        Kill dir_trabajo & BaseDeDato
'        fso.MoveFile dir_trabajo & "respaldo.mdb", dir_trabajo & BaseDeDato
'        MsgBox "Proceso finalizo sin problema, el sistema cerrar..." & Chr(13), vbExclamation + vbOKOnly, Msgtitulo
'        End
    End If
Case 1 '-------> Salir
    Me.Hide
    Unload Me
End Select
End Sub

Private Sub Form_Load()
fg_centra Me
Text1(0).text = ""
Text1(0).text = dir_trabajo & BaseDeDato
Msgtitulo = "Traspaso Información"
fpDateTime1(0).text = Format(Date, "dd/mm/yyyy")
Frame2.Top = 360
Command1(0).Caption = "&Aceptar"
Command1(1).Caption = "&Salir"
fpDateTime1(0).Visible = False
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
Case 0
    Cd.Filename = ""
    Cd.Filter = "Todos los archivos (dbgt*.mdb)|dbgt*.mdb"
    Cd.DefaultExt = "dbgt*.mdb"
    Cd.Flags = &H80000 Or &H400& Or &H1000& Or &H4&
    Cd.ShowOpen
    If Cd.Filename = "" Then Text1(0).text = "" Else Text1(0).text = Cd.Filename 'Dir(CD.FileName)
End Select
End Sub



