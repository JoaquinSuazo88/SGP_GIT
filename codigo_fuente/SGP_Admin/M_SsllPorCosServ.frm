VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form M_SsllPorCosServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Porcentaje Costo Servicio"
   ClientHeight    =   8160
   ClientLeft      =   3750
   ClientTop       =   1920
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame5 
      Height          =   1095
      Left            =   720
      TabIndex        =   2
      Top             =   600
      Width           =   8415
      Begin EditLib.fpLongInteger fpLongInteger1 
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1245
         _Version        =   196608
         _ExtentX        =   2196
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
         BackColor       =   16777215
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
         AlignTextH      =   2
         AlignTextV      =   0
         AllowNull       =   -1  'True
         NoSpecialKeys   =   3
         AutoAdvance     =   0   'False
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
         OnFocusPosition =   0
         ControlType     =   0
         Text            =   ""
         MaxValue        =   "2147483647"
         MinValue        =   "-2147483647"
         NegFormat       =   1
         NegToggle       =   0   'False
         Separator       =   ""
         UseSeparator    =   0   'False
         IncInt          =   1
         BorderGrayAreaColor=   -2147483637
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
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7800
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "M_SsllPorCosServ.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar Toolbar2 
         Height          =   360
         Left            =   3000
         TabIndex        =   4
         Top             =   600
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   635
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Cargar Información"
               ImageIndex      =   1
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3465
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label sombra 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3495
         TabIndex        =   8
         Top             =   290
         Width           =   4335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   3480
         TabIndex        =   7
         Top             =   720
         Width           =   75
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Centro de Costo"
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
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   315
         Width           =   1380
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   3000
         Picture         =   "M_SsllPorCosServ.frx":039A
         Top             =   120
         Width           =   480
      End
   End
   Begin VB.Frame Frame6 
      Height          =   6135
      Left            =   240
      TabIndex        =   0
      Top             =   1800
      Width           =   9135
      Begin FPSpread.vaSpread vaSpread2 
         Height          =   5610
         Left            =   480
         TabIndex        =   1
         Top             =   240
         Width           =   8055
         _Version        =   393216
         _ExtentX        =   14208
         _ExtentY        =   9895
         _StockProps     =   64
         BackColorStyle  =   3
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
         MaxCols         =   3
         SpreadDesigner  =   "M_SsllPorCosServ.frx":06A4
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   360
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   7440
      Width           =   855
   End
End
Attribute VB_Name = "M_SsllPorCosServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim RS2 As New ADODB.Recordset
Dim modo As String, codigo As String, Msgtitulo As String
Dim est As Boolean
Dim strSQL, CambRut As String
Dim TotalRegistros As Integer
Dim OpGr As Boolean
Dim varNuevaCol, varFlag As String
Dim wvarOriginalDesde, wvarOriginalHasta, TipoOp As String

Private Sub Form_Load()
    Dim itop As Integer
    
    Me.HelpContextID = vg_OpcM
    Me.Height = 9000
    Me.Width = 9700
    fg_centra Me
    Msgtitulo = "Porcentaje Costo Servicio"
        
    modo = ""
    est = True
    
    TotalRegistros = 0
    itop = 1
    'CallForm = Local_CallForm
    
    Gl_Mo_Botones Me, 1
    Gl_Ac_Botones Me, 1, 3, modo
        
    varNuevaCol = ""
    OpGr = False
    
    vaSpread2.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    TipoOp = ""
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        If Not (fpLongInteger1(1).text = "") Then
            MoverDatosGrilla
            Gl_Ac_Botones Me, 1, 15, modo
        End If
    End Select
    vaSpread2.Enabled = True
    'vaSpread2.SetFocus
    TotalRegistros = vaSpread2.MaxRows
    TipoOp = ""
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim codigo As Long, Nombre As String, NomCor As String
Dim vartxtCol1, vartxtCol3, vartxtCol5, vartxtCol6, wvarNewDesde, wvarNewHasta As String
Dim i As Integer

On Error GoTo Man_Error

Select Case Button.Index
Case 1
    TipoOp = "Agregar"
    vaSpread2.Enabled = True
    
    If (fpLongInteger1(1).text <> "") Then
    
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            
            vaSpread2.Col = 1
            vaSpread2.CellType = CellTypeStaticText
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 2
            vaSpread2.CellType = CellTypeStaticText
            vaSpread2.TypeHAlign = TypeHAlignRight
            
            vaSpread2.Col = 3
            vaSpread2.CellType = CellTypeStaticText
            vaSpread2.TypeHAlign = TypeHAlignRight
        Next i
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        
        vaSpread2.Row = vaSpread2.MaxRows: vaSpread2.Col = 2
        vaSpread2.SetActiveCell 1, vaSpread2.MaxRows: vaSpread2.SetFocus
                        
        vaSpread2.Col = 3
                
        varNuevaCol = 1
        
        modo = "A"
        Gl_Ac_Botones Me, 1, 0, modo
        fpLongInteger1(1).Enabled = False
        Image1(0).Enabled = False
        Toolbar2.Enabled = False
    Else
        If Toolbar1.Buttons(1).Visible = True Then Toolbar1.Buttons(1).Visible = False: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(3).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(7).Visible = False
    End If
Case 5
    'BORRAR un registro
    
    If (wvarOriginalDesde <> "" And wvarOriginalHasta <> "") Then
        wvarNewDesde = wvarOriginalDesde
        wvarNewHasta = wvarOriginalHasta
        If MsgBox("¿Desea Eliminar La Fila Seleccionada?   ", vbQuestion + vbYesNo, Msgtitulo) = vbYes Then
            strSQL = "DELETE b_ssll_pctcosto WHERE pcs_desde = " & wvarNewDesde & " AND pcs_hasta = " & wvarNewHasta & ""
            vg_db.Execute strSQL
            Cancela
            Exit Sub
            MsgBox "Borrando Registro"
        Else
            Cancela
        End If
    End If
    
Case 10
    'NO guarda la asociacion nueva
    
    If MsgBox("Cancela...", vbQuestion + vbYesNo, Msgtitulo) = vbNo Then Exit Sub
        modo = "Cancel"
        
        If modo = "Cancel" Then
            modo = ""
            Cancela
        Else
            Cancela
        End If
    
Case 12
    'SI guarda la asociacion nueva
    Dim DesdeUsuario, HastaUsuario, PorcentajeUsuario As String
    Dim DesdeGri, HastaGri, PorcentajeGri As String
    Dim wvarAnterior, wvarPosterior As String
    Dim CantReg, varFlagCorrecto As Integer
    Dim Diferencia As String
    Dim RangoVacio, varRango2 As String
    
    
    varFlag = 0
    vaSpread2.Row = vaSpread2.MaxRows

    vaSpread2.Col = 1
    DesdeUsuario = CDbl(fg_Quitachar(Trim(vaSpread2.text), ","))

    vaSpread2.Col = 2
    HastaUsuario = CDbl(fg_Quitachar(Trim(vaSpread2.text), ","))

    vaSpread2.Col = 3
    PorcentajeUsuario = CDbl(fg_Quitachar(Trim(vaSpread2.text), "%"))
    
    If (vaSpread2.MaxRows = 1) Then
        varFlagCorrecto = 1
    End If
       
    For i = 1 To vaSpread2.MaxRows - 1
        vaSpread2.Row = i
        vaSpread2.Col = 1
        DesdeGri = CDbl(fg_Quitachar(Trim(vaSpread2.text), ","))
        
        vaSpread2.Col = 2
        HastaGri = CDbl(fg_Quitachar(Trim(vaSpread2.text), ","))
        
        If Val(DesdeUsuario) > Val(HastaUsuario) Then
            varFlag = varFlag + 1
            Exit For
        ElseIf (Val(DesdeUsuario) >= Val(DesdeGri) And Val(DesdeUsuario) <= Val(HastaGri)) Then 'Or (Val(HastaUsuario) >= Val(DesdeGri) And Val(HastaUsuario) <= Val(HastaGri)) Then
            varFlag = varFlag + 1
            Exit For
        ElseIf (Val(HastaUsuario) >= Val(DesdeGri) And Val(HastaUsuario) <= Val(HastaGri)) Then
            varFlag = varFlag + 1
            Cancela
            Exit For
        Else
            vaSpread2.Row = i + 1
            vaSpread2.Col = 1
            varRango2 = fg_Quitachar(Trim(vaSpread2.text), ",")
            vaSpread2.Row = i
            
            'If (Val(HastaGri) < Val(DesdeUsuario) And Val(varRango2) > Val(HastaUsuario)) Then
                If (i = 1 And Val(HastaUsuario) < Val(DesdeGri)) Then
                    varFlagCorrecto = 1
                    Exit For
                ElseIf (Val(HastaUsuario) < Val(DesdeGri)) Then
                    If (Val(HastaGri) < Val(DesdeUsuario) And Val(varRango2) > Val(HastaUsuario)) Then
                        varFlagCorrecto = 1
                        Exit For
                    End If
                ElseIf (Val(varRango2) > Val(HastaUsuario) And Val(DesdeUsuario) > Val(DesdeGri)) Then
                    varFlagCorrecto = 1
                    Exit For
                ElseIf vaSpread2.MaxRows = Val(i) + 1 And Val(DesdeUsuario) > Val(HastaGri) Then
                    varFlagCorrecto = 1
                    Exit For
                
                ElseIf (i = vaSpread2.MaxRows - 1 And Val(HastaUsuario) < Val(HastaGri)) Then
                    varFlagCorrecto = 1
                    Exit For
                End If
            'End If
        End If
    Next
    
    If (varFlagCorrecto <> 1) Then
        MsgBox ("El Rango o Parte de Este Ya Está Ingresado"), vbExclamation + vbOKOnly, Msgtitulo
        Cancela
        fpLongInteger1(1).SetFocus
    Else
        strSQL = "INSERT INTO b_ssll_pctcosto(pcs_codcen, pcs_desde, pcs_hasta, pcs_pctcos) " & _
                 "VALUES('" & fpLongInteger1(1).Value & "'," & DesdeUsuario & "," & HastaUsuario & "," & _
                 "" & PorcentajeUsuario & ")"
        
        vg_db.Execute strSQL
        
        fpLongInteger1(1).Enabled = True
        Image1(0).Enabled = True
        Toolbar2.Enabled = True
        Cancela
        MsgBox "Se Han Guardado Los Datos", vbInformation + vbOKOnly, Msgtitulo
        TipoOp = ""
    End If

Case 13

Case 15
    If vaSpread2.MaxRows < 1 Then MsgBox "No Existe Datos Imprimir", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
    I_PorcCostoServ
Case 4
    
Case 18
    Me.Hide
    Unload Me
End Select

Exit Sub
Man_Error:
If Err = 3034 Then Exit Sub
If Err.Number = -2147467259 Then MsgBox "El dato esta asociado con otra tabla...", vbCritical, "Error" Else MsgBox "Error : " & Err & " " & Err.Description, vbCritical, "Error"
End Sub



Sub MoverDatosGrilla()
    fg_carga ""
    Dim x As Boolean

    vaSpread2.TextTip = 2
    vaSpread2.TextTipDelay = 250
    x = vaSpread2.SetTextTipAppearance("Arial", "11", False, False, &HFFFF&, &H800000)
    vaSpread2.Visible = False
    vaSpread2.MaxRows = 0
    vaSpread2.Row = -1
    vaSpread2.Col = -1
    
    If (fpLongInteger1(1).Value = "") Then
        strSQL = "SELECT pcs_codcen, pcs_desde, pcs_hasta, pcs_pctcos " & _
                 "FROM b_ssll_pctcosto, b_clientes " & _
                 "WHERE cli_tipo = 0 AND cli_activo = 1 AND pcs_codcen = cli_codigo " & _
                 "ORDER BY cli_codigo"
    Else
        strSQL = "SELECT pcs_codcen, pcs_desde, pcs_hasta, pcs_pctcos " & _
                 "FROM b_ssll_pctcosto, b_clientes " & _
                 "WHERE cli_tipo = 0 AND cli_activo = 1 AND pcs_codcen = cli_codigo " & _
                 "AND cli_codigo = '" & fpLongInteger1(1).Value & "' " & _
                 "ORDER BY cli_codigo"
    End If

    Set RS = vg_db.Execute(strSQL)
    
    Do While Not RS.EOF
        vaSpread2.MaxRows = vaSpread2.MaxRows + 1
        vaSpread2.Row = vaSpread2.MaxRows
        
        vaSpread2.Col = 1
        vaSpread2.text = Trim(RS!pcs_desde)
        vaSpread2.Col = 2
        vaSpread2.text = Trim(RS!pcs_hasta)
        vaSpread2.Col = 3
        vaSpread2.text = Trim(IIf(IsNull(RS!pcs_pctcos), "0", RS!pcs_pctcos))
                
        RS.MoveNext
    Loop
    
    RS.Close: Set RS = Nothing
    
    vaSpread2.Visible = True
    
    If vaSpread2.MaxRows > 0 Then
       vaSpread2.Row = 1
       vaSpread2.Col = 1
       codigo = ""
       codigo = Val(vaSpread2.text)
       vaSpread2.SetActiveCell 1, 1
    End If
    
    Toolbar1.Buttons(2).Visible = True
    Toolbar1.Buttons(2).Enabled = True
        
    Label2.Caption = Format(vaSpread2.MaxRows, fg_Pict(7, 0)) & " Registro"
    fg_descarga
End Sub


Private Sub fpLongInteger1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then Exit Sub
    SendKeys "{Tab}"
End Sub


Private Sub fpLongInteger1_Change(Index As Integer)
    If est Then Exit Sub
    fpayuda(0).Caption = ""
    If modo = "" Then modo = "M"
End Sub


Private Sub fpLongInteger1_LostFocus(Index As Integer)
    Dim codi As Long, Bd As String, Ul As String
    On Error GoTo Man_Error
    If fpLongInteger1(Index).Value = "" Then fpayuda(0).Caption = "": codi = 0: Exit Sub
    
    codi = fpLongInteger1(Index).Value
    Bd = IIf(Index = 1, "b_clientes", "")
    Ul = IIf(Bd = "b_clientes", "cli", "")
    
    Set RS1 = Nothing
    
    strSQL = "SELECT " & Ul & "_codigo, " & Ul & "_nombre FROM " & Bd & " WHERE " & Ul & "_codigo=" & IIf(Ul = "cli", "'" & codi & "'", codi) & ""
    RS1.Open strSQL, vg_db, adOpenStatic
    
    If Not RS1.EOF Then
        fpayuda(0).Caption = IIf(IsNull(Trim(RS1!cli_nombre) = ""), "", RS1!cli_nombre)
        vg_codigo = RS1!cli_codigo
        codi = 0
    Else
        MsgBox "No existe codigo en la tabla..."
        fpayuda(0).Caption = ""
        fpLongInteger1(Index).Value = ""
        codi = 0
        On Error Resume Next: fpLongInteger1(Index).SetFocus
    End If
    
    RS1.Close: Set RS1 = Nothing
    Exit Sub
    
Man_Error:
    If Err = 3034 Then Exit Sub
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub


Private Sub Image1_Click(Index As Integer)
    vg_left = fpayuda(0).Left + 2300
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_clientes", "cli_", "Mantenedor Centro Costo", "CentCost"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpLongInteger1(1).Value = Val(vg_codigo)
    fpayuda(0).Caption = vg_nombre
End Sub


Private Sub Cancela()
    OpGr = True
    vaSpread2.Row = vaSpread2.ActiveRow
    
    MoverDatosGrilla
    
    OpGr = False
    CambRut = ""
    
    modo = "": Gl_Ac_Botones Me, 1, 15, modo
    
    fpLongInteger1(1).Enabled = True
    Image1(0).Enabled = True
    Toolbar2.Enabled = True
    TipoOp = ""
End Sub


Private Sub GrabaRegistro(Fila)

    Dim PrvRut, PrvNombre As String
    Dim ProdCod, ProdNombre As String
    Dim FechaVigencia As String
    Dim Porcentaje, CodCentroCosto As String
    
    On Error GoTo Man_Error
    
    OpGr = True
    vaSpread2.Row = Fila
    
    CodCentroCosto = Trim(fpLongInteger1(1).text)
    
    vaSpread2.Col = 1: PrvRut = fg_DespintaRut(Trim(vaSpread2.Value))
    vaSpread2.Col = 2: PrvNombre = Trim(vaSpread2.Value)
    
    vaSpread2.Col = 3: ProdCod = Trim(vaSpread2.Value)
    vaSpread2.Col = 4: ProdNombre = Trim(vaSpread2.Value)
    
    vaSpread2.Col = 5
    If Trim(vaSpread2.text) = "" Then
        FechaVigencia = Format(Date, "yyyy/mm/dd")
    Else
        FechaVigencia = Format(CDate(Trim(vaSpread2.text)), "yyyy/mm/dd")
    End If
    
    vaSpread2.Col = 6
    If Trim(vaSpread2.text) = "" Then
        Porcentaje = 0
    Else
        Porcentaje = fg_Quitachar(Trim(vaSpread2.text), "%")
    End If
    
    If PrvNombre = "" Or ProdNombre = "" Or CodCentroCosto = "" Then MsgBox "Falta información...", vbExclamation + vbOKOnly, Msgtitulo: vaSpread2.Row = Fila: vaSpread2.SetActiveCell vaSpread2.ActiveCol, vaSpread2.ActiveRow: vaSpread2.SetFocus: OpGr = False: Cancela: Exit Sub
    
    If modo = "A" Then
    
        strSQL = "INSERT INTO b_ssll_dxv(dxv_codcen, dxv_rutpro, dxv_codfmc, dxv_fecvig, dxv_pctdxv) " & _
                 "VALUES('" & CodCentroCosto & "','" & PrvRut & "','" & ProdCod & "','" & _
                 "" & FechaVigencia & "'," & Porcentaje & ")"
            
        vg_db.Execute strSQL
    
    modo = "": Gl_Ac_Botones Me, 1, 1, modo
    OpGr = False

    MoverDatosGrilla

    Toolbar1.Buttons(1).Visible = True: Toolbar1.Buttons(2).Visible = False: Toolbar1.Buttons(3).Visible = False
    Toolbar1.Buttons(4).Visible = False: Toolbar1.Buttons(5).Visible = False: Toolbar1.Buttons(6).Visible = False
    Toolbar1.Buttons(7).Visible = False: Toolbar1.Buttons(8).Visible = False
    
    Toolbar1.Buttons(1).Enabled = True: Toolbar1.Buttons(11).Enabled = False: Toolbar1.Buttons(13).Enabled = False
    Toolbar1.Buttons(15).Enabled = False: Toolbar1.Buttons(15).Enabled = True
    
    End If
    
Man_Error:
    If Err = -2147467259 Then MsgBox "El dato esta asociado a otra tabla...", vbCritical, "Error": Exit Sub
    If Err = 3034 Then Exit Sub
    fg_descarga
    ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub


Private Sub vaSpread2_Click(ByVal Col As Long, ByVal Row As Long)
    vaSpread2.Row = Row
    vaSpread2.Col = 1
    wvarOriginalDesde = fg_Quitachar(Trim(vaSpread2.text), ",")
    vaSpread2.Col = 2
    wvarOriginalHasta = fg_Quitachar(Trim(vaSpread2.text), ",")
End Sub

Private Sub vaSpread2_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim GrillaDesde, GrillaHasta As String
    
    vaSpread2.Row = NewRow
    vaSpread2.Col = 1
    wvarOriginalDesde = fg_Quitachar(Trim(vaSpread2.text), ",")
    vaSpread2.Col = 2
    wvarOriginalHasta = fg_Quitachar(Trim(vaSpread2.text), ",")
End Sub


Private Sub vaSpread2_Change(ByVal Col As Long, ByVal Row As Long)
Dim GrillaDesde, GrillaHasta, GrillaPorcent As String
Dim NuevoDesde, NuevoHasta As String
Dim LimiteRango1, LimiteRango2 As String
Dim i, varFlagCorrecto, Existe As Integer

varFlag = 1
    
vaSpread2.Row = Row
vaSpread2.Col = 1
GrillaDesde = fg_Quitachar(Trim(vaSpread2.text), ",")
vaSpread2.Col = 2
GrillaHasta = fg_Quitachar(Trim(vaSpread2.text), ",")
vaSpread2.Col = 3
GrillaPorcent = fg_Quitachar(Trim(vaSpread2.text), "%")


If (Col = 1) Then
    vaSpread2.Row = Row - 1
    vaSpread2.Col = 2
    
    If vaSpread2.Row = 0 Then
        LimiteRango1 = 0
    Else
        LimiteRango1 = fg_Quitachar(Trim(vaSpread2.text), ",")
    End If
End If
    
If (Col = 2) Then
    vaSpread2.Row = Row + 1
    vaSpread2.Col = 1
    LimiteRango2 = fg_Quitachar(Trim(vaSpread2.text), ",")
End If

If (TipoOp <> "Agregar") Then

    varFlag = 0
    For i = 1 To vaSpread2.MaxRows
        If (Col = 3) Then
            varFlagCorrecto = 1
            Exit For
        End If
        
        If (Col = 1) Then
            If Val(GrillaDesde) > Val(GrillaHasta) Then
                varFlagCorrecto = 0
                Cancela
                Exit For

            ElseIf (Val(GrillaDesde) <= Val(LimiteRango1)) Then
                If (Val(GrillaDesde) <> 0 And Val(LimiteRango1) <> 0) Then
                    varFlagCorrecto = 0
                    Cancela
                    Exit For
                Else
                    varFlagCorrecto = 1
                    Exit For
                End If
            ElseIf Val(LimiteRango1) = 0 Then
                varFlagCorrecto = 1
            Else
                varFlagCorrecto = 1
            End If
        End If
            
            
        If (Col = 2) Then
            If Val(GrillaDesde) > Val(GrillaHasta) Then
                varFlagCorrecto = 0
                Cancela
                Exit For
            ElseIf Val(LimiteRango2) = 0 Then
                varFlagCorrecto = 1
                Cancela
                Exit For
            ElseIf (Val(GrillaHasta) >= Val(LimiteRango2)) Then
                varFlagCorrecto = 0
                Cancela
                Exit For
            Else
                varFlagCorrecto = 1
                Cancela
                Exit For
            End If
        End If
    Next i
    
    If varFlagCorrecto <> 1 Then
        MsgBox ("El Rango o Parte de Este Ya Está Ingresado"), vbExclamation + vbOKOnly, Msgtitulo
        Cancela
        fpLongInteger1(1).SetFocus
    Else
        Existe = 0
        
        For i = 1 To vaSpread2.MaxRows
            vaSpread2.Row = i
            vaSpread2.Col = Col
            If (fg_Quitachar(Trim(vaSpread2.text), ",") = GrillaDesde And vaSpread2.Row <> Row) Then
                Existe = 1
            End If
        Next i
        
        If Existe = 0 Then
            strSQL = "UPDATE b_ssll_pctcosto SET pcs_desde = " & GrillaDesde & ", " & _
                     "pcs_hasta = " & GrillaHasta & ", pcs_pctcos = " & GrillaPorcent & " " & _
                     "WHERE pcs_codcen = '" & fpLongInteger1(1).Value & "' AND pcs_desde = " & wvarOriginalDesde & " " & _
                     "AND pcs_hasta = " & wvarOriginalHasta
            
            vg_db.Execute strSQL
            
    
            fpLongInteger1(1).Enabled = True
            Image1(0).Enabled = True
            Toolbar2.Enabled = True
            Cancela
        End If
    End If
End If
End Sub
