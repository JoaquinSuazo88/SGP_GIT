VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form M_CamIRe 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habilitar Cambio Ingrediente en Receta SGP"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7845
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6975
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         ItemData        =   "M_CamIRe.frx":0000
         Left            =   2520
         List            =   "M_CamIRe.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   165
         Width           =   3765
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   2595
         TabIndex        =   6
         Top             =   220
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo Cambio Ingrediente"
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
         Index           =   4
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Height          =   6225
      Left            =   30
      TabIndex        =   0
      Top             =   1560
      Width           =   7125
      Begin VSFlex8LCtl.VSFlexGrid grC 
         Height          =   5775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6735
         _cx             =   11880
         _cy             =   10186
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   14745342
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   14745342
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"M_CamIRe.frx":0004
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BorderStyle     =   1
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   135
      Index           =   0
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   300
   End
End
Attribute VB_Name = "M_CamIRe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS As New ADODB.Recordset
Dim RS1 As New ADODB.Recordset
Dim modo As String, Msgtitulo As String, Est As Boolean


Private Sub Combo1_Click(Index As Integer)
If Est Then Exit Sub
If Combo1(0).ListIndex <> -1 Then ValidarChecked
End Sub

Private Sub Form_Activate()
fg_descarga
End Sub

Private Sub Form_Load()
Me.HelpContextID = vg_OpcM
Me.Height = 8325
Me.Width = 7365
Msgtitulo = "Habilitar Cambio Ingrediente en Receta SGP"
fg_centra Me
modo = ""
Gl_Mo_Botones Me, 1
Gl_Ac_Botones Me, 1, 5, modo
Est = True
'------->
Combo1(0).Clear
RS.Open "SELECT gci_codigo, gci_nombre FROM a_grupocambioing ORDER BY gci_codigo", vg_db, adOpenStatic
If Not RS.EOF Then
   Do While Not RS.EOF
      Combo1(0).AddItem Trim(RS!gci_nombre) & Space(150) & "(" & fg_pone_cero(RS!gci_codigo, 10) & ")"
      RS.MoveNext
   Loop
End If
RS.Close: Set RS = Nothing
MoverDatos
Est = False
End Sub

Sub MoverDatos()
On Error GoTo Man_Error
Dim SecT As Long, SecS As Long, vRow As Long
fg_carga ""
grC.Cols = 4 '24
grC.Rows = 1
grC.FixedCols = 1 '7
grC.ExtendLastCol = False
grC.TextMatrix(0, 0) = "Familia Producto"
grC.TextMatrix(0, 1) = "Checked"
grC.TextMatrix(0, 3) = "Prod. Asoc."
grC.ColWidth(0) = 500: grC.ColAlignment(0) = flexAlignGeneral '= flexAlignCenterCenter
grC.ColWidth(1) = 1300: grC.ColAlignment(1) = flexAlignCenterCenter
grC.ColWidth(3) = 1300: grC.ColAlignment(3) = flexAlignCenterCenter
grC.Editable = flexEDKbdMouse
grC.OutlineCol = 0
grC.OutlineBar = flexOutlineBarSimpleLeaf
grC.MergeCells = flexMergeNever
grC.AllowUserResizing = flexResizeNone
grC.AllowSelection = True
grC.GridLines = flexGridFlatVert
vRND = fgRND
Set RS = vg_db.Execute("sgpadm_p_armarfamproducto")
Do While Not RS.EOF
    If RS!tip_nivel = 0 Then
        grC.Rows = grC.Rows + 1: grC.Row = grC.Rows - 1
        grC.IsSubtotal(grC.Row) = True
        grC.RowOutlineLevel(grC.Row) = 1
        grC.IsCollapsed(grC.Row) = flexOutlineCollapsed
        grC.Col = 0: grC.text = Trim(RS!tip_nombre)
        grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H80000012 '&HFF0000
    End If
    If RS!tip_nivel = 1 Then
        grC.Rows = grC.Rows + 1: grC.Row = grC.Rows - 1
        grC.IsSubtotal(grC.Row) = True
        grC.RowOutlineLevel(grC.Row) = 2
        grC.IsCollapsed(grC.Row) = flexOutlineCollapsed
        grC.Col = 0: grC.text = Trim(RS!tip_nombre)
        grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &HFF0000
    End If
    If RS!tip_nivel = 2 Then
        grC.Rows = grC.Rows + 1: grC.Row = grC.Rows - 1
        grC.IsSubtotal(grC.Row) = True
        grC.RowOutlineLevel(grC.Row) = 3
        grC.IsCollapsed(grC.Row) = flexOutlineCollapsed
        grC.Col = 0: grC.text = Trim(RS!tip_nombre)
        grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &HFF&
    End If
    If RS!tip_nivel = 3 Then
        grC.Rows = grC.Rows + 1: grC.Row = grC.Rows - 1
        grC.IsSubtotal(grC.Row) = True
        grC.RowOutlineLevel(grC.Row) = 4
        grC.IsCollapsed(grC.Row) = flexOutlineCollapsed
        grC.Col = 0: grC.text = Trim(RS!tip_nombre) 'IIf(RS1!tienepu = 0, "I", "I-PU")
        grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H8080&
    End If
    If RS!tip_nivel = 4 Then
        grC.Rows = grC.Rows + 1: grC.Row = grC.Rows - 1
        grC.IsSubtotal(grC.Row) = True
        grC.RowOutlineLevel(grC.Row) = 5
        grC.IsCollapsed(grC.Row) = flexOutlineCollapsed
        grC.Col = 0: grC.text = Trim(RS!tip_nombre) 'IIf(RS1!tienepu = 0, "I", "I-PU")
        grC.Cell(flexcpForeColor, grC.Row, 0, grC.Row, grC.Cols - 1) = &H8080&
    End If
    
    If RS!tip_check > 0 Then
       grC.Cell(flexcpChecked, grC.Row, 1) = flexUnchecked
       grC.Col = 3: grC.text = RS!tip_check
       grC.Cell(flexcpForeColor, grC.Row, grC.Col, grC.Row, grC.Cols - 1) = &H80000012
    End If
    grC.Col = 2: grC.text = RS!tip_codigo

    RS.MoveNext
Loop
RS.Close: Set RS = Nothing
If grC.Rows > 1 Then
    grC.Outline 5: grC.Outline 4: grC.Outline 3: grC.Outline 2: grC.Outline 1
    grC.AutoSize 0, 0, False
    grC.Cell(flexcpPictureAlignment, 1, 1, grC.Rows - 1, 1) = flexPicAlignCenterCenter
End If
grC.ColHidden(1) = False
grC.ColHidden(3) = False
If Combo1(0).ListIndex <> -1 Then ValidarChecked
fg_descarga
Exit Sub
Man_Error:
If Err = 3034 Then RS.Close: Set RS = Nothing: Exit Sub
RS.Close: Set RS = Nothing
Resume Next
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub

Sub ValidarChecked()
Dim i As Long, x As Long, codgrp As String, codfam As String, ValLcntH As String, Parentesis As String
Dim veccta()
'-------> limpiar grilla
 With grC
    For i = 1 To .Rows - 1
        .Row = i
        .Col = 1
        If .Cell(flexcpChecked, .Row, 1) = flexChecked Then
           .Col = 1
           .CellChecked = flexUnchecked
        End If
    Next
 End With

codgrp = Trim(fg_codigocbo(Combo1, 0, 10, 0))
RS.Open "SELECT par_codigo, par_valor FROM a_param WHERE par_codigo='" & codgrp & "'", vg_db, adOpenStatic

If Not RS.EOF Then
    ReDim Preserve veccta(0)
    ValLcntH = "": Parentesis = "": i = 0
    If Not IsNull(RS!par_valor) Then Parentesis = RS!par_valor
    i = 0
    For x = 1 To Len(Parentesis)
        If Asc(Mid(Parentesis, x, 1)) <> 59 Then
           ValLcntH = ValLcntH + Mid(Parentesis, x, 1)
        Else
'              If codigo = Val(ValLcntH) Then encuentra = True: Exit For
           ReDim Preserve veccta(i): veccta(i) = ValLcntH: ValLcntH = "": i = i + 1
        End If
    Next x
    ReDim Preserve veccta(i): veccta(i) = ValLcntH
   With grC
      For i = 1 To .Rows - 1
          .Row = i
          .Col = 1
          If .Cell(flexcpChecked, .Row, 1) = flexUnchecked Then
             .Col = 2: codfam = .text
             For x = 0 To UBound(veccta)
                 If veccta(x) = codfam Then
                    .Col = 1
                    .CellChecked = flexChecked
                    Exit For
                 End If
             Next x
          End If
      Next
   End With
End If
RS.Close: Set RS = Nothing
End Sub

Private Sub grC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
grC.Col = Col
grC.Row = Row
If modo = "" Then modo = "M"
If Toolbar1.Buttons(12).Visible = False Then Gl_Ac_Botones Me, 1, 0, modo
End Sub

Private Sub grC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
grC.Editable = IIf(NewCol = 3, flexEDNone, flexEDKbd)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim i As Long, codgrp As String, codfam As String, indgra As Boolean
On Error GoTo Man_Error
Select Case Button.Index
Case 3 '-------> Modificar
    modo = "M"
    Gl_Ac_Botones Me, 1, 0, modo
Case 7, 10 '-------> Actualizar lista y cancelar
    MoverDatos
    Gl_Ac_Botones Me, 1, 5, modo
Case 12 '------> Confirmar
    If Combo1(0).ListIndex = -1 Then MsgBox "Debe seleccionar grupo cambio... ", vbInformation + vbOKOnly, Msgtitulo: Exit Sub
    fg_carga ""
    indgra = False
    codgrp = Trim(fg_codigocbo(Combo1, 0, 10, 0))
    codfam = ""
    With grC
         For i = 1 To .Rows - 1
             .Row = i
             .Col = 1
             If .CellChecked = flexChecked Then
                .Col = 2
                codfam = codfam & .text & ";"
                indgra = True
             End If
         Next
    End With
    '-------> Borrar registro
    vg_db.Execute "DELETE FROM a_param WHERE par_codigo = '" & codgrp & "'"
    
    If indgra Then
       '------>  Agregar registro
       codfam = Mid(codfam, 1, Len(codfam) - 1)
'       vg_db.Execute "INSERT INTO a_param VALUES ('" & codgrp & "', '" & Trim(Mid(Combo1(0).text, 1, 40)) & "', 'C', '" & codfam & "')"
       vg_db.Execute "sgpadm_iu_param 'A', '" & codgrp & "', '" & Trim(Mid(Combo1(0).text, 1, 40)) & "', 'C', '" & codfam & "'"
    End If
    modo = ""
    Gl_Ac_Botones Me, 1, 5, modo
    fg_descarga
    MsgBox "Generación grabado finalizado sin problema", vbInformation + vbOKOnly, Msgtitulo
Case 15 '-------> Imprimir
    I_AsociarListaPrecio
Case 18 '-------> Salir
   Me.Hide
   Unload Me
End Select
Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
ins_log_error Date & Time & Err & ":  " & Error$(Err)
End Sub
