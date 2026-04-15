VERSION 5.00
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form C_DetalleConveniosConErrorPel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle Convenios con Error PEL"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   20040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Aplicar y Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   18120
      TabIndex        =   14
      Top             =   7680
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   19815
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   6
         Left            =   2520
         TabIndex        =   15
         Top             =   6600
         Width           =   1020
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   3
            Left            =   45
            TabIndex        =   16
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   1
         Left            =   1440
         TabIndex        =   11
         Top             =   6600
         Width           =   1020
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   2
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   0
         Left            =   3600
         TabIndex        =   9
         Top             =   6600
         Width           =   3540
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   4
            Left            =   45
            TabIndex        =   10
            Top             =   135
            Width           =   3435
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   2
         Left            =   7200
         TabIndex        =   7
         Top             =   6600
         Width           =   1740
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   5
            Left            =   45
            TabIndex        =   8
            Top             =   135
            Width           =   1635
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   3
         Left            =   9000
         TabIndex        =   5
         Top             =   6600
         Width           =   3540
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   6
            Left            =   45
            TabIndex        =   6
            Top             =   135
            Width           =   3435
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   4
         Left            =   12600
         TabIndex        =   3
         Top             =   6600
         Width           =   1260
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   7
            Left            =   45
            TabIndex        =   4
            Top             =   135
            Width           =   1155
         End
      End
      Begin VB.Frame Frame3 
         Height          =   435
         Index           =   5
         Left            =   13920
         TabIndex        =   1
         Top             =   6600
         Width           =   3420
         Begin VB.TextBox TextDet2 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   240
            Index           =   8
            Left            =   45
            TabIndex        =   2
            Top             =   135
            Width           =   3315
         End
      End
      Begin FPSpread.vaSpread vaSpread1 
         Height          =   6255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   19575
         _Version        =   393216
         _ExtentX        =   34528
         _ExtentY        =   11033
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   11
         SpreadDesigner  =   "C_DetalleConveniosConErrorPel.frx":0000
      End
   End
End
Attribute VB_Name = "C_DetalleConveniosConErrorPel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Dim MsgTitulo          As String
'Public lc_Aux As String
Dim Est               As Boolean
Dim IdLoteDet         As Double
Dim Aux_HelpContextID As String

Private Sub Command1_Click()

On Error GoTo Man_Error

Dim i          As Long
Dim ISeleccion As Boolean
Dim MyBuffer   As String
Dim RS         As New ADODB.Recordset
Dim orgcom     As String
Dim Ceco       As String
Dim Prov       As String
Dim Mat        As String
Dim estado     As String
Dim Aux_i      As Long

ISeleccion = False

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
    
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
    
       ISeleccion = True
       Exit For
       
    End If

Next i

If Not ISeleccion Then

   MsgBox "Debe haber por lo menos un ítem seleccionado de la lista, para actualizar...", vbExclamation + vbOKOnly, MsgTitulo
   Exit Sub

End If

If MsgBox("Esta seguro realizar reenvio del lote a la PEL...", vbQuestion + vbYesNo, MsgTitulo) = vbNo Then Exit Sub

Aux_i = 0

Let MyBuffer = ""
Let MyBuffer = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "iso-8859-1" & Chr(34) & " ?>"
Let MyBuffer = MyBuffer & "<UpdateDetPel>"

For i = 1 To vaSpread1.MaxRows

    vaSpread1.Row = i
   
    vaSpread1.Col = 1
    
    If vaSpread1.text = "1" And vaSpread1.RowHidden = False Then
    
       vaSpread1.Col = 2
       orgcom = vaSpread1.text
       
       vaSpread1.Col = 3
       Ceco = vaSpread1.text
       
       vaSpread1.Col = 5
       Prov = vaSpread1.text
       
       vaSpread1.Col = 7
       Mat = vaSpread1.text
       
       MyBuffer = MyBuffer & " <DetPel"
       MyBuffer = MyBuffer & " OrgCom = " & Chr(34) & orgcom & Chr(34)
       MyBuffer = MyBuffer & " Ceco = " & Chr(34) & Ceco & Chr(34)
       MyBuffer = MyBuffer & " Prov = " & Chr(34) & Prov & Chr(34)
       MyBuffer = MyBuffer & " Mat = " & Chr(34) & Mat & Chr(34)
       MyBuffer = MyBuffer & "/>"

       Aux_i = Aux_i + 1
       
    End If
    
Next i

MyBuffer = MyBuffer & "</UpdateDetPel>"
      
If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient

Set RS = vg_db.Execute("sgpadm_Upd_XmlDetConveniosPel " & IdLoteDet & ", '" & MyBuffer & "'")

If Not RS.EOF Then

   If RS(0) > 0 Or RS(0) < 0 Then
        
     fg_descarga
      
     Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Error_Modificacion"), Aux_HelpContextID, "", "", "")

     MsgBox RS(1) & VgLinea, vbCritical, MsgTitulo
          
     RS.Close
     Set RS = Nothing
                 
     Exit Sub
              
   Else
        
      If RS(2) > 0 Then
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Modificado"), Aux_HelpContextID, "", "", "")

         MsgBox "Proceso finalizado sin problema...", vbInformation + vbOKOnly, Me.Caption
      
      Else
      
         Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_No_Encontraron_Datos_Actualizar"), Aux_HelpContextID, "", "", "")
         
         MsgBox "Proceso finalizado, no se encontraron datos que actualizar...", vbInformation + vbOKOnly, Me.Caption
         
      End If
              
   End If

End If

RS.Close
Set RS = Nothing

If Aux_i = vaSpread1.MaxRows Then

   vg_codigo = "X"
   
End If

fg_descarga

Call fg_GrabaLogSistema(vg_NUsr, fg_TraeLogConcepto("vg_logsis_Salir"), Aux_HelpContextID, "", "", "")

Me.Hide
Unload Me

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Activate()

On Error GoTo Man_Error

fg_descarga

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub Form_Load()

On Error GoTo Man_Error

Aux_HelpContextID = 1199200
vg_codigo = ""

fg_centra Me

Me.Caption = " Consultar & Actualizar Detalle Convenios con problema de envio - " & IdLoteDet

MsgTitulo = "Consultar & Actualizar Detalle Convenios con problema de envio"

Command1.Enabled = False
   
If Mid(ValidarUsuarioAcceso(1199200, vg_NUsr), 3, 1) = "1" Then

   Command1.Enabled = True

End If

Exit Sub
Man_Error:
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Sub LlenarDatos(IdLote As Double)

On Error GoTo Man_Error

Dim RS As New ADODB.Recordset

fg_carga ""

IdLoteDet = IdLote

vaSpread1.MaxRows = 0

If RS.State = 1 Then RS.Close
RS.CursorLocation = adUseClient
vg_db.CursorLocation = adUseClient
  
Set RS = vg_db.Execute("sgpadm_Sel_ConsultarDetalleConveniosPel " & IdLote & "")
If Not RS.EOF Then
   
   Do While Not RS.EOF
    
      vaSpread1.MaxRows = vaSpread1.MaxRows + 1
      vaSpread1.Row = vaSpread1.MaxRows
    
      vaSpread1.Col = 1
      vaSpread1.text = "0"
    
      'Org Compras
      vaSpread1.Col = 2
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(0)
    
      'Ceco
      vaSpread1.Col = 3
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = Trim(RS(1))
      
'      vaSpread2.Row = 1
'      vaSpread2.Col = IIf(IsNull(RS(2)) Or RS(2) = "C", 2, IIf(RS(2) = "X", 3, 1))
'
'      vaSpread1.Col = 4
'      vaSpread1.TypePictPicture = vaSpread2.TypePictPicture
      
      'Descripción Ceco
      vaSpread1.Col = 4
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(2)
    
      'Proveedor
      vaSpread1.Col = 5
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(3)
      
      'Descripción proveedor
      vaSpread1.Col = 6
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(4)
    
      'Material Sap
      vaSpread1.Col = 7
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(5)
    
      'Descripción material sap
      vaSpread1.Col = 8
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(6)

      'Fecha inicio validez
      vaSpread1.Col = 9
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(9)

      'Fecha fin validez
      vaSpread1.Col = 10
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = RS(10)
      
      vaSpread1.Col = 11
      vaSpread1.CellType = CellTypeStaticText
      vaSpread1.text = 0
         
      RS.MoveNext
      
   Loop
   
Else
   
   vaSpread1.MaxRows = 0
   fg_descarga
   MsgBox "No existe información requerida", vbExclamation + vbOKOnly, MsgTitulo

End If
RS.Close
Set RS = Nothing

fg_descarga


Exit Sub
Man_Error:
fg_descarga
MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub TextDet2_KeyPress(Index As Integer, KeyAscii As Integer)

On Error GoTo Man_Error

Dim i          As Long
Dim X          As Long
Dim indactivo  As Integer
Dim TexBus     As String
Dim EstBuq     As String
Dim wvarArr8020() As String
            
If KeyAscii <> 13 Then Exit Sub
'SendKeys "{Tab}"
    
wvarArr8020 = Split(TextDet2(Index).text, ",")

If Index = 2 Then
   
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 3 Then
   
   TextDet2(2).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 4 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(5).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 5 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 6 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(5).text = ""
   TextDet2(7).text = ""
   TextDet2(8).text = ""

ElseIf Index = 7 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(6).text = ""
   TextDet2(5).text = ""
   TextDet2(8).text = ""

ElseIf Index = 8 Then
   
   TextDet2(2).text = ""
   TextDet2(3).text = ""
   TextDet2(4).text = ""
   TextDet2(6).text = ""
   TextDet2(7).text = ""
   TextDet2(5).text = ""

End If

For i = 1 To vaSpread1.MaxRows
           
    vaSpread1.Row = i
    vaSpread1.Col = 11
    vaSpread1.text = 0
    
Next

Select Case Index

Case 2, 3, 4, 5, 6, 7, 8
    
    vaSpread1.Visible = False
    
    If Trim(TextDet2(Index).text) <> "" Then
       
       For X = 0 To UBound(wvarArr8020)
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           vaSpread1.Col = Index
           indactivo = UCase(Trim(vaSpread1.Value)) Like IIf(Index = 2 Or Index = 3 Or Index = 5 Or Index = 7, "" & UCase(wvarArr8020(X)) & "", "*" & UCase(wvarArr8020(X)) & "*")
           vaSpread1.Col = Index
           
           If indactivo = -1 And Trim(vaSpread1.text) <> "" Then
              
              vaSpread1.Col = 11
              
              If Val(vaSpread1.Value) <> 1 Then
                              
                 vaSpread1.Col = 1
              
                 If vaSpread1.RowHidden = True Then
                 
                    vaSpread1.RowHidden = False
                    vaSpread1.Col = 11
                    vaSpread1.text = 1
                 
                 Else
                 
                    vaSpread1.Col = 11
                    vaSpread1.text = 1
                 
                 End If
                 
              End If
              
           Else
              
              vaSpread1.Col = 11
              EstBuq = vaSpread1.Value
              vaSpread1.Col = 2
              
              If vaSpread1.RowHidden = False And EstBuq <> 1 Then
              
                 vaSpread1.RowHidden = True
                 
                 vaSpread1.Col = 11
                 vaSpread1.text = 0
                 
              End If
           
           End If
        
        Next i
        
        vaSpread1.SetActiveCell Index + 1, 1
        vaSpread1.ColUserSortIndicator(-1) = ColUserSortIndicatorNone
        vaSpread1.ColUserSortIndicator(IIf(Trim(wvarArr8020(X)) = "", 0, 0)) = ColUserSortIndicatorAscending
        vaSpread1.SortKey(1) = IIf(Trim(wvarArr8020(X)) = "", 0, 0): vaSpread1.SortKeyOrder(1) = SortKeyOrderAscending
        vaSpread1.Sort -1, -1, vaSpread1.maxcols, vaSpread1.MaxRows, SortByRow
        
        Next X
        
    End If
    
    If Trim(TextDet2(Index).text) = "" Then
       
       For i = 1 To vaSpread1.MaxRows
           
           vaSpread1.Row = i
           If vaSpread1.RowHidden = True Then vaSpread1.RowHidden = False
           
           vaSpread1.Col = 11
           vaSpread1.text = 0
       
       Next
       
       vaSpread1.SetActiveCell Index, vaSpread1.SearchCol(Index, 0, vaSpread1.MaxRows, Trim(TextDet2(Index).text), SearchFlagsGreaterOrEqual)
       vaSpread1.SetActiveCell Index, 1
    
    End If
    
    vaSpread1.Visible = True

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub

Private Sub vaSpread1_BlockSelected(ByVal BlockCol As Long, ByVal BlockRow As Long, ByVal BlockCol2 As Long, ByVal BlockRow2 As Long)

On Error GoTo Man_Error

Select Case BlockCol

Case 1
    
    Dim i As Long
    vaSpread1.Col = 1
       
    For i = 1 To vaSpread1.MaxRows
        
        vaSpread1.Row = i
        
        vaSpread1.Col = 1
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
        
        End If
    
    Next
    
Case Is <> 1

    If BlockRow = -1 Then Exit Sub

    vaSpread1.Col = 1
    
    For i = BlockRow To BlockRow2
        
        vaSpread1.Row = i
              
        vaSpread1.Col = 1
        
        If vaSpread1.RowHidden = False Then
            
           vaSpread1.Value = IIf(vaSpread1.Value = "1", "0", "1")
    
        End If
        
    Next
    
End Select

Exit Sub
Man_Error:
    Est = False
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, MsgTitulo

End Sub
