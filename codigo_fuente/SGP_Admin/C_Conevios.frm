VERSION 5.00
Object = "{CDF3B183-D408-11CE-AE2C-0080C786E37D}#3.0#0"; "Edt32x30.ocx"
Begin VB.Form C_Convenios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Convenios"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   8535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2220
      Left            =   210
      TabIndex        =   0
      Top             =   0
      Width           =   8205
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar Excel"
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
         Left            =   5355
         TabIndex        =   6
         Top             =   1575
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
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
         Left            =   6720
         TabIndex        =   5
         Top             =   1575
         Width           =   1275
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   1
         Left            =   1365
         TabIndex        =   2
         Top             =   315
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
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
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin EditLib.fpText fpText 
         Height          =   315
         Index           =   0
         Left            =   1365
         TabIndex        =   4
         Top             =   735
         Width           =   1275
         _Version        =   196608
         _ExtentX        =   2249
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
         ButtonStyle     =   0
         ButtonWidth     =   0
         ButtonWrap      =   -1  'True
         ButtonDefaultAction=   0   'False
         ThreeDText      =   0
         ThreeDTextHighlightColor=   -2147483637
         ThreeDTextShadowColor=   -2147483632
         ThreeDTextOffset=   1
         AlignTextH      =   2
         AlignTextV      =   2
         AllowNull       =   0   'False
         NoSpecialKeys   =   3
         AutoAdvance     =   -1  'True
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
         MaxLength       =   10
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
         ButtonAlign     =   1
         OLEDropMode     =   0
         OLEDragMode     =   0
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   480
         Index           =   0
         Left            =   2625
         Picture         =   "C_Conevios.frx":0000
         Top             =   210
         Width           =   480
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   0
         Left            =   3120
         TabIndex        =   9
         Top             =   300
         Width           =   4935
      End
      Begin VB.Label fpayuda 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   1
         Left            =   3120
         TabIndex        =   7
         Top             =   720
         Width           =   4935
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   3165
         TabIndex        =   8
         Top             =   750
         Width           =   4920
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   480
         Index           =   1
         Left            =   2625
         Picture         =   "C_Conevios.frx":030A
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cód. Ing."
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
         Left            =   210
         TabIndex        =   3
         Top             =   780
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Org.Compra"
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
         Left            =   210
         TabIndex        =   1
         Top             =   375
         Width           =   1005
      End
      Begin VB.Label label 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   3165
         TabIndex        =   10
         Top             =   330
         Width           =   4920
      End
   End
End
Attribute VB_Name = "C_Convenios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Msgtitulo As String

Private Sub Command1_Click()

Dim RS       As New ADODB.Recordset
Dim Sql      As String
Dim xlApp    As Object
Dim xlWb     As Object
Dim xlWs     As Object
Dim recArray As Variant
    
 On Error GoTo Man_Error
 
'-----> Validar Org. Compras
If Trim(fpayuda(0).text) = "" Then
   MsgBox "Debe ingresar Org. Compras", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
End If

'-------> Validar Ingrediente
If Trim(fpText1(0).text) <> "" Then
  Set RS = vg_db.Execute("SELECT ing_codigo, ing_nombre " & _
               "FROM b_ingrediente WITH (NOLOCK) " & _
               "WHERE ing_codigo = '" & LimpiaDato(Trim(fpText1(0).text)) & "' " & _
               "AND   ing_indppr   = 1 " & _
               "AND   ing_activo = '1'")
   If RS.EOF Then
      RS.Close
      Set RS = Nothing
      fpayuda(1).Caption = ""
      MsgBox "No existe Ingrediente", vbExclamation + vbOKOnly, Msgtitulo: Exit Sub
      Exit Sub
   End If
   fpayuda(1).Caption = Trim(RS!ing_nombre)
   fpText1(0).text = RS!ing_codigo
   RS.Close
   Set RS = Nothing
End If

    '-------> Lectura
    Sql = ""
    Sql = " sgp_Sel_ConveniosSap "
    sel = sel & " '" & fpText1(1).text & "' "
    sel = sel & " ,'" & fpText1(0).text & "' "
    Set RS = vg_db.Execute(Sql)

    '-------> Create an instance of Excel and add a workbook
    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Add
    Set xlWs = xlWb.Worksheets("Hoja1")
  
    '-------> Display Excel and give user control of Excel's lifetime
'    xlApp.Visible = True
    xlApp.UserControl = True
    
    '-------> Check version of Excel
    Call encabezado(rst, xlWs)
          
    xlWs.Cells(2, 1).CopyFromRecordset rst
    '-------> Auto-fit the column widths and row heights
    xlApp.Selection.CurrentRegion.Columns.AutoFit
    xlApp.Selection.CurrentRegion.Rows.AutoFit
    
    xlWb.Close True, NomArchivoExcel

    Dim XL As New excel.Application 'Crea el objeto excel
    XL.Workbooks.Open NomArchivoExcel, , True 'El true es para abrir el archivo en modo Solo lectura (False si lo quieres de otro modo)
    XL.Visible = True
    XL.WindowState = xlMaximized 'Para que la ventana aparezca maximizada
    
    '-------> Close ADO objects
    rst.Close
    Set rst = Nothing
    
    ' -- Cerrar Excel
    xlApp.Quit
    '-------> Release Excel references
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing




Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub Command2_Click()
 On Error GoTo Man_Error

   Me.Hide
   Unload Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub Form_Load()
 On Error GoTo Man_Error
 
Msgtitulo = "Consultar Convenios"
fg_centra Me

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Private Sub fpText_Change(Index As Integer)
 On Error GoTo Man_Error
 
Dim RS As New ADODB.Recordset
Dim Sql As String

Select Case Index
Case 0
    Set RS = vg_db.Execute("SELECT ing_codigo, ing_nombre " & _
                "FROM b_ingrediente WITH (NOLOCK) " & _
                "WHERE ing_codigo = '" & LimpiaDato(Trim(fpText1(0).text)) & "' " & _
                "AND   ing_indppr   = 1 " & _
                "AND   ing_activo = '1'")
    If RS.EOF Then
       RS.Close
       Set RS = Nothing
       fpayuda(1).Caption = ""
       Exit Sub
    End If
    fpayuda(1).Caption = Trim(RS!ing_nombre)
    fpText1(0).text = RS!ing_codigo
    RS.Close
    Set RS = Nothing
Case 1
    Sql = ""
    Sql = "sgpadm_Sel_OrgCompras "
    Sql = Sql & " '" & LimpiaDato(Trim(fpText1(1).text)) & "' "
    Set RS = vg_db.Execute(Sql)
    If RS.EOF Then
       RS.Close
       Set RS = Nothing
       fpayuda(4).Caption = ""
       Exit Sub
    End If
    fpayuda(0).Caption = RS!ID_ORGCOMPRA
    fpText1(1).text = RS!ID_ORGCOMPRA
    RS.Close
    Set RS = Nothing

End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub

Sub encabezado(ByRef rst As ADODB.Recordset, ByRef xlWs As Object)
On Error GoTo Man_Error

Dim fldCount As Long
Dim icol As Long

'-------> Copy field names to the first row of the worksheet
fldCount = rst.Fields.count
For icol = 1 To fldCount
    xlWs.Cells(1, icol).Value = rst.Fields(icol - 1).Name
Next

Exit Sub
Man_Error:
    MsgBox Err.Description, vbCritical, Msgtitulo
End Sub

Private Sub Image1_Click(Index As Integer)
 On Error GoTo Man_Error
 
Select Case Index
Case 0
    vg_left = fpayuda(4).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_centrologisticoceco_sap", "", "Organización de Compras", "Celo"
    B_TabEst.Show 1
    Me.Refresh
    If vg_codigo = "" Then Exit Sub
    fpText2.text = vg_codigo
Case 1
    vg_left = fpayuda(1).Left + 3000
    vg_nombre = "": vg_codigo = ""
    B_TabEst.LlenaDatos "b_ingrediente", "ing_", "INgrediente", "IngReal"
    B_TabEst.Show 1
    Me.Refresh
    Screen.MousePointer = 0
    If vg_codigo = "" Then Exit Sub
    fpText1(0).text = vg_codigo
    fpayuda(1).Caption = vg_nombre
End Select

Exit Sub
Man_Error:
    fg_descarga
    MsgBox Err & ":  " & Error$(Err), vbCritical, Msgtitulo
End Sub
